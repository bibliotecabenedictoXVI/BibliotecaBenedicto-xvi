from flask import Flask, render_template_string, request, redirect, url_for, session, flash
from datetime import datetime, timedelta
import re
import sqlite3
import unicodedata
import os
import hashlib
import pandas as pd
import requests
import smtplib
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "seminario_sjp2_secret_key")

EXCEL_DB_PATH = "BASE DATOS BIBLIOTECA.xlsx"
DB_PATH = "biblioteca_sjp2.db"
DATABASE_URL = (os.environ.get("DATABASE_URL") or "").strip()
ES_POSTGRES = bool(DATABASE_URL)

if ES_POSTGRES:
    import psycopg


def _convertir_sql_pg(sql):
    """Adapta consultas estilo SQLite (?, INSERT OR REPLACE/IGNORE, AUTOINCREMENT) a PostgreSQL."""
    s = sql.strip()
    s = re.sub(r"\bINTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT\b", "SERIAL PRIMARY KEY", s, flags=re.IGNORECASE)
    s = re.sub(r"\bBLOB\b", "BYTEA", s, flags=re.IGNORECASE)
    s = re.sub(r"\bINSERT\s+OR\s+IGNORE\s+INTO\s+", "INSERT INTO ", s, flags=re.IGNORECASE)

    m_replace = re.match(
        r"^\s*INSERT\s+OR\s+REPLACE\s+INTO\s+([a-zA-Z0-9_]+)(?:\s*\(([^)]+)\))?\s+VALUES\s*\((.+)\)\s*;?\s*$",
        s,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if m_replace:
        tabla = m_replace.group(1).lower()
        cols_raw = m_replace.group(2)
        vals_raw = m_replace.group(3)

        conflict_cols_map = {
            "libros": ["id"],
            "credenciales": ["usuario"],
            "usuarios_biblio": ["id"],
            "sync_excel": ["id"],
            "configuracion": ["id"],
        }

        if cols_raw:
            cols = [c.strip() for c in cols_raw.split(",")]
            cols_sql = f" ({', '.join(cols)})"
        else:
            orden_tablas = {
                "libros": ["id", "titulo", "autor", "categoria", "anio", "ejemplares", "ubicacion", "editorial", "ciudad", "llave_unica", "resumen"],
                "credenciales": ["usuario", "password", "rol", "nombre_completo", "correo"],
                "usuarios_biblio": ["id", "nombre", "tipo", "correo"],
                "sync_excel": ["id", "hash_md5"],
            }
            cols = orden_tablas.get(tabla, [])
            cols_sql = f" ({', '.join(cols)})" if cols else ""

        conflict_cols = conflict_cols_map.get(tabla, [cols[0]] if cols else ["id"])
        conflict_target = ", ".join(conflict_cols)

        if cols:
            update_cols = [c for c in cols if c not in conflict_cols]
            if update_cols:
                update_clause = "DO UPDATE SET " + ", ".join([f"{c} = EXCLUDED.{c}" for c in update_cols])
            else:
                update_clause = "DO NOTHING"
        else:
            update_clause = "DO NOTHING"

        s = f"INSERT INTO {tabla}{cols_sql} VALUES ({vals_raw}) ON CONFLICT ({conflict_target}) {update_clause}"

    elif re.match(r"^\s*INSERT\s+INTO\s+", s, flags=re.IGNORECASE) and "ON CONFLICT" not in s.upper():
        m_ins = re.match(r"^\s*INSERT\s+INTO\s+([a-zA-Z0-9_]+)", s, flags=re.IGNORECASE)
        if m_ins:
            tabla = m_ins.group(1).lower()
            if tabla == "libros_pendientes":
                s = s.rstrip(";") + " ON CONFLICT (id_usuario, id_libro) DO NOTHING"
            elif tabla == "credenciales" and "INSERT OR IGNORE" in sql.upper():
                s = s.rstrip(";") + " ON CONFLICT (usuario) DO NOTHING"

    s = s.replace("?", "%s")
    return s


class _CursorAdapter:
    def __init__(self, raw_cursor, es_pg):
        self._cur = raw_cursor
        self._es_pg = es_pg

    def execute(self, sql, params=None):
        consulta = _convertir_sql_pg(sql) if self._es_pg else sql
        if params is None:
            return self._cur.execute(consulta)
        return self._cur.execute(consulta, params)

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    @property
    def rowcount(self):
        return self._cur.rowcount

    def close(self):
        return self._cur.close()


class _ConnAdapter:
    def __init__(self, raw_conn, es_pg):
        self._conn = raw_conn
        self._es_pg = es_pg

    def cursor(self):
        return _CursorAdapter(self._conn.cursor(), self._es_pg)

    def commit(self):
        return self._conn.commit()

    def rollback(self):
        return self._conn.rollback()

    def close(self):
        return self._conn.close()


def get_db_connection():
    if ES_POSTGRES:
        raw = psycopg.connect(DATABASE_URL, connect_timeout=10)
        return _ConnAdapter(raw, True)
    raw = sqlite3.connect(DB_PATH, timeout=15)
    return _ConnAdapter(raw, False)


def normalizar_texto(texto):
    if not texto or pd.isna(texto):
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper()


def _obtener_dewey_limpio(valor):
    """Retorna el código Dewey limpio tal como viene, o '' si está vacío/inválido."""
    if valor is None or pd.isna(valor):
        return ""
    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return str(valor).strip()
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return ""
    return texto


def formatear_autor_apa(autor_str):
    """Convierte un autor a formato APA (Apellido, N.)."""
    if not autor_str or autor_str.strip() == "":
        return "Autor Desconocido"
    partes = autor_str.strip().split()
    if len(partes) >= 2:
        apellido = partes[-1]
        nombres = " ".join([p[0] + "." for p in partes[:-1]])
        return f"{apellido}, {nombres}"
    return autor_str


def formatear_autor_chicago(autor_str):
    """Convierte un autor a formato Chicago (Nombre Apellido)."""
    if not autor_str or autor_str.strip() == "":
        return "Autor Desconocido"
    return autor_str.title()


def _hash_archivo(ruta):
    h = hashlib.md5()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(8192), b""):
            h.update(bloque)
    return h.hexdigest()


def sincronizar_excel_a_bd(conn, cursor):
    if not os.path.exists(EXCEL_DB_PATH):
        return

    hash_actual = _hash_archivo(EXCEL_DB_PATH)
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS sync_excel (id INTEGER PRIMARY KEY CHECK (id = 1), hash_md5 TEXT)"
    )
    conn.commit()

    cursor.execute("SELECT hash_md5 FROM sync_excel WHERE id = 1")
    row_sync = cursor.fetchone()
    ultimo_hash = row_sync[0] if row_sync else None

    cursor.execute("SELECT COUNT(*) FROM libros")
    total_libros_bd = cursor.fetchone()[0] or 0

    if ultimo_hash == hash_actual and total_libros_bd > 0:
        return

    try:
        excel_df = pd.read_excel(EXCEL_DB_PATH, sheet_name=0)
        for idx, row in excel_df.iterrows():
            clasificacion_dewey = _obtener_dewey_limpio(row.get('Clasificación Dewey'))
            if not clasificacion_dewey:
                continue

            titulo = str(row.get('Título completo', '')).strip()
            if not titulo or titulo.lower() == 'nan':
                continue

            llave_unica = str(row.get('Llave Única', '')).strip()
            if not llave_unica or llave_unica.lower() == 'nan':
                llave_unica = f"LIB-{idx + 1}"

            autor = str(row.get('Autor(es)', 'S/A')).strip()
            if not autor or autor.lower() == 'nan':
                autor = 'S/A'

            try:
                anio = int(row.get('Año de publicación', 2000))
            except Exception:
                anio = 2000

            try:
                ejem_excel = int(row.get('Ejemplar', 1))
                if ejem_excel <= 0:
                    ejem_excel = 1
            except Exception:
                ejem_excel = 1

            editorial = str(row.get('Editorial', 'S/E')).strip()
            if not editorial or editorial.lower() == 'nan':
                editorial = 'S/E'

            ubicacion_excel = str(row.get('Ubicación', 'GENERAL')).strip()
            if not ubicacion_excel or ubicacion_excel.lower() == 'nan':
                ubicacion_excel = 'GENERAL'

            resumen_excel = str(row.get('RESUMEN LIBROS', '')).strip()
            if resumen_excel.lower() == 'nan':
                resumen_excel = ''

            cursor.execute(
                "SELECT COUNT(*) FROM prestamos WHERE id_libro = ?",
                (clasificacion_dewey,)
            )
            prestados_libro = cursor.fetchone()[0] or 0
            ejem_disponibles = max(0, ejem_excel - prestados_libro)

            cursor.execute(
                "INSERT OR REPLACE INTO libros VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    clasificacion_dewey,
                    titulo,
                    autor,
                    clasificacion_dewey,
                    anio,
                    ejem_disponibles,
                    ubicacion_excel,
                    editorial,
                    "MEDELLIN",
                    llave_unica,
                    resumen_excel,
                )
            )

        cursor.execute(
            "INSERT OR REPLACE INTO sync_excel (id, hash_md5) VALUES (1, ?)",
            (hash_actual,)
        )
        conn.commit()
        print("¡Base de datos sincronizada automáticamente desde BASE DATOS BIBLIOTECA.xlsx!")
    except Exception as e:
        conn.rollback()
        print(f"Error sincronizando Excel: {e}")


def inicializar_bd():
    conn = get_db_connection()
    cursor = conn.cursor()

    tablas = [
        """
        CREATE TABLE IF NOT EXISTS credenciales (
            usuario TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            rol TEXT NOT NULL,
            nombre_completo TEXT,
            correo TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS libros (
            id TEXT PRIMARY KEY,
            titulo TEXT NOT NULL,
            autor TEXT NOT NULL,
            categoria TEXT NOT NULL,
            anio INTEGER NOT NULL,
            ejemplares INTEGER NOT NULL,
            ubicacion TEXT NOT NULL,
            editorial TEXT,
            ciudad TEXT,
            llave_unica TEXT,
            resumen TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS usuarios_biblio (
            id TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            tipo TEXT NOT NULL,
            correo TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS prestamos (
            id_libro TEXT,
            id_usuario TEXT,
            fecha TEXT,
            fecha_devolucion_limite TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS historial_movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo_movimiento TEXT,
            id_usuario TEXT,
            nombre_usuario TEXT,
            id_libro TEXT,
            titulo_libro TEXT,
            fecha TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS libros_pendientes (
            id_usuario TEXT,
            id_libro TEXT,
            PRIMARY KEY (id_usuario, id_libro)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS notificaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            destinatario TEXT,
            mensaje TEXT,
            fecha TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS eliminados_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            admin TEXT,
            id_libro TEXT,
            titulo TEXT,
            fecha TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS configuracion (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            titulo TEXT,
            subtitulo TEXT,
            logo BLOB,
            smtp_correo TEXT,
            smtp_password TEXT
        )
        """,
    ]

    try:
        for ddl in tablas:
            try:
                cursor.execute(ddl)
                conn.commit()
            except Exception as e_ddl:
                conn.rollback()
                print(f"Aviso creando tabla: {e_ddl}")

        try:
            cursor.execute("ALTER TABLE prestamos ADD COLUMN fecha_devolucion_limite TEXT")
            conn.commit()
        except Exception:
            conn.rollback()

        cursor.execute("SELECT COUNT(*) FROM credenciales")
        if cursor.fetchone()[0] == 0:
            admin_usr = (os.environ.get("ADMIN_USUARIO") or "").strip().upper()
            admin_pwd = (os.environ.get("ADMIN_PASSWORD") or "").strip()
            admin_nom = (os.environ.get("ADMIN_NOMBRE") or "ADMINISTRADOR").strip().upper()
            admin_mail = (os.environ.get("ADMIN_CORREO") or "").strip()
            if admin_usr and admin_pwd:
                cursor.execute(
                    "INSERT INTO credenciales VALUES (?, ?, 'Administrador', ?, ?)",
                    (admin_usr, admin_pwd, admin_nom, admin_mail)
                )
        conn.commit()

        cursor.execute("SELECT COUNT(*) FROM configuracion")
        if cursor.fetchone()[0] == 0:
            cursor.execute(
                "INSERT INTO configuracion (id, titulo, subtitulo) VALUES (1, 'BIBLIOTECA BENEDICTO XVI', 'SEMINARIO SAN JUAN PABLO II - ARQUIDIÓCESIS DE MEDELLÍN')"
            )
            conn.commit()

        sincronizar_excel_a_bd(conn, cursor)
    finally:
        conn.close()


def _inicializar_bd_segura():
    try:
        inicializar_bd()
    except Exception as e:
        print(f"Error inicializando BD en segundo plano: {e}")


threading.Thread(target=_inicializar_bd_segura, daemon=True).start()


def _obtener_credenciales_correo():
    remitente = (os.environ.get("SMTP_CORREO") or "").strip()
    password = (os.environ.get("SMTP_PASSWORD") or "").strip().replace(" ", "")

    if not remitente or not password:
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT smtp_correo, smtp_password FROM configuracion WHERE id = 1")
            res = cursor.fetchone()
            conn.close()
            if res:
                if not remitente and res[0]:
                    remitente = res[0].strip()
                if not password and res[1]:
                    password = res[1].strip().replace(" ", "")
        except Exception as e:
            print(f"Error leyendo config SMTP en BD: {e}")

    return remitente, password


def _enviar_por_brevo_api(remitente, destinatario, asunto, html_content, texto_plano):
    api_key = (os.environ.get("BREVO_API_KEY") or "").strip()
    if not api_key:
        return False

    try:
        resp = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            headers={
                "accept": "application/json",
                "api-key": api_key,
                "content-type": "application/json",
            },
            json={
                "sender": {
                    "name": "Biblioteca Benedicto XVI",
                    "email": remitente,
                },
                "to": [{"email": destinatario.strip()}],
                "subject": asunto,
                "htmlContent": html_content,
                "textContent": texto_plano,
            },
            timeout=12,
        )
        if resp.status_code in (200, 201, 202):
            print(f"¡Correo enviado vía Brevo API a {destinatario}!")
            return True
        print(f"Aviso Brevo API ({resp.status_code}): {resp.text}")
        return False
    except Exception as e:
        print(f"Error Brevo API: {e}")
        return False


def _enviar_hilo(destinatario_correo, asunto, cuerpo_texto):
    remitente, password = _obtener_credenciales_correo()
    if not remitente and not (os.environ.get("BREVO_API_KEY") or "").strip():
        print("Aviso correo: faltan SMTP_CORREO o BREVO_API_KEY.")
        return

    remitente_efectivo = remitente or "bibliotecabenedictoxvissjpii@gmail.com"

    html_content = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333333; line-height: 1.5;">
        <p>{cuerpo_texto.replace(chr(10), '<br>')}</p>
        <br>
        <hr style="border: none; border-top: 1px solid #cccccc; width: 100%; margin: 20px 0;">
        <span style="font-size: 18px; font-weight: bold; color: #1d3557; display: block;">Biblioteca Benedicto XVI</span>
        <span style="font-size: 14px; font-weight: bold; color: #000000; display: block; margin-top: 2px;">Seminario San Juan Pablo II - Arquidiócesis de Medellín</span>
        <div style="font-size: 13px; color: #333333; margin-top: 2px;">Medellín, Antioquia</div>
      </body>
    </html>
    """

    if _enviar_por_brevo_api(remitente_efectivo, destinatario_correo, asunto, html_content, cuerpo_texto):
        return

    if not remitente or not password:
        print("Aviso correo: faltan SMTP_CORREO o SMTP_PASSWORD (en entorno o configuración).")
        return

    msg = MIMEMultipart('alternative')
    msg['From'] = remitente
    msg['To'] = destinatario_correo.strip()
    msg['Subject'] = asunto
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    try:
        servidor = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=10)
        servidor.login(remitente, password)
        servidor.sendmail(remitente, destinatario_correo.strip(), msg.as_string())
        servidor.quit()
        print(f"¡Correo enviado exitosamente (SSL 465) a {destinatario_correo}!")
        return
    except Exception as e_ssl:
        print(f"Aviso SMTP SSL 465: {e_ssl}")

    try:
        servidor = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        servidor.ehlo()
        servidor.starttls()
        servidor.ehlo()
        servidor.login(remitente, password)
        servidor.sendmail(remitente, destinatario_correo.strip(), msg.as_string())
        servidor.quit()
        print(f"¡Correo enviado exitosamente (TLS 587) a {destinatario_correo}!")
        return
    except Exception as e_tls:
        print(f"Aviso SMTP TLS 587: {e_tls}")

    print(f"ERROR FINAL envío a {destinatario_correo}: no fue posible enviar por Brevo ni SMTP.")


def enviar_correo_automatico(destinatario_correo, asunto, cuerpo_texto):
    if not destinatario_correo or "@" not in destinatario_correo or "example.com" in destinatario_correo:
        return
    threading.Thread(target=_enviar_hilo, args=(destinatario_correo, asunto, cuerpo_texto), daemon=True).start()


LAYOUT_HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Biblioteca Benedicto XVI - Seminario San Juan Pablo II</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        html, body {
            background-color: #f4f6f8;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            overflow-x: hidden;
            max-width: 100%;
        }
        .navbar-custom { background-color: #1d3557; }
        .card-stats { border-left: 4px solid #b78a2c; }
        .badge-dewey {
            white-space: normal !important;
            word-break: break-word;
            overflow-wrap: anywhere;
            max-width: 75%;
            text-align: left;
            line-height: 1.3;
        }
        @media (max-width: 767px) {
            .desktop-table { display: none !important; }
            .mobile-cards { display: block !important; width: 100%; }
            .container { padding-left: 12px; padding-right: 12px; max-width: 100%; }
            h4 { font-size: 1.15rem; }
            .card { max-width: 100%; overflow-x: hidden; }
        }
        @media (min-width: 768px) {
            .desktop-table { display: block !important; }
            .mobile-cards { display: none !important; }
        }
    </style>
</head>
<body>
    {% if session.get('usuario') %}
    <nav class="navbar navbar-expand-lg navbar-dark navbar-custom px-3 shadow-sm">
        <a class="navbar-brand fw-bold fs-6 d-flex align-items-center" href="/">
            <i class="fa-solid fa-cross text-warning me-2"></i> Biblioteca Benedicto XVI
        </a>
        <button class="navbar-toggler border-0" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
            <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNav">
            <ul class="navbar-nav me-auto mb-2 mb-lg-0 mt-2">
                <li class="nav-item"><a class="nav-link text-white" href="/"><i class="fa-solid fa-book"></i> Catálogo Dewey</a></li>
                {% if session.get('rol') == 'Administrador' %}
                <li class="nav-item"><a class="nav-link text-white" href="/admin/miembros"><i class="fa-solid fa-users"></i> Miembros</a></li>
                {% endif %}
                <li class="nav-item"><a class="nav-link text-white" href="/importador_abierto"><i class="fa-solid fa-globe"></i> Buscador Masivo</a></li>
                <li class="nav-item"><a class="nav-link text-white" href="/prestamos"><i class="fa-solid fa-rotate"></i> Préstamos</a></li>
                <li class="nav-item"><a class="nav-link text-white" href="/buscar"><i class="fa-solid fa-magnifying-glass"></i> Búsqueda</a></li>
                {% if session.get('rol') == 'Seminarista' %}
                <li class="nav-item"><a class="nav-link text-white" href="/pendientes"><i class="fa-solid fa-thumbtack"></i> Pendientes</a></li>
                <li class="nav-item"><a class="nav-link text-white" href="/notificaciones"><i class="fa-solid fa-bell"></i> Notificaciones</a></li>
                {% endif %}
                {% if session.get('rol') == 'Administrador' %}
                <li class="nav-item"><a class="nav-link text-white" href="/admin/notificaciones"><i class="fa-solid fa-paper-plane"></i> Avisos</a></li>
                <li class="nav-item"><a class="nav-link text-white" href="/admin/historial"><i class="fa-solid fa-clipboard-list"></i> Registro</a></li>
                <li class="nav-item"><a class="nav-link text-white" href="/admin/config"><i class="fa-solid fa-gear"></i> Admin</a></li>
                {% endif %}
            </ul>
            <div class="text-white d-flex flex-column flex-lg-row align-items-lg-center pt-2 pt-lg-0 border-top border-light border-opacity-25 border-lg-0">
                <span class="mb-2 mb-lg-0 me-lg-3 small"><i class="fa-solid fa-user"></i> {{ session.get('nombre_completo', session['usuario']) }} ({{ session['rol'] }})</span>
                <a href="/logout" class="btn btn-danger btn-sm"><i class="fa-solid fa-right-from-bracket"></i> Salir</a>
            </div>
        </div>
    </nav>
    {% endif %}

    <div class="container my-3 my-md-4">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'danger' if category == 'error' else 'success' }} alert-dismissible fade show shadow-sm small" role="alert">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        {{ contenido | safe }}
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = (request.form.get('usuario') or '').strip().upper()
        password = (request.form.get('password') or '').strip()

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT usuario, password, rol, nombre_completo FROM credenciales WHERE UPPER(usuario) = ?", (usuario,))
        row = cursor.fetchone()
        conn.close()

        if row and row[1] == password:
            session['usuario'] = row[0]
            session['rol'] = row[2]
            session['nombre_completo'] = row[3] if row[3] else row[0]
            return redirect(url_for('index'))
        else:
            flash("Usuario o contraseña incorrectos.", "error")

    contenido = """
    <div class="row justify-content-center mt-4 mt-md-5">
        <div class="col-12 col-md-6 col-lg-5">
            <div class="card shadow border-0 p-4 text-center bg-white">
                <div class="mb-3">
                    <div class="display-6 text-primary mb-1"><i class="fa-solid fa-church"></i></div>
                    <h5 class="fw-bold text-dark mb-0">SEMINARIO SAN JUAN PABLO II</h5>
                    <small class="text-muted">Arquidiócesis de Medellín</small>
                </div>
                <hr class="my-2">
                <h6 class="fw-bold text-secondary mb-3">Biblioteca Benedicto XVI</h6>
                <form method="POST" class="text-start">
                    <div class="mb-3">
                        <label class="form-label fw-bold small">Tipo de usuario</label>
                        <select name="tipo" class="form-select form-select-sm">
                            <option value="Administrador">Administrador</option>
                            <option value="Seminarista">Seminarista</option>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-bold small">Usuario</label>
                        <input type="text" name="usuario" class="form-control form-control-sm" required style="text-transform: uppercase;" autocomplete="username">
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-bold small">Contraseña</label>
                        <input type="password" name="password" class="form-control form-control-sm" required autocomplete="current-password">
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold py-2" style="background-color: #1d3557;">INGRESAR AL SISTEMA</button>
                </form>
            </div>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
def index():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, titulo, autor, categoria, anio, ejemplares, ubicacion, resumen FROM libros")
    libros = cursor.fetchall()

    cursor.execute("SELECT id_libro, id_usuario, fecha FROM prestamos")
    prestamos = cursor.fetchall()
    conn.close()

    total_titulos = len(libros)
    total_ejemplares_disp = sum(l[5] for l in libros)
    total_prestados = len(prestamos)
    total_copias = total_ejemplares_disp + total_prestados

    filas_pc = ""
    tarjetas_celular = ""

    for l in libros:
        badge_color = "success" if l[5] > 0 else "danger"

        acciones = f'<a href="/libro/{l[0]}" class="btn btn-sm btn-outline-primary" title="Ver Detalle"><i class="fa-solid fa-eye"></i></a>'
        acciones += f' <a href="/marcar_pendiente/{l[0]}" class="btn btn-sm btn-outline-warning" title="Marcar Pendiente"><i class="fa-solid fa-thumbtack"></i></a>'
        if session.get('rol') == 'Administrador':
            acciones += f' <a href="/admin/editar_libro/{l[0]}" class="btn btn-sm btn-outline-secondary" title="Editar"><i class="fa-solid fa-pen"></i></a>'
            acciones += f' <a href="/admin/eliminar_libro/{l[0]}" class="btn btn-sm btn-outline-danger" title="Eliminar" onclick="return confirm(\'¿Eliminar este libro?\')"><i class="fa-solid fa-trash"></i></a>'

        filas_pc += f"""
        <tr>
            <td><strong>{l[0]}</strong></td>
            <td><a href="/libro/{l[0]}" class="text-primary text-decoration-none fw-bold">{l[1]}</a></td>
            <td>{l[2]}</td>
            <td>{l[4]}</td>
            <td><span class="badge bg-{badge_color}">{l[5]}</span></td>
            <td>{l[6]}</td>
            <td>{acciones}</td>
        </tr>
        """

        tarjetas_celular += f"""
        <div class="card mb-3 shadow-sm border-0 bg-white p-3">
            <div class="d-flex justify-content-between align-items-start gap-2 mb-2">
                <span class="badge bg-dark badge-dewey">Dewey: {l[0]}</span>
                <span class="badge bg-{badge_color} flex-shrink-0">Disp: {l[5]}</span>
            </div>
            <h6 class="fw-bold mb-1" style="overflow-wrap: anywhere;">
                <a href="/libro/{l[0]}" class="text-primary text-decoration-none">{l[1]}</a>
            </h6>
            <p class="small text-muted mb-1"><i class="fa-solid fa-user"></i> <strong>Autor:</strong> {l[2]}</p>
            <p class="small text-muted mb-1"><i class="fa-solid fa-calendar"></i> <strong>Año:</strong> {l[4]} | <i class="fa-solid fa-location-dot"></i> <strong>Ubicación:</strong> {l[6]}</p>
            <div class="d-flex justify-content-end gap-1 mt-2 pt-2 border-top">
                {acciones}
            </div>
        </div>
        """

    contenido = f"""
    <div class="card shadow-sm border-0 p-4 bg-white mb-4 text-center">
        <div class="d-flex flex-column align-items-center justify-content-center">
            <div class="text-primary mb-2" style="font-size: 2.5rem;">
                <i class="fa-solid fa-church"></i>
            </div>
            <h4 class="fw-bold text-dark mb-1">SEMINARIO SAN JUAN PABLO II</h4>
            <h6 class="text-muted fw-semibold mb-2">Arquidiócesis de Medellín — Biblioteca Benedicto XVI</h6>
            <hr class="w-25 mx-auto my-2">
            <p class="text-secondary small mb-0">Sistema integral de gestión bibliotecaria, catálogo Dewey y recursos académicos.</p>
        </div>
    </div>

    <div class="row text-center mb-3 g-2">
        <div class="col-6 col-md-3"><div class="card card-stats shadow-sm p-2 bg-white"><h6 class="small text-muted mb-1">TÍTULOS</h6><h5 class="fw-bold text-primary mb-0">{total_titulos}</h5></div></div>
        <div class="col-6 col-md-3"><div class="card card-stats shadow-sm p-2 bg-white"><h6 class="small text-muted mb-1">EJEMPLARES</h6><h5 class="fw-bold text-primary mb-0">{total_copias}</h5></div></div>
        <div class="col-6 col-md-3"><div class="card card-stats shadow-sm p-2 bg-white"><h6 class="small text-muted mb-1">PRESTADOS</h6><h5 class="fw-bold text-danger mb-0">{total_prestados}</h5></div></div>
        <div class="col-6 col-md-3"><div class="card card-stats shadow-sm p-2 bg-white"><h6 class="small text-muted mb-1">DISPONIBLES</h6><h5 class="fw-bold text-success mb-0">{total_ejemplares_disp}</h5></div></div>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <div class="d-flex justify-content-between align-items-center">
            <h4 class="fw-bold text-secondary mb-0 fs-6"><i class="fa-solid fa-book"></i> Catálogo Dewey</h4>
            {'<a href="/admin/nuevo_libro" class="btn btn-success btn-sm fw-bold"><i class="fa-solid fa-plus"></i> Nuevo Libro</a>' if session.get('rol') == 'Administrador' else ''}
        </div>
    </div>

    <div class="desktop-table card shadow-sm border-0 p-3 bg-white">
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Dewey</th><th>Título</th><th>Autor</th><th>Año</th><th>Disp.</th><th>Ubicación</th><th>Acción</th></tr>
                </thead>
                <tbody>
                    {filas_pc}
                </tbody>
            </table>
        </div>
    </div>

    <div class="mobile-cards">
        {tarjetas_celular if tarjetas_celular else '<p class="text-center text-muted">No hay libros registrados.</p>'}
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/marcar_pendiente/<id_libro>')
def marcar_pendiente(id_libro):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO libros_pendientes VALUES (?, ?)", (session['usuario'], id_libro))
    conn.commit()
    conn.close()
    flash("Libro agregado a tu lista de pendientes.", "success")
    return redirect(url_for('pendientes') if session.get('rol') == 'Seminarista' else url_for('index'))


@app.route('/libro/<id_libro>')
def detalle_libro(id_libro):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, titulo, autor, categoria, anio, ejemplares, ubicacion, editorial, ciudad, resumen FROM libros WHERE id = ?", (id_libro,))
    l = cursor.fetchone()
    conn.close()
    if not l:
        return redirect(url_for('index'))

    autor_apa = formatear_autor_apa(l[2])
    autor_chi = formatear_autor_chicago(l[2])

    cita_apa = f"{autor_apa} ({l[4]}). <i>{l[1]}</i>. {l[7]}."
    cita_chi = f"{autor_chi}. ({l[4]}). <i>{l[1]}</i>. {l[8]}: {l[7]}."

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 p-md-4 bg-white">
        <h3 class="fw-bold text-primary mb-3 fs-5">{l[1]}</h3>
        <p class="small mb-1"><strong>Clasificación Dewey:</strong> {l[0]}</p>
        <p class="small mb-1"><strong>Autor:</strong> {l[2]}</p>
        <p class="small mb-1"><strong>Editorial y Año:</strong> {l[7]} ({l[4]})</p>
        <p class="small mb-3"><strong>Ubicación:</strong> {l[6]} | <strong>Disponibles:</strong> {l[5]}</p>
        <hr>
        <h5 class="fw-bold text-secondary fs-6">Resumen / De qué trata</h5>
        <p class="bg-light p-2 rounded small">{"Sin resumen registrado." if not l[9] else l[9]}</p>
        <h5 class="fw-bold text-secondary mt-3 fs-6">Normas APA (7.ª ed.) y Chicago (17.ª ed.)</h5>
        <div class="mb-2">
            <label class="form-label small fw-bold text-muted">APA 7ed:</label>
            <textarea class="form-control form-control-sm bg-light" rows="2" readonly>{cita_apa}</textarea>
        </div>
        <div class="mb-3">
            <label class="form-label small fw-bold text-muted">Chicago:</label>
            <textarea class="form-control form-control-sm bg-light" rows="2" readonly>{cita_chi}</textarea>
        </div>
        <div class="d-flex gap-2">
            <a href="/marcar_pendiente/{l[0]}" class="btn btn-warning btn-sm fw-bold"><i class="fa-solid fa-thumbtack"></i> Marcar Pendiente</a>
            <a href="/" class="btn btn-secondary btn-sm"><i class="fa-solid fa-arrow-left"></i> Volver al Catálogo</a>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/miembros', methods=['GET', 'POST'])
def miembros():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    if request.method == 'POST':
        id_m = (request.form.get('id') or '').strip().upper()
        nombre = (request.form.get('nombre') or '').strip().upper()
        correo = (request.form.get('correo') or '').strip()
        tipo = (request.form.get('tipo') or 'Seminarista').strip()

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO usuarios_biblio VALUES (?, ?, ?, ?)", (id_m, nombre, tipo, correo))
        conn.commit()
        conn.close()
        flash("Miembro registrado correctamente.", "success")
        return redirect(url_for('miembros'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, tipo, correo FROM usuarios_biblio")
    lista_m = cursor.fetchall()
    conn.close()

    filas = ""
    for m in lista_m:
        filas += f"""
        <tr>
            <td><strong>{m[0]}</strong></td>
            <td>{m[1]}</td>
            <td>{m[2]}</td>
            <td>{m[3]}</td>
            <td><a href="/admin/eliminar_miembro/{m[0]}" class="btn btn-danger btn-sm" onclick="return confirm(\'¿Eliminar miembro?\')"><i class="fa-solid fa-trash"></i></a></td>
        </tr>
        """

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-users"></i> Registrar Miembros</h4>
        <form method="POST" class="row g-2">
            <div class="col-12 col-md-2">
                <input type="text" name="id" class="form-control form-control-sm" required placeholder="ID / Carnet" style="text-transform: uppercase;">
            </div>
            <div class="col-12 col-md-3">
                <input type="text" name="nombre" class="form-control form-control-sm" required placeholder="Nombre completo" style="text-transform: uppercase;">
            </div>
            <div class="col-12 col-md-3">
                <input type="email" name="correo" class="form-control form-control-sm" placeholder="Correo electrónico">
            </div>
            <div class="col-12 col-md-2">
                <select name="tipo" class="form-select form-select-sm">
                    <option value="Seminarista">Seminarista</option>
                    <option value="Profesor">Profesor</option>
                    <option value="Sacerdote">Sacerdote</option>
                </select>
            </div>
            <div class="col-12 col-md-2">
                <button type="submit" class="btn btn-success btn-sm w-100 fw-bold">Registrar</button>
            </div>
        </form>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6">Miembros Registrados</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>ID</th><th>Nombre</th><th>Perfil</th><th>Correo</th><th>Acción</th></tr>
                </thead>
                <tbody>
                    {filas if filas else '<tr><td colspan="5" class="text-center text-muted">No hay miembros.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/eliminar_miembro/<id_m>')
def eliminar_miembro(id_m):
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM usuarios_biblio WHERE id = ?", (id_m,))
    conn.commit()
    conn.close()
    flash("Miembro eliminado.", "success")
    return redirect(url_for('miembros'))


@app.route('/admin/nuevo_libro', methods=['GET', 'POST'])
def nuevo_libro():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    if request.method == 'POST':
        id_dewey = (request.form.get('id') or '').strip().upper()
        titulo = (request.form.get('titulo') or '').strip().upper()
        autor = (request.form.get('autor') or '').strip().upper()
        categoria = (request.form.get('categoria') or '').strip().upper()
        try:
            anio = int(request.form.get('anio', 2026))
        except Exception:
            anio = 2026
        try:
            ejemplares = int(request.form.get('ejemplares', 1))
        except Exception:
            ejemplares = 1
        ubicacion = (request.form.get('ubicacion') or '').strip().upper()
        editorial = (request.form.get('editorial') or '').strip().upper()
        resumen = (request.form.get('resumen') or '').strip()

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO libros VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (id_dewey, titulo, autor, categoria, anio, ejemplares, ubicacion, editorial, "MEDELLIN", id_dewey, resumen)
        )
        conn.commit()
        conn.close()
        flash("Libro registrado exitosamente.", "success")
        return redirect(url_for('index'))

    contenido = """
    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-5"><i class="fa-solid fa-plus"></i> Registrar Nuevo Libro</h4>
        <form method="POST">
            <div class="row mb-2 g-2">
                <div class="col-12 col-md-4">
                    <label class="form-label small fw-bold">Clasificación Dewey</label>
                    <input type="text" name="id" class="form-control form-control-sm" required style="text-transform: uppercase;">
                </div>
                <div class="col-12 col-md-8">
                    <label class="form-label small fw-bold">Título Completo</label>
                    <input type="text" name="titulo" class="form-control form-control-sm" required style="text-transform: uppercase;">
                </div>
            </div>
            <div class="row mb-2 g-2">
                <div class="col-12 col-md-6">
                    <label class="form-label small fw-bold">Autor(es)</label>
                    <input type="text" name="autor" class="form-control form-control-sm" required style="text-transform: uppercase;">
                </div>
                <div class="col-12 col-md-6">
                    <label class="form-label small fw-bold">Tema / Categoría</label>
                    <input type="text" name="categoria" class="form-control form-control-sm" required style="text-transform: uppercase;">
                </div>
            </div>
            <div class="row mb-2 g-2">
                <div class="col-4">
                    <label class="form-label small fw-bold">Año</label>
                    <input type="number" name="anio" class="form-control form-control-sm" value="2026" required>
                </div>
                <div class="col-4">
                    <label class="form-label small fw-bold">Ejemplares</label>
                    <input type="number" name="ejemplares" class="form-control form-control-sm" value="1" required>
                </div>
                <div class="col-4">
                    <label class="form-label small fw-bold">Ubicación</label>
                    <input type="text" name="ubicacion" class="form-control form-control-sm" required style="text-transform: uppercase;">
                </div>
            </div>
            <div class="mb-2">
                <label class="form-label small fw-bold">Editorial</label>
                <input type="text" name="editorial" class="form-control form-control-sm" style="text-transform: uppercase;">
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Resumen / De qué trata</label>
                <textarea name="resumen" class="form-control form-control-sm" rows="3"></textarea>
            </div>
            <button type="submit" class="btn btn-success btn-sm fw-bold">Guardar Libro</button>
            <a href="/" class="btn btn-secondary btn-sm">Cancelar</a>
        </form>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/editar_libro/<id_libro>', methods=['GET', 'POST'])
def editar_libro(id_libro):
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    conn = get_db_connection()
    cursor = conn.cursor()

    if request.method == 'POST':
        titulo = (request.form.get('titulo') or '').strip().upper()
        autor = (request.form.get('autor') or '').strip().upper()
        categoria = (request.form.get('categoria') or '').strip().upper()
        try:
            anio = int(request.form.get('anio', 2026))
        except Exception:
            anio = 2026
        try:
            ejemplares = int(request.form.get('ejemplares', 1))
        except Exception:
            ejemplares = 1
        ubicacion = (request.form.get('ubicacion') or '').strip().upper()
        editorial = (request.form.get('editorial') or '').strip().upper()
        resumen = (request.form.get('resumen') or '').strip()

        cursor.execute("""
            UPDATE libros SET titulo = ?, autor = ?, categoria = ?, anio = ?, ejemplares = ?, ubicacion = ?, editorial = ?, resumen = ?
            WHERE id = ?
        """, (titulo, autor, categoria, anio, ejemplares, ubicacion, editorial, resumen, id_libro))
        conn.commit()
        conn.close()
        flash("Libro actualizado correctamente.", "success")
        return redirect(url_for('index'))

    cursor.execute("SELECT id, titulo, autor, categoria, anio, ejemplares, ubicacion, editorial, resumen FROM libros WHERE id = ?", (id_libro,))
    l = cursor.fetchone()
    conn.close()
    if not l:
        return redirect(url_for('index'))

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-5"><i class="fa-solid fa-pen"></i> Editar Libro: {l[0]}</h4>
        <form method="POST">
            <div class="mb-2">
                <label class="form-label small fw-bold">Título Completo</label>
                <input type="text" name="titulo" class="form-control form-control-sm" value="{l[1]}" required style="text-transform: uppercase;">
            </div>
            <div class="row mb-2 g-2">
                <div class="col-12 col-md-6">
                    <label class="form-label small fw-bold">Autor(es)</label>
                    <input type="text" name="autor" class="form-control form-control-sm" value="{l[2]}" required style="text-transform: uppercase;">
                </div>
                <div class="col-12 col-md-6">
                    <label class="form-label small fw-bold">Tema / Categoría</label>
                    <input type="text" name="categoria" class="form-control form-control-sm" value="{l[3]}" required style="text-transform: uppercase;">
                </div>
            </div>
            <div class="row mb-2 g-2">
                <div class="col-4">
                    <label class="form-label small fw-bold">Año</label>
                    <input type="number" name="anio" class="form-control form-control-sm" value="{l[4]}" required>
                </div>
                <div class="col-4">
                    <label class="form-label small fw-bold">Ejemplares</label>
                    <input type="number" name="ejemplares" class="form-control form-control-sm" value="{l[5]}" required>
                </div>
                <div class="col-4">
                    <label class="form-label small fw-bold">Ubicación</label>
                    <input type="text" name="ubicacion" class="form-control form-control-sm" value="{l[6]}" required style="text-transform: uppercase;">
                </div>
            </div>
            <div class="mb-2">
                <label class="form-label small fw-bold">Editorial</label>
                <input type="text" name="editorial" class="form-control form-control-sm" value="{l[7]}" style="text-transform: uppercase;">
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Resumen / De qué trata</label>
                <textarea name="resumen" class="form-control form-control-sm" rows="3">{l[8] if l[8] else ''}</textarea>
            </div>
            <button type="submit" class="btn btn-primary btn-sm fw-bold">Actualizar</button>
            <a href="/" class="btn btn-secondary btn-sm">Cancelar</a>
        </form>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/eliminar_libro/<id_libro>')
def eliminar_libro(id_libro):
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT titulo FROM libros WHERE id = ?", (id_libro,))
    res = cursor.fetchone()
    titulo_libro = res[0] if res else "Desconocido"

    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
    cursor.execute("DELETE FROM libros WHERE id = ?", (id_libro,))
    cursor.execute(
        "INSERT INTO eliminados_log (admin, id_libro, titulo, fecha) VALUES (?, ?, ?, ?)",
        (session['usuario'], id_libro, titulo_libro, fecha_actual)
    )
    conn.commit()
    conn.close()

    flash("Libro eliminado y registrado en auditoría.", "success")
    return redirect(url_for('index'))


@app.route('/prestamos', methods=['GET', 'POST'])
def prestamos():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = get_db_connection()
    cursor = conn.cursor()

    if request.method == 'POST':
        id_libro_input = (request.form.get('id_libro') or '').strip().upper()
        if session.get('rol') == 'Seminarista':
            id_usuario = session['usuario']
        else:
            id_usuario = (request.form.get('id_usuario') or '').strip().upper()

        if not id_libro_input or not id_usuario:
            flash("Debes ingresar el código del libro y el ID del usuario.", "error")
            conn.close()
            return redirect(url_for('prestamos'))

        cursor.execute(
            "SELECT id, ejemplares, titulo, autor FROM libros WHERE UPPER(id) LIKE ? OR UPPER(llave_unica) LIKE ?",
            (f'%{id_libro_input}%', f'%{id_libro_input}%')
        )
        libro = cursor.fetchone()

        cursor.execute("SELECT correo, nombre_completo FROM credenciales WHERE UPPER(usuario) = ?", (id_usuario,))
        usr = cursor.fetchone()
        correo_destino = ""
        nombre_destino = id_usuario

        if usr:
            correo_destino = usr[0] or ""
            nombre_destino = usr[1] if usr[1] else id_usuario
        else:
            cursor.execute("SELECT correo, nombre FROM usuarios_biblio WHERE UPPER(id) = ?", (id_usuario,))
            usr_biblio = cursor.fetchone()
            if usr_biblio:
                correo_destino = usr_biblio[0] or ""
                nombre_destino = usr_biblio[1] if usr_biblio[1] else id_usuario

        if not libro:
            flash("El libro con ese código no existe.", "error")
        elif libro[1] <= 0:
            flash("No hay ejemplares disponibles.", "error")
        else:
            id_libro_real = libro[0]
            titulo_libro = libro[2]

            ahora = datetime.now()
            fecha_actual = ahora.strftime("%Y-%m-%d %H:%M")
            fecha_limite = (ahora + timedelta(days=10)).strftime("%Y-%m-%d %H:%M")

            cursor.execute("UPDATE libros SET ejemplares = ejemplares - 1 WHERE id = ?", (id_libro_real,))
            cursor.execute("INSERT INTO prestamos VALUES (?, ?, ?, ?)", (id_libro_real, id_usuario, fecha_actual, fecha_limite))
            cursor.execute(
                "INSERT INTO historial_movimientos (tipo_movimiento, id_usuario, nombre_usuario, id_libro, titulo_libro, fecha) VALUES (?, ?, ?, ?, ?, ?)",
                ("PRESTAMO", id_usuario, nombre_destino, id_libro_real, titulo_libro, fecha_actual)
            )
            conn.commit()

            if correo_destino and "@" in correo_destino and "example.com" not in correo_destino:
                asunto = f"Préstamo Registrado (Plazo 10 días): {titulo_libro}"
                cuerpo = f"Estimado/a {nombre_destino},\n\nSe ha registrado el préstamo del libro '{titulo_libro}' ({id_libro_real}).\nFecha de préstamo: {fecha_actual}\nFecha límite de devolución (10 días): {fecha_limite}\n\nBiblioteca Benedicto XVI"
                enviar_correo_automatico(correo_destino, asunto, cuerpo)

            flash("Préstamo registrado con éxito (Plazo: 10 días).", "success")
        conn.close()
        return redirect(url_for('prestamos'))

    cursor.execute("SELECT id FROM libros")
    lista_dewey = [r[0] for r in cursor.fetchall()]

    cursor.execute("SELECT usuario FROM credenciales")
    lista_usuarios = [r[0] for r in cursor.fetchall()]

    if session.get('rol') == 'Seminarista':
        cursor.execute("SELECT p.id_libro, l.titulo, p.id_usuario, p.fecha, p.fecha_devolucion_limite FROM prestamos p JOIN libros l ON p.id_libro = l.id WHERE p.id_usuario = ?", (session['usuario'],))
    else:
        cursor.execute("SELECT p.id_libro, l.titulo, p.id_usuario, p.fecha, p.fecha_devolucion_limite FROM prestamos p JOIN libros l ON p.id_libro = l.id")
    lista_prestamos = cursor.fetchall()
    conn.close()

    filas = ""
    for p in lista_prestamos:
        limite_str = p[4] if p[4] else "No asignada"
        puede_devolver = (
            session.get('rol') == 'Administrador'
            or (session.get('rol') == 'Seminarista' and str(p[2]).upper() == str(session['usuario']).upper())
        )
        accion_devolver = (
            f'<a href="/devolver/{p[0]}/{p[2]}" class="btn btn-warning btn-sm"><i class="fa-solid fa-rotate-left"></i> Devolver</a>'
            if puede_devolver
            else '<span class="badge bg-info text-dark">Activo</span>'
        )

        filas += f"""
        <tr>
            <td><strong>{p[0]}</strong></td>
            <td>{p[1]}</td>
            <td>{p[2]}</td>
            <td>{p[3]}</td>
            <td><span class="text-danger fw-bold">{limite_str}</span></td>
            <td>{accion_devolver}</td>
        </tr>
        """

    options_dewey = "".join([f'<option value="{d}">' for d in lista_dewey])
    options_usuarios = "".join([f'<option value="{u}">' for u in lista_usuarios])

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-rotate"></i> Préstamos y Devoluciones (Plazo: 10 días)</h4>
        <form method="POST" class="row g-2">
            <div class="col-12 col-md-5">
                <input type="text" name="id_libro" class="form-control form-control-sm" list="sugerencias_dewey" required style="text-transform: uppercase;" placeholder="Código o Dewey">
                <datalist id="sugerencias_dewey">
                    {options_dewey}
                </datalist>
            </div>
            {'<input type="hidden" name="id_usuario" value="' + session['usuario'] + '">' if session.get('rol') == 'Seminarista' else '''
            <div class="col-12 col-md-5">
                <input type="text" name="id_usuario" class="form-control form-control-sm" list="sugerencias_usuarios" required style="text-transform: uppercase;" placeholder="ID Seminarista">
                <datalist id="sugerencias_usuarios">
                    ''' + options_usuarios + '''
                </datalist>
            </div>
            '''}
            <div class="col-12 col-md-2">
                <button type="submit" class="btn btn-success btn-sm w-100 fw-bold">Registrar</button>
            </div>
        </form>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-book-bookmark"></i> Préstamos Activos</h4>
        <div class="table-responsive">
            <table class="table table-sm table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Libro</th><th>Título</th><th>Usuario</th><th>F. Préstamo</th><th>Límite</th><th>Acción</th></tr>
                </thead>
                <tbody>
                    {filas if filas else '<tr><td colspan="6" class="text-center text-muted">No hay préstamos activos.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/devolver/<id_libro>/<id_usuario>')
def devolver(id_libro, id_usuario):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    es_admin = session.get('rol') == 'Administrador'
    es_dueno_seminarista = (
        session.get('rol') == 'Seminarista'
        and str(id_usuario).upper() == str(session['usuario']).upper()
    )
    if not (es_admin or es_dueno_seminarista):
        flash("No tienes permiso para devolver este préstamo.", "error")
        return redirect(url_for('prestamos'))

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT 1 FROM prestamos WHERE id_libro = ? AND UPPER(id_usuario) = ?",
        (id_libro, str(id_usuario).upper())
    )
    if not cursor.fetchone():
        conn.close()
        flash("Ese préstamo activo no existe o ya fue devuelto.", "error")
        return redirect(url_for('prestamos'))

    cursor.execute("SELECT titulo FROM libros WHERE id = ?", (id_libro,))
    res_l = cursor.fetchone()
    titulo_libro = res_l[0] if res_l else "Libro"

    cursor.execute("SELECT correo, nombre_completo FROM credenciales WHERE UPPER(usuario) = ?", (str(id_usuario).upper(),))
    usr = cursor.fetchone()
    correo_destino = usr[0] if usr else ""
    nombre_destino = usr[1] if usr and usr[1] else id_usuario

    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")

    cursor.execute("UPDATE libros SET ejemplares = ejemplares + 1 WHERE id = ?", (id_libro,))
    cursor.execute("DELETE FROM prestamos WHERE id_libro = ? AND UPPER(id_usuario) = ?", (id_libro, str(id_usuario).upper()))
    cursor.execute(
        "INSERT INTO historial_movimientos (tipo_movimiento, id_usuario, nombre_usuario, id_libro, titulo_libro, fecha) VALUES (?, ?, ?, ?, ?, ?)",
        ("DEVOLUCION", id_usuario, nombre_destino, id_libro, titulo_libro, fecha_actual)
    )

    cursor.execute("SELECT id_usuario FROM libros_pendientes WHERE id_libro = ?", (id_libro,))
    seminaristas_pendientes = cursor.fetchall()

    for sem in seminaristas_pendientes:
        id_sem_pendiente = sem[0]
        cursor.execute("SELECT correo, nombre_completo FROM credenciales WHERE usuario = ?", (id_sem_pendiente,))
        datos_sem = cursor.fetchone()
        if datos_sem:
            correo_sem = datos_sem[0]
            nombre_sem = datos_sem[1] if datos_sem[1] else id_sem_pendiente

            if correo_sem and "@" in correo_sem and "example.com" not in correo_sem:
                asunto_aviso = f"Libro Disponible: '{titulo_libro}'"
                cuerpo_aviso = f"Estimado/a {nombre_sem},\n\nTe informamos que el libro '{titulo_libro}' (Código: {id_libro}), el cual tenías marcado en tu lista de pendientes, acaba de ser devuelto a la biblioteca y ya se encuentra disponible para préstamo.\n\nBiblioteca Benedicto XVI"
                enviar_correo_automatico(correo_sem, asunto_aviso, cuerpo_aviso)

            cursor.execute(
                "INSERT INTO notificaciones (destinatario, mensaje, fecha) VALUES (?, ?, ?)",
                (id_sem_pendiente, f"¡El libro '{titulo_libro}' que tenías en pendientes ya fue devuelto y está disponible!", fecha_actual)
            )

    conn.commit()
    conn.close()

    if correo_destino and "@" in correo_destino and "example.com" not in correo_destino:
        asunto = f"Devolución Exitosa: {titulo_libro}"
        cuerpo = f"Estimado/a {nombre_destino},\n\nSe ha registrado correctamente la devolución del libro '{titulo_libro}'.\nGracias por cumplir con el plazo del préstamo.\n\nBiblioteca Benedicto XVI"
        enviar_correo_automatico(correo_destino, asunto, cuerpo)

    flash("Devolución registrada correctamente y avisos de disponibilidad enviados.", "success")
    return redirect(url_for('prestamos'))


@app.route('/buscar')
def buscar():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    query = (request.args.get('q') or '').strip().upper()

    conn = get_db_connection()
    cursor = conn.cursor()
    if query:
        cursor.execute(
            "SELECT id, titulo, autor, categoria, anio, ejemplares, ubicacion FROM libros WHERE id LIKE ? OR titulo LIKE ? OR autor LIKE ? OR categoria LIKE ? OR ubicacion LIKE ?",
            (f'%{query}%', f'%{query}%', f'%{query}%', f'%{query}%', f'%{query}%')
        )
    else:
        cursor.execute("SELECT id, titulo, autor, categoria, anio, ejemplares, ubicacion FROM libros")
    resultados = cursor.fetchall()
    conn.close()

    filas = ""
    for l in resultados:
        acciones = f'<a href="/libro/{l[0]}" class="btn btn-sm btn-outline-primary" title="Ver Detalle"><i class="fa-solid fa-eye"></i></a>'
        acciones += f' <a href="/marcar_pendiente/{l[0]}" class="btn btn-sm btn-outline-warning" title="Marcar Pendiente"><i class="fa-solid fa-thumbtack"></i></a>'
        if session.get('rol') == 'Administrador':
            acciones += f' <a href="/admin/editar_libro/{l[0]}" class="btn btn-sm btn-outline-secondary" title="Editar"><i class="fa-solid fa-pen"></i></a>'
            acciones += f' <a href="/admin/eliminar_libro/{l[0]}" class="btn btn-sm btn-outline-danger" title="Eliminar" onclick="return confirm(\'¿Eliminar este libro?\')"><i class="fa-solid fa-trash"></i></a>'

        filas += f"""
        <tr>
            <td><strong>{l[0]}</strong></td>
            <td><a href="/libro/{l[0]}" class="text-primary text-decoration-none fw-bold">{l[1]}</a></td>
            <td>{l[2]}</td>
            <td>{l[3]}</td>
            <td>{l[5]}</td>
            <td>{l[6]}</td>
            <td>{acciones}</td>
        </tr>
        """

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-magnifying-glass"></i> Búsqueda Global</h4>
        <form method="GET" class="input-group input-group-sm">
            <input type="text" name="q" class="form-control" placeholder="Buscar por título, autor, clasificación..." value="{query}">
            <button class="btn btn-primary" type="submit">Buscar</button>
        </form>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Dewey</th><th>Título</th><th>Autor</th><th>Categoría</th><th>Disp.</th><th>Ubicación</th><th>Acción</th></tr>
                </thead>
                <tbody>
                    {filas if filas else '<tr><td colspan="7" class="text-center text-muted">No se encontraron libros.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/pendientes')
def pendientes():
    if 'usuario' not in session or session.get('rol') != 'Seminarista':
        return redirect(url_for('index'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id_libro FROM libros_pendientes WHERE id_usuario = ?", (session['usuario'],))
    ids = [r[0] for r in cursor.fetchall()]

    libros = []
    if ids:
        placeholders = ','.join(['?'] * len(ids))
        cursor.execute(f"SELECT id, titulo, autor, ejemplares, ubicacion FROM libros WHERE id IN ({placeholders})", ids)
        libros = cursor.fetchall()
    conn.close()

    filas = ""
    for l in libros:
        filas += f"""
        <tr>
            <td><strong>{l[0]}</strong></td>
            <td><a href="/libro/{l[0]}" class="text-primary text-decoration-none fw-bold">{l[1]}</a></td>
            <td>{l[2]}</td>
            <td>{l[3]}</td>
            <td>
                <a href="/libro/{l[0]}" class="btn btn-sm btn-outline-primary"><i class="fa-solid fa-eye"></i></a>
                <a href="/quitar_pendiente/{l[0]}" class="btn btn-danger btn-sm"><i class="fa-solid fa-trash"></i></a>
            </td>
        </tr>
        """

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-thumbtack"></i> Libros Pendientes</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Dewey</th><th>Título</th><th>Autor</th><th>Disp.</th><th>Acción</th></tr>
                </thead>
                <tbody>
                    {filas if filas else '<tr><td colspan="5" class="text-center text-muted">No tienes libros pendientes.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/quitar_pendiente/<id_libro>')
def quitar_pendiente(id_libro):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM libros_pendientes WHERE id_usuario = ? AND id_libro = ?", (session['usuario'], id_libro))
    conn.commit()
    conn.close()
    flash("Libro removido de pendientes.", "success")
    return redirect(url_for('pendientes'))


@app.route('/notificaciones')
def notificaciones():
    if 'usuario' not in session or session.get('rol') != 'Seminarista':
        return redirect(url_for('index'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT fecha, mensaje FROM notificaciones WHERE destinatario = ? OR destinatario = 'TODOS' ORDER BY id DESC", (session['usuario'],))
    notifs = cursor.fetchall()
    conn.close()

    filas = ""
    for n in notifs:
        filas += f"<tr><td>{n[0]}</td><td>{n[1]}</td></tr>"

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-bell"></i> Notificaciones y Avisos</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Fecha y Hora</th><th>Aviso Recibido</th></tr>
                </thead>
                <tbody>
                    {filas if filas else '<tr><td colspan="2" class="text-center text-muted">No hay notificaciones.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/notificaciones', methods=['GET', 'POST'])
def admin_notificaciones():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    if request.method == 'POST':
        dest = (request.form.get('destinatario') or '').strip()
        msg = (request.form.get('mensaje') or '').strip()
        if dest and msg:
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO notificaciones (destinatario, mensaje, fecha) VALUES (?, ?, ?)", (dest, msg, fecha_actual))
            cursor.execute("SELECT correo FROM credenciales WHERE rol = 'Seminarista'")
            seminaristas = cursor.fetchall()
            conn.commit()
            conn.close()

            for sem in seminaristas:
                if sem[0]:
                    enviar_correo_automatico(sem[0], "Aviso Institucional - Biblioteca Benedicto XVI", msg)

            flash("Notificación enviada con éxito.", "success")
        return redirect(url_for('admin_notificaciones'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT destinatario, fecha, mensaje FROM notificaciones ORDER BY id DESC")
    historial = cursor.fetchall()
    conn.close()

    filas = ""
    for h in historial:
        filas += f"<tr><td>{h[0]}</td><td>{h[1]}</td><td>{h[2]}</td></tr>"

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-paper-plane"></i> Enviar Notificaciones y Avisos</h4>
        <form method="POST">
            <div class="mb-2">
                <label class="form-label small fw-bold">Destinatario (ID o 'TODOS')</label>
                <input type="text" name="destinatario" class="form-control form-control-sm" value="TODOS" required style="text-transform: uppercase;">
            </div>
            <div class="mb-2">
                <label class="form-label small fw-bold">Mensaje o Aviso</label>
                <textarea name="mensaje" class="form-control form-control-sm" rows="3" required></textarea>
            </div>
            <button type="submit" class="btn btn-success btn-sm fw-bold">Enviar</button>
        </form>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6">Historial de Notificaciones</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Destinatario</th><th>Fecha</th><th>Mensaje</th></tr>
                </thead>
                <tbody>
                    {filas if filas else '<tr><td colspan="3" class="text-center text-muted">No hay historial.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/historial')
def admin_historial():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT tipo_movimiento, id_usuario, nombre_usuario, id_libro, titulo_libro, fecha FROM historial_movimientos ORDER BY id DESC")
    movimientos = cursor.fetchall()

    cursor.execute("SELECT admin, id_libro, titulo, fecha FROM eliminados_log ORDER BY id DESC")
    eliminados = cursor.fetchall()
    conn.close()

    filas_m = ""
    for m in movimientos:
        badge = "success" if m[0] == "PRESTAMO" else "warning"
        filas_m += f"""
        <tr>
            <td><span class="badge bg-{badge}">{m[0]}</span></td>
            <td>{m[1]}</td>
            <td>{m[2]}</td>
            <td><strong>{m[3]}</strong></td>
            <td>{m[4]}</td>
            <td>{m[5]}</td>
        </tr>
        """

    filas_e = ""
    for e in eliminados:
        filas_e += f"<tr><td>{e[0]}</td><td><strong>{e[1]}</strong></td><td>{e[2]}</td><td>{e[3]}</td></tr>"

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-clipboard-list"></i> Historial de Movimientos</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Mov.</th><th>ID</th><th>Nombre</th><th>Dewey</th><th>Título</th><th>Fecha y Hora</th></tr>
                </thead>
                <tbody>
                    {filas_m if filas_m else '<tr><td colspan="6" class="text-center text-muted">No hay movimientos.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-trash-can"></i> Libros Eliminados (Auditoría)</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Admin</th><th>Dewey</th><th>Título</th><th>Fecha</th></tr>
                </thead>
                <tbody>
                    {filas_e if filas_e else '<tr><td colspan="4" class="text-center text-muted">No hay registros.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/config', methods=['GET', 'POST'])
def admin_config():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT smtp_correo, smtp_password FROM configuracion WHERE id = 1")
    cfg = cursor.fetchone()
    conn.close()

    correo_actual = (os.environ.get("SMTP_CORREO") or (cfg[0] if cfg and cfg[0] else "") or "")
    pass_actual = (os.environ.get("SMTP_PASSWORD") or (cfg[1] if cfg and cfg[1] else "") or "")

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-gear"></i> Configuración Servidor de Correo (SMTP SSL)</h4>
        <form method="POST" action="/admin/config_smtp">
            <div class="mb-2">
                <label class="form-label small fw-bold">Correo Remitente (Gmail)</label>
                <input type="email" name="smtp_correo" class="form-control form-control-sm" value="{correo_actual}" required>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Contraseña de Aplicación de Google</label>
                <input type="password" name="smtp_password" class="form-control form-control-sm" value="{pass_actual}" required>
            </div>
            <button type="submit" class="btn btn-primary btn-sm fw-bold">Guardar</button>
        </form>
    </div>

    <div class="text-center">
        <a href="/admin/cuentas" class="btn btn-dark btn-sm fw-bold px-4 py-2 shadow"><i class="fa-solid fa-users-gear"></i> Gestión de Cuentas y Accesos</a>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/config_smtp', methods=['POST'])
def config_smtp():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))
    smtp_correo = (request.form.get('smtp_correo') or '').strip()
    smtp_password = (request.form.get('smtp_password') or '').strip().replace(" ", "")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE configuracion SET smtp_correo = ?, smtp_password = ? WHERE id = 1", (smtp_correo, smtp_password))
    conn.commit()
    conn.close()
    flash("Configuración SMTP guardada correctamente.", "success")
    return redirect(url_for('admin_config'))


@app.route('/admin/cuentas', methods=['GET', 'POST'])
def admin_cuentas():
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))

    if request.method == 'POST':
        usuario = (request.form.get('usuario') or '').strip().upper()
        nombre_completo = (request.form.get('nombre_completo') or '').strip().upper()
        password = (request.form.get('password') or '').strip()
        correo = (request.form.get('correo') or '').strip()
        rol = (request.form.get('rol') or 'Seminarista').strip()

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO credenciales VALUES (?, ?, ?, ?, ?)", (usuario, password, rol, nombre_completo, correo))
        conn.commit()
        conn.close()
        flash("Cuenta registrada con éxito.", "success")
        return redirect(url_for('admin_cuentas'))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT usuario, nombre_completo, correo, rol FROM credenciales")
    cuentas = cursor.fetchall()
    conn.close()

    filas_c = ""
    for c in cuentas:
        filas_c += f"""
        <tr>
            <td><strong>{c[0]}</strong></td>
            <td>{c[1]}</td>
            <td>{c[2]}</td>
            <td>{c[3]}</td>
            <td>
                <span class="badge bg-success">Activo</span>
                <a href="/admin/eliminar_cuenta/{c[0]}" class="btn btn-danger btn-sm ms-2" onclick="return confirm(\'¿Eliminar esta cuenta de usuario?\')"><i class="fa-solid fa-trash"></i></a>
            </td>
        </tr>
        """

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-3 fs-6"><i class="fa-solid fa-users-gear"></i> Administración de Cuentas</h4>
        <form method="POST" class="row g-2">
            <div class="col-12 col-md-3">
                <input type="text" name="usuario" class="form-control form-control-sm" required placeholder="Usuario / ID" style="text-transform: uppercase;">
            </div>
            <div class="col-12 col-md-5">
                <input type="text" name="nombre_completo" class="form-control form-control-sm" required placeholder="Nombre completo" style="text-transform: uppercase;">
            </div>
            <div class="col-12 col-md-4">
                <select name="rol" class="form-select form-select-sm">
                    <option value="Seminarista">Seminarista</option>
                    <option value="Administrador">Administrador</option>
                </select>
            </div>
            <div class="col-12 col-md-4">
                <input type="password" name="password" class="form-control form-control-sm" required placeholder="Contraseña">
            </div>
            <div class="col-12 col-md-5">
                <input type="email" name="correo" class="form-control form-control-sm" placeholder="Correo electrónico">
            </div>
            <div class="col-12 col-md-3">
                <button type="submit" class="btn btn-success btn-sm w-100 fw-bold"><i class="fa-solid fa-user-plus"></i> Registrar</button>
            </div>
        </form>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6">Cuentas Registradas</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Usuario</th><th>Nombre</th><th>Correo</th><th>Rol</th><th>Acción</th></tr>
                </thead>
                <tbody>
                    {filas_c}
                </tbody>
            </table>
        </div>
        <div class="mt-3">
            <a href="/admin/config" class="btn btn-secondary btn-sm"><i class="fa-solid fa-arrow-left"></i> Volver</a>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/admin/eliminar_cuenta/<usuario>')
def eliminar_cuenta(usuario):
    if 'usuario' not in session or session.get('rol') != 'Administrador':
        return redirect(url_for('index'))
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM credenciales WHERE usuario = ?", (usuario,))
    conn.commit()
    conn.close()
    flash("Cuenta de usuario eliminada correctamente.", "success")
    return redirect(url_for('admin_cuentas'))


@app.route('/importador_abierto', methods=['GET', 'POST'])
def importador_abierto():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    resultados_articulos = []
    query = ""

    if request.method == 'POST':
        query = (request.form.get('query') or '').strip()
        if query:
            try:
                url_crossref = f"https://api.crossref.org/works?query={query}&rows=15"
                headers = {"User-Agent": "BibliotecaSeminaristaSJP2/1.0 (mailto:admin@seminario.edu)"}
                resp_cr = requests.get(url_crossref, headers=headers, timeout=8)
                if resp_cr.status_code == 200:
                    for item in resp_cr.json().get('message', {}).get('items', []):
                        titulo_art = item.get('title', ['Sin título'])[0]
                        autores_lista = item.get('author', [])
                        nombres_autores = ", ".join([f"{a.get('given', '')} {a.get('family', '')}".strip() for a in autores_lista]) or "Autor desconocido"
                        try:
                            anio_pub = item.get('published-print', {}).get('date-parts', [[2024]])[0][0]
                        except Exception:
                            anio_pub = 2024
                        link_pdf = item.get('URL', f"https://scholar.google.com/scholar?q={query.replace(' ', '+')}")
                        resultados_articulos.append({'titulo': titulo_art, 'autor': nombres_autores, 'anio': anio_pub, 'enlace': link_pdf, 'fuente': 'Crossref'})
            except Exception as e:
                print(f"Error Crossref: {e}")

            try:
                url_doaj = f"https://doaj.org/api/v3/search/articles/{query}?pageSize=15"
                resp_doaj = requests.get(url_doaj, timeout=8)
                if resp_doaj.status_code == 200:
                    for entry in resp_doaj.json().get('results', []):
                        biblio = entry.get('bibjson', {})
                        titulo_art = biblio.get('title', 'Sin título')
                        autores_lista = biblio.get('author', [])
                        nombres_autores = ", ".join([a.get('name', 'Desconocido') for a in autores_lista]) or "Autor institucional"
                        try:
                            anio_pub = int(biblio.get('year', 2024))
                        except Exception:
                            anio_pub = 2024
                        enlaces = biblio.get('link', [])
                        link_pdf = enlaces[0].get('url', '#') if enlaces else f"https://scholar.google.com/scholar?q={query.replace(' ', '+')}"
                        resultados_articulos.append({'titulo': titulo_art, 'autor': nombres_autores, 'anio': anio_pub, 'enlace': link_pdf, 'fuente': 'DOAJ'})
            except Exception as e:
                print(f"Error DOAJ: {e}")

    filas_res = ""
    for art in resultados_articulos:
        bg_col = "info" if art['fuente'] == 'Crossref' else "warning text-dark"
        filas_res += f"""
        <tr>
            <td><strong>{art['titulo']}</strong><br><span class="badge bg-{bg_col}">{art['fuente']}</span></td>
            <td>{art['autor']}</td>
            <td>{art['anio']}</td>
            <td><a href="{art['enlace']}" target="_blank" class="btn btn-sm btn-outline-primary mb-1"><i class="fa-solid fa-external-link"></i></a></td>
            <td>
                <form action="/guardar_importado" method="POST" style="display:inline;">
                    <input type="hidden" name="titulo" value="{art['titulo']}">
                    <input type="hidden" name="autor" value="{art['autor']}">
                    <input type="hidden" name="anio" value="{art['anio']}">
                    <input type="hidden" name="enlace" value="{art['enlace']}">
                    <button type="submit" class="btn btn-sm btn-success fw-bold mb-1"><i class="fa-solid fa-download"></i></button>
                </form>
            </td>
        </tr>
        """

    contenido = f"""
    <div class="card shadow-sm border-0 p-3 bg-white mb-3">
        <h4 class="fw-bold text-secondary mb-2 fs-6"><i class="fa-solid fa-globe"></i> Buscador Masivo (Crossref + DOAJ)</h4>
        <form method="POST" class="input-group input-group-sm">
            <input type="text" name="query" class="form-control" placeholder="Ej: Teología dogmática, Ética..." value="{query}" required>
            <button class="btn btn-primary" type="submit"><i class="fa-solid fa-magnifying-glass"></i></button>
        </form>
    </div>

    <div class="card shadow-sm border-0 p-3 bg-white">
        <h4 class="fw-bold text-secondary mb-3 fs-6">Resultados ({len(resultados_articulos)})</h4>
        <div class="table-responsive">
            <table class="table table-hover align-middle small">
                <thead class="table-dark">
                    <tr><th>Título / Fuente</th><th>Autor(es)</th><th>Año</th><th>Enlace</th><th>Importar</th></tr>
                </thead>
                <tbody>
                    {filas_res if filas_res else '<tr><td colspan="5" class="text-center text-muted">No ingresaste un término de búsqueda.</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_template_string(LAYOUT_HTML, contenido=contenido)


@app.route('/guardar_importado', methods=['POST'])
def guardar_importado():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    titulo = (request.form.get('titulo') or '').strip().upper()
    autor = (request.form.get('autor') or '').strip().upper()
    try:
        anio = int(request.form.get('anio', 2024))
    except Exception:
        anio = 2024
    enlace = (request.form.get('enlace') or '').strip()

    id_virtual = f"MS-{abs(hash(titulo)) % 10000}"

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO libros VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (id_virtual, titulo, autor, "INVESTIGACIÓN ACADÉMICA", anio, 1, "ACCESO ABIERTO MASIVO", "Crossref & DOAJ", "MEDELLIN", id_virtual, f"Recurso académico masivo importado. Enlace web: {enlace}")
    )
    conn.commit()
    conn.close()

    flash("¡Artículo importado con éxito!", "success")
    return redirect(url_for('index'))


if __name__ == '__main__':
    inicializar_bd()
    app.run(host='0.0.0.0', port=5000, debug=True)