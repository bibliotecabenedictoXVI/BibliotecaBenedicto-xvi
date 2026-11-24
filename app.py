import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime
import sqlite3
import unicodedata
from io import BytesIO
import pandas as pd
import os
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
import webbrowser

EXCEL_DB_PATH = "BASE DATOS BIBLIOTECA.xlsx"

def normalizar_texto(texto):
    if not texto or pd.isna(texto):
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper()

def inicializar_bd():
    conn = sqlite3.connect("biblioteca_sjp2.db")
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS credenciales (
            usuario TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            rol TEXT NOT NULL,
            nombre_completo TEXT,
            correo TEXT
        )
    """)
    
    cursor.execute("""
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
    """)
    
    try:
        cursor.execute("ALTER TABLE libros ADD COLUMN resumen TEXT")
        conn.commit()
    except:
        pass
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios_biblio (
            id TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            tipo TEXT NOT NULL,
            correo TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS prestamos (
            id_libro TEXT,
            id_usuario TEXT,
            fecha TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historial_movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo_movimiento TEXT,
            id_usuario TEXT,
            nombre_usuario TEXT,
            id_libro TEXT,
            titulo_libro TEXT,
            fecha TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS libros_pendientes (
            id_usuario TEXT,
            id_libro TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notificaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            destinatario TEXT,
            mensaje TEXT,
            fecha TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS eliminados_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            admin TEXT,
            id_libro TEXT,
            titulo TEXT,
            fecha TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS configuracion (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            titulo TEXT,
            subtitulo TEXT,
            logo BLOB,
            smtp_correo TEXT,
            smtp_password TEXT
        )
    """)
    
    cursor.execute("SELECT COUNT(*) FROM credenciales")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO credenciales VALUES ('GG', '1209', 'Administrador', 'ADMINISTRADOR GG', 'admin@example.com')")
        cursor.execute("INSERT INTO credenciales VALUES ('seminarista', '1234', 'Seminarista', 'SEMINARISTA PRUEBA', 'seminarista@example.com')")
    else:
        try:
            cursor.execute("ALTER TABLE credenciales ADD COLUMN nombre_completo TEXT")
            cursor.execute("ALTER TABLE credenciales ADD COLUMN correo TEXT")
            conn.commit()
        except:
            pass
        
    if os.path.exists(EXCEL_DB_PATH):
        try:
            excel_df = pd.read_excel(EXCEL_DB_PATH, sheet_name=0)
            print(f"Sincronizando {len(excel_df)} filas desde '{EXCEL_DB_PATH}'...")
            for _, row in excel_df.iterrows():
                llave_unica = str(row.get('Llave Única', '')).strip()
                if not llave_unica or llave_unica == 'nan': continue
                
                titulo = str(row.get('Título completo', 'S/T')).strip()
                autor = str(row.get('Autor(es)', 'S/A')).strip()
                clasificacion_dewey = str(row.get('Clasificación Dewey', 'General')).strip()
                
                try:
                    anio = int(row.get('Año de publicación', 2000))
                except:
                    anio = 2000
                    
                try:
                    ejem = int(row.get('Ejemplar', 1))
                    if ejem <= 0: ejem = 1
                except:
                    ejem = 1
                    
                editorial = str(row.get('Editorial', 'S/E')).strip()
                ciudad = "MEDELLIN"
                
                ubicacion_excel = str(row.get('Ubicación', 'GENERAL')).strip()
                if not ubicacion_excel or ubicacion_excel == 'nan':
                    ubicacion_excel = "GENERAL"

                resumen_excel = str(row.get('RESUMEN LIBROS', '')).strip()
                if not resumen_excel or resumen_excel == 'nan':
                    resumen_excel = ""
                
                cursor.execute("INSERT OR REPLACE INTO libros VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (
                    clasificacion_dewey, titulo, autor, clasificacion_dewey, anio, ejem, ubicacion_excel, editorial, ciudad, llave_unica, resumen_excel
                ))
            conn.commit()
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")

    cursor.execute("SELECT COUNT(*) FROM usuarios_biblio")
    if cursor.fetchone()[0] == 0:
        usuarios_iniciales = [
            ("U001", "JUAN PEREZ", "Seminarista", "juan.perez@example.com"),
            ("U002", "MARIA GOMEZ", "Profesor", "maria.gomez@example.com"),
        ]
        cursor.executemany("INSERT INTO usuarios_biblio VALUES (?, ?, ?, ?)", usuarios_iniciales)
        conn.commit()
    else:
        try:
            cursor.execute("ALTER TABLE usuarios_biblio ADD COLUMN correo TEXT")
            conn.commit()
        except:
            pass
        
    cursor.execute("SELECT COUNT(*) FROM configuracion")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO configuracion (id, titulo, subtitulo) VALUES (1, 'BIBLIOTECA BENEDICTO XVI', 'SEMINARIO SAN JUAN PABLO II')")
    else:
        try:
            cursor.execute("ALTER TABLE configuracion ADD COLUMN smtp_correo TEXT")
            cursor.execute("ALTER TABLE configuracion ADD COLUMN smtp_password TEXT")
            conn.commit()
        except:
            pass
    
    conn.commit()
    conn.close()


def enviar_correo_automatico(destinatario_correo, asunto, cuerpo_texto):
    if not destinatario_correo or "@" not in destinatario_correo:
        return False
        
    try:
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT smtp_correo, smtp_password FROM configuracion WHERE id = 1")
        res = cursor.fetchone()
        conn.close()

        if not res or not res[0] or not res[1]:
            return False

        remitente = res[0]
        password = res[1]

        img_base64_str = ""
        ruta_imagen = "WhatsApp Image 2026-09-21 at 10.47.47.jpeg"
        if os.path.exists(ruta_imagen):
            with open(ruta_imagen, "rb") as img_file:
                img_base64_str = base64.b64encode(img_file.read()).decode('utf-8')

        html_content = f"""
        <html>
          <body style="font-family: Arial, sans-serif; color: #333333; line-height: 1.5;">
            <p>{cuerpo_texto.replace(chr(10), '<br>')}</p>
            <br>
            <hr style="border: none; border-top: 1px solid #cccccc; width: 100%; margin: 20px 0;">
            <table cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
              <tr>
                <td style="vertical-align: top; padding-right: 15px;">
                  <img src="data:image/jpeg;base64,{img_base64_str}" width="90" style="display: block; border: 1px solid #dddddd;" alt="Benedicto XVI">
                  <div style="font-size: 11px; font-style: italic; color: #555555; margin-top: 4px; text-align: center;">Benedictus PP XVI</div>
                </td>
                <td style="vertical-align: middle; border-left: 2px solid #1d3557; padding-left: 15px;">
                  <span style="font-size: 18px; font-weight: bold; color: #1d3557; display: block;">Biblioteca Benedicto XVI</span>
                  <span style="font-size: 14px; font-weight: bold; color: #000000; display: block; margin-top: 2px;">Seminario san Juan Pablo II</span>
                  <div style="margin-top: 8px;">
                    <a href="mailto:bibliotecabenedictoxvissjpii@gmail.com" style="color: #1a0dab; font-size: 13px; text-decoration: underline;">bibliotecabenedictoxvissjpii@gmail.com</a>
                    <div style="font-size: 13px; color: #333333; margin-top: 2px;">Medellín, Antioquia</div>
                  </div>
                </td>
              </tr>
            </table>
          </body>
        </html>
        """

        msg = MIMEMultipart('alternative')
        msg['From'] = remitente
        msg['To'] = destinatario_correo
        msg['Subject'] = asunto
        
        msg.attach(MIMEText(html_content, 'html', 'utf-8'))

        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(remitente, password)
        servidor.sendmail(remitente, destinatario_correo, msg.as_string())
        servidor.quit()
        return True
    except Exception as e:
        print(f"Error al enviar correo electrónico: {e}")
        return False


class LoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Acceso al Sistema - Seminario Juan Pablo II")
        self.root.state("zoomed")
        self.root.minsize(1050, 680)
        self.root.configure(bg="#16283f")

        inicializar_bd()
        self.cargar_credenciales_bd()
        self.crear_pantalla_login()
        self.root.protocol("WM_DELETE_WINDOW", self.salir_aplicacion)

    def cargar_credenciales_bd(self):
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT usuario, password, rol FROM credenciales")
        filas = cursor.fetchall()
        conn.close()

        self.admins = {}
        self.seminaristas = {}
        for u, p, r in filas:
            if r == "Administrador":
                self.admins[u] = p
            else:
                self.seminaristas[u] = p

    def crear_pantalla_login(self):
        card = tk.Frame(self.root, bg="#ffffff", bd=0)
        card.place(relx=.5, rely=.5, anchor="center", width=420, height=480)

        tk.Label(card, text="🛡️", font=("Segoe UI Emoji", 30),
                 fg="#b78a2c", bg="white").pack(pady=(12, 0))
        tk.Label(card, text="SEMINARIO JUAN PABLO II",
                 font=("Segoe UI", 13, "bold"), fg="#16283f",
                 bg="white").pack()
        tk.Label(card, text="Medellín - Colombia",
                 font=("Segoe UI", 8, "italic"), fg="#718096",
                 bg="white").pack(pady=(0, 10))

        form = tk.Frame(card, bg="white")
        form.pack(fill="x", padx=40)

        tk.Label(form, text="Tipo de usuario", font=("Segoe UI", 9, "bold"),
                 fg="#344054", bg="white").pack(anchor="w")
        self.combo_tipo = ttk.Combobox(
            form, values=["Administrador", "Seminarista"],
            state="readonly", font=("Segoe UI", 10)
        )
        self.combo_tipo.pack(fill="x", pady=(2, 8))
        self.combo_tipo.current(0)

        tk.Label(form, text="Usuario", font=("Segoe UI", 9, "bold"),
                 fg="#344054", bg="white").pack(anchor="w")
        self.txt_user = tk.Entry(form, font=("Segoe UI", 10), relief="solid", bd=1)
        self.txt_user.pack(fill="x", ipady=4, pady=(2, 8))

        tk.Label(form, text="Contraseña", font=("Segoe UI", 9, "bold"),
                 fg="#344054", bg="white").pack(anchor="w")
        self.txt_pass = tk.Entry(form, show="•", font=("Segoe UI", 10),
                                 relief="solid", bd=1)
        self.txt_pass.pack(fill="x", ipady=4, pady=(2, 10))
        self.txt_pass.bind("<Return>", lambda e: self.verificar_login())

        tk.Button(form, text="INGRESAR", command=self.verificar_login,
                  bg="#1d3557", fg="white", activebackground="#27496d",
                  activeforeground="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), cursor="hand2").pack(
                      fill="x", ipady=6)

        tk.Label(card, text="Control bibliotecario con notificaciones SMTP.",
                 font=("Segoe UI", 8), fg="#98a2b3", bg="white").pack(pady=(20, 0))

    def verificar_login(self):
        tipo = self.combo_tipo.get()
        user = self.txt_user.get().strip()
        pwd = self.txt_pass.get().strip()

        if not user or not pwd:
            messagebox.showerror("Campos vacíos", "Ingrese usuario y contraseña.")
            return

        self.cargar_credenciales_bd()
        base = self.admins if tipo == "Administrador" else self.seminaristas
        if user in base and base[user] == pwd:
            self.lanzar_sistema(tipo, user)
        else:
            messagebox.showerror("Acceso denegado",
                                 f"Usuario o contraseña de {tipo.lower()} incorrectos.")

    def lanzar_sistema(self, rol, usuario):
        self.root.destroy()
        root_app = tk.Tk()
        BibliotecaApp(root_app, rol, usuario)
        root_app.mainloop()

    def salir_aplicacion(self):
        if messagebox.askyesno("Salir", "¿Desea cerrar el sistema?"):
            self.root.destroy()


class BibliotecaApp:
    def __init__(self, root, rol, usuario):
        self.root = root
        self.rol = rol
        self.usuario = usuario

        self.cargar_configuracion_bd()
        self.color_primario = "#1d3557"
        self.color_secundario = "#b78a2c"
        self.fondo = "#f4f6f8"

        self.root.title(f"{self.titulo_texto} · {rol}: {usuario}")
        self.root.state("zoomed")
        self.root.minsize(1050, 680)
        self.root.configure(bg=self.fondo)
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar_sesion)

        self.configurar_estilos()
        self.crear_interfaz()

    def cargar_configuracion_bd(self):
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT titulo, subtitulo, logo FROM configuracion WHERE id = 1")
        res = cursor.fetchone()
        conn.close()
        if res:
            self.titulo_texto, self.subtitulo_texto, self.logo_blob = res[0], res[1], res[2]
        else:
            self.titulo_texto = "BIBLIOTECA BENEDICTO XVI"
            self.subtitulo_texto = "SEMINARIO SAN JUAN PABLO II"
            self.logo_blob = None

    def configurar_estilos(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10),
                        background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"),
                        background="#e9eef4", foreground="#1d3557")
        style.map("Treeview", background=[("selected", "#dce9f7")],
                  foreground=[("selected", "#16283f")])
        style.configure("TNotebook", background=self.fondo, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(14, 8), font=("Segoe UI", 10, "bold"))

    def crear_interfaz(self):
        header = tk.Frame(self.root, bg=self.color_primario, height=84)
        header.pack(fill="x")
        header.pack_propagate(False)

        logo_frame = tk.Frame(header, bg=self.color_secundario, width=48, height=48)
        logo_frame.pack(side="left", padx=(18, 10), pady=18)
        logo_frame.pack_propagate(False)

        self.logo_img_tk = None
        if self.logo_blob:
            try:
                img_data = BytesIO(self.logo_blob)
                self.logo_img_tk = tk.PhotoImage(data=img_data.getvalue())
                tk.Label(logo_frame, image=self.logo_img_tk, bg=self.color_secundario).place(relx=0.5, rely=0.5, anchor="center")
            except Exception:
                try:
                    from PIL import Image, ImageTk
                    pil_img = Image.open(BytesIO(self.logo_blob)).resize((40, 40), Image.Resampling.LANCZOS)
                    self.logo_img_tk = ImageTk.PhotoImage(pil_img)
                    tk.Label(logo_frame, image=self.logo_img_tk, bg=self.color_secundario).place(relx=0.5, rely=0.5, anchor="center")
                except Exception:
                    tk.Label(logo_frame, text="🛡️", font=("Segoe UI Emoji", 18), bg=self.color_secundario, fg="white").place(relx=0.5, rely=0.5, anchor="center")
        else:
            tk.Label(logo_frame, text="🛡️", font=("Segoe UI Emoji", 18), bg=self.color_secundario, fg="white").place(relx=0.5, rely=0.5, anchor="center")

        title_box = tk.Frame(header, bg=self.color_primario)
        title_box.pack(side="left", pady=12)
        self.lbl_titulo = tk.Label(title_box, text=self.titulo_texto,
                                   font=("Segoe UI", 16, "bold"),
                                   fg="white", bg=self.color_primario)
        self.lbl_titulo.pack(anchor="w")
        self.lbl_sub = tk.Label(title_box, text=self.subtitulo_texto,
                                font=("Segoe UI", 9, "bold"), fg="#d7c18a",
                                bg=self.color_primario)
        self.lbl_sub.pack(anchor="w")

        right = tk.Frame(header, bg=self.color_primario)
        right.pack(side="right", padx=16, pady=10)
        tk.Label(right, text=f"●  {self.usuario}\n    {self.rol}",
                 justify="right", font=("Segoe UI", 10, "bold"),
                 fg="white", bg=self.color_primario).pack(side="left", padx=8)

        tk.Button(right, text="⎋  Cerrar sesión",
                  command=self.cerrar_sesion, bg="#8f2d35", fg="white",
                  activebackground="#a83a43", activeforeground="white",
                  relief="flat", font=("Segoe UI", 10, "bold"),
                  cursor="hand2", padx=10, pady=5).pack(side="right")

        dash = tk.Frame(self.root, bg=self.fondo)
        dash.pack(fill="x", padx=16, pady=(10, 4))

        self.card_total = self.crear_tarjeta(dash, "TÍTULOS", "0", 0)
        self.card_copias = self.crear_tarjeta(dash, "EJEMPLARES", "0", 1)
        self.card_prestados = self.crear_tarjeta(dash, "PRESTADOS", "0", 2)
        self.card_disponibles = self.crear_tarjeta(dash, "DISPONIBLES", "0", 3)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=16, pady=4)

        self.tab_libros = ttk.Frame(self.notebook)
        if self.rol == "Administrador":
            self.tab_usuarios = ttk.Frame(self.notebook)
            self.tab_abierto = ttk.Frame(self.notebook)
        self.tab_prestamos = ttk.Frame(self.notebook)
        self.tab_buscar = ttk.Frame(self.notebook)
        
        if self.rol == "Seminarista":
            self.tab_pendientes = ttk.Frame(self.notebook)
            self.tab_notificaciones = ttk.Frame(self.notebook)
        else:
            self.tab_notif_admin = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_libros, text="  📚 Catálogo Dewey  ")
        if self.rol == "Administrador":
            self.notebook.add(self.tab_usuarios, text="  👥 Miembros  ")
            self.notebook.add(self.tab_abierto, text="  🌍 Importador Open Access  ")
        self.notebook.add(self.tab_prestamos, text="  🔄 Préstamos Sincronizados  ")
        self.notebook.add(self.tab_buscar, text="  🔎 Búsqueda Global  ")

        if self.rol == "Seminarista":
            self.notebook.add(self.tab_pendientes, text="  📌 Libros Pendientes  ")
            
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM notificaciones WHERE destinatario = ? OR destinatario = 'TODOS'", (self.usuario,))
            num_notif = cursor.fetchone()[0]
            conn.close()

            texto_pestana_notif = f"  🔔 Notificaciones ({num_notif})  " if num_notif > 0 else "  🔔 Notificaciones  "
            self.notebook.add(self.tab_notificaciones, text=texto_pestana_notif)
        else:
            self.notebook.add(self.tab_notif_admin, text="  🔔 Enviar Notificaciones  ")
            self.tab_admin_logs = ttk.Frame(self.notebook)
            self.tab_config = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_admin_logs, text="  📋 Registro General  ")
            self.notebook.add(self.tab_config, text="  ⚙ Administración  ")

        self.construir_seccion_libros()
        if self.rol == "Administrador":
            self.construir_seccion_usuarios()
            self.construir_seccion_abierto()
        self.construir_seccion_prestamos()
        self.construir_seccion_buscar()
        
        if self.rol == "Seminarista":
            self.construir_seccion_pendientes()
            self.construir_seccion_notificaciones_seminarista()
        else:
            self.construir_seccion_notificaciones_admin()
            self.construir_seccion_admin_logs()
            self.construir_seccion_configuracion()

        self.actualizar_dashboard()

    def crear_tarjeta(self, parent, titulo, valor, columna):
        card = tk.Frame(parent, bg="white", bd=1, relief="solid")
        card.grid(row=0, column=columna, sticky="nsew", padx=3)
        parent.grid_columnconfigure(columna, weight=1)
        tk.Label(card, text=titulo, font=("Segoe UI", 9, "bold"),
                 fg="#667085", bg="white").pack(anchor="w", padx=12, pady=(6, 0))
        lbl = tk.Label(card, text=valor, font=("Segoe UI", 18, "bold"),
                       fg=self.color_primario, bg="white")
        lbl.pack(anchor="w", padx=12, pady=(0, 6))
        return lbl

    def obtener_libros_bd(self):
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, titulo, autor, categoria, anio, ejemplares, ubicacion, editorial, ciudad, llave_unica, resumen FROM libros")
        filas = cursor.fetchall()
        conn.close()
        libros = []
        for f in filas:
            libros.append({
                "id": f[0], "titulo": f[1], "autor": f[2],
                "categoria": f[3], "anio": f[4], "ejemplares": f[5],
                "ubicacion": f[6] if len(f) > 6 and f[6] else "GENERAL", 
                "editorial": f[7] if f[7] else "S/E", 
                "ciudad": f[8] if f[8] else "S/L", 
                "llave_unica": f[9] if len(f) > 9 and f[9] else f[0],
                "resumen": f[10] if len(f) > 10 and f[10] else ""
            })
        return libros

    def obtener_usuarios_bd(self):
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre, tipo, correo FROM usuarios_biblio")
        filas_u = cursor.fetchall()
        
        cursor.execute("SELECT usuario, rol, nombre_completo, correo FROM credenciales WHERE rol = 'Seminarista'")
        filas_c = cursor.fetchall()
        conn.close()

        usuarios = {}
        for f in filas_u:
            usuarios[f[0]] = {"id": f[0], "nombre": f[1], "tipo": f[2], "correo": f[3] if len(f) > 3 else ""}
        for u, r, nom_c, correo_c in filas_c:
            if u not in usuarios:
                nombre_mostrar = nom_c if nom_c else u.upper()
                usuarios[u] = {"id": u, "nombre": nombre_mostrar, "tipo": "Seminarista", "correo": correo_c if correo_c else ""}
        return list(usuarios.values())

    def obtener_prestamos_bd(self):
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id_libro, id_usuario, fecha FROM prestamos")
        filas = cursor.fetchall()
        conn.close()
        prestamos = []
        for f in filas:
            prestamos.append({"libro": f[0], "usuario": f[1], "fecha": f[2]})
        return prestamos

    def actualizar_dashboard(self):
        libros = self.obtener_libros_bd()
        prestamos = self.obtener_prestamos_bd()
        total = len(libros)
        copias = sum(l["ejemplares"] for l in libros) + len(prestamos)
        prestados = len(prestamos)
        disponibles = sum(l["ejemplares"] for l in libros)
        self.card_total.config(text=str(total))
        self.card_copias.config(text=str(copias))
        self.card_prestados.config(text=str(prestados))
        self.card_disponibles.config(text=str(disponibles))

    def configurar_mayusculas(self, entry):
        def al_escribir(event):
            val = entry.get()
            mayus = val.upper()
            if val != mayus:
                pos = entry.index(tk.INSERT)
                entry.delete(0, tk.END)
                entry.insert(0, mayus)
                try:
                    entry.icursor(pos)
                except:
                    pass
        entry.bind("<KeyRelease>", al_escribir)

    def actualizar_tabla_libros(self):
        for row in self.tabla_libros.get_children():
            self.tabla_libros.delete(row)
        libros = self.obtener_libros_bd()
        for l in libros:
            self.tabla_libros.insert("", "end",
                values=(l["id"], l["titulo"], l["autor"], l["categoria"],
                        l["anio"], l["ejemplares"], l["ubicacion"]))
        if hasattr(self, "lbl_catalogo_count"):
            self.lbl_catalogo_count.config(text=f"{len(libros)} títulos sincronizados desde Excel")

    def construir_seccion_libros(self):
        cont = tk.Frame(self.tab_libros, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=12, pady=10)

        if self.rol == "Administrador":
            left = tk.LabelFrame(cont, text=" Gestión de Catálogo Dewey ",
                                 font=("Segoe UI", 10, "bold"),
                                 bg=self.fondo, padx=10, pady=8)
            left.pack(side="left", fill="y", padx=(0, 10))

            tk.Button(left, text="＋ Nuevo Libro (Escanear)", command=self.abrir_ventana_nuevo_libro,
                      bg="#2f855a", fg="white", relief="flat",
                      font=("Segoe UI", 10, "bold"), pady=6).pack(fill="x", pady=4)
            tk.Button(left, text="✎ Modificar Libro", command=self.abrir_ventana_modificar_libro,
                      bg="#286090", fg="white", relief="flat",
                      font=("Segoe UI", 10, "bold"), pady=6).pack(fill="x", pady=4)
            tk.Button(left, text="🗑 Eliminar Libro", command=self.abrir_ventana_eliminar_libro,
                      bg="#9f2d35", fg="white", relief="flat",
                      font=("Segoe UI", 10, "bold"), pady=6).pack(fill="x", pady=4)
        else:
            info = tk.Frame(cont, bg="#fff8e7", bd=1, relief="solid")
            info.pack(fill="x", pady=(0, 6))
            tk.Label(info, text=f"👤 Seminarista Activo: {self.usuario}",
                     font=("Segoe UI", 10, "bold"), fg="#7a5b12",
                     bg="#fff8e7").pack(side="left", padx=10, pady=6)
            tk.Label(info, text="Catálogo de consulta disponible. Diríjase a Préstamos Sincronizados.",
                     font=("Segoe UI", 9), fg="#7a5b12",
                     bg="#fff8e7").pack(side="left")

        right = tk.Frame(cont, bg=self.fondo)
        right.pack(side="left", fill="both", expand=True)

        top = tk.Frame(right, bg=self.fondo)
        top.pack(fill="x", pady=(0, 4))
        tk.Label(top, text="Catálogo Clasificación Dewey - Sincronizado con Excel",
                 font=("Segoe UI", 12, "bold"), fg=self.color_primario,
                 bg=self.fondo).pack(side="left")
        self.lbl_catalogo_count = tk.Label(top, text="",
                                           font=("Segoe UI", 9),
                                           fg="#667085", bg=self.fondo)
        self.lbl_catalogo_count.pack(side="right")

        table = tk.Frame(right, bg="white")
        table.pack(fill="both", expand=True)
        cols = ("Clasificación Dewey", "Título", "Autor", "Tema / Categoría", "Año", "Disp.", "Ubicación General")
        self.tabla_libros = ttk.Treeview(table, columns=cols, show="headings")
        widths = {"Clasificación Dewey": 130, "Título": 280, "Autor": 180, "Tema / Categoría": 160,
                  "Año": 60, "Disp.": 55, "Ubicación General": 160}
        for c in cols:
            self.tabla_libros.heading(c, text=c)
            self.tabla_libros.column(c, width=widths[c], anchor="center" if c in
                                     ("Clasificación Dewey", "Año", "Disp.") else "w")
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tabla_libros.yview)
        self.tabla_libros.configure(yscrollcommand=sb.set)
        self.tabla_libros.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        
        self.tabla_libros.bind("<Double-1>", lambda e: self.ver_subventana_detalle_desde_tabla())
        self.actualizar_tabla_libros()

    def ver_subventana_detalle_desde_tabla(self):
        sel = self.tabla_libros.selection()
        if not sel:
            return
        vals = self.tabla_libros.item(sel[0], "values")
        lid = vals[0]
        libros = self.obtener_libros_bd()
        libro = next((l for l in libros if l["id"] == lid), None)
        if libro:
            self.mostrar_subventana_detalle(libro)

    def mostrar_subventana_detalle(self, libro):
        sub = tk.Toplevel(self.root)
        sub.title(f"Detalles, Resumen y Cita Bibliográfica: {libro['titulo']}")
        sub.geometry("640x700")
        sub.configure(bg="white")
        sub.transient(self.root)
        sub.grab_set()

        tk.Label(sub, text="📖 Información, Resumen y Citas Sistematizadas", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg="white").pack(pady=8)
        
        info_frame = tk.Frame(sub, bg="white", padx=16)
        info_frame.pack(fill="both", expand=True)

        detalles = [
            ("Clasificación Dewey:", libro["id"]),
            ("Llave Única:", libro["llave_unica"]),
            ("Título:", libro["titulo"]),
            ("Autor:", libro["autor"]),
            ("Ubicación Estantería:", libro["ubicacion"]),
            ("Año:", str(libro["anio"])),
            ("Editorial:", libro["editorial"]),
            ("Ejemplares Disponibles:", str(libro["ejemplares"]))
        ]

        for lbl, val in detalles:
            row = tk.Frame(info_frame, bg="white")
            row.pack(fill="x", pady=1)
            tk.Label(row, text=lbl, font=("Segoe UI", 9, "bold"), fg="#475467", bg="white", width=22, anchor="w").pack(side="left")
            tk.Label(row, text=val, font=("Segoe UI", 9), fg="#1d2939", bg="white", anchor="w").pack(side="left", fill="x", expand=True)

        res_frame = tk.LabelFrame(info_frame, text=" Resumen / De qué trata el libro ", font=("Segoe UI", 9, "bold"), bg="white", fg="#1d3557", padx=8, pady=6)
        res_frame.pack(fill="x", pady=(8, 4))

        txt_resumen = tk.Text(res_frame, height=4, wrap="word", font=("Segoe UI", 9), bg="#f9fafb", relief="solid", bd=1)
        txt_resumen.insert("1.0", libro["resumen"] if libro["resumen"] else "Sin resumen registrado para este ejemplar.")
        txt_resumen.config(state="disabled")
        txt_resumen.pack(fill="x", expand=True)

        citas_frame = tk.LabelFrame(info_frame, text=" Sistematización de Cita Automática ", font=("Segoe UI", 9, "bold"), bg="white", fg="#b78a2c", padx=8, pady=6)
        citas_frame.pack(fill="x", pady=(6, 4))

        cita_apa = f'{libro["autor"]}. ({libro["anio"]}). "{libro["titulo"]}". {libro["editorial"]}.'
        cita_chi = f'{libro["autor"]}. "{libro["titulo"]}". {libro["ciudad"]}: {libro["editorial"]}, {libro["anio"]}.'

        tk.Label(citas_frame, text="Formato APA 7ed:", font=("Segoe UI", 9, "bold"), bg="white", fg="#344054").pack(anchor="w")
        lbl_apa = tk.Entry(citas_frame, font=("Segoe UI", 9), bg="#f4f6f8", relief="solid", bd=1)
        lbl_apa.insert(0, cita_apa)
        lbl_apa.pack(fill="x", pady=(0, 4))

        tk.Label(citas_frame, text="Formato Chicago/De la Sota:", font=("Segoe UI", 9, "bold"), bg="white", fg="#344054").pack(anchor="w")
        lbl_chi = tk.Entry(citas_frame, font=("Segoe UI", 9), bg="#f4f6f8", relief="solid", bd=1)
        lbl_chi.insert(0, cita_chi)
        lbl_chi.pack(fill="x", pady=(0, 2))

        def copiar_cita(texto_cita):
            sub.clipboard_clear()
            sub.clipboard_append(texto_cita)
            messagebox.showinfo("Cita Copiada", "La cita bibliográfica se ha copiado al portapapeles.", parent=sub)

        btn_box = tk.Frame(citas_frame, bg="white")
        btn_box.pack(fill="x", pady=(2, 0))
        tk.Button(btn_box, text="📋 Copiar APA", command=lambda: copiar_cita(cita_apa), bg="#286090", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=8, pady=2).pack(side="left", padx=2)
        tk.Button(btn_box, text="📋 Copiar Chicago", command=lambda: copiar_cita(cita_chi), bg="#286090", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=8, pady=2).pack(side="left", padx=2)

        f_acciones = tk.Frame(sub, bg="white", pady=6)
        f_acciones.pack(fill="x", padx=16)
        
        if self.rol == "Seminarista":
            def marcar_como_pendiente():
                conn = sqlite3.connect("biblioteca_sjp2.db")
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM libros_pendientes WHERE id_usuario = ? AND id_libro = ?", (self.usuario, libro["id"]))
                if cursor.fetchone():
                    messagebox.showwarning("Aviso", "Este libro ya está en tu lista de pendientes.", parent=sub)
                else:
                    cursor.execute("INSERT INTO libros_pendientes VALUES (?, ?)", (self.usuario, libro["id"]))
                    conn.commit()
                    messagebox.showinfo("Éxito", "Libro marcado como pendiente en tu sesión.", parent=sub)
                    if hasattr(self, "actualizar_tabla_pendientes"):
                        self.actualizar_tabla_pendientes()
                conn.close()

            tk.Button(f_acciones, text="📌 Marcar Pendiente", command=marcar_como_pendiente, bg="#b78a2c", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=5).pack(side="left")

        if self.rol == "Administrador":
            def editar_desde_detalle():
                sub.destroy()
                self.abrir_ventana_modificar_libro_por_id(libro["id"])

            def eliminar_desde_detalle():
                if messagebox.askyesno("Confirmar", f"¿Eliminar permanentemente el libro con Clasificación '{libro['id']}'?", parent=sub):
                    prestamos = self.obtener_prestamos_bd()
                    if any(p["libro"] == libro["id"] for p in prestamos):
                        messagebox.showwarning("No permitido", "El libro tiene préstamos activos asociados.", parent=sub)
                        return

                    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
                    conn = sqlite3.connect("biblioteca_sjp2.db")
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM libros WHERE id = ?", (libro["id"],))
                    cursor.execute("INSERT INTO eliminados_log (admin, id_libro, titulo, fecha) VALUES (?, ?, ?, ?)", (self.usuario, libro["id"], libro["titulo"], fecha_actual))
                    conn.commit()
                    conn.close()

                    self.actualizar_tabla_libros()
                    self.actualizar_dashboard()
                    if hasattr(self, "actualizar_tabla_eliminados"):
                        self.actualizar_tabla_eliminados()
                    if hasattr(self, "mostrar_todos"):
                        self.mostrar_todos()
                    sub.destroy()
                    messagebox.showinfo("Eliminado", "El libro fue eliminado del sistema y registrado en auditoría.")

            tk.Button(f_acciones, text="✎ Editar", command=editar_desde_detalle, bg="#286090", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=5).pack(side="left", padx=5)
            tk.Button(f_acciones, text="🗑 Eliminar", command=eliminar_desde_detalle, bg="#9f2d35", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=5).pack(side="left", padx=2)

        tk.Button(f_acciones, text="Cerrar", command=sub.destroy, bg=self.color_primario, fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=14, pady=5).pack(side="right")

    def abrir_ventana_modificar_libro_por_id(self, lid):
        libros = self.obtener_libros_bd()
        libro_sel = next((l for l in libros if l["id"] == lid), None)
        if not libro_sel: return

        win = tk.Toplevel(self.root)
        win.title("Modificar Libro")
        win.geometry("440x600")
        win.configure(bg="white")
        win.transient(self.root)
        win.grab_set()

        form = tk.Frame(win, bg="white", padx=16)
        form.pack(fill="both", expand=True)

        campos = [
            ("Título:", "e_titulo", libro_sel["titulo"]),
            ("Autor:", "e_autor", libro_sel["autor"]),
            ("Categoría / Tema:", "e_cat", libro_sel["categoria"]),
            ("Año:", "e_anio", str(libro_sel["anio"])),
            ("Editorial:", "e_edit", libro_sel["editorial"]),
            ("Ciudad:", "e_ciudad", libro_sel["ciudad"]),
            ("Ejemplares:", "e_ejem", str(libro_sel["ejemplares"])),
            ("Ubicación Estantería:", "e_ubi", libro_sel["ubicacion"]),
            ("Resumen:", "e_res", libro_sel["resumen"])
        ]

        entries = {}
        tk.Label(form, text=f"Clasificación Dewey (Fija): {lid}", font=("Segoe UI", 10, "bold"), bg="white", fg="#7a5b12").pack(anchor="w", pady=4)

        for lbl, key, val in campos:
            tk.Label(form, text=lbl, font=("Segoe UI", 9, "bold"), bg="white", fg="#344054").pack(anchor="w")
            if key == "e_res":
                e = tk.Text(form, width=44, height=3, font=("Segoe UI", 9))
                e.insert("1.0", val)
            else:
                e = tk.Entry(form, width=48, font=("Segoe UI", 10))
                e.insert(0, val)
                self.configurar_mayusculas(e)
            e.pack(pady=2)
            entries[key] = e

        def actualizar():
            data = {}
            for k, v in entries.items():
                if k == "e_res":
                    data[k] = v.get("1.0", tk.END).strip()
                else:
                    data[k] = v.get().strip()

            try:
                anio = int(data["e_anio"])
                ejem = int(data["e_ejem"])
                if ejem < 0: raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Año y Ejemplares deben ser numéricos.", parent=win)
                return

            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE libros SET titulo = ?, autor = ?, categoria = ?, anio = ?, ejemplares = ?, ubicacion = ?, editorial = ?, ciudad = ?, resumen = ?
                WHERE id = ?
            """, (data["e_titulo"], data["e_autor"], data["e_cat"], anio, ejem, data["e_ubi"], data["e_edit"], data["e_ciudad"], data["e_res"], lid))
            conn.commit()
            conn.close()

            self.actualizar_tabla_libros()
            if hasattr(self, "mostrar_todos"):
                self.mostrar_todos()
            self.actualizar_dashboard()
            win.destroy()
            messagebox.showinfo("Actualizado", "Modificación guardada correctamente.")

        tk.Button(win, text="Actualizar", command=actualizar, bg="#286090", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=6).pack(pady=8)

    def abrir_ventana_nuevo_libro(self):
        win = tk.Toplevel(self.root)
        win.title("Registrar Nuevo Libro")
        win.geometry("440x620")
        win.configure(bg="white")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="Registro Manual de Libro", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg="white").pack(pady=8)
        
        form = tk.Frame(win, bg="white", padx=16)
        form.pack(fill="both", expand=True)

        campos = [
            ("Clasificación Dewey (ID):", "e_id"),
            ("Título:", "e_titulo"),
            ("Autor:", "e_autor"),
            ("Categoría / Tema:", "e_cat"),
            ("Año de Publicación:", "e_anio"),
            ("Editorial:", "e_edit"),
            ("Ciudad / Lugar:", "e_ciudad"),
            ("Ejemplares:", "e_ejem"),
            ("Ubicación Estantería (Ej: ANTIGUO TESTAMENTO):", "e_ubi"),
            ("Resumen / De qué trata:", "e_res")
        ]

        entries = {}
        for lbl, key in campos:
            tk.Label(win, text=lbl, font=("Segoe UI", 9, "bold"), bg="white", fg="#344054").pack(anchor="w")
            if key == "e_res":
                e = tk.Text(form, width=44, height=3, font=("Segoe UI", 9))
            else:
                e = tk.Entry(form, width=48, font=("Segoe UI", 10))
                self.configurar_mayusculas(e)
            e.pack(pady=2)
            entries[key] = e

        def guardar():
            vals = {}
            for k, v in entries.items():
                if k == "e_res":
                    vals[k] = v.get("1.0", tk.END).strip()
                else:
                    vals[k] = v.get().strip()

            if any(not v for k, v in vals.items() if k != "e_res"):
                messagebox.showerror("Error", "Complete los campos obligatorios.", parent=win)
                return
            
            libros = self.obtener_libros_bd()
            if any(l["id"] == vals["e_id"] for l in libros):
                messagebox.showerror("Error", "La clasificación Dewey ya existe.", parent=win)
                return

            try:
                anio = int(vals["e_anio"])
                ejem = int(vals["e_ejem"])
                if ejem < 0: raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Año y Ejemplares deben ser numéricos.", parent=win)
                return

            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO libros VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (
                vals["e_id"], vals["e_titulo"], vals["e_autor"],
                vals["e_cat"], anio, ejem, vals["e_ubi"], vals["e_edit"], vals["e_ciudad"], vals["e_id"], vals["e_res"]
            ))
            conn.commit()
            conn.close()

            self.actualizar_tabla_libros()
            if hasattr(self, "mostrar_todos"):
                self.mostrar_todos()
            self.actualizar_dashboard()
            win.destroy()
            messagebox.showinfo("Éxito", "Libro registrado correctamente.")

        tk.Button(win, text="Guardar Libro", command=guardar, bg="#2f855a", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=6).pack(pady=8)

    def abrir_ventana_modificar_libro(self):
        sel = self.tabla_libros.selection()
        if not sel:
            messagebox.showwarning("Selección", "Seleccione un libro de la tabla para modificar.")
            return
        vals = self.tabla_libros.item(sel[0], "values")
        self.abrir_ventana_modificar_libro_por_id(vals[0])

    def abrir_ventana_eliminar_libro(self):
        sel = self.tabla_libros.selection()
        if not sel:
            messagebox.showwarning("Selección", "Seleccione un libro de la tabla para eliminar.")
            return
        vals = self.tabla_libros.item(sel[0], "values")
        lid = vals[0]
        titulo_libro = vals[1]
        
        if messagebox.askyesno("Confirmar", f"¿Eliminar permanentemente el libro con Clasificación '{lid}'?"):
            prestamos = self.obtener_prestamos_bd()
            if any(p["libro"] == lid for p in prestamos):
                messagebox.showwarning("No permitido", "El libro tiene préstamos activos asociados.")
                return

            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM libros WHERE id = ?", (lid,))
            cursor.execute("INSERT INTO eliminados_log (admin, id_libro, titulo, fecha) VALUES (?, ?, ?, ?)", (self.usuario, lid, titulo_libro, fecha_actual))
            conn.commit()
            conn.close()

            self.actualizar_tabla_libros()
            if hasattr(self, "mostrar_todos"):
                self.mostrar_todos()
            self.actualizar_dashboard()
            if hasattr(self, "actualizar_tabla_eliminados"):
                self.actualizar_tabla_eliminados()
            messagebox.showinfo("Eliminado", "El libro fue eliminado del sistema y registrado en auditoría.")

    # --- MÓDULO: IMPORTADOR DE ARTÍCULOS OPEN ACCESS CON DOBLE CLIC Y VALIDACIÓN PDF ---
    def construir_seccion_abierto(self):
        cont = tk.Frame(self.tab_abierto, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(cont, text="🌍 Buscador e Importador de Artículos de Acceso Abierto (Open Access)", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg=self.fondo).pack(anchor="w", pady=(0, 4))
        tk.Label(cont, text="Busca literatura en repositorios globales. Haz DOBLE CLIC en cualquier artículo para abrir su enlace web o descargar PDF.", font=("Segoe UI", 9), fg="#667085", bg=self.fondo).pack(anchor="w", pady=(0, 10))

        f_busq = tk.Frame(cont, bg=self.fondo)
        f_busq.pack(fill="x", pady=(0, 10))

        self.txt_busq_abierto = tk.Entry(f_busq, width=60, font=("Segoe UI", 10))
        self.configurar_mayusculas(self.txt_busq_abierto)
        self.txt_busq_abierto.pack(side="left", ipady=4, padx=(0, 8))
        self.txt_busq_abierto.bind("<Return>", lambda e: self.ejecutar_busqueda_abierta())

        tk.Button(f_busq, text="🔍 Buscar en Repositorios", command=self.ejecutar_busqueda_abierta, bg=self.color_primario, fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=4).pack(side="left")

        table_frame = tk.Frame(cont, bg="white")
        table_frame.pack(fill="both", expand=True)

        cols = ("Título del Artículo / Documento", "Autor(es)", "Año", "DOI / Enlace")
        self.tabla_abierto = ttk.Treeview(table_frame, columns=cols, show="headings")
        widths = {"Título del Artículo / Documento": 400, "Autor(es)": 240, "Año": 70, "DOI / Enlace": 220}
        for c in cols:
            self.tabla_abierto.heading(c, text=c)
            self.tabla_abierto.column(c, width=widths[c], anchor="center" if c in ("Año",) else "w")

        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tabla_abierto.yview)
        self.tabla_abierto.configure(yscrollcommand=sb.set)
        self.tabla_abierto.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tabla_abierto.bind("<Double-1>", lambda e: self.ver_detalle_articulo_abierto())

        f_acc = tk.Frame(cont, bg=self.fondo, pady=8)
        f_acc.pack(fill="x")

        tk.Button(f_acc, text="📥 Importar Artículo Seleccionado al Catálogo", command=self.importar_articulo_seleccionado, bg="#2f855a", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=6).pack(anchor="w")

        self.resultados_abiertos_cache = []

    def ejecutar_busqueda_abierta(self):
        query = self.txt_busq_abierto.get().strip()
        if not query:
            messagebox.showwarning("Aviso", "Ingrese un término de búsqueda (ej. antropología filosófica, teología...).")
            return

        for row in self.tabla_abierto.get_children():
            self.tabla_abierto.delete(row)
        self.resultados_abiertos_cache = []

        try:
            url_api = f"https://api.crossref.org/works?query={query}&rows=15"
            headers = {"User-Agent": "BibliotecaSeminaristaSJP2/1.0 (mailto:admin@seminario.edu)"}
            response = requests.get(url_api, headers=headers, timeout=10)

            if response.status_code == 200:
                data = response.json()
                items = data.get('message', {}).get('items', [])
                
                for item in items:
                    titulo_art = item.get('title', ['Sin título'])[0]
                    autores_lista = item.get('author', [])
                    nombres_autores = ", ".join([f"{a.get('given', '')} {a.get('family', '')}" for a in autores_lista]) or "Autor institucional / Desconocido"
                    
                    try:
                        anio_pub = item.get('published-print', {}).get('date-parts', [[2024]])[0][0]
                    except:
                        anio_pub = 2024
                        
                    enlace_url = item.get('URL', '#')
                    doi_val = item.get('DOI', 'OPEN-ACCESS')

                    pdf_url = ""
                    links_lista = item.get('link', [])
                    for link_item in links_lista:
                        content_type = link_item.get('content-type', '')
                        if 'pdf' in content_type or 'application/pdf' in content_type:
                            pdf_url = link_item.get('URL', '')
                            break
                    
                    if not pdf_url and doi_val != 'OPEN-ACCESS':
                        pdf_url = f"https://doi.org/{doi_val}"

                    self.resultados_abiertos_cache.append({
                        'titulo': titulo_art,
                        'autor': nombres_autores,
                        'anio': anio_pub,
                        'enlace': enlace_url,
                        'doi': doi_val,
                        'pdf': pdf_url if pdf_url else enlace_url
                    })

                    self.tabla_abierto.insert("", "end", values=(titulo_art, nombres_autores, anio_pub, doi_val))

                if not self.resultados_abiertos_cache:
                    messagebox.showinfo("Sin resultados", "No se encontraron artículos abiertos para este término.")
            else:
                messagebox.showerror("Error", "Error al conectar con los repositorios externos de Crossref.")
        except Exception as e:
            messagebox.showerror("Error de Conexión", f"No se pudo realizar la consulta:\n{e}")

    def ver_detalle_articulo_abierto(self):
        sel = self.tabla_abierto.selection()
        if not sel:
            return
        idx = self.tabla_abierto.index(sel[0])
        if idx >= len(self.resultados_abiertos_cache):
            return
            
        art = self.resultados_abiertos_cache[idx]

        win = tk.Toplevel(self.root)
        win.title("Enlace Web y Descarga de Artículo Open Access")
        win.geometry("560x420")
        win.configure(bg="white")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="🔗 Acceso Directo al Repositorio y PDF", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg="white").pack(pady=10)

        f_info = tk.Frame(win, bg="white", padx=16)
        f_info.pack(fill="both", expand=True)

        tk.Label(f_info, text="Título:", font=("Segoe UI", 9, "bold"), fg="#475467", bg="white").pack(anchor="w")
        lbl_tit = tk.Entry(f_info, font=("Segoe UI", 9), bg="#f4f6f8", relief="solid", bd=1)
        lbl_tit.insert(0, art['titulo'])
        lbl_tit.config(state="readonly")
        lbl_tit.pack(fill="x", pady=(0, 8))

        tk.Label(f_info, text="Autor(es):", font=("Segoe UI", 9, "bold"), fg="#475467", bg="white").pack(anchor="w")
        lbl_aut = tk.Entry(f_info, font=("Segoe UI", 9), bg="#f4f6f8", relief="solid", bd=1)
        lbl_aut.insert(0, art['autor'])
        lbl_aut.config(state="readonly")
        lbl_aut.pack(fill="x", pady=(0, 8))

        tk.Label(f_info, text="Enlace Web del Artículo (DOI / Página oficial):", font=("Segoe UI", 9, "bold"), fg="#1d3557", bg="white").pack(anchor="w")
        txt_link = tk.Entry(f_info, font=("Segoe UI", 9), bg="white", relief="solid", bd=1, fg="blue")
        txt_link.insert(0, art['enlace'])
        txt_link.pack(fill="x", pady=(0, 8))

        tk.Label(f_info, text="Enlace Directo / PDF:", font=("Segoe UI", 9, "bold"), fg="#b78a2c", bg="white").pack(anchor="w")
        txt_pdf = tk.Entry(f_info, font=("Segoe UI", 9), bg="white", relief="solid", bd=1, fg="blue")
        txt_pdf.insert(0, art['pdf'])
        txt_pdf.pack(fill="x", pady=(0, 12))

        def abrir_enlace(url):
            try:
                webbrowser.open(url)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el navegador:\n{e}", parent=win)

        def descargar_pdf_archivo(url_pdf):
            try:
                archivo_destino = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")],
                    initialfile=f"Articulo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    parent=win
                )
                if not archivo_destino:
                    return

                headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                resp = requests.get(url_pdf, headers=headers, stream=True, timeout=15)
                
                if resp.status_code == 200:
                    contenido_inicial = resp.raw.read(1024)
                    
                    # Validación estricta: Si el servidor devuelve código HTML o error en vez de un PDF real
                    if b"<html" in contenido_inicial.lower() or b"<!doctype" in contenido_inicial.lower():
                        webbrowser.open(url_pdf)
                        messagebox.showinfo("Redirección Web", "El repositorio editorial requiere autenticación o visualización interactiva. Se ha abierto el enlace original en el navegador web para su lectura y descarga.", parent=win)
                        return

                    with open(archivo_destino, 'wb') as f:
                        f.write(contenido_inicial)
                        for chunk in resp.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                
                    messagebox.showinfo("Descarga Exitosa", f"El archivo PDF se ha guardado correctamente en:\n{archivo_destino}", parent=win)
                else:
                    webbrowser.open(url_pdf)
                    messagebox.showinfo("Aviso del Navegador", "El servidor externo requiere redirección web. Se ha abierto la página oficial.", parent=win)
            except Exception as e:
                webbrowser.open(url_pdf)
                messagebox.showinfo("Aviso", f"Se abrió el enlace en el navegador web predeterminado.", parent=win)

        f_btns = tk.Frame(win, bg="white", pady=10)
        f_btns.pack(fill="x", padx=16)

        tk.Button(f_btns, text="🌐 Abrir Página Web", command=lambda: abrir_enlace(art['enlace']), bg="#286090", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6).pack(side="left", padx=2)
        tk.Button(f_btns, text="📥 Descargar / Ver PDF", command=lambda: descargar_pdf_archivo(art['pdf']), bg="#2f855a", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6).pack(side="left", padx=6)
        tk.Button(f_btns, text="Cerrar", command=win.destroy, bg="#667085", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6).pack(side="right", padx=2)

    def importar_articulo_seleccionado(self):
        sel = self.tabla_abierto.selection()
        if not sel:
            messagebox.showwarning("Selección", "Seleccione un artículo de la tabla para importar.")
            return

        idx = self.tabla_abierto.index(sel[0])
        art = self.resultados_abiertos_cache[idx]

        titulo = art['titulo'].upper()
        autor = art['autor'].upper()
        anio = int(art['anio'])
        enlace = art['enlace']

        id_virtual = f"OA-{abs(hash(titulo)) % 10000}"

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO libros VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                       (id_virtual, titulo, autor, "ACCESO ABIERTO", anio, 99, "REPOSITORIO DIGITAL", "Open Access Global", "MEDELLIN", id_virtual, f"Recurso digital abierto importado desde repositorio externo. Enlace web: {enlace}"))
        conn.commit()
        conn.close()

        self.actualizar_tabla_libros()
        if hasattr(self, "mostrar_todos"):
            self.mostrar_todos()
        self.actualizar_dashboard()
        messagebox.showinfo("Éxito", f"¡Artículo '{titulo}' importado correctamente al Catálogo General con ID {id_virtual}!")

    def construir_seccion_usuarios(self):
        cont = tk.Frame(self.tab_usuarios, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        if self.rol == "Administrador":
            form = tk.LabelFrame(cont, text=" Registrar o Gestionar Miembros / Seminaristas ",
                                 bg=self.fondo, font=("Segoe UI", 10, "bold"),
                                 padx=12, pady=8)
            form.pack(fill="x")
            
            self.txt_id_usuario = tk.Entry(form, width=15, font=("Segoe UI", 10))
            self.txt_nombre_usuario = tk.Entry(form, width=24, font=("Segoe UI", 10))
            self.configurar_mayusculas(self.txt_nombre_usuario)
            
            self.txt_correo_usuario = tk.Entry(form, width=28, font=("Segoe UI", 10))
            
            self.combo_rol_usr = ttk.Combobox(
                form, values=["Seminarista", "Profesor", "Sacerdote"],
                state="readonly", width=15, font=("Segoe UI", 10))

            for col, label, widget in [
                (0, "ID / Carnet", self.txt_id_usuario),
                (2, "Nombre completo", self.txt_nombre_usuario),
                (4, "Correo Electrónico", self.txt_correo_usuario),
                (6, "Perfil", self.combo_rol_usr)
            ]:
                tk.Label(form, text=label, bg=self.fondo,
                         font=("Segoe UI", 9, "bold")).grid(
                             row=0, column=col, sticky="w", padx=4)
                widget.grid(row=1, column=col, padx=4, pady=4)
            self.combo_rol_usr.current(0)
            
            tk.Button(form, text="Registrar", command=self.agregar_usuario,
                      bg="#2f855a", fg="white", relief="flat",
                      font=("Segoe UI", 10, "bold")).grid(
                          row=1, column=8, padx=6, ipadx=8, ipady=3)
                          
            tk.Button(form, text="🗑 Eliminar", command=self.eliminar_usuario,
                      bg="#9f2d35", fg="white", relief="flat",
                      font=("Segoe UI", 10, "bold")).grid(
                          row=1, column=9, padx=6, ipadx=8, ipady=3)

        table = tk.Frame(cont)
        table.pack(fill="both", expand=True, pady=8)
        self.tabla_usuarios = ttk.Treeview(
            table, columns=("ID", "Nombre", "Correo Electrónico", "Perfil"), show="headings")
        for c, w in [("ID", 100), ("Nombre", 300), ("Correo Electrónico", 300), ("Perfil", 180)]:
            self.tabla_usuarios.heading(c, text=c)
            self.tabla_usuarios.column(c, width=w)
        self.tabla_usuarios.pack(fill="both", expand=True)
        self.actualizar_tabla_usuarios()

    def agregar_usuario(self):
        uid = self.txt_id_usuario.get().strip()
        nombre = self.txt_nombre_usuario.get().strip()
        correo = self.txt_correo_usuario.get().strip()
        rol = self.combo_rol_usr.get()
        if not uid or not nombre:
            messagebox.showerror("Error", "Complete ID/Carnet y nombre.")
            return
        
        usuarios = self.obtener_usuarios_bd()
        if any(u["id"] == uid for u in usuarios):
            messagebox.showerror("Error", "El ID/Carnet ya existe.")
            return

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("INSERT INTO usuarios_biblio VALUES (?, ?, ?, ?)", (uid, nombre, rol, correo))
        conn.commit()
        conn.close()

        self.actualizar_tabla_usuarios()
        self.txt_id_usuario.delete(0, tk.END)
        self.txt_nombre_usuario.delete(0, tk.END)
        self.txt_correo_usuario.delete(0, tk.END)
        messagebox.showinfo("Éxito", "Miembro registrado correctamente.")

    def eliminar_usuario(self):
        sel = self.tabla_usuarios.selection()
        if not sel:
            messagebox.showwarning("Selección", "Seleccione un miembro de la tabla para eliminar.")
            return
        vals = self.tabla_usuarios.item(sel[0], "values")
        uid = vals[0]

        prestamos = self.obtener_prestamos_bd()
        if any(p["usuario"] == uid for p in prestamos):
            messagebox.showwarning("No permitido", "El miembro tiene préstamos activos asociados.", parent=self.root)
            return

        if not messagebox.askyesno("Confirmar", f"¿Desea eliminar al miembro con ID '{uid}'?"):
            return

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM usuarios_biblio WHERE id = ?", (uid,))
        conn.commit()
        conn.close()

        self.actualizar_tabla_usuarios()
        messagebox.showinfo("Eliminado", "Miembro eliminado del sistema con éxito.")

    def actualizar_tabla_usuarios(self):
        if not hasattr(self, "tabla_usuarios"): return
        for row in self.tabla_usuarios.get_children():
            self.tabla_usuarios.delete(row)
        usuarios = self.obtener_usuarios_bd()
        for u in usuarios:
            self.tabla_usuarios.insert("", "end",
                                       values=(u["id"], u["nombre"], u["correo"], u["tipo"]))

    def construir_seccion_prestamos(self):
        cont = tk.Frame(self.tab_prestamos, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        form = tk.LabelFrame(cont, text=" Sincronización y Escáner de Préstamos / Devoluciones (Con Envío Automático de Correo) ",
                             bg=self.fondo, font=("Segoe UI", 10, "bold"),
                             padx=12, pady=10)
        form.pack(fill="x")
        
        tk.Label(form, text="Escanear Código o Clasificación Dewey:", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4)
        
        f_entry_list = tk.Frame(form, bg=self.fondo)
        f_entry_list.grid(row=1, column=0, padx=4, pady=4, sticky="w")

        self.txt_p_libro = tk.Entry(f_entry_list, width=35, font=("Segoe UI", 10))
        self.txt_p_libro.pack(side="top", anchor="w")
        self.configurar_mayusculas(self.txt_p_libro)
        self.txt_p_libro.bind("<Return>", lambda e: self.registrar_prestamo())

        self.lista_sugerencias_prestamo = tk.Listbox(self.root, width=110, height=8, font=("Segoe UI", 10), bg="white", relief="solid", bd=1)
        self.txt_p_libro.bind("<KeyRelease>", self.actualizar_sugerencias_libros_prestamo)
        self.lista_sugerencias_prestamo.bind("<<ListboxSelect>>", self.seleccionar_sugerencia_libro_prestamo)

        self.txt_p_libro.bind("<Down>", self.enfocar_sugerencia_prestamo)
        self.lista_sugerencias_prestamo.bind("<Return>", self.seleccionar_sugerencia_con_enter_prestamo)
        self.lista_sugerencias_prestamo.bind("<Up>", self.verificar_regresar_input_prestamo)

        tk.Label(form, text="Escanear Carnet / ID de Seminarista:", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)
        
        f_entry_usr = tk.Frame(form, bg=self.fondo)
        f_entry_usr.grid(row=1, column=1, padx=4, pady=4, sticky="nw")

        self.txt_p_usuario = tk.Entry(f_entry_usr, width=25, font=("Segoe UI", 10))
        self.txt_p_usuario.pack(side="top", anchor="w")
        self.configurar_mayusculas(self.txt_p_usuario)
        self.txt_p_usuario.bind("<Return>", lambda e: self.registrar_prestamo())

        self.lista_sugerencias_usuario = tk.Listbox(self.root, width=40, height=4, font=("Segoe UI", 9), bg="white", relief="solid", bd=1)
        self.txt_p_usuario.bind("<KeyRelease>", self.actualizar_sugerencias_usuarios)
        self.lista_sugerencias_usuario.bind("<<ListboxSelect>>", self.seleccionar_sugerencia_usuario)

        if self.rol == "Seminarista":
            self.txt_p_usuario.insert(0, self.usuario)
            self.txt_p_usuario.config(state="disabled")

        tk.Button(form, text="📥 Registrar Préstamo",
                  command=self.registrar_prestamo, bg="#2f855a", fg="white",
                  relief="flat", font=("Segoe UI", 10, "bold"),
                  padx=10, pady=5).grid(row=1, column=2, padx=10, sticky="nw")
        tk.Button(form, text="📤 Registrar Devolución",
                  command=self.registrar_devolucion, bg="#c56b21", fg="white",
                  relief="flat", font=("Segoe UI", 10, "bold"),
                  padx=10, pady=5).grid(row=1, column=3, padx=10, sticky="nw")

        tk.Label(cont, text="Préstamos Activos Sincronizados",
                 font=("Segoe UI", 12, "bold"), fg=self.color_primario,
                 bg=self.fondo).pack(anchor="w", pady=(12, 6))
        
        self.tabla_prestamos = ttk.Treeview(
            cont, columns=("Libro", "Título", "Usuario", "Fecha"),
            show="headings")
        for c, w in [("Libro", 130), ("Título", 320), ("Usuario", 220), ("Fecha", 180)]:
            self.tabla_prestamos.heading(c, text=c)
            self.tabla_prestamos.column(c, width=w)
        self.tabla_prestamos.pack(fill="both", expand=True)
        self.actualizar_tabla_prestamos()

    def enfocar_sugerencia_prestamo(self, event):
        if self.lista_sugerencias_prestamo.winfo_ismapped() and self.lista_sugerencias_prestamo.size() > 0:
            self.lista_sugerencias_prestamo.focus_set()
            self.lista_sugerencias_prestamo.selection_set(0)
            return "break"

    def seleccionar_sugerencia_con_enter_prestamo(self, event):
        sel = self.lista_sugerencias_prestamo.curselection()
        if sel:
            val = self.lista_sugerencias_prestamo.get(sel[0])
            id_dewey = val.split(" - ")[0]
            self.txt_p_libro.delete(0, tk.END)
            self.txt_p_libro.insert(0, id_dewey)
            self.lista_sugerencias_prestamo.place_forget()
            self.txt_p_libro.focus_set()
        return "break"

    def verificar_regresar_input_prestamo(self, event):
        sel = self.lista_sugerencias_prestamo.curselection()
        if not sel or sel[0] == 0:
            self.txt_p_libro.focus_set()
            return "break"

    def actualizar_sugerencias_libros_prestamo(self, event):
        if event.keysym in ("Up", "Down", "Return", "Tab"):
            return

        texto = self.txt_p_libro.get().strip()
        self.lista_sugerencias_prestamo.place_forget()
        if not texto:
            return

        libros = self.obtener_libros_bd()
        norm_texto = normalizar_texto(texto)
        coincidencias = [f"{l['id']} - {l['titulo']} ({l['autor']})" for l in libros if norm_texto in normalizar_texto(l['titulo']) or norm_texto in normalizar_texto(l['autor']) or norm_texto in normalizar_texto(l['id'])]

        if coincidencias:
            self.lista_sugerencias_prestamo.delete(0, tk.END)
            for c in coincidencias:
                self.lista_sugerencias_prestamo.insert(tk.END, c)
            
            x = self.txt_p_libro.winfo_rootx() - self.root.winfo_rootx()
            y = self.txt_p_libro.winfo_rooty() - self.root.winfo_rooty() + self.txt_p_libro.winfo_height() + 2
            self.lista_sugerencias_prestamo.place(x=x, y=y, width=850)
            self.lista_sugerencias_prestamo.lift()

    def seleccionar_sugerencia_libro_prestamo(self, event):
        sel = self.lista_sugerencias_prestamo.curselection()
        if sel:
            val = self.lista_sugerencias_prestamo.get(sel[0])
            id_dewey = val.split(" - ")[0]
            self.txt_p_libro.delete(0, tk.END)
            self.txt_p_libro.insert(0, id_dewey)
            self.lista_sugerencias_prestamo.place_forget()
            self.txt_p_libro.focus_set()

    def actualizar_sugerencias_usuarios(self, event):
        if event.keysym in ("Up", "Down", "Return", "Tab"):
            return
        if self.rol == "Seminarista": return
        texto = self.txt_p_usuario.get().strip()
        self.lista_sugerencias_usuario.place_forget()
        if not texto:
            return

        usuarios = self.obtener_usuarios_bd()
        norm_texto = normalizar_texto(texto)
        coincidencias = [f"{u['id']} - {u['nombre']}" for u in usuarios if norm_texto in normalizar_texto(u['nombre']) or norm_texto in normalizar_texto(u['id'])]

        if coincidencias:
            self.lista_sugerencias_usuario.delete(0, tk.END)
            for c in coincidencias:
                self.lista_sugerencias_usuario.insert(tk.END, c)
            
            x = self.txt_p_usuario.winfo_rootx() - self.root.winfo_rootx()
            y = self.txt_p_usuario.winfo_rooty() - self.root.winfo_rooty() + self.txt_p_usuario.winfo_height() + 2
            self.lista_sugerencias_usuario.place(x=x, y=y, width=320)
            self.lista_sugerencias_usuario.lift()

    def seleccionar_sugerencia_usuario(self, event):
        sel = self.lista_sugerencias_usuario.curselection()
        if sel:
            val = self.lista_sugerencias_usuario.get(sel[0])
            uid = val.split(" - ")[0]
            self.txt_p_usuario.delete(0, tk.END)
            self.txt_p_usuario.insert(0, uid)
            self.lista_sugerencias_usuario.place_forget()

    def registrar_prestamo(self):
        if hasattr(self, "lista_sugerencias_prestamo"):
            self.lista_sugerencias_prestamo.place_forget()
        if hasattr(self, "lista_sugerencias_usuario"):
            self.lista_sugerencias_usuario.place_forget()

        input_libro = self.txt_p_libro.get().strip()
        input_usr = self.usuario if self.rol == "Seminarista" else self.txt_p_usuario.get().strip()

        if not input_libro or not input_usr:
            return 

        libros = self.obtener_libros_bd()
        usuarios = self.obtener_usuarios_bd()
        prestamos = self.obtener_prestamos_bd()

        libro = next((l for l in libros if l["id"] == input_libro or normalizar_texto(l["titulo"]) == normalizar_texto(input_libro)), None)
        if not libro:
            libro = next((l for l in libros if normalizar_texto(input_libro) in normalizar_texto(l["titulo"])), None)

        norm_usr = normalizar_texto(input_usr)
        usuario = next((u for u in usuarios if u["id"] == input_usr or normalizar_texto(u["nombre"]) == norm_usr), None)
        if not usuario:
            usuario = next((u for u in usuarios if norm_usr in normalizar_texto(u["nombre"])), None)

        if not usuario:
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("SELECT usuario, rol, nombre_completo, correo FROM credenciales WHERE usuario = ? COLLATE NOCASE", (input_usr,))
            res_c = cursor.fetchone()
            conn.close()
            if res_c:
                nom = res_c[2] if res_c[2] else res_c[0].upper()
                usuario = {"id": res_c[0], "nombre": nom, "tipo": res_c[1] if res_c[1] else "Seminarista", "correo": res_c[3] if res_c[3] else ""}
            else:
                messagebox.showerror("Error", "Seminarista o miembro no encontrado en el sistema.")
                return

        if not libro:
            messagebox.showerror("Error", "No se encontró ningún libro con ese identificador.")
            return
            
        if libro["ejemplares"] <= 0:
            messagebox.showwarning("Stock Agotado", f"No se puede realizar el préstamo. El libro '{libro['titulo']}' no tiene ejemplares disponibles.")
            return
            
        if any(p["libro"] == libro["id"] and p["usuario"] == usuario["id"] for p in prestamos):
            messagebox.showwarning("Préstamo existente", "Este miembro ya tiene prestado este libro.")
            return

        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("UPDATE libros SET ejemplares = ejemplares - 1 WHERE id = ?", (libro["id"],))
        cursor.execute("INSERT INTO prestamos VALUES (?, ?, ?)", (libro["id"], usuario["id"], fecha_actual))
        cursor.execute("INSERT INTO historial_movimientos (tipo_movimiento, id_usuario, nombre_usuario, id_libro, titulo_libro, fecha) VALUES (?, ?, ?, ?, ?, ?)", 
                       ("PRESTAMO", usuario["id"], usuario["nombre"], libro["id"], libro["titulo"], fecha_actual))
        conn.commit()
        conn.close()

        if usuario.get("correo"):
            asunto = f"Préstamo Registrado: {libro['titulo']}"
            cuerpo = f"Estimado/a {usuario['nombre']},\n\nSe ha registrado exitosamente el préstamo del libro:\n- Título: {libro['titulo']}\n- Autor: {libro['autor']}\n- Clasificación Dewey: {libro['id']}\n- Fecha: {fecha_actual}\n\nBiblioteca Benedicto XVI - Seminario San Juan Pablo II"
            enviar_correo_automatico(usuario["correo"], asunto, cuerpo)

        self.actualizar_tabla_libros()
        if hasattr(self, "mostrar_todos"):
            self.mostrar_todos()
        self.actualizar_tabla_prestamos()
        if hasattr(self, "actualizar_tabla_admin_logs"):
            self.actualizar_tabla_admin_logs()
        self.actualizar_dashboard()
        
        self.txt_p_libro.delete(0, tk.END)
        if self.rol != "Seminarista":
            self.txt_p_usuario.delete(0, tk.END)

        messagebox.showinfo("Préstamo Sincronizado", f"Préstamo del libro '{libro['titulo']}' asignado a {usuario['nombre']} correctamente.")

    def registrar_devolucion(self):
        sel = self.tabla_prestamos.selection()
        if not sel:
            messagebox.showwarning("Selección", "Seleccione un préstamo activo de la tabla.")
            return
        values = self.tabla_prestamos.item(sel[0], "values")
        lid, titulo_libro, uid_col = str(values[0]), str(values[1]), str(values[2])
        
        if "(" in uid_col and ")" in uid_col:
            uid = uid_col.split("(")[-1].replace(")", "").strip()
        else:
            uid = uid_col

        usuarios = self.obtener_usuarios_bd()
        usr_obj = next((u for u in usuarios if u["id"] == uid), None)
        nombre_usr = usr_obj["nombre"] if usr_obj else uid
        correo_usr = usr_obj["correo"] if usr_obj and "correo" in usr_obj else ""

        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("UPDATE libros SET ejemplares = ejemplares + 1 WHERE id = ?", (lid,))
        cursor.execute("DELETE FROM prestamos WHERE id_libro = ? AND id_usuario = ?", (lid, uid))
        cursor.execute("INSERT INTO historial_movimientos (tipo_movimiento, id_usuario, nombre_usuario, id_libro, titulo_libro, fecha) VALUES (?, ?, ?, ?, ?, ?)", 
                       ("DEVOLUCION", uid, nombre_usr, lid, titulo_libro, fecha_actual))
        conn.commit()
        conn.close()

        if correo_usr:
            asunto = f"Devolución Exitosa: {titulo_libro}"
            cuerpo = f"Estimado/a {nombre_usr},\n\nSe ha registrado correctamente la devolución del libro:\n- Título: {titulo_libro}\n- Clasificación Dewey: {lid}\n- Fecha de devolución: {fecha_actual}\n\nGracias por cumplir con el préstamo.\nBiblioteca Benedicto XVI - Seminario San Juan Pablo II"
            enviar_correo_automatico(correo_usr, asunto, cuerpo)

        self.actualizar_tabla_libros()
        if hasattr(self, "mostrar_todos"):
            self.mostrar_todos()
        self.actualizar_tabla_prestamos()
        if hasattr(self, "actualizar_tabla_admin_logs"):
            self.actualizar_tabla_admin_logs()
        self.actualizar_dashboard()
        messagebox.showinfo("Devolución", "Devolución registrada con éxito y almacenada en el Registro General.")

    def actualizar_tabla_prestamos(self):
        for row in self.tabla_prestamos.get_children():
            self.tabla_prestamos.delete(row)
        prestamos = self.obtener_prestamos_bd()
        libros = self.obtener_libros_bd()
        usuarios = self.obtener_usuarios_bd()
        
        for p in prestamos:
            if self.rol == "Seminarista":
                if normalizar_texto(p["usuario"]) != normalizar_texto(self.usuario):
                    continue

            libro = next((l for l in libros if l["id"] == p["libro"]), None)
            titulo = libro["titulo"] if libro else "Libro eliminado"
            
            usr_obj = next((u for u in usuarios if u["id"] == p["usuario"]), None)
            if not usr_obj:
                conn = sqlite3.connect("biblioteca_sjp2.db")
                cursor = conn.cursor()
                cursor.execute("SELECT usuario, nombre_completo FROM credenciales WHERE usuario = ?", (p["usuario"],))
                res_c = cursor.fetchone()
                conn.close()
                nombre_usr = res_c[1] if res_c and res_c[1] else (res_c[0].upper() if res_c else p["usuario"])
            else:
                nombre_usr = usr_obj["nombre"]

            texto_usuario = f"{nombre_usr} ({p['usuario']})"

            self.tabla_prestamos.insert("", "end",
                values=(p["libro"], titulo, texto_usuario, p["fecha"]))

    def construir_seccion_buscar(self):
        cont = tk.Frame(self.tab_buscar, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        top = tk.Frame(cont, bg=self.fondo)
        top.pack(fill="x")
        tk.Label(top, text="Búsqueda Global por Clasificación o Texto (Doble clic para ver detalles, editar o eliminar)",
                 font=("Segoe UI", 11, "bold"), fg=self.color_primario,
                 bg=self.fondo).pack(side="left")
        
        f_busq = tk.Frame(top, bg=self.fondo)
        f_busq.pack(side="left", padx=10)

        self.txt_busqueda = tk.Entry(f_busq, width=40, font=("Segoe UI", 10))
        self.configurar_mayusculas(self.txt_busqueda)
        self.txt_busqueda.pack(side="top", anchor="w", ipady=3)
        self.txt_busqueda.bind("<Return>", lambda e: self.buscar_libros())

        self.lista_sugerencias_busqueda = tk.Listbox(self.root, width=70, height=5, font=("Segoe UI", 9), bg="white", relief="solid", bd=1)
        self.txt_busqueda.bind("<KeyRelease>", self.actualizar_sugerencias_busqueda_global)
        self.lista_sugerencias_busqueda.bind("<<ListboxSelect>>", self.seleccionar_sugerencia_busqueda_global)

        tk.Button(top, text="🔎 Buscar", command=self.buscar_libros,
                  bg=self.color_primario, fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=12, pady=5).pack(side="left", padx=5)
        tk.Button(top, text="Mostrar Todo", command=self.mostrar_todos,
                  bg="#667085", fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=10, pady=5).pack(
                      side="left", padx=4)

        table = tk.Frame(cont)
        table.pack(fill="both", expand=True, pady=10)
        self.tabla_busqueda = ttk.Treeview(
            table, columns=("Clasificación Dewey", "Título", "Autor", "Categoría",
                            "Disp.", "Ubicación General"), show="headings")
        for c, w in [("Clasificación Dewey", 130), ("Título", 280), ("Autor", 190),
                     ("Categoría", 140), ("Disp.", 60),
                     ("Ubicación General", 160)]:
            self.tabla_busqueda.heading(c, text=c)
            self.tabla_busqueda.column(c, width=w)
        self.tabla_busqueda.pack(fill="both", expand=True)
        
        self.tabla_busqueda.bind("<Double-1>", lambda e: self.ver_subventana_detalle_desde_busqueda())
        self.mostrar_todos()

    def ver_subventana_detalle_desde_busqueda(self):
        sel = self.tabla_busqueda.selection()
        if not sel: return
        vals = self.tabla_busqueda.item(sel[0], "values")
        lid = vals[0]
        libros = self.obtener_libros_bd()
        libro = next((l for l in libros if l["id"] == lid), None)
        if libro:
            self.mostrar_subventana_detalle(libro)

    def actualizar_sugerencias_busqueda_global(self, event):
        if event.keysym in ("Up", "Down", "Return", "Tab"):
            return
        texto = self.txt_busqueda.get().strip()
        self.lista_sugerencias_busqueda.place_forget()
        if not texto:
            return

        libros = self.obtener_libros_bd()
        norm_texto = normalizar_texto(texto)
        coincidencias = [f"{l['id']} - {l['titulo']} ({l['autor']})" for l in libros if norm_texto in normalizar_texto(l['titulo']) or norm_texto in normalizar_texto(l['autor']) or norm_texto in normalizar_texto(l['id'])]

        if coincidencias:
            self.lista_sugerencias_busqueda.delete(0, tk.END)
            for c in coincidencias:
                self.lista_sugerencias_busqueda.insert(tk.END, c)
            
            x = self.txt_busqueda.winfo_rootx() - self.root.winfo_rootx()
            y = self.txt_busqueda.winfo_rooty() - self.root.winfo_rooty() + self.txt_busqueda.winfo_height() + 2
            self.lista_sugerencias_busqueda.place(x=x, y=y, width=650)
            self.lista_sugerencias_busqueda.lift()

    def seleccionar_sugerencia_busqueda_global(self, event):
        sel = self.lista_sugerencias_busqueda.curselection()
        if sel:
            val = self.lista_sugerencias_busqueda.get(sel[0])
            id_dewey = val.split(" - ")[0]
            self.txt_busqueda.delete(0, tk.END)
            self.txt_busqueda.insert(0, id_dewey)
            self.lista_sugerencias_busqueda.place_forget()
            self.buscar_libros()

    def mostrar_todos(self):
        if hasattr(self, "lista_sugerencias_busqueda"):
            self.lista_sugerencias_busqueda.place_forget()
        self.txt_busqueda.delete(0, tk.END)
        self._llenar_busqueda(self.obtener_libros_bd())

    def buscar_libros(self):
        if hasattr(self, "lista_sugerencias_busqueda"):
            self.lista_sugerencias_busqueda.place_forget()
            
        termino = normalizar_texto(self.txt_busqueda.get().strip())
        libros = self.obtener_libros_bd()
        
        resultados = []
        for l in libros:
            if (termino in normalizar_texto(l["titulo"]) or
                termino in normalizar_texto(l["autor"]) or
                termino in normalizar_texto(l["categoria"]) or
                termino in normalizar_texto(l["id"]) or
                termino in normalizar_texto(l["ubicacion"]) or
                termino in normalizar_texto(l["resumen"])):
                resultados.append(l)

        self._llenar_busqueda(resultados)
        if not resultados:
            messagebox.showinfo("Búsqueda", "No se encontraron coincidencias.")

    def _llenar_busqueda(self, resultados):
        for row in self.tabla_busqueda.get_children():
            self.tabla_busqueda.delete(row)
        for l in resultados:
            self.tabla_busqueda.insert("", "end",
                values=(l["id"], l["titulo"], l["autor"], l["categoria"],
                        l["ejemplares"], l["ubicacion"]))

    def construir_seccion_pendientes(self):
        cont = tk.Frame(self.tab_pendientes, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(cont, text="Libros Marcados como Pendientes (Recordatorio Personal)", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg=self.fondo).pack(anchor="w", pady=(0, 6))

        table_frame = tk.Frame(cont, bg="white")
        table_frame.pack(fill="both", expand=True, pady=(0, 10))

        cols = ("Clasificación Dewey", "Título del Libro", "Autor", "Ejemplares Disp.")
        self.tabla_pendientes = ttk.Treeview(table_frame, columns=cols, show="headings")
        widths = {"Clasificación Dewey": 150, "Título del Libro": 350, "Autor": 220, "Ejemplares Disp.": 120}
        for c in cols:
            self.tabla_pendientes.heading(c, text=c)
            self.tabla_pendientes.column(c, width=widths[c], anchor="center" if c in ("Clasificación Dewey", "Ejemplares Disp.") else "w")

        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tabla_pendientes.yview)
        self.tabla_pendientes.configure(yscrollcommand=sb.set)
        self.tabla_pendientes.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def quitar_pendiente():
            sel = self.tabla_pendientes.selection()
            if not sel:
                messagebox.showwarning("Selección", "Seleccione un libro pendiente para quitar de la lista.")
                return
            vals = self.tabla_pendientes.item(sel[0], "values")
            lid = vals[0]
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM libros_pendientes WHERE id_usuario = ? AND id_libro = ?", (self.usuario, lid))
            conn.commit()
            conn.close()
            self.actualizar_tabla_pendientes()
            messagebox.showinfo("Éxito", "Libro removido de pendientes.")

        tk.Button(cont, text="🗑 Quitar de Pendientes", command=quitar_pendiente, bg="#9f2d35", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=6).pack(anchor="w")
        self.actualizar_tabla_pendientes()

    def actualizar_tabla_pendientes(self):
        if not hasattr(self, "tabla_pendientes"): return
        for row in self.tabla_pendientes.get_children():
            self.tabla_pendientes.delete(row)

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id_libro FROM libros_pendientes WHERE id_usuario = ?", (self.usuario,))
        filas = cursor.fetchall()
        conn.close()

        libros = self.obtener_libros_bd()
        ids_pendientes = [f[0] for f in filas]
        for l in libros:
            if l["id"] in ids_pendientes:
                self.tabla_pendientes.insert("", "end", values=(l["id"], l["titulo"], l["autor"], l["ejemplares"]))

    def construir_seccion_notificaciones_seminarista(self):
        cont = tk.Frame(self.tab_notificaciones, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(cont, text="Bandeja de Notificaciones y Avisos de la Administración", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg=self.fondo).pack(anchor="w", pady=(0, 6))

        table_frame = tk.Frame(cont, bg="white")
        table_frame.pack(fill="both", expand=True)

        cols = ("Fecha y Hora", "Mensaje / Aviso Recibido")
        self.tabla_notif_sem = ttk.Treeview(table_frame, columns=cols, show="headings")
        widths = {"Fecha y Hora": 180, "Mensaje / Aviso Recibido": 650}
        for c in cols:
            self.tabla_notif_sem.heading(c, text=c)
            self.tabla_notif_sem.column(c, width=widths[c], anchor="w")

        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tabla_notif_sem.yview)
        self.tabla_notif_sem.configure(yscrollcommand=sb.set)
        self.tabla_notif_sem.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.actualizar_tabla_notif_seminarista()

    def actualizar_tabla_notif_seminarista(self):
        if not hasattr(self, "tabla_notif_sem"): return
        for row in self.tabla_notif_sem.get_children():
            self.tabla_notif_sem.delete(row)

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT fecha, mensaje FROM notificaciones WHERE destinatario = ? OR destinatario = 'TODOS' ORDER BY id DESC", (self.usuario,))
        filas = cursor.fetchall()
        conn.close()

        for fecha, msg in filas:
            self.tabla_notif_sem.insert("", "end", values=(fecha, msg))

    def construir_seccion_notificaciones_admin(self):
        cont = tk.Frame(self.tab_notif_admin, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(cont, text="Centro de Envío de Notificaciones y Avisos Personalizados", font=("Segoe UI", 12, "bold"), fg=self.color_primario, bg=self.fondo).pack(anchor="w", pady=(0, 8))

        form = tk.LabelFrame(cont, text=" Redactar Notificación ", bg=self.fondo, font=("Segoe UI", 10, "bold"), padx=14, pady=10)
        form.pack(fill="x", pady=(0, 12))

        tk.Label(form, text="Destinatario (ID de Seminarista o 'TODOS'):", bg=self.fondo, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        self.txt_notif_dest = tk.Entry(form, width=35, font=("Segoe UI", 10))
        self.txt_notif_dest.insert(0, "TODOS")
        self.txt_notif_dest.pack(anchor="w", pady=4)
        self.configurar_mayusculas(self.txt_notif_dest)

        tk.Label(form, text="Mensaje o Aviso:", bg=self.fondo, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 0))
        self.txt_notif_msg = tk.Text(form, width=70, height=4, font=("Segoe UI", 10))
        self.txt_notif_msg.pack(anchor="w", pady=4)

        def enviar_notificacion():
            dest = self.txt_notif_dest.get().strip()
            msg = self.txt_notif_msg.get("1.0", tk.END).strip()
            if not dest or not msg:
                messagebox.showerror("Error", "Complete el destinatario y el mensaje.")
                return

            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO notificaciones (destinatario, mensaje, fecha) VALUES (?, ?, ?)", (dest, msg, fecha_actual))
            conn.commit()
            conn.close()

            usuarios = self.obtener_usuarios_bd()
            if dest == "TODOS":
                for u in usuarios:
                    if u.get("correo"):
                        enviar_correo_automatico(u["correo"], "Aviso Institucional - Biblioteca SJP2", msg)
            else:
                usr_obj = next((u for u in usuarios if u["id"] == dest or normalizar_texto(u["nombre"]) == normalizar_texto(dest)), None)
                if usr_obj and usr_obj.get("correo"):
                    enviar_correo_automatico(usr_obj["correo"], "Aviso Institucional - Biblioteca SJP2", msg)

            self.txt_notif_msg.delete("1.0", tk.END)
            messagebox.showinfo("Éxito", "Notificación guardada y enviada correctamente por correo.")
            self.actualizar_tabla_historial_notif()

        tk.Button(form, text="📤 Enviar Notificación y Correo", command=enviar_notificacion, bg="#2f855a", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=6).pack(anchor="w", pady=6)

        f_hist_top = tk.Frame(cont, bg=self.fondo)
        f_hist_top.pack(fill="x", pady=(6, 4))
        
        tk.Label(f_hist_top, text="Historial de Notificaciones Enviadas", font=("Segoe UI", 11, "bold"), fg=self.color_primario, bg=self.fondo).pack(side="left")

        def eliminar_notificacion_admin():
            sel = self.tabla_historial_notif.selection()
            if not sel:
                messagebox.showwarning("Selección", "Seleccione una notificación de la tabla para eliminar.")
                return
            
            vals = self.tabla_historial_notif.item(sel[0], "values")
            dest, fecha, msg = vals[0], vals[1], vals[2]

            if messagebox.askyesno("Confirmar", "¿Desea eliminar permanentemente esta notificación del historial?"):
                conn = sqlite3.connect("biblioteca_sjp2.db")
                cursor = conn.cursor()
                cursor.execute("DELETE FROM notificaciones WHERE destinatario = ? AND fecha = ? AND mensaje = ?", (dest, fecha, msg))
                conn.commit()
                conn.close()

                self.actualizar_tabla_historial_notif()
                messagebox.showinfo("Eliminado", "Notificación eliminada correctamente.")

        tk.Button(f_hist_top, text="🗑 Eliminar Notificación Seleccionada", command=eliminar_notificacion_admin, bg="#9f2d35", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=4).pack(side="right")
        
        table_frame = tk.Frame(cont, bg="white")
        table_frame.pack(fill="both", expand=True)

        cols = ("Destinatario", "Fecha y Hora", "Mensaje Enviado")
        self.tabla_historial_notif = ttk.Treeview(table_frame, columns=cols, show="headings")
        widths = {"Destinatario": 150, "Fecha y Hora": 170, "Mensaje Enviado": 520}
        for c in cols:
            self.tabla_historial_notif.heading(c, text=c)
            self.tabla_historial_notif.column(c, width=widths[c], anchor="w")

        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tabla_historial_notif.yview)
        self.tabla_historial_notif.configure(yscrollcommand=sb.set)
        self.tabla_historial_notif.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.actualizar_tabla_historial_notif()

    def actualizar_tabla_historial_notif(self):
        if not hasattr(self, "tabla_historial_notif"): return
        for row in self.tabla_historial_notif.get_children():
            self.tabla_historial_notif.delete(row)

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT destinatario, fecha, mensaje FROM notificaciones ORDER BY id DESC")
        filas = cursor.fetchall()
        conn.close()

        for dest, fecha, msg in filas:
            self.tabla_historial_notif.insert("", "end", values=(dest, fecha, msg))

    def construir_seccion_admin_logs(self):
        cont = tk.Frame(self.tab_admin_logs, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        notebook_logs = ttk.Notebook(cont)
        notebook_logs.pack(fill="both", expand=True)

        tab_log_prestamos = ttk.Frame(notebook_logs)
        tab_log_eliminados = ttk.Frame(notebook_logs)

        notebook_logs.add(tab_log_prestamos, text="  📋 Registro General de Movimientos  ")
        notebook_logs.add(tab_log_eliminados, text="  🗑 Auditoría de Libros Eliminados  ")

        f1 = tk.Frame(tab_log_prestamos, bg=self.fondo)
        f1.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(f1, text="Historial permanente de préstamos y devoluciones (Sin eliminación automática)", font=("Segoe UI", 11, "bold"), fg=self.color_primario, bg=self.fondo).pack(anchor="w", pady=(0, 6))
        
        cols1 = ("Movimiento", "Seminarista ID", "Nombre Miembro", "Clasificación Dewey", "Título del Libro", "Fecha", "Hora", "Año")
        self.tabla_admin_logs = ttk.Treeview(f1, columns=cols1, show="headings")
        widths1 = {"Movimiento": 110, "Seminarista ID": 110, "Nombre Miembro": 180, "Clasificación Dewey": 130, "Título del Libro": 280, "Fecha": 100, "Hora": 70, "Año": 60}
        for c in cols1:
            self.tabla_admin_logs.heading(c, text=c)
            self.tabla_admin_logs.column(c, width=widths1[c], anchor="center" if c in ("Movimiento", "Seminarista ID", "Clasificación Dewey", "Fecha", "Hora", "Año") else "w")
        self.tabla_admin_logs.pack(fill="both", expand=True)

        f2 = tk.Frame(tab_log_eliminados, bg=self.fondo)
        f2.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(f2, text="Registro Histórico de Libros Eliminados por Administradores", font=("Segoe UI", 11, "bold"), fg=self.color_primario, bg=self.fondo).pack(anchor="w", pady=(0, 6))

        cols2 = ("Administrador", "Clasificación Dewey (ID)", "Título del Libro", "Fecha y Hora de Eliminación")
        self.tabla_eliminados = ttk.Treeview(f2, columns=cols2, show="headings")
        widths2 = {"Administrador": 160, "Clasificación Dewey (ID)": 160, "Título del Libro": 380, "Fecha y Hora de Eliminación": 200}
        for c in cols2:
            self.tabla_eliminados.heading(c, text=c)
            self.tabla_eliminados.column(c, width=widths2[c], anchor="center" if c in ("Clasificación Dewey (ID)", "Fecha y Hora de Eliminación") else "w")
        self.tabla_eliminados.pack(fill="both", expand=True)

        self.actualizar_tabla_admin_logs()
        self.actualizar_tabla_eliminados()

    def actualizar_tabla_admin_logs(self):
        if not hasattr(self, "tabla_admin_logs"):
            return
        for row in self.tabla_admin_logs.get_children():
            self.tabla_admin_logs.delete(row)
            
        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT tipo_movimiento, id_usuario, nombre_usuario, id_libro, titulo_libro, fecha 
            FROM historial_movimientos 
            ORDER BY id DESC
        """)
        filas = cursor.fetchall()
        conn.close()

        for tipo_mov, uid, nombre_usr, lid, titulo_libro, fecha_str in filas:
            mov_label = "📥 PRÉSTAMO" if tipo_mov == "PRESTAMO" else "📤 DEVOLUCIÓN"
            titulo = titulo_libro if titulo_libro else "Libro no especificado"
            
            partes = fecha_str.split(" ") if fecha_str else ["", ""]
            fecha_puro = partes[0]
            hora_puro = partes[1] if len(partes) > 1 else ""
            anio_puro = fecha_puro.split("-")[0] if "-" in fecha_puro else ""

            self.tabla_admin_logs.insert("", "end", values=(mov_label, uid, nombre_usr, lid, titulo, fecha_puro, hora_puro, anio_puro))

    def actualizar_tabla_eliminados(self):
        if not hasattr(self, "tabla_eliminados"):
            return
        for row in self.tabla_eliminados.get_children():
            self.tabla_eliminados.delete(row)

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT admin, id_libro, titulo, fecha FROM eliminados_log ORDER BY id DESC")
        filas = cursor.fetchall()
        conn.close()

        for admin, lid, titulo, fecha in filas:
            self.tabla_eliminados.insert("", "end", values=(admin, lid, titulo, fecha))

    def construir_seccion_configuracion(self):
        container = tk.Frame(self.tab_config, bg=self.fondo)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=self.fondo, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.fondo)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cont = tk.Frame(scrollable_frame, bg=self.fondo)
        cont.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(cont, text="Administración del Sistema",
                 font=("Segoe UI", 14, "bold"), fg=self.color_primario,
                 bg=self.fondo).pack(anchor="w")
        tk.Label(cont, text="Personalización institucional, servidor de correo (SMTP) y gestión de cuentas.",
                 font=("Segoe UI", 9), fg="#667085", bg=self.fondo).pack(
                     anchor="w", pady=(0, 10))

        box = tk.LabelFrame(cont, text=" Identidad Institucional y Logotipo ",
                            bg=self.fondo, font=("Segoe UI", 10, "bold"),
                            padx=12, pady=10)
        box.pack(fill="x", pady=(0, 10))

        tk.Label(box, text="Institución:", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        self.txt_edit_titulo = tk.Entry(box, width=50, font=("Segoe UI", 10))
        self.txt_edit_titulo.insert(0, self.titulo_texto)
        self.txt_edit_titulo.grid(row=0, column=1, padx=8)

        tk.Label(box, text="Subtítulo:", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=4)
        self.txt_edit_sub = tk.Entry(box, width=50, font=("Segoe UI", 10))
        self.txt_edit_sub.insert(0, self.subtitulo_texto)
        self.txt_edit_sub.grid(row=1, column=1, padx=8)

        tk.Label(box, text="Logotipo (PC):", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=4)
        
        f_logo_btn = tk.Frame(box, bg=self.fondo)
        f_logo_btn.grid(row=2, column=1, sticky="w", padx=8)
        
        self.lbl_ruta_logo = tk.Label(f_logo_btn, text="Ninguna imagen seleccionada", font=("Segoe UI", 9, "italic"), fg="#667085", bg=self.fondo)
        self.lbl_ruta_logo.pack(side="left", padx=(0, 10))
        
        tk.Button(f_logo_btn, text="📁 Cargar Imagen...", command=self.seleccionar_imagen_logo,
                  bg="#286090", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=3).pack(side="left")

        tk.Button(box, text="Guardar Identidad y Logo",
                  command=self.aplicar_personalizacion, bg="#2f855a",
                  fg="white", relief="flat", font=("Segoe UI", 10, "bold"),
                  padx=12, pady=5).grid(row=3, column=1, sticky="w", pady=8)

        box_smtp = tk.LabelFrame(cont, text=" Configuración de Correo Electrónico Automático (SMTP) ",
                                 bg=self.fondo, font=("Segoe UI", 10, "bold"),
                                 padx=12, pady=10)
        box_smtp.pack(fill="x", pady=(0, 10))

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        cursor.execute("SELECT smtp_correo, smtp_password FROM configuracion WHERE id = 1")
        smtp_res = cursor.fetchone()
        conn.close()

        smtp_mail_val = smtp_res[0] if smtp_res and smtp_res[0] else ""
        smtp_pass_val = smtp_res[1] if smtp_res and smtp_res[1] else ""

        tk.Label(box_smtp, text="Correo Remitente (Gmail):", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        self.txt_smtp_mail = tk.Entry(box_smtp, width=40, font=("Segoe UI", 10))
        self.txt_smtp_mail.insert(0, smtp_mail_val)
        self.txt_smtp_mail.grid(row=0, column=1, padx=8)

        tk.Label(box_smtp, text="Contraseña de Aplicación:", bg=self.fondo, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=4)
        self.txt_smtp_pass = tk.Entry(box_smtp, width=40, show="•", font=("Segoe UI", 10))
        self.txt_smtp_pass.insert(0, smtp_pass_val)
        self.txt_smtp_pass.grid(row=1, column=1, padx=8, pady=4)

        def guardar_smtp():
            mail = self.txt_smtp_mail.get().strip()
            pwd = self.txt_smtp_pass.get().strip()
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("UPDATE configuracion SET smtp_correo = ?, smtp_password = ? WHERE id = 1", (mail, pwd))
            conn.commit()
            conn.close()
            messagebox.showinfo("Guardado", "Credenciales SMTP de correo actualizadas exitosamente.")

        tk.Button(box_smtp, text="Guardar Configuración Correo", command=guardar_smtp, bg="#286090", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=10, pady=4).grid(row=2, column=1, sticky="w", padx=8, pady=6)

        box_cuentas = tk.LabelFrame(cont, text=" Gestión de Credenciales y Accesos ",
                                    bg=self.fondo, font=("Segoe UI", 10, "bold"),
                                    padx=12, pady=12)
        box_cuentas.pack(fill="x", pady=(0, 16))

        tk.Label(box_cuentas, text="Administre las cuentas de acceso, nombres completos y correos electrónicos en una ventana emergente dedicada.", font=("Segoe UI", 9), fg="#667085", bg=self.fondo).pack(anchor="w", pady=(0, 8))

        tk.Button(box_cuentas, text="⚙ Abrir Ventana de Gestión de Cuentas", command=self.abrir_ventana_gestionar_cuentas,
                  bg="#1d3557", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=14, pady=8).pack(anchor="w")

    def abrir_ventana_gestionar_cuentas(self):
        win = tk.Toplevel(self.root)
        win.title("Gestión de Credenciales de Acceso")
        win.geometry("780x540")
        win.configure(bg="white")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="👥 Administración de Cuentas y Accesos", font=("Segoe UI", 13, "bold"), fg=self.color_primario, bg="white").pack(pady=10)

        f_form = tk.LabelFrame(win, text=" Registrar Nueva Cuenta ", bg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=10)
        f_form.pack(fill="x", padx=16, pady=6)

        tk.Label(f_form, text="Usuario / ID:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4)
        txt_u = tk.Entry(f_form, width=22, font=("Segoe UI", 10))
        txt_u.grid(row=1, column=0, padx=4, pady=2)

        tk.Label(f_form, text="Nombre Completo:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)
        txt_nom = tk.Entry(f_form, width=32, font=("Segoe UI", 10))
        self.configurar_mayusculas(txt_nom)
        txt_nom.grid(row=1, column=1, padx=4, pady=2)

        tk.Label(f_form, text="Contraseña:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", padx=4, pady=(6,0))
        txt_c = tk.Entry(f_form, width=22, show="•", font=("Segoe UI", 10))
        txt_c.grid(row=3, column=0, padx=4, pady=2)

        tk.Label(f_form, text="Correo Electrónico:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=2, column=1, sticky="w", padx=4, pady=(6,0))
        txt_email = tk.Entry(f_form, width=32, font=("Segoe UI", 10))
        txt_email.grid(row=3, column=1, padx=4, pady=2)

        tk.Label(f_form, text="Rol:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=2, column=2, sticky="w", padx=4, pady=(6,0))
        combo_rol = ttk.Combobox(f_form, values=["Seminarista", "Administrador"], state="readonly", width=18, font=("Segoe UI", 10))
        combo_rol.grid(row=3, column=2, padx=4, pady=2)
        combo_rol.current(0)

        def crear_cuenta_pop():
            u = txt_u.get().strip()
            nom = txt_nom.get().strip()
            c = txt_c.get().strip()
            email = txt_email.get().strip()
            tipo = combo_rol.get()

            if not u or not c or not nom:
                messagebox.showerror("Error", "Ingrese usuario, nombre completo y contraseña.", parent=win)
                return

            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("SELECT usuario FROM credenciales WHERE usuario = ?", (u,))
            if cursor.fetchone():
                messagebox.showerror("Existente", f"El usuario '{u}' ya existe.", parent=win)
                conn.close()
                return

            cursor.execute("INSERT INTO credenciales VALUES (?, ?, ?, ?, ?)", (u, c, tipo, nom, email))
            conn.commit()
            conn.close()

            txt_u.delete(0, tk.END)
            txt_nom.delete(0, tk.END)
            txt_c.delete(0, tk.END)
            txt_email.delete(0, tk.END)
            combo_rol.current(0)
            actualizar_tabla_pop()
            messagebox.showinfo("Éxito", "Cuenta creada correctamente.", parent=win)

        tk.Button(f_form, text="＋ Registrar Cuenta", command=crear_cuenta_pop, bg="#2f855a", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=14, pady=6).grid(row=3, column=3, padx=12, pady=2)

        table_frame = tk.Frame(win, bg="white")
        table_frame.pack(fill="both", expand=True, padx=16, pady=10)

        cols = ("Usuario", "Nombre Completo", "Correo Electrónico", "Rol", "Estado")
        tabla_cuentas_win = ttk.Treeview(table_frame, columns=cols, show="headings")
        widths = {"Usuario": 110, "Nombre Completo": 180, "Correo Electrónico": 200, "Rol": 130, "Estado": 90}
        for col in cols:
            tabla_cuentas_win.heading(col, text=col)
            tabla_cuentas_win.column(col, width=widths[col], anchor="w" if col in ("Nombre Completo", "Correo Electrónico") else "center")
        
        sb = ttk.Scrollbar(table_frame, orient="vertical", command=tabla_cuentas_win.yview)
        tabla_cuentas_win.configure(yscrollcommand=sb.set)
        tabla_cuentas_win.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def actualizar_tabla_pop():
            for row in tabla_cuentas_win.get_children():
                tabla_cuentas_win.delete(row)
            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("SELECT usuario, nombre_completo, correo, rol FROM credenciales")
            filas = cursor.fetchall()
            conn.close()
            for u, nom, correo, rol in sorted(filas):
                tabla_cuentas_win.insert("", "end", values=(u, nom if nom else "", correo if correo else "", rol if rol else "Seminarista", "Activo"))

        def eliminar_cuenta_pop():
            sel = tabla_cuentas_win.selection()
            if not sel:
                messagebox.showwarning("Selección", "Seleccione una cuenta de la lista para eliminar.", parent=win)
                return

            vals = tabla_cuentas_win.item(sel[0], "values")
            u, tipo = vals[0], vals[3]

            if u == self.usuario and tipo == self.rol:
                messagebox.showwarning("No permitido", "No puede eliminar la sesión actual activa.", parent=win)
                return

            conn = sqlite3.connect("biblioteca_sjp2.db")
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM credenciales WHERE rol = 'Administrador'")
            total_admins = cursor.fetchone()[0]

            if tipo == "Administrador" and total_admins <= 1:
                messagebox.showwarning("No permitido", "Debe existir al menos un administrador en el sistema.", parent=win)
                conn.close()
                return

            if not messagebox.askyesno("Confirmar", f"¿Eliminar la cuenta '{u}'?", parent=win):
                conn.close()
                return

            cursor.execute("DELETE FROM credenciales WHERE usuario = ?", (u,))
            conn.commit()
            conn.close()

            actualizar_tabla_pop()
            messagebox.showinfo("Eliminado", "Eliminado del sistema.", parent=win)

        f_btns = tk.Frame(win, bg="white")
        f_btns.pack(fill="x", padx=16, pady=10)

        tk.Button(f_btns, text="🗑 Eliminar Cuenta Seleccionada", command=eliminar_cuenta_pop, bg="#9f2d35", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6).pack(side="left", padx=4)
        tk.Button(f_btns, text="Cerrar Ventana", command=win.destroy, bg=self.color_primario, fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6).pack(side="right", padx=4)

        actualizar_tabla_pop()

    def seleccionar_imagen_logo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Logotipo",
            filetypes=[("Archivos de Imagen", "*.png *.jpg *.jpeg *.gif")]
        )
        if ruta:
            self.ruta_logo_temp = ruta
            self.lbl_ruta_logo.config(text=ruta.split("/")[-1])

    def aplicar_personalizacion(self):
        titulo = self.txt_edit_titulo.get().strip()
        sub = self.txt_edit_sub.get().strip()
        if not titulo or not sub:
            messagebox.showerror("Error", "Complete los campos de título y subtítulo.")
            return

        conn = sqlite3.connect("biblioteca_sjp2.db")
        cursor = conn.cursor()
        
        if hasattr(self, "ruta_logo_temp") and self.ruta_logo_temp:
            with open(self.ruta_logo_temp, "rb") as f:
                logo_bytes = f.read()
            cursor.execute("UPDATE configuracion SET titulo = ?, subtitulo = ?, logo = ? WHERE id = 1", (titulo, sub, logo_bytes))
            self.logo_blob = logo_bytes
        else:
            cursor.execute("UPDATE configuracion SET titulo = ?, subtitulo = ? WHERE id = 1", (titulo, sub))
            
        conn.commit()
        conn.close()

        messagebox.showinfo("Éxito", "Identidad y configuración guardadas correctamente. Reinicie sesión para ver reflejado el nuevo logotipo.")

    def cerrar_sesion(self):
        if not messagebox.askyesno("Cerrar sesión", f"¿Desea cerrar sesión, {self.usuario}?"):
            return
        self.root.destroy()
        login_root = tk.Tk()
        LoginApp(login_root)
        login_root.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    LoginApp(root)
    root.mainloop()