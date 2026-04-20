import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pandas as pd
import sticker as sticker
import printer as printer
from pathlib import Path
from datetime import datetime


class ProcesadorDatos:
    @staticmethod
    def normalizar_fecha(fecha):
        fecha = fecha.strip()
        if not fecha:
            return ""

        formatos = ("%Y-%m-%d", "%Y%m%d", "%d/%m/%Y", "%d-%m-%Y")
        for formato in formatos:
            try:
                return datetime.strptime(fecha, formato).strftime("%Y-%m-%d")
            except ValueError:
                continue

        raise ValueError(
            "Fecha inválida. Use AAAA-MM-DD, AAAAMMDD, DD/MM/AAAA o DD-MM-AAAA."
        )

    @staticmethod
    def procesar_datos_escaneados(datos):
        partes = datos.strip().split()

        if len(partes) < 5:
            raise ValueError("Datos escaneados incompletos.")

        documento = partes[0]
        primer_apellido = partes[1]
        segundo_apellido = partes[2]

        nombres = []
        sexo = None

        for parte in partes[3:]:
            if parte in ("M", "F"):
                sexo = parte
                break
            nombres.append(parte)

        if not sexo:
            raise ValueError("No se encontró el campo sexo en los datos escaneados.")

        resto = partes[partes.index(sexo) + 1:]
        fecha_nacimiento = resto[0] if len(resto) > 0 else ""
        rh = resto[1] if len(resto) > 1 else ""

        try:
            fecha_nacimiento = ProcesadorDatos.normalizar_fecha(fecha_nacimiento) if fecha_nacimiento else ""
        except ValueError:
            fecha_nacimiento = "Formato invalido"

        return {
            "DOCUMENTO": documento,
            "APELLIDOS": f"{primer_apellido} {segundo_apellido}",
            "NOMBRES": " ".join(nombres),
            "SEXO": sexo,
            "FECHA DE NACIMIENTO": fecha_nacimiento,
            "RH": rh,
        }


class DynamicWindow:
    PLACEHOLDER_SCANNER = "Ingrese los datos en el orden mostrado arriba..."

    def __init__(self, campos_cedula, columnas_adicionales, campos_plantilla, path_database, evento_nombre, imprimir_con_titulos, imprimir_con_qr):
        self.campos_cedula = campos_cedula
        self.columnas_adicionales = columnas_adicionales
        self.campos_plantilla = campos_plantilla
        self.path_database = path_database
        self.evento_nombre = evento_nombre
        self.imprimir_con_titulos = imprimir_con_titulos
        self.imprimir_con_qr = imprimir_con_qr

        self.campos_manuales = [
            "DOCUMENTO",
            "APELLIDOS",
            "NOMBRES",
            "SEXO",
            "FECHA DE NACIMIENTO",
            "RH",
        ]

        self.ventana = tk.Tk()
        self.configure_window()
        self.load_assets()
        self.create_widgets()
        self.setup_bindings()
        self.ventana.mainloop()

    def configure_window(self):
        self.ventana.title(f"⭐ Registros - {self.evento_nombre}")
        self.ventana.geometry("820x760")
        self.ventana.configure(bg="#F0F2F5")
        self.ventana.resizable(True, True)
        self.ventana.protocol("WM_DELETE_WINDOW", self.cerrar_app)

        style = ttk.Style()
        style.theme_use("clam")
        self.configure_styles(style)

    def configure_styles(self, style):
        primary_color = "#1E88E5"
        secondary_color = "#1565C0"
        background_color = "#F0F2F5"
        text_color = "#2D3436"
        accent_color = "#FFC107"
        error_color = "#D32F2F"

        style.configure(
            "Clean.TLabelframe",
            background="#F0F2F5",
            bordercolor="#F0F2F5",
            labelmargins=0,
            relief="flat",
        )

        style.configure(
            "Clean.TLabelframe.Label",
            background="#F0F2F5",
            foreground="#333",
            font=("Arial", 12, "italic"),
        )

        style.configure("TFrame", background=background_color)
        style.configure(
            "TLabel",
            background=background_color,
            foreground=text_color,
            font=("Segoe UI", 11),
        )

        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 13),
            foreground="black",
            background="#E0E0E0",
            padding=(10, 5),
            borderwidth=1,
            relief="flat",
        )

        style.map(
            "Accent.TButton",
            background=[("active", "#BDBDBD"), ("pressed", "#BDBDBD")],
            relief=[("pressed", "sunken")],
        )

        style.configure(
            "Input.TEntry",
            font=("Segoe UI", 11),
            fieldbackground="white",
            bordercolor="#CCCCCC",
            lightcolor="#CCCCCC",
            darkcolor="#CCCCCC",
            padding=5,
        )

        style.configure(
            "Highlight.TFrame",
            background=background_color,
            borderwidth=2,
            relief="flat",
        )

        style.configure(
            "Error.TLabel",
            background=background_color,
            foreground=error_color,
            font=("Segoe UI", 11, "bold"),
        )

        style.configure(
            "Status.TLabel",
            background="#E3F2FD",
            foreground=text_color,
            font=("Segoe UI", 12),
        )

    def load_assets(self):
        self.logo_tk = None
        try:
            self.logo_img = Image.open("_internal/logo.png").resize((250, 60))
            self.logo_tk = ImageTk.PhotoImage(self.logo_img)
            self.ventana.iconbitmap("_internal/logo2.ico")
        except Exception as e:
            print("Error cargando recursos:", str(e))

    def create_widgets(self):
        main_container = ttk.Frame(self.ventana)
        main_container.pack(fill="both", expand=True, padx=20, pady=15)

        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 5), ipady=10)

        if self.logo_tk:
            ttk.Label(header_frame, image=self.logo_tk).pack(pady=5, anchor="center")

        ttk.Label(
            header_frame,
            text=f"Registro de Asistentes de {self.evento_nombre}",
            font=("Arial", 14, "bold"),
        ).pack(pady=(10, 0), anchor="center")

        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill="both", expand=True)

        self.create_scanner_section(content_frame)
        self.create_manual_fields(content_frame)

        self.entradas_adicionales = []
        if self.columnas_adicionales:
            self.create_additional_fields(content_frame)

        self.create_action_buttons(main_container)

        self.status_bar = ttk.Label(
            main_container,
            text="Estado: Listo para registrar",
            anchor="center",
            font=("Arial", 12, "bold"),
            background="#F0F2F5",
            foreground="#2D3436",
            padding=(10, 5),
        )
        self.status_bar.pack(fill="y", pady=(0, 10))

    def create_scanner_section(self, parent):
        scanner_frame = ttk.LabelFrame(
            parent,
            text=" Sección de Escaneo de Datos ",
            padding=5,
            style="Clean.TLabelframe",
        )
        scanner_frame.pack(fill="x", pady=5)

        indicacion_label = ttk.Label(
            scanner_frame,
            text="El asterisco (*) significa campos esperados en el escaneo. Para ingreso manual puede dejar campos vacíos.",
            font=("Segoe UI", 11, "italic"),
            foreground="#666666",
        )
        indicacion_label.pack(pady=5)

        format_frame = ttk.Frame(scanner_frame)
        format_frame.pack(fill="x", pady=(0, 8))

        campos = [
            ("* DOCUMENTO", "#B3E5FC"),
            ("* APELLIDOS", "#F0F4C3"),
            ("* NOMBRES", "#C5CAE9"),
            ("* SEXO (M ó F)", "#FFCDD2"),
            ("* F. NACIMIENTO", "#BBDEFB"),
            ("RH", "#FFCDD2"),
        ]

        for texto, color in campos:
            lbl = ttk.Label(
                format_frame,
                text=texto,
                font=("Segoe UI", 8, "bold"),
                background=color,
                padding=(5, 2),
                anchor="center",
            )
            lbl.pack(side="left", expand=True, fill="x", padx=0)

        self.input_text = tk.Text(
            scanner_frame,
            height=2,  # más pequeño, como pidió el usuario
            wrap="word",
            font=("Consolas", 12),
            bg="white",
            relief="flat",
            borderwidth=1,
            padx=10,
            pady=10,
            fg="#666666",
        )
        self.input_text.pack(fill="x", expand=True)
        self.input_text.insert("1.0", self.PLACEHOLDER_SCANNER)

        self.input_text.bind("<FocusIn>", self.clear_placeholder)
        self.input_text.bind("<FocusOut>", self.restore_placeholder)
        self.input_text.bind("<Return>", lambda e: "break")

    def create_manual_fields(self, parent):
        manual_frame = ttk.LabelFrame(
            parent,
            text=" Sección de Ingreso Manual ",
            padding=10,
            style="Clean.TLabelframe",
        )
        manual_frame.pack(fill="x", pady=(5, 5))

        ttk.Label(
            manual_frame,
            text="Puede llenar solo los campos que necesite. Los demás se guardarán vacíos.",
            font=("Segoe UI", 11, "italic"),
            foreground="#666666",
        ).pack(anchor="w", pady=(0, 8))

        self.entradas_manuales = {}
        grid_frame = ttk.Frame(manual_frame)
        grid_frame.pack(fill="x", expand=True)

        definicion_campos = [
            ("DOCUMENTO", "entry"),
            ("APELLIDOS", "entry"),
            ("NOMBRES", "entry"),
            ("SEXO", "combo"),
            ("FECHA DE NACIMIENTO", "entry"),
            ("RH", "combo"),
        ]

        for i, (campo, tipo) in enumerate(definicion_campos):
            row = i // 2
            col = (i % 2) * 2

            ttk.Label(
                grid_frame,
                text=f"{campo}:",
                font=("Segoe UI", 10, "bold"),
            ).grid(row=row, column=col, padx=5, pady=5, sticky="e")

            if tipo == "combo" and campo == "SEXO":
                widget = ttk.Combobox(
                    grid_frame,
                    values=["", "M", "F"],
                    state="readonly",
                    font=("Segoe UI", 10),
                )
                widget.set("")
            elif tipo == "combo" and campo == "RH":
                widget = ttk.Combobox(
                    grid_frame,
                    values=["", "O+", "O-", "A+", "A-", "B+", "B-", "AB+", "AB-"],
                    font=("Segoe UI", 10),
                )
                widget.set("")
            else:
                widget = ttk.Entry(grid_frame, style="Input.TEntry")

            widget.grid(row=row, column=col + 1, padx=5, pady=5, sticky="ew")
            self.entradas_manuales[campo] = widget
            grid_frame.columnconfigure(col + 1, weight=1)

    def create_additional_fields(self, parent):
        fields_frame = ttk.LabelFrame(
            parent,
            text=" Sección de Campos Adicionales ",
            padding=10,
            style="Clean.TLabelframe",
        )
        fields_frame.pack(fill="x", pady=0)

        grid_frame = ttk.Frame(fields_frame)
        grid_frame.pack(fill="x", expand=True)

        nro_columnas = 2 if len(self.columnas_adicionales) > 1 else 1

        for i, columna in enumerate(self.columnas_adicionales):
            col = i % nro_columnas
            row = i // nro_columnas

            ttk.Label(
                grid_frame,
                text=f"{columna.upper()}:",
                font=("Segoe UI", 10, "bold"),
            ).grid(row=row, column=col * 2, padx=5, pady=3, sticky="e")

            entry = ttk.Entry(grid_frame, style="Input.TEntry")
            entry.grid(row=row, column=col * 2 + 1, padx=5, pady=3, sticky="ew")
            self.entradas_adicionales.append(entry)

            grid_frame.columnconfigure(col * 2 + 1, weight=1)

    def create_action_buttons(self, parent):
        btn_container = ttk.Frame(parent)
        btn_container.pack(fill="both", pady=(10, 10))

        actions = [
            ("🖨️ Guardar Datos e Imprimir", self.guardar_datos_e_imprimir),
            ("💾 Guardar Datos", self.guardar_datos),
            ("❌ Cerrar Evento", self.cerrar_app),
        ]

        for text, command in actions:
            btn = ttk.Button(
                btn_container,
                text=text,
                command=command,
                style="Accent.TButton",
            )
            btn.pack(side="left", padx=5, expand=True)

    def setup_bindings(self):
        self.input_text.bind("<Control-a>", self.select_all_text)

    def select_all_text(self, event):
        self.input_text.tag_add("sel", "1.0", "end")
        return "break"

    def clear_placeholder(self, event):
        if self.input_text.get("1.0", "end-1c") == self.PLACEHOLDER_SCANNER:
            self.input_text.delete("1.0", "end")
            self.input_text.config(fg="black")

    def restore_placeholder(self, event):
        if not self.input_text.get("1.0", "end-1c").strip():
            self.input_text.insert("1.0", self.PLACEHOLDER_SCANNER)
            self.input_text.config(fg="#666666")

    def obtener_texto_scanner(self):
        texto = self.input_text.get("1.0", "end-1c").strip()
        if texto == self.PLACEHOLDER_SCANNER:
            return ""
        return texto

    def obtener_datos_manuales(self):
        datos = {campo: "" for campo in self.campos_manuales}

        for campo, widget in self.entradas_manuales.items():
            valor = widget.get().strip()

            if campo == "FECHA DE NACIMIENTO" and valor:
                valor = ProcesadorDatos.normalizar_fecha(valor)

            datos[campo] = valor

        return datos

    def hay_datos_manuales(self, datos_manuales):
        return any(valor.strip() for valor in datos_manuales.values())

    def obtener_datos_principales(self):
        raw_scan = self.obtener_texto_scanner()
        datos_scan = {}
        datos_manuales = self.obtener_datos_manuales()

        if raw_scan:
            datos_scan = ProcesadorDatos.procesar_datos_escaneados(raw_scan)

        if not raw_scan and not self.hay_datos_manuales(datos_manuales):
            raise ValueError("Debe ingresar datos escaneados o al menos un dato manual.")

        datos_finales = {campo: "" for campo in self.campos_manuales}
        datos_finales.update(datos_scan)

        # Lo manual sobrescribe lo escaneado solo si el usuario escribió algo
        for campo, valor in datos_manuales.items():
            if valor != "":
                datos_finales[campo] = valor

        return datos_finales

    def update_status(self, message, error=False):
        self.status_bar.config(
            text=message,
            foreground="red" if error else "green",
            font=("Arial", 12, "bold" if error else "normal"),
        )

    def reset_status(self):
        self.status_bar.config(
            text="Estado: Listo para registrar",
            foreground="black",
            font=("Arial", 12, "bold"),
        )

    def procesar_guardado(self, imprimir=False):
        try:
            datos_principales = self.obtener_datos_principales()
            registro = self.generar_registro(datos_principales)
            self.guardar_en_excel(registro)

            if imprimir:
                self.generar_sticker(registro, self.imprimir_con_titulos, self.imprimir_con_qr)

                ruta_base = Path(__file__).parent if "__file__" in globals() else Path.cwd()
                ruta_salida_word = ruta_base / "sticker_a_imprimir.docx"

                self.update_status("Estado: Imprimiendo registro...", error=False)
                self.ventana.update_idletasks()
                self.imprimir_sticker(ruta_salida_word)

                self.update_status("Estado: ¡Proceso completado con éxito!", error=False)
            else:
                self.update_status("Estado: Datos almacenados con éxito!", error=False)

            self.limpiar_campos()
            self.ventana.after(3000 if imprimir else 2000, self.reset_status)

        except Exception as e:
            self.update_status(f"Estado: Error, {str(e)}", error=True)
            self.ventana.after(4000, self.reset_status)

    def guardar_datos_e_imprimir(self):
        self.procesar_guardado(imprimir=True)

    def guardar_datos(self):
        self.procesar_guardado(imprimir=False)

    def generar_registro(self, datos_cedula):
        registro = {campo: datos_cedula.get(campo, "") for campo in self.campos_cedula}

        for columna, entrada in zip(self.columnas_adicionales, self.entradas_adicionales):
            registro[columna] = entrada.get().strip()

        return registro

    def guardar_en_excel(self, registro):
        for key in registro:
            registro[key] = str(registro[key])

        try:
            df = pd.read_excel(self.path_database, dtype=str)
            df = df.fillna("")
        except FileNotFoundError:
            df = pd.DataFrame(columns=self.campos_cedula + self.columnas_adicionales)

        df = pd.concat([df, pd.DataFrame([registro])], ignore_index=True)
        df.to_excel(self.path_database, index=False)

    def generar_sticker(self, registro, imprimir_con_titulos, imprimir_con_qr):
        datos_plantilla = {
            campo: registro.get(campo, "")
            for campo in self.campos_plantilla
        }
        # Enviar el documetno aparte
        documento_qr = registro.get("DOCUMENTO", "")
        sticker.crear_documento(datos_plantilla, imprimir_con_titulos, imprimir_con_qr, documento_qr)

    def imprimir_sticker(self, ruta_salida_word):
        printer.imprimir_documento(str(ruta_salida_word))

    def limpiar_campos(self):
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert("1.0", self.PLACEHOLDER_SCANNER)
        self.input_text.config(fg="#666666", highlightbackground="#CCCCCC", highlightthickness=1)

        for campo, widget in self.entradas_manuales.items():
            if isinstance(widget, ttk.Combobox):
                widget.set("")
            else:
                widget.delete(0, tk.END)

        for entrada in self.entradas_adicionales:
            entrada.delete(0, tk.END)

    def cerrar_app(self):
        if messagebox.askyesno("Confirmar", "¿Está seguro de cerrar el evento?"):
            self.ventana.quit()
            self.ventana.destroy()