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
    def procesar_datos_escaneados(datos):
        partes = datos.strip().split()

        if len(partes) < 5:
            raise ValueError("Datos escaneados incompletos.")

        # El primer campo siempre es el documento
        documento = partes[0]

        # Los apellidos son las siguientes dos columnas
        primer_apellido = partes[1]
        segundo_apellido = partes[2]

        # Los nombres están entre el segundo apellido y el sexo (M/F)
        nombres = []
        for parte in partes[3:]:
            if parte in ("M", "F"):
                sexo = parte
                break
            nombres.append(parte)
        else:
            raise ValueError("No se encontró el campo sexo en los datos escaneados.")

        # Fecha de nacimiento y RH están después del sexo
        resto = partes[partes.index(sexo) + 1:]
        fecha_nacimiento = resto[0] if len(resto) > 0 else ""
        rh = resto[1] if len(resto) > 1 else ""

        # Convertir fecha de nacimiento a formato AAAA-MM-DD
        try:
            if '-' in fecha_nacimiento:
                fecha_nacimiento = datetime.strptime(fecha_nacimiento, "%Y-%m-%d").strftime("%Y-%m-%d")
            else:
                fecha_nacimiento = datetime.strptime(fecha_nacimiento, "%Y%m%d").strftime("%Y-%m-%d")
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
    def __init__(self, campos_cedula, columnas_adicionales, campos_plantilla, path_database, evento_nombre, imprimir_con_titulos):
        self.campos_cedula = campos_cedula
        self.columnas_adicionales = columnas_adicionales
        self.campos_plantilla = campos_plantilla
        self.path_database = path_database
        self.evento_nombre = evento_nombre
        self.imprimir_con_titulos = imprimir_con_titulos
        
        self.ventana = tk.Tk()
        self.configure_window()
        self.load_assets()
        self.create_widgets()
        self.setup_bindings()
        self.ventana.mainloop()

    def configure_window(self):
        self.ventana.title(f"⭐ Registros - {self.evento_nombre}")
        self.ventana.geometry("820x660")
        self.ventana.configure(bg="#F0F2F5")
        self.ventana.resizable(True, True)
        
        style = ttk.Style()
        style.theme_use('clam')
        self.configure_styles(style)

    def configure_styles(self, style):
        primary_color = "#1E88E5"
        secondary_color = "#1565C0"
        background_color = "#F0F2F5"
        text_color = "#2D3436"
        accent_color = "#FFC107"
        error_color = "#D32F2F"

        # Añade este nuevo estilo para LabelFrame sin fondo ni borde
        style.configure('Clean.TLabelframe', 
                    background='#F0F2F5',
                    bordercolor='#F0F2F5',  
                    labelmargins=0,      
                    relief='flat')       
        
        style.configure('Clean.TLabelframe.Label', 
                    background='#F0F2F5', 
                    foreground='#333',     
                    font=('Arial', 12, 'italic')) 

        style.configure('TFrame', background=background_color)
        style.configure('TLabel', 
                background=background_color, 
                foreground=text_color, 
                font=("Segoe UI", 11))
        
        style.configure('Accent.TButton', 
            font=("Segoe UI", 13,), 
            foreground="black", 
            background="#E0E0E0",
            padding=(10, 5),
            borderwidth=1,
            relief="flat")
        
        style.map('Accent.TButton',
              background=[('active', "#BDBDBD"), ('pressed', "#BDBDBD")],
              relief=[('pressed', 'sunken')])
        
        style.configure('Input.TEntry', 
                font=("Segoe UI", 11),
                fieldbackground="white",
                bordercolor="#CCCCCC",
                lightcolor="#CCCCCC",
                darkcolor="#CCCCCC",
                padding=5)
        
        style.configure('Highlight.TFrame', 
                background=background_color,
                borderwidth=2,
                relief="flat")
        
        style.configure('Error.TLabel', 
                background=background_color, 
                foreground=error_color, 
                font=("Segoe UI", 11, "bold"))
        
        style.configure('Status.TLabel', 
                background="#E3F2FD", 
                foreground=text_color, 
                font=("Segoe UI", 12))

    def load_assets(self):
        try:
            self.logo_img = Image.open("_internal/logo.png").resize((250, 60))
            self.logo_tk = ImageTk.PhotoImage(self.logo_img)
            self.ventana.iconbitmap("_internal/logo2.ico")
        except Exception as e:
            print("Error cargando recursos:", str(e))

    def create_widgets(self):
        main_container = ttk.Frame(self.ventana)
        main_container.pack(fill='both', expand=True, padx=20, pady=15)
        
        # Header Section
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill='x', pady=(0, 5), ipady=10)
        
        ttk.Label(header_frame, image=self.logo_tk).pack(pady=5, anchor='center')
        ttk.Label(header_frame, 
             text=f"Registro de Asistentes de {self.evento_nombre}",
             font=("Arial", 14, "bold")).pack(pady=(10, 0), anchor='center')
        
        # Main Content
        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill='both', expand=True)
        
        # Scanner Section
        self.create_scanner_section(content_frame)

        # Additional Fields
        self.entradas_adicionales = []
        if self.columnas_adicionales:
            self.create_additional_fields(content_frame)
        
        # Action Buttons
        self.create_action_buttons(main_container)

                # Status Bar en la parte superior
        self.status_bar = ttk.Label(main_container, 
                   text="Estado: Listo para registrar",
                   anchor='center',
                   font=("Arial", 12, "bold"),
                   background="#F0F2F5",
                   foreground="#2D3436",
                   padding=(10, 5))
        self.status_bar.pack(fill='y', pady=(0, 10))

    def create_scanner_section(self, parent):
        scanner_frame = ttk.LabelFrame(parent, 
                                      text=" Sección de Escaneo de Datos ",
                                      padding=5,
                                      style='Clean.TLabelframe')
        scanner_frame.pack(fill='x', pady=5)

        # Texto de indicacion
        indicacion_label = ttk.Label(scanner_frame, 
                                text="El asterico (*) Significa que son Campos obligatorios",
                                font=("Segoe UI", 12, "italic"),
                                foreground="#666666")
        indicacion_label.pack(pady=(5))

        # Sección de formato esperado
        format_frame = ttk.Frame(scanner_frame)
        format_frame.pack(fill='x', pady=(0, 10))
        
        # Etiquetas de campos en orden
        campos = [
            ("* DOCUMENTO", "#B3E5FC"),
            ("* APELLIDOS", "#F0F4C3"),
            ("* NOMBRES", "#C5CAE9"),
            ("* SEXO (M ó F)", "#FFCDD2"),
            ("* F. NACIMIENTO", "#BBDEFB"),
            ("RH", "#FFCDD2")
        ]
        
        for texto, color in campos:
            lbl = ttk.Label(format_frame, 
                           text=texto,
                           font=("Segoe UI", 8, "bold"),
                           background=color,
                           padding=(5, 2),
                           anchor='center')
            lbl.pack(side='left', expand=True, fill='x', padx=0)

        # Área de texto con placeholder
        self.input_text = tk.Text(scanner_frame, 
                                 height=4,
                                 wrap='word',
                                 font=("Consolas", 12),
                                 bg="white",
                                 relief="flat",
                                 borderwidth=1,
                                 padx=10,
                                 pady=15,
                                 fg='#666666')
        self.input_text.pack(fill='x', expand=True)
        self.input_text.insert('1.0', "Ingrese los datos en el orden mostrado arriba...")
        
        # Configurar eventos para el placeholder
        self.input_text.bind('<FocusIn>', self.clear_placeholder)
        self.input_text.bind('<FocusOut>', self.restore_placeholder)
        self.input_text.bind('<Return>', lambda e: "break") 

        # Texto de ejemplo
        ejemplo_label = ttk.Label(scanner_frame, 
                                text="- EJEMPLO - 1234567890 PEREZ GOMEZ JUAN DAVID M 1990-06-27 A+",
                                font=("Segoe UI", 12, "italic"),
                                foreground="#666666")
        ejemplo_label.pack(pady=(5))

    def clear_placeholder(self, event):
        if self.input_text.get("1.0", "end-1c") == "Ingrese los datos en el orden mostrado arriba...":
            self.input_text.delete("1.0", "end")
            self.input_text.config(fg='black')

    def restore_placeholder(self, event):
        if not self.input_text.get("1.0", "end-1c").strip():
            self.input_text.insert('1.0', "Ingrese los datos en el orden mostrado arriba...")
            self.input_text.config(fg='#666666')


    def create_additional_fields(self, parent):
        fields_frame = ttk.LabelFrame(parent, 
                                     text=" Sección de Campos Adicionales ",
                                     padding=10,
                                     style='Clean.TLabelframe')
        fields_frame.pack(fill='x', pady=0)
        
        grid_frame = ttk.Frame(fields_frame)
        grid_frame.pack(fill='x', expand=True)
        
        # Calcular el nro de columnas
        nro_columnas = 2 if len(self.columnas_adicionales) > 1 else 1
        
        for i, columna in enumerate(self.columnas_adicionales):
            col = i % nro_columnas
            row = i // nro_columnas
            ttk.Label(grid_frame, 
                     text=f"{columna.upper()}:",
                     font=("Segoe UI", 10, "bold")).grid(
                         row=row, column=col*2, 
                         padx=5, pady=3, sticky='e')
            
            entry = ttk.Entry(grid_frame, style='Input.TEntry')
            entry.grid(row=row, column=col*2+1, 
                      padx=5, pady=3, sticky='ew')
            self.entradas_adicionales.append(entry)
            
            grid_frame.columnconfigure(col*2+1, weight=1)

    def create_action_buttons(self, parent):
        btn_container = ttk.Frame(parent)
        btn_container.pack(fill='both', pady=(0, 10))
        
        actions = [
            ("🖨️ Guardar Datos e Imprimir", self.guardar_datos_e_imprimir),
            ("💾 Guardar Datos", self.guardar_datos),
            ("❌ Cerrar Evento", self.cerrar_app)
        ]
        
        for text, command in actions:
            btn = ttk.Button(btn_container,
                            text=text,
                            command=command,
                            style='Accent.TButton')
            btn.pack(side='left', padx=5, expand=True)

    def setup_bindings(self):
        # self.ventana.bind('<Return>', lambda e: self.guardar_datos())
        self.input_text.bind('<Control-a>', self.select_all_text)

    def select_all_text(self, event):
        self.input_text.tag_add('sel', '1.0', 'end')
        return 'break'

    def update_status(self, message, error=False):
        self.status_bar.config(
            text=message,
            foreground='red' if error else 'green',
            font=("Arial", 12, "bold" if error else "normal")
        )
        
    def reset_status(self):
        self.status_bar.config(
            text="Estado: Listo para registrar",
            foreground="black",
            font=("Arial", 12, "bold")
        )

    def guardar_datos_e_imprimir(self):
        try:
            raw_data = self.input_text.get("1.0", tk.END).strip()
            if not raw_data:
                raise ValueError("No se han ingresado datos")
                
            datos_cedula = ProcesadorDatos.procesar_datos_escaneados(raw_data)
            registro = self.generar_registro(datos_cedula)
            self.guardar_en_excel(registro)
            
            self.generar_sticker(registro, self.imprimir_con_titulos)

            ruta_base = Path(__file__).parent if "__file__" in locals() else Path.cwd()
            ruta_salida_word = ruta_base / "sticker_a_imprimir.docx"
            self.update_status(f"Estado: Imprimiendo registro...", error=False)
            self.ventana.update_idletasks()
            self.imprimir_sticker(ruta_salida_word)

            self.update_status(f"Estado: ¡Proceso completado con exito!", error=False)
            self.limpiar_campos()
            self.ventana.after(3000, self.reset_status)
        except Exception as e:
            self.update_status(f"Estado: Error, {str(e)}", error=True)
            self.input_text.config(highlightbackground='red', highlightthickness=1)
            self.ventana.after(4000, self.reset_status)
    
    def guardar_datos(self):
        try:
            raw_data = self.input_text.get("1.0", tk.END).strip()
            if not raw_data:
                raise ValueError("No se han ingresado datos")
                
            datos_cedula = ProcesadorDatos.procesar_datos_escaneados(raw_data)
            registro = self.generar_registro(datos_cedula)
            self.guardar_en_excel(registro)

            self.update_status(f"Estado: Datos almacenados con exito!")
            self.limpiar_campos()
            self.ventana.after(2000, self.reset_status)
        except Exception as e:
            self.update_status(f"Estado: Error, {str(e)}", error=True)
            self.input_text.config(highlightbackground='red', highlightthickness=1)
            self.ventana.after(4000, self.reset_status)

    def generar_registro(self, datos_cedula):
        registro = {campo: datos_cedula[campo] for campo in self.campos_cedula if campo in datos_cedula}
        for columna, entrada in zip(self.columnas_adicionales, self.entradas_adicionales):
            registro[columna] = entrada.get()
        return registro

    def guardar_en_excel(self, registro):
        # Poner como String todos los campos para evitar problemas de formato en Excel
        for key in registro:
            registro[key] = str(registro[key])
        try:
            df = pd.read_excel(self.path_database)
        except FileNotFoundError:
            df = pd.DataFrame(columns=self.campos_cedula + self.columnas_adicionales)
        
        df = pd.concat([df, pd.DataFrame([registro])], ignore_index=True)
        df.to_excel(self.path_database, index=False)

    def generar_sticker(self, registro, imprimir_con_titulos):
        datos_plantilla = {campo: registro[campo] for campo in self.campos_plantilla if campo in registro}
        sticker.crear_documento(datos_plantilla, imprimir_con_titulos)

    def imprimir_sticker(self, ruta_salida_word):
        printer.imprimir_documento(str(ruta_salida_word))
        
    def limpiar_campos(self):
        self.input_text.delete("1.0", tk.END)
        self.input_text.config(highlightbackground='#CCCCCC', highlightthickness=1)
        for entrada in self.entradas_adicionales:
            entrada.delete(0, tk.END)

    def cerrar_app(self):
        if messagebox.askyesno("Confirmar", "¿Está seguro de cerrar el evento?"):
            self.ventana.quit()
            self.ventana.destroy()
