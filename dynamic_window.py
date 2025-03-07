import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from PIL import Image, ImageTk
import sticker as sticker
import printer as printer
from pathlib import Path

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

        return {
            "DOCUMENTO": documento,
            "APELLIDOS": f"{primer_apellido} {segundo_apellido}",
            "NOMBRES": " ".join(nombres),
            "SEXO": sexo,
            "FECHA DE NACIMIENTO": fecha_nacimiento,
            "RH": rh,
        }

class DynamicWindow:
    def __init__(self, campos_cedula, columnas_adicionales, campos_plantilla, path_database, evento_nombre):
        self.campos_cedula = campos_cedula
        self.columnas_adicionales = columnas_adicionales
        self.campos_plantilla = campos_plantilla
        self.path_database = path_database
        self.evento_nombre = evento_nombre
        
        self.ventana = tk.Tk()
        self.configure_window()
        self.load_assets()
        self.create_widgets()
        self.setup_bindings()
        self.ventana.mainloop()

    def configure_window(self):
        self.ventana.title(f"⭐ Registros - {self.evento_nombre}")
        self.ventana.geometry("800x640")
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
                relief="solid")
        
        style.configure('Error.TLabel', 
                background=background_color, 
                foreground=error_color, 
                font=("Segoe UI", 11, "bold"))
        
        style.configure('Status.TLabel', 
                background="#E3F2FD", 
                foreground=text_color, 
                font=("Segoe UI", 12))
        
        style.configure('Highlight.TFrame', 
                       background="#F0F2F5",
                       borderwidth=2,
                       relief="solid")

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
                                      text=" Escaneo de Documento ",
                                      padding=15,
                                      style='Highlight.TFrame')
        scanner_frame.pack(fill='x', pady=10)
        
        self.input_text = tk.Text(scanner_frame, 
                                 height=4,
                                 wrap='word',
                                 font=("Consolas", 10),
                                 bg="white",
                                 relief="solid",
                                 borderwidth=1,
                                 padx=10,
                                 pady=15)
        self.input_text.pack(fill='x', expand=True)
                
        # Add a label for instructions
        instruction_label = ttk.Label(scanner_frame, 
                                      text="Por favor, escanee el documento en el área de texto.",
                                      font=("Segoe UI", 10, "italic"),
                                      foreground="#2D3436")
        instruction_label.pack(pady=(5, 0))

    def create_additional_fields(self, parent):
        fields_frame = ttk.LabelFrame(parent, 
                                     text=" Campos Adicionales ",
                                     padding=10,
                                     style='Highlight.TFrame')
        fields_frame.pack(fill='x', pady=10)
        
        grid_frame = ttk.Frame(fields_frame)
        grid_frame.pack(fill='x', expand=True)
        
        for i, columna in enumerate(self.columnas_adicionales):
            ttk.Label(grid_frame, 
                     text=f"{columna.upper()}:",
                     font=("Segoe UI", 10, "bold")).grid(
                         row=i, column=0, 
                         padx=5, pady=3, sticky='e')
            
            entry = ttk.Entry(grid_frame, style='Input.TEntry')
            entry.grid(row=i, column=1, 
                      padx=5, pady=3, sticky='ew')
            self.entradas_adicionales.append(entry)
            
            grid_frame.columnconfigure(1, weight=1)

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
        self.ventana.bind('<Return>', lambda e: self.guardar_datos())
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
            
            self.generar_sticker(registro)

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
        try:
            df = pd.read_excel(self.path_database)
        except FileNotFoundError:
            df = pd.DataFrame(columns=self.campos_cedula + self.columnas_adicionales)
        
        df = pd.concat([df, pd.DataFrame([registro])], ignore_index=True)
        df.to_excel(self.path_database, index=False)

    def generar_sticker(self, registro):
        datos_plantilla = {campo: registro[campo] for campo in self.campos_plantilla if campo in registro}
        sticker.crear_documento(datos_plantilla)

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

if __name__ == "__main__":
    campos_cedula = ["DOCUMENTO", "APELLIDOS", "NOMBRES", "SEXO", "FECHA DE NACIMIENTO", "RH"]
    columnas_adicionales = ["EMAIL", "TELÉFONO"]
    campos_plantilla = ["DOCUMENTO", "APELLIDOS", "NOMBRES"]
    path_database = "bd_evento.xlsx"
    evento_nombre = "Evento de Prueba"
    
    app = DynamicWindow(campos_cedula, columnas_adicionales, campos_plantilla, path_database, evento_nombre)