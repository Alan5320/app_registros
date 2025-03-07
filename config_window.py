import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from config_plantilla import ConfiguradorPlantilla

class ConfiguradorEvento:
    def __init__(self):
        self.evento_nombre = None
        self.path_database = None
        self.campos_seleccionados = []
        self.columnas_adicionales = []
        
        # Configuración de estilos
        self.estilo = {
            "fuente_principal": ("Segoe UI", 12),
            "color_fondo": "#F0F2F5",
            "color_botones": "#4A90E2",
            "color_texto": "#2D3436",
            "padding_general": 20,
            "ancho_minimo": 1000,  # Aumentado para mejor distribución
            "alto_minimo": 600
        }

        self.inicializar_ventana()
        
    def inicializar_ventana(self):
        self.ventana_config = tk.Tk()
        self.ventana_config.title("Registros - ServiEventos Col")
        self.ventana_config.configure(bg=self.estilo["color_fondo"])
        self.ventana_config.minsize(self.estilo["ancho_minimo"], self.estilo["alto_minimo"])
        
        try:
            self.ventana_config.iconbitmap("_internal/logo2.ico")
        except:
            pass
        
        self.cargar_logo()
        self.crear_widgets()
        self.ventana_config.mainloop()

    def cargar_logo(self):
        try:
            logo = Image.open("_internal/logo.png")
            logo = logo.resize((250, 60))
            self.logo_tk = ImageTk.PhotoImage(logo)
            label_logo = ttk.Label(self.ventana_config, image=self.logo_tk)
            label_logo.pack(pady=self.estilo["padding_general"])
        except FileNotFoundError:
            pass

    def crear_widgets(self):
        # Contenedor principal
        main_container = ttk.Frame(self.ventana_config)
        main_container.pack(fill="both", expand=True, padx=40, pady=20)
        
        # Nombre del evento
        ttk.Label(main_container, text="Nombre del Evento", font=("Segoe UI", 12, "bold")).pack(anchor="center", pady=(0, 10))
        self.entry_evento = ttk.Entry(main_container, font=("Segoe UI", 12))
        self.entry_evento.pack(fill="x", pady=(0, 20))
        
        # Contenedor de dos columnas (50-50)
        columns_container = ttk.Frame(main_container)
        columns_container.pack(fill="both", expand=True)
        
        # Configurar grid 50-50
        columns_container.grid_columnconfigure(0, weight=1, uniform="group1")
        columns_container.grid_columnconfigure(1, weight=1, uniform="group1")
        columns_container.grid_rowconfigure(0, weight=1)
        
        # Sección izquierda - Campos de cédula
        frame_cedula = ttk.LabelFrame(columns_container, text="Campos de la cédula", padding=15)
        frame_cedula.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Sección derecha - Campos adicionales
        frame_adicionales = ttk.LabelFrame(columns_container, text="Campos adicionales", padding=15)
        frame_adicionales.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        
        # Campos de cédula
        self.crear_campos_cedula(frame_cedula)
        
        # Campos adicionales
        self.crear_campos_adicionales(frame_adicionales)
        
        # Footer con botones
        self.crear_footer(main_container)

    def crear_campos_cedula(self, parent):
        campos_base = [
            "DOCUMENTO", "APELLIDOS", "NOMBRES", 
            "SEXO", "FECHA DE NACIMIENTO", "RH"
        ]
        
        self.checkbox_vars = []
        for campo in campos_base:
            var = tk.BooleanVar()
            cb = ttk.Checkbutton(
                parent, 
                text=campo, 
                variable=var,
                style="TCheckbutton"
            )
            cb.pack(anchor="w", padx=10, pady=5)
            self.checkbox_vars.append((var, campo))

    def crear_campos_adicionales(self, parent):
        # Contenedor con scroll
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        ))
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 3 campos iniciales
        for _ in range(3):
            self.agregar_campo_adicional()

    def agregar_campo_adicional(self):
        frame_campo = ttk.Frame(self.scrollable_frame)
        frame_campo.pack(fill="x", pady=3)
        
        entry = ttk.Entry(frame_campo, font=self.estilo["fuente_principal"], width=38)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_eliminar = ttk.Button(
            frame_campo, 
            text="✖", 
            width=3,
            command=lambda f=frame_campo: f.destroy()
        )
        btn_eliminar.pack(side="right")

    def crear_footer(self, parent):
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill="x", pady=20)
        
        btn_agregar = ttk.Button(
            footer_frame,
            text="➕ Agregar Campo",
            command=self.agregar_campo_adicional,
            style="Accent.TButton"
        )
        btn_agregar.pack(side="left", padx=5)
        
        # Botón Guardar con texto negro
        btn_guardar = ttk.Button(
            footer_frame,
            text="💾 Guardar Configuración",
            command=self.validar_configuracion,
            style="Accent.TButton"
        )
        btn_guardar.pack(side="right")
        
        # Configurar estilo del botón
        style = ttk.Style()
        style.configure("Accent.TButton", 
                      font=("Arial", 12),
                      foreground="black", 
                      background="#4A90E2",
                      padding=10)
        
        style.map("Accent.TButton",
                background=[('active', '#fff'), ('!disabled', '#fff')])

    def validar_configuracion(self):
        self.evento_nombre = self.entry_evento.get().strip()
        self.campos_seleccionados = [campo for var, campo in self.checkbox_vars if var.get()]
        
        # Obtener campos adicionales
        self.columnas_adicionales = []
        for child in self.scrollable_frame.winfo_children():
            for widget in child.winfo_children():
                if isinstance(widget, ttk.Entry):
                    value = widget.get().strip().upper()
                    if value:
                        self.columnas_adicionales.append(value)

        # Validaciones
        if not self.evento_nombre:
            self.mostrar_error("Nombre del evento requerido", self.entry_evento)
            return
            
        if not self.campos_seleccionados:
            messagebox.showwarning("Validación", "Debe seleccionar al menos un campo de la cédula")
            return
            
        self.guardar_configuracion()

    def mostrar_error(self, mensaje, widget):
        messagebox.showwarning("Validación", mensaje)
        widget.focus_set()
        widget.config(foreground="red")

    def guardar_configuracion(self):
        self.path_database = f"bd_evento_{self.evento_nombre}.xlsx"
        self.ventana_config.destroy()
        ConfiguradorPlantilla(self.campos_seleccionados, self.columnas_adicionales, self.path_database, self.evento_nombre)

if __name__ == "__main__":
    app = ConfiguradorEvento()