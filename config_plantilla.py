import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from dynamic_window import DynamicWindow

class ConfiguradorPlantilla:
    def __init__(self, campos_seleccionados, columnas_adicionales, path_database, evento_nombre):
        self.evento_nombre = evento_nombre
        self.path_database = path_database
        self.campos_seleccionados = campos_seleccionados
        self.columnas_adicionales = columnas_adicionales
        self.campos_totales = campos_seleccionados + columnas_adicionales

        # Configuración de estilos
        self.estilo = {
            "fuente_principal": ("Segoe UI", 12),
            "color_fondo": "#F0F2F5",
            "color_botones": "#4A90E2",
            "color_texto": "#2D3436",
            "padding_general": 20,
            "ancho_ventana": 800,  # Aumentado para 3 columnas
            "alto_ventana": 520    # Un poco más alto para incluir el selector de títulos
        }

        self.inicializar_ventana()

    def inicializar_ventana(self):
        self.ventana = tk.Tk()
        self.ventana.title("Datos plantilla - ServiEventos Col")
        self.ventana.configure(bg=self.estilo["color_fondo"])
        self.ventana.minsize(self.estilo["ancho_ventana"], self.estilo["alto_ventana"])

        # ✅ Crear variables Tk DESPUÉS de tener root
        self.imprimir_con_titulos_var = tk.BooleanVar(master=self.ventana, value=True)

        try:
            self.ventana.iconbitmap("_internal/logo2.ico")
        except:
            pass

        self.cargar_logo()
        self.crear_widgets()
        self.ventana.mainloop()

    def cargar_logo(self):
        try:
            logo = Image.open("_internal/logo.png")
            logo = logo.resize((200, 50))
            self.logo_tk = ImageTk.PhotoImage(logo)
            label_logo = ttk.Label(self.ventana, image=self.logo_tk)
            label_logo.pack(pady=10)
        except FileNotFoundError:
            pass

    def crear_widgets(self):
        main_frame = ttk.Frame(self.ventana)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Cabecera
        ttk.Label(
            main_frame,
            text="Selecciona los campos a imprimir en el sticker",
            font=("Segoe UI", 12, "bold"),
            anchor="center"
        ).pack(fill="x", pady=10)

        # ✅ Selector títulos
        titulos_frame = ttk.Frame(main_frame)
        titulos_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            titulos_frame,
            text="¿Desea imprimir con títulos en el sticker? Ejemplo: 'DOCUMENTO: 1234567890'",
            font=("Segoe UI", 12)
        ).pack(anchor="w")

        ttk.Checkbutton(
            titulos_frame,
            text="Sí, incluir títulos",
            variable=self.imprimir_con_titulos_var,
        ).pack(anchor="w", pady=5)

        # Contenedor principal de campos
        campos_frame = ttk.Frame(main_frame)
        campos_frame.pack(fill="both", expand=True)

        # Dividir en 3 columnas
        total_campos = len(self.campos_totales)
        por_columna = max(6, total_campos // 3 + (1 if total_campos % 3 != 0 else 0))

        self.plantilla_vars = []

        for col in range(3):
            frame_columna = ttk.Frame(campos_frame)
            frame_columna.grid(row=0, column=col, sticky="nsew", padx=10)
            campos_frame.grid_columnconfigure(col, weight=1)

            inicio = col * por_columna
            fin = inicio + por_columna

            for campo in self.campos_totales[inicio:fin]:
                # ✅ Asignar master para evitar "Too early..."
                var = tk.BooleanVar(master=self.ventana, value=False)

                cb = ttk.Checkbutton(
                    frame_columna,
                    text=campo,
                    variable=var,
                    style="TCheckbutton"
                )
                cb.pack(anchor="w", pady=3, fill="x")

                self.plantilla_vars.append((var, campo))

        # Footer con botón
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill="x", pady=15)

        self.crear_boton_guardar(footer_frame)

    def crear_boton_guardar(self, parent):
        style = ttk.Style()

        style.configure(
            "Accent.TButton",
            font=("Arial", 12),
            foreground="black",
            background="#4A90E2",
            padding=8
        )

        # Nota: ttk/Style puede ignorar background en algunos temas.
        style.map(
            "Accent.TButton",
            background=[("active", "#357ABD"), ("!disabled", "#4A90E2")]
        )

        ttk.Button(
            parent,
            text="💾 Guardar Selección",
            command=self.validar_seleccion,
            style="Accent.TButton",
            width=18
        ).pack(side="bottom", pady=5)

    def validar_seleccion(self):
        campos_plantilla = [campo for var, campo in self.plantilla_vars if var.get()]

        if not campos_plantilla:
            messagebox.showwarning("Validación", "Debe seleccionar al menos un campo para la plantilla")
            return

        imprimir_con_titulos = self.imprimir_con_titulos_var.get()

        self.ventana.destroy()
        DynamicWindow(
            self.campos_seleccionados,
            self.columnas_adicionales,
            campos_plantilla,
            self.path_database,
            self.evento_nombre,
            imprimir_con_titulos  # ✅ nuevo parámetro
        )
