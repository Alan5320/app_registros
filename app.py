import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

# Variable global para almacenar el nombre del evento
evento_nombre = ""
campos_seleccionados = ""
columnas_adicionales = ""

import tkinter as tk
from tkinter import messagebox

def configurar_evento():
    def guardar_configuracion():
        global evento_nombre

        # Obtener el nombre del evento
        evento_nombre = entry_evento.get().strip()
        if not evento_nombre:
            messagebox.showwarning("Advertencia", "Debe proporcionar un nombre para el evento.")
            return

        # Generar ruta dinámica para la base de datos
        global path_database
        path_database = f"bd_evento_{evento_nombre}.xlsx"

        # Obtener las selecciones del usuario
        global campos_seleccionados
        global columnas_adicionales
        campos_seleccionados = [var.get() for var in checkbox_vars if var.get()]
        columnas_adicionales = [entry.get().upper() for entry in entradas_campos_adicionales if entry.get().strip()]

        if not campos_seleccionados:
            messagebox.showwarning("Advertencia", "Debe seleccionar al menos un campo de la cédula.")
            return

        # Guardar la configuración para usarla en la siguiente ventana
        ventana_config.destroy()
        configurar_plantilla(campos_seleccionados, columnas_adicionales)

    # Crear ventana inicial para configuración
    ventana_config = tk.Tk()
    ventana_config.title("Registros - ServiEventos Col")
    ventana_config.configure(bg="#f5f5f5")

    ventana_config.iconbitmap("_internal/logo2.ico")
    ventana_config.resizable(False, False)

    from PIL import Image, ImageTk
    # Cargar el logo
    logo = Image.open("_internal/logo.png")
    logo = logo.resize((160, 40))  # Redimensionar la imagen si es necesario
    logo_tk = ImageTk.PhotoImage(logo)

    # Crear un Label para mostrar el logo
    label_logo = tk.Label(ventana_config, image=logo_tk, bg="#f5f5f5")
    label_logo.pack(pady=5)

    estilo_label = {"font": ("Arial", 12, "bold"), "bg": "#f5f5f5", "fg": "#333333"}
    estilo_boton = {"font": ("Arial", 12, "bold"), "bg": "#4CAF50", "fg": "white", "relief": "flat", "activebackground": "#45a049"}

    tk.Label(ventana_config, text="Nombre del Evento", **estilo_label).pack(pady=5)
    entry_evento = tk.Entry(ventana_config, font=("Arial", 12), relief="solid", highlightbackground="#cccccc", highlightthickness=1)
    entry_evento.pack(pady=5, padx=100, fill="x")

    tk.Label(ventana_config, text="Selecciona los campos de la cédula a guardar", **estilo_label).pack(pady=20, padx=150)

    # Crear un Frame con scroll
    frame_canvas = tk.Frame(ventana_config, bg="#f5f5f5")
    frame_canvas.pack(fill="both", expand=False, pady=5)

    canvas = tk.Canvas(frame_canvas, bg="#f5f5f5", highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)

    # Crear un Frame dentro del Canvas para los elementos
    frame_checkboxes = tk.Frame(canvas, bg="#f5f5f5")

    # Configurar el frame dentro del canvas
    canvas.create_window((0, 0), window=frame_checkboxes, anchor="center")

    # Ajustar el tamaño del canvas al contenido
    def ajustar_scroll(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        

    frame_checkboxes.bind("<Configure>", ajustar_scroll)

    # Campos base escaneados de la cédula
    campos_base = [
        "DOCUMENTO", "APELLIDOS", "NOMBRES", "SEXO", "FECHA DE NACIMIENTO", "RH"
    ]

    # Crear checkboxes
    checkbox_vars = []
    for campo in campos_base:
        var = tk.StringVar()
        checkbox = tk.Checkbutton(frame_checkboxes, text=campo, variable=var, onvalue=campo, offvalue="", font=("Arial", 11), bg="#f5f5f5", fg="#333333", activebackground="#4CAF50")
        checkbox.pack(anchor="w", padx=250, pady=2)
        checkbox_vars.append(var)

    # Entrada para agregar campos adicionales
    tk.Label(ventana_config, text="Agrega columnas adicionales para este evento", **estilo_label).pack(pady=5)

    entradas_campos_adicionales = []

    def agregar_campo_adicional():
        entrada = tk.Entry(frame_checkboxes, font=("Arial", 12), relief="solid", highlightbackground="#cccccc", highlightthickness=1)
        entrada.pack(pady=5, padx=250, anchor="w")
        entradas_campos_adicionales.append(entrada)

    agregar_button = tk.Button(ventana_config, text="Agregar Campo", command=agregar_campo_adicional, **estilo_boton)
    agregar_button.pack(pady=5)

    guardar_button = tk.Button(ventana_config, text="Guardar Configuración", command=guardar_configuracion, **estilo_boton)
    guardar_button.pack(pady=20)

    ventana_config.mainloop()


# Función para configurar los datos para la plantilla de Word
def configurar_plantilla(campos_seleccionados, columnas_adicionales):
    def guardar_seleccion_plantilla():
        campos_plantilla = [var.get() for var in plantilla_vars if var.get()]

        if not campos_plantilla:
            messagebox.showwarning("Advertencia", "Debe seleccionar al menos un campo para la plantilla.")
            return

        ventana_plantilla.destroy()
        iniciar_ventana_dinamica(campos_seleccionados, columnas_adicionales, campos_plantilla)

    ventana_plantilla = tk.Tk()
    ventana_plantilla.title("Datos plantilla - ServiEventos Col")
    ventana_plantilla.iconbitmap("_internal/logo2.ico")
    ventana_plantilla.resizable(False, False)

    from PIL import Image, ImageTk
    logo = Image.open("_internal/logo.png")
    logo = logo.resize((160, 40)) 
    logo_tk = ImageTk.PhotoImage(logo)
    tk.Label(ventana_plantilla, image=logo_tk).pack(pady=10)

    tk.Label(ventana_plantilla, text="Selecciona los campos para la plantilla de Word", font=("Arial", 12, "bold")).pack(pady=10, padx=40)

    plantilla_vars = []
    campos_totales = campos_seleccionados + columnas_adicionales
    for campo in campos_totales:
        var = tk.StringVar()
        checkbox = tk.Checkbutton(ventana_plantilla, text=campo, variable=var, onvalue=campo, offvalue="", font=("Arial", 10), anchor="center", padx=10)
        checkbox.pack(fill="x", pady=2)
        plantilla_vars.append(var)

    guardar_button = tk.Button(ventana_plantilla, text="Guardar Selección", command=guardar_seleccion_plantilla, font=("Arial", 12), bg="#4CAF50", fg="white")
    guardar_button.pack(pady=20)

    ventana_plantilla.mainloop()

# Función para procesar los datos escaneados
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


def iniciar_ventana_dinamica(campos_cedula, columnas_adicionales, campos_plantilla):
    def guardar_datos():
        try:
            datos_cedula = procesar_datos_escaneados(input_text.get("1.0", tk.END))
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return

        # Filtrar los campos seleccionados
        registro = {campo: datos_cedula[campo] for campo in campos_cedula if campo in datos_cedula}

        # Agregar datos de las columnas adicionales
        for columna, entrada in zip(columnas_adicionales, entradas_adicionales):
            registro[columna] = entrada.get()

        # Leer o crear la base de datos
        try:
            df = pd.read_excel(path_database)
        except FileNotFoundError:
            df = pd.DataFrame(columns=campos_cedula + columnas_adicionales)

        # Agregar el nuevo registro
        df = pd.concat([df, pd.DataFrame([registro])], ignore_index=True)
        df.to_excel(path_database, index=False)

        messagebox.showinfo("Éxito", "Datos registrados correctamente. EL DOCUMENTO VA A IMPRIMIRSE.")

        # Mostrar solo los campos seleccionados para la plantilla
        datos_plantilla = {campo: registro[campo] for campo in campos_plantilla if campo in registro}
        print("Datos para la plantilla", datos_plantilla)

        import sticker
        sticker.crear_documento(datos_plantilla)

        # Limpiar entradas
        input_text.delete("1.0", tk.END)
        for entrada in entradas_adicionales:
            entrada.delete(0, tk.END)

    # Crear ventana dinámica
    ventana_dinamica = tk.Tk()
    ventana_dinamica.title(f"Registros - {evento_nombre}")
    ventana_dinamica.iconbitmap("_internal/logo2.ico")
    ventana_dinamica.resizable(False, False)

    from PIL import Image, ImageTk
    # Cargar el logo
    logo = Image.open("_internal/logo.png")
    logo = logo.resize((160, 40))  # Redimensionar la imagen si es necesario
    logo_tk = ImageTk.PhotoImage(logo)

    # Crear un Label para mostrar el logo
    label_logo = tk.Label(ventana_dinamica, image=logo_tk)
    label_logo.pack(pady=5)

    # Estilos generales
    label_font = ("Arial", 12, "bold")
    input_font = ("Arial", 12)
    button_font = ("Arial", 12, "bold")
    label_bg = "#f2f2f2"
    label_fg = "#333"
    button_bg = "#4caf50"
    button_fg = "#fff"

    tk.Label(ventana_dinamica, text="Escanea los datos de la cédula", font=label_font, bg=label_bg, fg=label_fg).pack(pady=10)

    input_text = tk.Text(ventana_dinamica, height=6, width=50, font=input_font, bg="#fff", fg="#000", bd=2, relief="solid")
    input_text.pack(pady=10, padx=20)

    tk.Label(ventana_dinamica, text="Verifica si hay campos adicionales y llenalos", font=label_font, bg=label_bg, fg=label_fg).pack(pady=10) 

    # Crear entradas dinámicas para los campos adicionales
    entradas_adicionales = []
    def on_focus_in(event, entry, placeholder_text):
        if entry.get() == placeholder_text:
            entry.delete(0, tk.END)
            entry.config(fg="black", font=("Arial", 10))

    def on_focus_out(event, entry, placeholder_text):
        if not entry.get():
            entry.insert(0, placeholder_text)
            entry.config(fg="black", font=("Arial", 10))

    for columna in columnas_adicionales:
        frame = tk.Frame(ventana_dinamica, bg=label_bg)
        frame.pack(pady=5, fill="x")

        # Placeholder personalizado
        placeholder_text = f"INGRESE EL CAMPO '{columna.upper()}'"

        entrada = tk.Entry(frame, font=input_font, justify="center", bg="#fff", fg="#000000", bd=2, relief="solid")
        entrada.insert(0, placeholder_text)

        # Eventos para manejar el placeholder
        entrada.bind("<FocusIn>", lambda event, e=entrada, p=placeholder_text: on_focus_in(event, e, p))
        entrada.bind("<FocusOut>", lambda event, e=entrada, p=placeholder_text: on_focus_out(event, e, p))

        entrada.pack(side="left", padx=5, fill="x", expand=True)

        entradas_adicionales.append(entrada)
    
    def cerrar_app():
        ventana_dinamica.quit()
        ventana_dinamica.destroy()
    
    def configurar_plantilla_button():
        ventana_dinamica.destroy()
        configurar_plantilla(campos_cedula, columnas_adicionales)

    guardar_button = tk.Button(ventana_dinamica, text="Guardar Datos", command=guardar_datos, font=button_font, bg=button_bg, fg=button_fg, padx=10, pady=5)
    guardar_button.pack(pady=10)
    configurar_button = tk.Button(ventana_dinamica, text="Configurar Plantilla", command=configurar_plantilla_button, font=button_font, bg=button_bg, fg=button_fg, padx=10, pady=5)
    configurar_button.pack(pady=0)
    cerrar_button = tk.Button(ventana_dinamica, text="Cerrar Evento", command=cerrar_app, font=button_font, bg=button_bg, fg=button_fg, padx=10, pady=5)
    cerrar_button.pack(pady=15)

    ventana_dinamica.mainloop()


# Iniciar el programa
configurar_evento()