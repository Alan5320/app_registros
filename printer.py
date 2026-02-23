import win32com.client
from tkinter  import messagebox
def imprimir_documento(ruta_archivo):
    try:
        # Crear una instancia de la aplicación de Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # No mostrar la ventana de Word

        # Abrir el documento
        documento = word.Documents.Open(ruta_archivo)

        # Imprimir en la impresora predeterminada
        documento.PrintOut()

        # Cerrar el documento sin guardar
        documento.Close(SaveChanges=False)
        word.Quit()
    except Exception as e:
        messagebox.showerror("Error", f"¡Ups! Hubo un error al intentar imprimir: {e}")
