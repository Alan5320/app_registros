from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Cm
from tkinter import messagebox
import traceback
from pathlib import Path

def crear_documento(datos_json):
    try:
        # Validación de datos de entrada
        if not datos_json or not isinstance(datos_json, dict):
            raise ValueError("El JSON de entrada no es válido o está vacío")

        # Configuración de rutas robusta
        try:
            ruta_base = Path(__file__).parent if "__file__" in locals() else Path.cwd()
            ruta_plantilla = ruta_base / "_internal/plantilla_fija.docx"
            ruta_salida_word = ruta_base / "sticker_a_imprimir.docx"
            
            # Convertir a rutas absolutas y verificar existencia de plantilla
            ruta_plantilla = ruta_plantilla.resolve(strict=True)
            ruta_salida_word = ruta_salida_word.resolve()
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"No se encontró la plantilla: {e}")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Error configurando rutas: {e}")
            return False

        # Verificar si el archivo de salida está en uso
        if ruta_salida_word.exists():
            try:
                # Intentar abrir el archivo en modo append para verificar bloqueo
                with open(ruta_salida_word, 'a', encoding='utf-8'):
                    pass
            except IOError:
                messagebox.showerror(
                    "Archivo en uso",
                    f"El archivo {ruta_salida_word.name} está abierto. Ciérralo y vuelve a intentar."
                )
                return False

        # Cargar plantilla
        doc = Document(ruta_plantilla)

        # Ajustar márgenes según cantidad de datos
        cantidad_datos = len(datos_json)
        margen_superior = {
            1: Cm(2),
            2: Cm(1.4),
            3: Cm(1),
            4: Cm(1),
        }.get(cantidad_datos, Cm(0.4))  # Default para 5+ elementos

        for section in doc.sections:
            section.top_margin = margen_superior

        # Determinar tamaño de página
        seccion = doc.sections[0]
        ancho_pagina = seccion.page_width
        alto_pagina = seccion.page_height

        # Configurar tamaños base de fuente según dimensiones del documento
        if ancho_pagina > Cm(43.18) and alto_pagina > Cm(27.94):  # Póster/A3
            tamaños = (Pt(36), Pt(28))
        elif ancho_pagina > Cm(30.48) or alto_pagina > Cm(21.59):  # Carta/A4
            tamaños = (Pt(30), Pt(20))
        elif ancho_pagina > Cm(21.59) and alto_pagina > Cm(13.97):  # A5
            tamaños = (Pt(20), Pt(16))
        elif ancho_pagina > Cm(9) and alto_pagina > Cm(9):  # 10x10
            tamaños = (Pt(16), Pt(14))
        else:  # 10x5
            tamaños = (Pt(15), Pt(11))
        
        tamaño_base_nombres, tamaño_base_general = tamaños

        # Ajuste dinámico de tamaños de fuente
        ajustes_fuente = {
            1: (Pt(8), Pt(6)),
            2: (Pt(6), Pt(4)),
            3: (Pt(4), Pt(2)),
            4: (Pt(2), Pt(1))
        }
        ajuste = ajustes_fuente.get(cantidad_datos, (Pt(0), Pt(0)))
        tamaño_fuente_nombres = tamaño_base_nombres + ajuste[0]
        tamaño_fuente_general = tamaño_base_general + ajuste[1]

        # Limpiar párrafos vacíos existentes
        for paragraph in list(doc.paragraphs):  # Usar list() para copia segura
            if not paragraph.text.strip():
                p_elemento = paragraph._element
                p_elemento.getparent().remove(p_elemento)

        # Procesar campo NOMBRES
        nombres = datos_json.pop("NOMBRES", None)
        if nombres:
            parrafo_nombres = doc.add_paragraph()
            parrafo_nombres.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_nombres = parrafo_nombres.add_run(nombres)
            run_nombres.font.size = tamaño_fuente_nombres
            run_nombres.font.bold = True

        # Agregar demás campos
        parrafo = doc.add_paragraph()
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        formato_parrafo = parrafo.paragraph_format
        formato_parrafo.space_before = Pt(0)
        formato_parrafo.space_after = Pt(0)

        for clave, valor in datos_json.items():
            # Texto en negrita para las claves
            run_clave = parrafo.add_run(f"{clave}: ")
            run_clave.bold = True
            run_clave.font.size = tamaño_fuente_general
            
            # Valor normal
            run_valor = parrafo.add_run(f"{valor}\n")
            run_valor.font.size = tamaño_fuente_general

        # Eliminar párrafos vacíos finales
        while doc.paragraphs and not doc.paragraphs[-1].text.strip():
            ultimo_parrafo = doc.paragraphs[-1]._element
            ultimo_parrafo.getparent().remove(ultimo_parrafo)

        # Guardar documento con manejo de errores
        try:
            doc.save(ruta_salida_word)
        except PermissionError:
            messagebox.showerror(
                "Error de permisos",
                f"No se pudo guardar el documento. Asegúrate de tener permisos de escritura en:\n{ruta_salida_word}"
            )
            return False
        return True

    except Exception as e:
        traceback.print_exc() 
        messagebox.showerror(
            "Error crítico",
            f"Error inesperado: {str(e)}\nDetalles en la consola."
        )
        return False