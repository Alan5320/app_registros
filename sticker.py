from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Cm
from tkinter import messagebox

def crear_documento(datos_json):
    try:
        from pathlib import Path
        ruta_base = Path(__file__).parent if "__file__" in locals() else Path.cwd()
        ruta_plantilla = ruta_base / "_internal/plantilla_fija.docx"
        ruta_salida_word = ruta_base / "sticker_a_imprimir.docx"

        doc = Document(ruta_plantilla)

        # Verificar cuántos datos hay en el JSON
        cantidad_datos = len(datos_json)

        # Ajustar el margen superior dinámicamente según la cantidad de datos
        if cantidad_datos >= 5:
            for section in doc.sections:
                section.top_margin = Cm(0.4)  # Margen superior reducido a 0.5 cm
        elif ((cantidad_datos == 3) | (cantidad_datos == 4)):
            for section in doc.sections:
                section.top_margin = Cm(1)  # Margen superior aumentado a 1 cm
        elif cantidad_datos == 2:
            for section in doc.sections:
                section.top_margin = Cm(1.4)  # Margen superior aumentado a 1.4 cm
        elif cantidad_datos == 1:
            for section in doc.sections:
                section.top_margin = Cm(2)  # Margen superior aumentado a 1.4 cm
        else:
            # Mantener los márgenes existentes (de la plantilla)
            pass

        # Determinar tamaño de la página
        ancho_pagina = doc.sections[0].page_width
        alto_pagina = doc.sections[0].page_height

        # Ajustar el tamaño base de la fuente según el tamaño del documento
        if ancho_pagina > Cm(43.18) and alto_pagina > Cm(27.94):  # * Sticker Tamaño muy grande (póster o A3) * #
            tamaño_base_nombres = Pt(36)
            tamaño_base_general = Pt(28)
        elif ancho_pagina > Cm(30.48) or alto_pagina > Cm(21.59):  # * Sticker Tamaño grande (carta o A4) # *
            tamaño_base_nombres = Pt(30)
            tamaño_base_general = Pt(20)
        elif ancho_pagina > Cm(21.59) and alto_pagina > Cm(13.97):  # * Sticker Tamaño mediano (A5) * #
            tamaño_base_nombres = Pt(20)
            tamaño_base_general = Pt(16)
        elif ancho_pagina > Cm(9) and alto_pagina > Cm(9): # * Sticker 10x10 * #
            tamaño_base_nombres = Pt(16)
            tamaño_base_general = Pt(14)
        elif ancho_pagina > Cm(9) and alto_pagina > Cm(4): # * Sticker 10x5 * #
            tamaño_base_nombres = Pt(15)
            tamaño_base_general = Pt(11)

        # Ajustar tamaños dinámicamente según la cantidad de datos
        if cantidad_datos == 1:
            tamaño_fuente_nombres = tamaño_base_nombres + Pt(8)
            tamaño_fuente_general = tamaño_base_general + Pt(6)
        elif cantidad_datos == 2:
            tamaño_fuente_nombres = tamaño_base_nombres + Pt(6)
            tamaño_fuente_general = tamaño_base_general + Pt(4)        
        elif cantidad_datos == 3:
            tamaño_fuente_nombres = tamaño_base_nombres + Pt(4)
            tamaño_fuente_general = tamaño_base_general + Pt(2)
        elif cantidad_datos == 4:
            tamaño_fuente_nombres = tamaño_base_nombres + Pt(2)
            tamaño_fuente_general = tamaño_base_general + Pt(1)
        else:
            tamaño_fuente_nombres = tamaño_base_nombres
            tamaño_fuente_general = tamaño_base_general

        # Eliminar todos los párrafos vacíos existentes
        for paragraph in doc.paragraphs:
            if not paragraph.text.strip():
                p = paragraph._element
                p.getparent().remove(p)
                p._p = p._element = None

        # Verificar si "NOMBRES" está en el JSON
        nombres = datos_json.pop("NOMBRES", None)

        if nombres:
            # Crear un párrafo para "NOMBRES" con tamaño mayor y centrado
            parrafo_nombres = doc.add_paragraph()
            parrafo_nombres.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_nombres = parrafo_nombres.add_run(nombres)
            run_nombres.font.size = tamaño_fuente_nombres
            run_nombres.font.bold = True

        # Crear un párrafo para los demás datos
        parrafo = doc.add_paragraph()
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Configuración de espaciado
        parrafo_format = parrafo.paragraph_format
        parrafo_format.space_before = Pt(0)
        parrafo_format.space_after = Pt(0)

        # Iterar sobre los datos restantes del JSON y agregarlos al documento
        for clave, valor in datos_json.items():
            run_titulo = parrafo.add_run(f"{clave}: ")
            run_titulo.bold = True
            run_titulo.font.size = tamaño_fuente_general
            parrafo.add_run(f"{valor}\n").font.size = tamaño_fuente_general

        # Eliminar cualquier párrafo vacío final
        while doc.paragraphs and not doc.paragraphs[-1].text.strip():
            p = doc.paragraphs[-1]._element
            p.getparent().remove(p)
            p._p = p._element = None

        # Guardar el documento generado
        doc.save(ruta_salida_word)

        import print
        print.imprimir_documento(f"{ruta_salida_word}")

    except Exception as e:
        messagebox.showerror(
            "Error al crear el documento",
            "INTENTA CERRAR EL DOCUMENTO DE WORD O BORRALO Y VUELVELO A INTENTAR"
        )
        return False

