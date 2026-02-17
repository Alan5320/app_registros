from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Cm
from tkinter import messagebox
import traceback
from pathlib import Path

def _remove_paragraph(paragraph):
    """Elimina un párrafo del documento (uso interno)."""
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def crear_documento(datos_json, imprimir_con_titulos):
    try:
        # Validación de datos de entrada
        if not datos_json or not isinstance(datos_json, dict):
            raise ValueError("El JSON de entrada no es válido o está vacío")

        # Trabajar con copia para NO mutar el dict original
        datos = dict(datos_json)

        # Configuración de rutas robusta
        try:
            ruta_base = Path(__file__).parent if "__file__" in locals() else Path.cwd()
            ruta_plantilla = (ruta_base / "_internal/plantilla_fija.docx").resolve(strict=True)
            ruta_salida_word = (ruta_base / "sticker_a_imprimir.docx").resolve()
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"No se encontró la plantilla: {e}")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Error configurando rutas: {e}")
            return False

        # Verificar si el archivo de salida está en uso (mejor enfoque para .docx)
        # Nota: .docx es binario; no usar encoding.
        if ruta_salida_word.exists():
            try:
                # En Windows, renombrar falla si está abierto en Word.
                tmp = ruta_salida_word.with_suffix(".tmp_check")
                ruta_salida_word.rename(tmp)
                tmp.rename(ruta_salida_word)
            except Exception:
                messagebox.showerror(
                    "Archivo en uso",
                    f"El archivo {ruta_salida_word.name} está abierto. Ciérralo y vuelve a intentar."
                )
                return False

        # Cargar plantilla
        doc = Document(ruta_plantilla)

        #  Evitar herencia rara: tamaño de fuente 0 a todo.
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(0)

        # Poner márgenes a 0 para evitar herencia de estilos no deseados
        for section in doc.sections:
            section.top_margin = Cm(0)
            section.bottom_margin = Cm(0)
            section.left_margin = Cm(0)
            section.right_margin = Cm(0)

        # Ajustar márgenes según cantidad de datos (usa copia "datos")
        cantidad_datos = len(datos)

        # Determinar tamaño de página
        seccion = doc.sections[0]
        ancho_pagina = seccion.page_width
        alto_pagina = seccion.page_height

        # * --- Ajuste dinámico de márgenes y fuentes según tamaño de página y cantidad de datos ---
        # ----------------------------- Tamaño (10x6) -----------------------------
        if ancho_pagina > Cm(10) and alto_pagina > Cm(6):

            # Ajuste dinámico de márgenes para 10x6 según cantidad de datos
            margen_superior = {
                1: Cm(1.8), 2: Cm(1.5), 3: Cm(1.5), 4: Cm(1.3),
            }.get(cantidad_datos, Cm(1.1)) # * Para +5 elementos

            # * --- Ajustar márgenes para 10x6 ---
            for section in doc.sections:
                section.top_margin = margen_superior
                section.bottom_margin = Cm(0.2)

            # Ajuste dinámico de tamaños de fuente para 10x6 según cantidad de datos
            ajustes_fuente = {
                1: (Pt(22), Pt(22)),
                2: (Pt(20), Pt(17)),
                3: (Pt(18), Pt(15)),
                4: (Pt(17), Pt(14)),
                5: (Pt(16), Pt(13)),
            }
            # * --- Ajustar tamaño de fuente para 10x6 ---
            if cantidad_datos >= 6:
                tamaño_fuente_nombres, tamaño_fuente_general = (Pt(15), Pt(11))
            else:
                tamaño_fuente_nombres, tamaño_fuente_general = ajustes_fuente[cantidad_datos]

        # ----------------------------- Tamaño (5x3) -----------------------------
        elif ancho_pagina > Cm(5) and alto_pagina > Cm(3):

            # Ajuste dinámico de márgenes para 5x3 según cantidad de datos
            margen_superior = {
                1: Cm(0.8), 2: Cm(0.7), 3: Cm(0.6), 4: Cm(0.5),
            }.get(cantidad_datos, Cm(0.4)) # * Para +5 elementos

            # * --- Ajustar márgenes para 5x3 ---
            for section in doc.sections:
                section.top_margin = margen_superior
                section.bottom_margin = Cm(0.1)

            # Ajuste dinámico de tamaños de fuente para 5x3 según cantidad de datos
            ajustes_fuente = {
                1: (Pt(14), Pt(12)),
                2: (Pt(13), Pt(11)),
                3: (Pt(12), Pt(10)),
                4: (Pt(11), Pt(9)),
                5: (Pt(10), Pt(8)),
            }

            # * --- Ajustar tamaño de fuente para 5x3 ---
            if cantidad_datos >= 6:
                tamaño_fuente_nombres, tamaño_fuente_general = (Pt(9), Pt(7))
            else:
                tamaño_fuente_nombres, tamaño_fuente_general = ajustes_fuente[cantidad_datos]

        # Limpiar párrafos vacíos existentes (sin pelear con el “último párrafo”)
        # Eliminamos vacíos, pero si el doc quedara sin párrafos, Word igual deja uno.
        for p in list(doc.paragraphs):
            if not p.text.strip():
                _remove_paragraph(p)

        # Procesar campo NOMBRES (sin mutar original)
        nombres = datos.pop("NOMBRES", None)
        if nombres:
            parrafo_nombres = doc.add_paragraph()
            parrafo_nombres.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            parrafo_nombres.paragraph_format.space_before = Pt(0)
            parrafo_nombres.paragraph_format.space_after = Pt(0)

            run_nombres = parrafo_nombres.add_run(str(nombres))
            run_nombres.font.size = tamaño_fuente_nombres
            run_nombres.font.bold = True

        # Agregar demás campos
        parrafo = doc.add_paragraph()
        parrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        formato_parrafo = parrafo.paragraph_format
        formato_parrafo.space_before = Pt(0)
        formato_parrafo.space_after = Pt(0)

        def agregar_linea(parrafo, clave, valor):
            if imprimir_con_titulos:
                run_clave = parrafo.add_run(f"{clave}: ")
                run_clave.bold = True
                run_clave.font.size = tamaño_fuente_general

                run_valor = parrafo.add_run(str(valor))
                run_valor.font.size = tamaño_fuente_general
            else:
                run_valor = parrafo.add_run(str(valor))
                run_valor.font.size = tamaño_fuente_general

            parrafo.add_run("\n")  # salto de línea controlado

        # Orden específico primero
        for clave in ["APELLIDOS", "DOCUMENTO"]:
            if clave in datos:
                valor = datos.pop(clave)
                agregar_linea(parrafo, clave, valor)

        # Luego el resto de campos
        for clave, valor in datos.items():
            agregar_linea(parrafo, clave, valor)

        # Limpieza final: si se quedó algún párrafo vacío, lo removemos
        # (pero si Word lo re-crea, al menos ya no hay \n extra generado por nosotros).
        for p in list(doc.paragraphs):
            if not p.text.strip() and len(doc.paragraphs) > 1:
                _remove_paragraph(p)

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
