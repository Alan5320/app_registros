# app_registros, proyecto "Registro y Validación de Documentos"

# Descripción
Este programa permite registrar documentos, validar si un documento ya está registrado en una base de datos, y generar un sticker en formato Word con los datos registrados. Incluye funcionalidades para escanear documentos, registrar automáticamente los datos esenciales en un archivo Excel, generar documentos a partir de una plantilla predefinida y enviarlos a impresión en una impresora predeterminada.

# Requisitos Previos
* Plantilla de Word Fija:
Debe contar con un archivo llamado plantilla_fija.docx. Este archivo será utilizado como base para generar el sticker con los datos escaneados. Coloque este archivo en la misma carpeta donde se ejecutará el programa.

* Definir una Impresora Predeterminada:
Asegúrese de tener una impresora predeterminada configurada en el sistema, ya que el programa enviará automáticamente el sticker a imprimir.

* Archivos Excel:
documentos_registrados.xlsx: Este archivo almacena todos los documentos registrados. Si no existe, el programa lo creará automáticamente.
documentos_a_validar.xlsx: Este archivo debe contener una columna Documento con los documentos existentes para validación.

* Datos Mínimos Requeridos: 
Cada documento debe contener al menos 5 datos esenciales para su registro y procesamiento:
Documento (número único)
Apellidos
Sexo (M o F)
Fecha de nacimiento (formato AAAAMMDD)
Al menos un nombre

# Instalación
* Requisitos de Python:
Instale las dependencias necesarias ejecutando el siguiente comando:

Copiar código

pip install pandas openpyxl python-docx pywin32

* Descargue el Código:
Descargue el archivo Python del programa y colóquelo en la misma carpeta que los archivos mencionados anteriormente (plantilla_fija.docx, documentos_registrados.xlsx, documentos_a_validar.xlsx).

* Ejecución:
Ejecute el programa con el siguiente comando:

Copiar código

python <nombre_del_archivo>.py


# Uso
* Registrar un Documento:
Ingrese los datos completos del documento en el cuadro de texto siguiendo el formato:

Documento, Apellido1, Apellido2, Nombre1, Sexo (M/F), Fecha de Nacimiento (AAAAMMDD), 
POR LO GENERAL AL ESCANEAR UN DOCUMENTO CON UN ESCANER ESTOS DATOS SE COLOCAN AUTOMATICAMENTE EN EL CUADRO DE TEXTO,
Haga clic en Enviar Documento para registrar y procesar los datos. Esto generará un archivo Word con los datos y lo enviará a la impresora.

* Validar un Documento:
Ingrese los mismos datos anteioresa este paso en el cuadro de texto.
Haga clic en Validar Documento para verificar si está registrado en documentos_a_validar.xlsx.

* Borrar Entrada:
Haga clic en Borrar para limpiar el cuadro de texto.

* Cerrar Aplicación:
Haga clic en Cerrar para salir del programa.

# Funcionalidades
* Registro de Documento:
Los datos ingresados se procesan y se almacenan en el archivo documentos_registrados.xlsx. El programa valida automáticamente:
Formato correcto de la fecha de nacimiento.
Presencia del campo de sexo.
Suficiencia de datos (mínimo 5 campos).

* Generación de Documento Word:
Utiliza los datos más recientes para rellenar una plantilla de Word (plantilla_fija.docx) y genera un archivo sticker_a_imprimir.docx.

* Impresión Automática:
El sticker generado se envía automáticamente a la impresora predeterminada del sistema.

* Validación de Documento:
Comprueba si un documento está registrado en la base de datos documentos_a_validar.xlsx.

# Solución de Problemas
* El documento Word no se genera o no se imprime:
Asegúrese de que plantilla_fija.docx no esté abierto durante la ejecución del programa. Si el problema persiste, cierre manualmente cualquier instancia de Microsoft Word.

* El archivo Excel no se encuentra:
Verifique que los archivos documentos_registrados.xlsx y documentos_a_validar.xlsx estén en la carpeta de ejecución. De no existir, cree un archivo vacío con las columnas correspondientes.

* Error con la impresora:
Verifique que la impresora esté correctamente configurada como predeterminada en su sistema operativo.

Este programa es una herramienta flexible para gestionar documentos escaneados, adecuada para aplicaciones en oficinas, empresas o proyectos que requieren registro y validación de datos personales.
