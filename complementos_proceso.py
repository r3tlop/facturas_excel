from tkinter import messagebox
from openpyxl.styles import Border, Side

class mensajes:
    @staticmethod
    def registro_guardado(fila):
         messagebox.showinfo(
             "Registro guardado",
             f"El registro de la fila {fila} se guardo correctamente"
         )
    @staticmethod
    def registro_no_seleccionado():
         messagebox.showwarning(
             "Registro no seleccionado",
             f"Primero selecciona un registro"
         )
    @staticmethod
    def campos_busqueda_vacios():
         messagebox.showwarning(
             "Campos vacios",
             f"Ingresa algun dato para realizar la busqueda"
         )
    @staticmethod
    def archivo_no_encontrado():
         messagebox.showerror(
             "Archivo no encontrado",
             f"no se pudo encontrar el archivo de Excel"
         )
    @staticmethod
    def archivo_en_uso():
         messagebox.showerror(
             "Excel en uso",
             "No se pudo acceder al archivo\n"
             "Comprueba que el archivo de Excel no este abierto"
         )
    @staticmethod
    def hoja_no_encontrada(nombre_hoja):
         messagebox.showerror(
             "Hoja no encontrada",
             f"No se encontro la hoja '{nombre_hoja}'"
         )
    @staticmethod
    def error_inesperado(error):
         messagebox.showerror(
             "Error inesperado",
             f"Ocurrio un error inesperado\n"
             f"{type(error).__name__}: {error}"
         )
    @staticmethod
    def registro_actualizado(fila):
         messagebox.showinfo(
             "Registro actualizado",
             f"El registro de la fila '{fila}' se actualizo correctamente"
         )
    @staticmethod
    def confirmar_actualizacion(fila):
         return messagebox.askyesno(
              "Confirmar actualizacion",
              f"¿Estas seguro que deseas actualizar la fila '{fila}'?"
         )
    @staticmethod
    def error_comprobar_guardado():
         return messagebox.showerror(
              "Error al comprobar",
              "El registro fue guardado, pero los datos encontrados "
              "en Excel no coinciden con los datos ingresados."
    )
    @staticmethod
    def error_comprobar_actualizacion():
        messagebox.showerror(
            "Error al comprobar",
            "El registro fue actualizado, pero los datos encontrados "
            "en Excel no coinciden con los datos actualizados."
    )
    @staticmethod
    def archivos_factura_no_encontrados(folio):
        return messagebox.showwarning(
             "Archivos encontrados",
             f"No se encontraron archivos para el folio: \n\n{folio}"
        )
    @staticmethod
    def error_al_abrir_carpeta(error):
        return messagebox.showerror(
            "Error al abrir carpeta",
            "No se pudo abrir la carpeta del folio.\n\n"
            f"{type(error).__name__}: {error}"
        )
    @staticmethod
    def error_al_abrir_factura(error):
         messagebox.showerror(
              "Error al abrir factura",
              f"No se pudo abrir el achivo.\n\n"
              f"{type(error).__name__}: {error}"
         )
    @staticmethod
    def factura_no_seleccionada():
         messagebox.showwarning(
              "Archivo no seleccionado",
              "Selecciona uno o mas archivos para abrir"
         )
    @staticmethod
    def confirmar_abrir_factura(nombre_archivo):
        return messagebox.askyesno(
              "Abrir Factura",
              f"¿Desdeas abrir el siguiente archivo?\n\n{nombre_archivo}"
        )
    @staticmethod
    def campos_obligatorios_vacios(campos):
         messagebox.showwarning(
              "Campos obligatorios",
              "Faltan los siguientes campos obligatorios:\n\n"
              + "\n".join(campos)
         )
    @staticmethod
    def folio_repetido_guardado(folio, hoja, fila):
         messagebox.showwarning(
              "Folio repetido",
              f"El folio {folio} ya estaba registrado.\n\n"
              f"Se guardo el nuevo registro en:\n"
              f"HOJA: {hoja}\n"
              f"FILA: {fila}"
         )
    @staticmethod
    def agregar_complemento():
        return messagebox.askyesno(
            "Agregar complemento",
            "La factura se guardo correctamente.\n\n"
            "¿Deseas agregar un comeplemnto de pago?"
         )
    @staticmethod
    def guardar_complemento():
         messagebox.showinfo(
              "Complemento de pago",
              "El boton de guardar funciona.\n"
              "El guardado en Excel esta pendiente"
         )
    @staticmethod
    def mensaje_error(error):
        messagebox.showerror(
            "Error al abrir complemento",
            f"{type(error).__name__}: {error}"
        )


def columnas_busqueda():
    columnas = [
        "Archivo",
        "Hoja",
        "Fila",
        "Fecha",
        "Numero de Cotizacion",
        "Numero Orden de Compra",
        "Cliente",
        "Folio Factura",
        "Concepto",
        "Monto Facturado",
        "Estatus",
        "Complemento 1",
        "Complemento 2"
        ]
    return columnas
    
def estado_factura():
    return [
        "PENDIENTE",
        "FACTURADO",
        "PAGADA",
        "CANCELADA"
    ]
    
def busqueda_valores(datos_busqueda):
    return datos_busqueda
 
def nombres_columnas():
    cols = [
        "Fecha",
        "Numero de Cotizacion",
        "Numero Orden de Compra",
        "Cliente",
        "Folio Factura",
        "Concepto",
        "Monto Facturado",
        "Estatus",
        "Complemento 1",
        "Complemento 2"
    ]
    return cols

def mapa_columnas():
    map_cols = {
        "fecha" : 0,
        "cotizacion" : 1,
        "orden_cotizacion" : 2,
        "cliente" : 3,
        "folio_factura" : 4,
        "concepto" : 5,
        "monto_facturado" : 6,
        "estatus" : 7,
        "complemento1" : 8,
        "complemento2" : 9
    }
    return map_cols

def encabezados_excel():
    return [
        "FECHA",
        "NO. DE COT.",
        "NO. DE O.C.",
        "CLIENTE",
        "FOLIO FACTURA",
        "CONCEPTO",
        "MONTO FACTURADO",
        "ESTATUS",
        "COMP",
        "COMP2"
            ]

def anchos_columnas():
     return {
          "A": 13,
          "B": 16,
          "C": 18,
          "D": 25,
          "E": 18, 
          "F": 45,
          "G": 18,
          "H": 15,
          "I": 20,
          "J": 20,
     }

def borde_celdas():
    return Border(
        left = Side(style = "thin"),
        right = Side(style = "thin"),
        top = Side(style = "thin"),
        bottom = Side(style = "thin")
        )

def colores_estatus():
     return {
          "PENDIENTE" : ("#F5D79A", "#422006"),
          "FACTURADO" : ("#B9D9F3", "#17324D"),
          "PAGADA" : ("#BCE5C8", "#14532D"),
          "CANCELADA" : ("#F2B6BC", "#7F1D1D")
     }

def meses_excel():
     return {
            1: "ENERO",
            2: "FEBRERO",
            3: "MARZO",
            4: "ABRIL",
            5: "MAYO",
            6: "JUNIO",
            7: "JULIO",
            8: "AGOSTO",
            9: "SEPTIEMBRE",
            10: "OCTUBRE",
            11: "NOVIEMBRE",
            12: "DICIEMBRE"
        }

def formatos_permitidos():
     return [
        ".pdf",
        ".png",
        ".jpg",
        ".jpeg",
        ".webp",
        ".tif",
        ".tiff",
        ".bmp"
     ]

def columnas_info_facturas():
    return {
    "Folio": ("FOLIO", 100),
    "Creacion": ("FECHA DE CREACION", 175),
    "Incorporacion": ("FECHA DE INCORPORACION", 190),
    "Archivo": ("NOMBRE DEL ARCHIVO", 240)
}

def secciones_manual(): 
     return [
          "Introduccion",
          "Registrar Factura",
          "Buscar Registros",
          "Actualizar Factura",
          "Archivos de Facturas",
          "Recomendaciones",
          "Carpetas del Programa",
          "Agregar Excel Historicos",
          "Folios y Documentos",
          "Registros e Incorporaciones",
          "Acerca del Programa"
     ]

def obtener_contenido_manual():
    return {
          "Introduccion": """
BIENVENIDO AL SISTEMA DE FACTURAS    
    El Sistema  permite registrar, buscar y actualizar informacion 
    relacionada 
    con las facturas mediante archivos de Excel.
    Desde la ventana principal puedes capturar los datos de una factura, 
    consultar registros existentes y actualizar la informacion de un registro 
    seleccionado.
    El sistema tambien permite consultar documentos relacionados con cada 
    folio de factura y organizar los archivos encontrados.

RECOMENDACIONES INICIALES
    ▸ Verifica los datos antes de guardar.
    ▸ Utiliza los folios correctos para relacionar documnetos.
    ▸ Comprueba el estatus de la factura antes de actualizar.
    ▸ No modifiques manualmente los archivos de Excel mientras
      el sistema esta trabajando con ellos.
""",
          "Registrar Factura": """
¿COMO AGREGO UN REGISTRO?
    1. Abre una ventana principal del programa.
    2. Revisa la fecha, que se establece automaticamente.
    3. Captura el numero de Cotizacion.
    4. Ingresa el numero de Orden de Compra.
    5. Escribe el nombre del Cliente.
    6. Ingresa el Folio de Factura.
    7. Captrua el Monto Facturado.
    8. Selecciona el Estatus correspondiente.
    9. Completa el Concepto y los si es que le corresponde alguno.
    10. Pulsa AGREGAR

CAMPOS OBLIGATORIOS
    Los campos requeridos son:
        ▸ Numero de Cotizacion
        ▸ Numero Orden de Compra
        ▸ Cliente
        ▸ Folio Factura
        ▸ Monto Facturado
        ▸ Estatus
    El Concepto y los Complementos pueden quedar vacios.

FOLIOS REPETIDOS (IMPORTANTE)
    Si el folio ya esta registrado, el sistema utiliza una hoja especial 
    para guardar el registro repetido y conservarlo separado de los 
    registros normales.
""",
          "Buscar Registros": """
¿COMO BUSCO UN REGISTRO?
    1. Escribe uno o varios datos en los campos de ventana 
       principal.
    2. Puedes utilizar datos como Cotizacion, Orden de Compra,
       Cliente, Folio, Concepto, Monto, Estatus o Complementos.
    3. Pulsa BUSCAR
    4. Revisa los resultados que aparecen en la ventana secundaria.

CONSEJOS
    ▸ No es necesario llenar todos los campos.
    ▸ Ingresa al menos un dato para realizar la busqueda. 
    ▸ Revisa el cliente y el folio para identificar el registro correcto.
    ▸ Si aparecen varios resultados, selecciona el que corresponda antes de 
      realizar otra operacion.
    ▸ En la ventana de resultados puedes seleccionar un registro, abrir su fila en 
      Excel o consultar los archivos relacionados.
""",
          "Actualizar Factura": """
¿COMO ACTUALIZO UN REGISTRO?
    1. Busca la factura utilizando los datos disponibles.
    2. En la ventana de resultados, selecciona el registro deseado.
    3. Pulsa SELECCIONAR.
    4. Verifica que los datos se hayan cargado en la ventana principal. 
    5. Modifica los campos que necesites corregir.
    6. Pulsa ACTUALIZAR.
    7. Confirma la operacion cuando el sistema lo solicite.

IMPORTANTE
    ▸ EL registro debe estar seleccionado antes de actualizarlo.
    ▸ Comprueba que el Folio, Cliente y Numero de Compra correspondan a la 
      factura correcta.
    ▸ Verifica los datos antes de confirmar para evitar cambios 
      no deseados.
""",
          "Archivos de Facturas": """
    El sistema dispone de una carpeta llamada folios_registros para organizar 
    los documentos relacionados con los folios registrados.
    Cuando encuentra archivos cuyo nombre contiene numeros coincidentes 
    con un folio registrado, puede organizarlos en una carpeta 
    correspondiente a ese folio.

CONSULTAR DOCUMENTOS
    ▸ Buscar la factura.
    ▸ Selecciona el registro correcto.
    ▸ Pulsa ABRIR REGISTRO.
    ▸ Revisa los documentos que aparecen en la ventana de informacion de 
      facturas.
    ▸ Selecciona la accion disponible ppara consultar el archivo.
    ▸ Los archivos admitidos dependen de los formatos configurados en el 
      programa.

FORMATOS PERMITIDOS POR EL PROGRAMA
    ▸   .pdf
    ▸   .png
    ▸   .jpg
    ▸   .jpeg
    ▸   .webp
    ▸   .tif
    ▸   .tiff
    ▸   .bmp

RECOMENDACION
    ▸ Procura que los nombre de los documentos incluyan el folio 
      correspondiente para facilitar su identificacion.
""",
          "Carpetas del Programa": """
ORGANIZACION DE ARCHIVOS DEL SISTEMA
    ▸ archivos_excel
        Contiene los libros de Excel que utiliza el sistema para guardar y 
        consultar la informacion de las facturas.
        El archivo anual actual sigue el formato FACTURAS_AÑO.xlsx
    ▸ folios_facturas
        Contiene los documentos relacionados con las facturas.
        El sistema puede organizar los coumentos en subcarpetas cuyo nombre 
        corresponde al folio de la factura.
    ▸ registros
        Guarda la bitacora de actividad y errores del programa.
        Los archivos de registro permiten revisar operaciones realizadas y 
        diagnosticar problemas
    
RECOMENDACION
    ▸ No elimines ni cambies de lugar estas carpetas.
""",
          "Agregar Excel Historicos": """
INCORPORAR LIBROS DE EXCEL DE OTROS AÑOS
    Los libros historicos permiten consultar facturas de periodos anterirores 
    desde el sistema.

Antes de incorporar un libro:
    1. Haz una copias de seguridad del archivo original.
    2. Comprueba que el libro tenga los datos de las facturas.
    3. Verifica que las columnas esten en el orden esperado.
    4. Guarda el archivo dentro de archivos_excel.
    5. Comprueba que le nombre y la estructura del Excel sean compatibles.
    
ORDEN DE LAS COLUMNAS
    La hoja debe tener los encabezados en la primera fila y los registros a 
    partir de la segunda fila:
        A. Fecha
        B. Numero de Cotizacion
        C. Numero Orden de Compra
        D. Cliente
        E. Folio Factura
        F. Concepto
        G. Monto Facturado
        H. Estatus
        I. Complemento 1
        J. Complemento 2

IMPORTANTE
    ▸ No basta con cambiar el nombre del archivo si las columnas o la 
      estructura de las hojas son diferentes.
    ▸ Conserva los encabezados y el orden indicados.
    ▸ Antes de utilizar un libro historico, verifica que sus hojas
      sean compatibles con el sistema y que los datos se 
      interpreten correctamente.
""",
          "Folios y Documentos": """
FOLIOS DE FACTURA Y DOCUMENTOS RELACIONADOS
    El folio identifica la factura y permite relacionarla con sus documentos.
    La carpreta utilizada por el sistema es folios_facturas.
    Los documentos compatibles pueden organizarse en una subcarpeta con el 
    nombre del folio correspondiente.

CONSULTAR DOCUMENTOS
    1. Busca la factura
    2. Selecciona el registro correcto.
    3. Abre la informacion del registro.
    4. Selecciona el documento que quieras consultar.
    5. Utiliza ABRIR FACTURA o ABRIR CARPETA.

¿DONDE DESCARGAR LOS FOLIOS DE FACTURA?
    1. Descarga el documento
        Localiza la factura que necesitas y descarga el documento 
        correspondiente. 
        Guarda el archivo con cualquier extension mientras el sistema reconozca 
        la extension del archivo descargado.
    2. Localizar el archivo descargado
        Abre la carpeta de 'Descargas' de Windows y comprueba que los 
        documentos se hayan guardado correctamente. 
        Identifica el folio de la factura para evitar confundirlo con otro registro.
    3. Copiar el documento a folios_facturas
        Abre la carpeta del Sistema de Facturas y entra en folios_facturas.
        Copia alli los documentos descargados, o clocalos en la subcarpeta que 
        corresponda el folio si tu programa ya utiliza esa organizacion.
    4. Verificar el nombre y la ubicacion.
        Comprueba que el nombre del documento permita identificar el folio y 
        que el archivo este en la ubicacion esperada por el programa. 
        Evita cambiar nombre o mover documentos que ya esten registrados sin 
        verificar como los identifica el sistema.
    5. Abrir el documento desde el programa
        Inicia el Sistema de Facturas, Busca la factura por folio o cliente, 
        selecciona el registro correcto y utiliza la opcion disponible para abrir 
        el registro, la factura o su carpeta de documentos.

RECOMENDACIONES
    ▸ Verifica que el folio del documento coincida con el folio registrado en 
      Excel.
    ▸ Conserva juntos los archivos que correspondan a una misma factura.
    ▸ No elimines documentos ni carpetas con facturas ya registradas.
    ▸ Si el programa no encuentra un documento, revisa primero el nombre del 
      archivo y su ubicacion.
""",
          "Registros e Incorporaciones": """
CARPETA REGISTROS
    La carpeta registros almacena la bitacora del programa.
    Puede contener archivos como REGISTRO_AÑO.log, donde se guardan 
    eventos y errores para facilitar su revision.
    Si ocurre un problema, conserva el archivo de registro correspondiente 
    para poder identificar lo sucedido.

ARCHIVO incorporaciones_facturas.json
    Este archivo se encuentra en la carpeta principal del
    programa y mantiene informacion sobre incorporaciones
    de documentos relacionadas con los folios.
    Cuando el sistema mueve un documento a su carpeta de folio,
    puede registrar la incorporacion en este archivo.

RECOMENDACIONES
    ▸ No borres el archivo JSON
    ▸ No lo edites manualmente si no conoces su estructura.
    ▸ Incluyelo en las copias de seguridad.
    ▸ Si falta o presente arrores, revisa la bitacora del sistema.    
""",
          "Recomendaciones": """
BUENAS PRACTICAS
    ▸ Verfica los datos anets de registrar una factura.
    ▸ Utiliza siempre el folio correcto.
    ▸ Evita registrar la informacion incompleta en los campos obligatorios.
    ▸ Confirma que seleccionaste el registro correcto antes de actualizarlo.
    ▸ Comprueba que el monto y el estatus sean corrctos.
    ▸ Manten organizados los documentos de cada factura.
    ▸ Realiza copias de seguridad periodicas de los archivos de Excel y 
      de los documentos importantes.

COSAS QUE NO DEBES HACER
    ▸ No modificar los archivos de Excel mientras el programa está en uso.
        Podrías provocar conflictos al guardar, buscar o actualizar registros.
    ▸ No eliminar ni renombrar archivos de Excel sin verificar su función.
        Algunos archivos pueden contener registros históricos necesarios para 
        consultar facturas anteriores.
    ▸ No cambiar los nombres ni el orden de las columnas de los archivos 
      históricos.   
        El programa necesita reconocer los encabezados y la estructura de los 
        datos para leerlos correctamente.
    ▸ No mover ni eliminar carpetas del programa.
        Carpetas como `folios_facturas, `archivos_excel` y `registros` pueden 
        ser necesarias para acceder a documentos, consultar información o 
        conservar el historial.
    ▸  No borrar ni editar manualmente `incorporaciones.json`.
        Este archivo puede contener información utilizada para controlar las 
        incorporaciones de archivos históricos.
    ▸ No guardar archivos ajenos a las facturas dentro de las carpetas de 
      documentos.
        Evita mezclar archivos personales, temporales o de otros procesos con 
        los documentos que administra el sistema.
    ▸ No agregar archivos duplicados sin comprobar su contenido.
        Podrías generar confusiones al localizar folios o documentos 
        relacionados con una factura.
    ▸ No cambiar la extensión de un archivo manualmente.
        Cambiar el nombre de un PDF o Excel para que parezca otro tipo de 
        archivo no convierte su contenido al nuevo formato.
    ▸ No cerrar el programa a la fuerza mientras guarda información.
        Espera a que termine la operación para reducir el riesgo de archivos 
        incompletos o cambios sin guardar.
    ▸ No eliminar archivos antiguos solo porque ya no aparecen en la ventana 
        principal.
        Antes de borrar información histórica, verifica si todavía se necesita para 
        consultas o auditorías.

USO DE EXCEL
    El sistema puede trabajar con libros de Excel abiertos.
    Para reducir problemas de acceso, evita editar manualmente un registro 
    mientras el programa lo esta actualizando.
    Si ocurre un error, revisa el mensaje que aparece en pantalla y verifica que 
    el archivo no este bloqueado por otro proceso.
""",
          "Acerca del Programa": """
    SISTEMA DE FACTURAS | Version 1.4.4

PROPOSITO
    Facilitar el registro, la consulta y actualizacion de informacion de facturas, 
    asi como la organizacion de los documentos relacionados con sus folios.

FUNCIONES PRINCIPALES
    ▸ Registro de facturas.
    ▸ Busca de registros.
    ▸ Actualizacion de informacion.
    ▸ Consulta de documentos relacionados.
    ▸ Organizacion de archivos por folio.
    ▸ Registro de eventos y errores del programa.

    La informacion se gestiona mediante la interfaz del programa y los 
    archivos de Excel configurados para su funcionamiento.
"""
        }

def encabezados_complementos(self):
    return [
         "FECHA EMITIDA",
         "CLIENTE",
         "FOLIO FACTURA",
         "COMPLEMENTO",
         "MONTO FACTURADO",
         "SALDO ANTERIOR",
         "MONTO PAGADO",
         "FECHA DE PAGO",
         "ESTATUS"
    ]


     
    