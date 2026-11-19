

\# Sistema de Facturas



Aplicación de escritorio desarrollada en Python con Tkinter para facilitar el registro, consulta y actualización de información de facturas mediante archivos de Excel.



El sistema permite organizar la información por año y mes, administrar complementos de pago y consultar documentos relacionados con cada factura.



\## Características principales



\- \*\*Registro de facturas:\*\* permite ingresar información como fecha, número de cotización, orden de cotización, cliente, folio, concepto, monto y estatus.

\- \*\*Búsqueda de registros:\*\* permite localizar facturas utilizando diferentes criterios de búsqueda.

\- \*\*Actualización de información:\*\* permite modificar registros previamente almacenados.

\- \*\*Administración de complementos de pago:\*\* permite registrar y consultar información relacionada con los pagos de las facturas.

\- \*\*Organización automática:\*\* genera archivos de Excel por año y organiza la información mediante hojas correspondientes a los meses.

\- \*\*Consulta de documentos:\*\* permite trabajar con documentos relacionados con las facturas.

\- \*\*Registro de actividad:\*\* genera archivos de registro para almacenar información sobre la actividad del sistema.

\- \*\*Interfaz gráfica:\*\* cuenta con ventanas independientes para facilitar las operaciones y la consulta de información.



\## Tecnologías utilizadas



\- Python

\- Tkinter

\- OpenPyXL

\- PyWin32

\- Microsoft Excel



\## Estructura del proyecto



```text

ProgramasExcel/

│

├── iconos/

│   ├── icono\_busqueda.ico

│   ├── icono\_facturas.ico

│   ├── icono\_manual.ico

│   ├── icono\_manual.png

│   ├── icono\_principal.ico

│   └── icono\_principal.png

│

├── main.py

├── proceso\_excel.py

├── complementos\_proceso.py

├── requirements.txt

├── README.md

└── .gitignore

```



Durante la ejecución, el sistema genera y administra otras carpetas y archivos locales.



\## Requisitos



Para ejecutar el proyecto necesitas:



\- Windows.

\- Python 3.10 o superior.

\- Microsoft Excel instalado.

\- Pip para instalar las dependencias.



\## Instalación



\### 1. Clonar el repositorio



Abre una terminal y ejecuta:



```bash

git clone https://github.com/r3tlop/facturas\_excel.git

```



\### 2. Entrar a la carpeta del proyecto



```bash

cd ProgramasExcel

```



\### 3. Instalar las dependencias



```bash

pip install -r requirements.txt

```



\### 4. Ejecutar la aplicación



```bash

python main.py

```



\## Dependencias



Las principales dependencias externas utilizadas son:



\- `openpyxl`: permite crear, leer y modificar archivos de Excel.

\- `pywin32`: permite interactuar con Microsoft Excel mediante la interfaz COM de Windows.



\## Almacenamiento y organización de archivos



El sistema utiliza la ubicación del programa como referencia para organizar sus archivos.



Durante su funcionamiento administra las siguientes carpetas:



| Carpeta o archivo | Descripción |

|---|---|

| `archivos\_excel/` | Almacena los libros de Excel con los registros de facturas. |

| `folios\_facturas/` | Almacena documentos relacionados con las facturas. |

| `registros/` | Contiene los archivos de registro de actividad del sistema. |

| `incorporaciones\_facturas.json` | Almacena información relacionada con las incorporaciones de documentos. |



\### Organización de los archivos Excel



Los registros se organizan en libros de Excel por año, utilizando el siguiente formato:



```text

FACTURAS\_AÑO.xlsx

```



Dentro de cada libro, la información se distribuye mediante hojas correspondientes a los meses del año.



Los archivos de Excel, documentos de facturas, registros y datos locales generados durante la ejecución no forman parte del repositorio público.



\## Estado del proyecto



Proyecto en desarrollo.



El objetivo es facilitar la administración de información de facturas mediante una herramienta de escritorio, reduciendo tareas manuales y manteniendo los registros organizados.



\## Autor



Proyecto personal desarrollado como práctica de programación en Python, automatización de procesos y administración de información mediante Excel.



