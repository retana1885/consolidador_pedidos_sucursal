CONSOLIDADOR DE HOJAS DE ENVÍO A SUCURSALES

Aplicación de escritorio desarrollada en Python + Tkinter para consolidar
Notas de Salida de Traspaso (TMA = 16) desde BODEGA (origen fijo) hacia
sucursales, generando un archivo Excel apilado a partir de plantillas editables.

El sistema se conecta a SQL Server (NOVACAJA), permite seleccionar múltiples
notas por fecha y sucursal destino, y genera automáticamente un archivo Excel
con el formato correcto según la sucursal.

FUNCIONALIDADES PRINCIPALES

Conexión directa a SQL Server mediante ODBC

Origen fijo: BODEGA = FolTda_Codigo = 2

Filtro por:

Fecha

Sucursal destino (AlmTda_Codigo)

Doble listado:

Notas disponibles

Notas seleccionadas

Ordenamiento por:

F2

F1

Hora

Generación de Excel apilado:

Un solo encabezado

Múltiples notas consecutivas

Respeta formato (fuente, bordes, alineación)

Detección automática de plantilla:

GENERAL

COLIMA

MANZANILLO

TECOMAN

Plantillas externas (no requiere recompilar)

Archivo config.json editable

Compilable a .exe para uso sin Python instalado

ESTRUCTURA DEL PROYECTO

ConsolidadorPedidos
|
|-- main.py Código principal
|-- config.json Configuración editable
|-- icon.ico Icono del sistema
|
|-- Plantillas
| |-- Plantilla GENERAL.xlsx
| |-- Plantilla COLIMA.xlsx
| |-- Plantilla MANZANILLO.xlsx
| |-- Plantilla TECOMAN.xlsx
|
|-- README_ConsolidadorPedidos.txt

REQUISITOS

DESARROLLO:

Python 3.10 o superior

Paquetes:
pyodbc
openpyxl
pyinstaller

PRODUCCIÓN:

Windows 10 / 11

ODBC Driver 17 for SQL Server

Acceso a base de datos SQL Server

CONFIGURACIÓN (config.json)

Ejemplo:

{
"db": {
"driver": "ODBC Driver 17 for SQL Server",
"server": "IP_SERVIDOR",
"database": "NOVACAJA5",
"username": "USUARIO",
"password": "PASSWORD",
"trust_server_certificate": true
},
"app": {
"tma_traspaso": 16,
"origen_bodega_foltda": 2,
"templates_dir": "Plantillas",
"default_output_name": "Traspasos_Apilado.xlsx"
}
}

USO DE LA APLICACIÓN

Seleccionar fecha

Seleccionar sucursal destino

Presionar "Buscar notas"

Mover notas a "Seleccionadas"

Presionar "Generar Excel (apilado)"

Elegir ubicación del archivo

LÓGICA DE NEGOCIO

Solo procesa:

TMA_Codigo = 16

FolTda_Codigo = 2 (BODEGA)

Folio formado como:
FolTda-FolEst-FolDoc-FolConsecutivo

Referencia 2 NO se imprime (por diseño)

El formato de la primera nota (F1) se replica
en todas las notas siguientes

COMPILACIÓN A .EXE (RESUMEN)

Comando de compilación:

python -m PyInstaller --noconfirm --clean --onefile --windowed --name ConsolidadorPedidos --icon icon.ico main.py

Después de compilar, colocar junto al .exe:

config.json

carpeta Plantillas

icon.ico

TECNOLOGÍAS UTILIZADAS

Python

Tkinter

pyodbc

openpyxl

SQL Server

PyInstaller

NOTAS IMPORTANTES

Las plantillas pueden modificarse sin recompilar

El sistema detecta automáticamente la plantilla correcta

No hay credenciales hardcodeadas

Pensado para uso operativo real

AUTOR

Juan Pablo Retana Larios
Desarrollo interno – Automatización operativa y analítica
Grupo Canave / Farmacias San Camilo

LICENCIA

Uso interno / privado
No redistribuible sin autorización
