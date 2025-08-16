Bot de extracción de expedientes UNAM

Descripción:
Este bot automatiza la extracción de información de expedientes desde la plataforma de seguimiento de titulación de la UNAM. Realiza login manual, aplica filtros, recorre los expedientes disponibles (paginando cuando es necesario) y exporta los resultados a Excel o CSV.

Requisitos:
-Python
-Selenium
-webdriver-manager
-pandas
-openpyxl

Instalación:
pip install selenium webdriver-manager pandas openpyxl

Uso:
python bot.py
1.Se abrirá una ventana de Chrome.
2.Inicia sesión manualmente en la plataforma.
3.El script detectará el login y aplicará automáticamente el filtro deseado.
4.Se abrirá cada expediente en una nueva pestaña y se extraerán los datos.
5.Los datos se exportarán a un archivo Excel (expedientes-YYYYMMDD-HHMMSS.xlsx) y a expedientes.json

Funcionalidades:
Aplicación automática del filtro por estado (valor por default: "Entrega electrónica y física de documentos").
Ajuste del número de registros mostrados por página a 100.
Extracción de datos clave del expediente (cuenta, nombre, carrera, plantel, etc.).
Detección de expedientes sin cita (omitidos).
Exportación a Excel y JSON.