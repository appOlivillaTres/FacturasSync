import os

SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_KEY = "TU_SERVICE_ROLE_KEY"

CARPETA_OUTLOOK = "FACTURAS"

# Cuentas de Outlook separadas por tipo de documento: facturas@ recibe
# facturas, almacen@ recibe albaranes. Cada una se lee de forma
# independiente (ver outlook_service.conectar). El tipo de documento ya no
# se decide adivinando por el contenido del PDF: lo decide el buzón de
# origen (tipo_documento() se mantiene solo como aviso/red de seguridad —
# ver main.py).
CORREO_FACTURAS = "facturas@olivillatres.com"
CORREO_ALBARANES = "almacen@olivillatres.com"

TEMP_FOLDER = "temp"

LOG_FOLDER = "logs"

ESTADO_FILE = "estado.json"

INTERVALO = 30

# Cada cuántos segundos se revisa si hay recordatorios de presupuestos que
# enviar (antes se miraba solo una vez al día; ahora se comprueba de forma
# periódica según este intervalo). Por defecto, cada hora.
INTERVALO_RECORDATORIOS = 3600
import os
from dotenv import load_dotenv

load_dotenv()

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

# Facturas de este NIF o de este proveedor van SIEMPRE a facturas_csg,
# sin pasar por la lógica normal de destino (obra/almacén).
NIF_CSG = "B10708832"
NOMBRE_CSG = "CSG REFORMAS"

# NIF de nuestra propia empresa. Muchas facturas lo mencionan (porque somos
# el cliente/destinatario), y a veces la IA o el regex lo confunden con el
# NIF del proveedor que emite la factura. Se usa para descartarlo y seguir
# buscando el NIF real del proveedor.
NIF_PROPIO = "B45813284"
NOMBRE_PROPIO = "OLIVILLA TRES"

# Cuenta de Outlook desde la que se envían los recordatorios de presupuestos
# a los clientes. Debe ser una cuenta que ya esté configurada/añadida en el
# mismo Outlook donde corre FacturaSync (aunque el buzón que se lee para
# facturas/albaranes sea otro distinto).
CORREO_RECORDATORIOS_PRESUPUESTOS = "almacen@olivillatres.com"

# Ruta al ejecutable de Tesseract OCR, necesaria para leer facturas que son
# fotos/escaneos sin capa de texto (pdfplumber y PyMuPDF no pueden leerlas).
# Tesseract es un programa aparte, no se instala con pip — hay que instalarlo
# en el PC de Windows donde corre FacturaSync (ver instrucciones abajo).
# Déjalo en None si tesseract.exe ya está en el PATH del sistema.
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"