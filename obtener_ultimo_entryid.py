# obtener_ultimo_entryid_facturas.py
from outlook_service import conectar
from config import CORREO_FACTURAS

carpeta = conectar(CORREO_FACTURAS)
mensajes = carpeta.Items
mensajes.Sort("[ReceivedTime]", True)

ultimo = mensajes.GetFirst()
print("EntryID del correo más reciente en", CORREO_FACTURAS, ":")
print(ultimo.EntryID)
print("Asunto:", ultimo.Subject, "| Recibido:", ultimo.ReceivedTime)