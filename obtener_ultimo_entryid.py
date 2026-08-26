# obtener_ultimo_entryid.py — ejecútalo UNA vez desde la carpeta FacturasSync
from outlook_service import conectar
from config import CORREO_ALBARANES

carpeta = conectar(CORREO_ALBARANES)
mensajes = carpeta.Items
mensajes.Sort("[ReceivedTime]", True)  # más reciente primero

ultimo = mensajes.GetFirst()
print("EntryID del correo más reciente en", CORREO_ALBARANES, ":")
print(ultimo.EntryID)
print("Asunto:", ultimo.Subject, "| Recibido:", ultimo.ReceivedTime)