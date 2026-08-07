from datetime import datetime
import os

os.makedirs("logs", exist_ok=True)

LOG_FILE = "logs/facturasync.log"


def log(texto):

    linea = f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {texto}"

    print(linea)

    with open(LOG_FILE, "a", encoding="utf8") as f:
        f.write(linea + "\n")
        