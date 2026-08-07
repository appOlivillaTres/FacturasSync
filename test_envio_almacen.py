"""
Script de prueba. Manda UN correo de prueba usando la cuenta de almacen,
para comprobar si el problema esta en la logica de recordatorios o en el
envio por Outlook en si.

Uso:
    python test_envio_almacen.py tu_email_personal@gmail.com
"""

import sys

from outlook_service import enviar_correo
from config import CORREO_RECORDATORIOS_PRESUPUESTOS

if len(sys.argv) < 2:
    print("Uso: python test_envio_almacen.py destinatario@ejemplo.com")
    sys.exit(1)

destinatario = sys.argv[1]

print(f"Cuenta configurada para enviar (CORREO_RECORDATORIOS_PRESUPUESTOS): {CORREO_RECORDATORIOS_PRESUPUESTOS}")
print(f"Enviando correo de prueba a: {destinatario}\n")

enviar_correo(
    destinatario=destinatario,
    asunto="Prueba de envio desde almacen",
    cuerpo="Esto es un correo de prueba para comprobar desde que cuenta llega.",
    cuenta_remitente=CORREO_RECORDATORIOS_PRESUPUESTOS,
)

print("\nListo. Revisa el archivo logs/facturasync.log para ver el detalle,")
print("y sobre todo mira el campo 'De' / 'From' del correo que te llegue")
print("(no donde se haya guardado en 'Elementos enviados' de Outlook).")