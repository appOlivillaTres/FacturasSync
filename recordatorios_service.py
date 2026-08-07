from datetime import date, datetime, timezone

import requests

from config import SUPABASE_URL, SUPABASE_KEY, CORREO_RECORDATORIOS_PRESUPUESTOS
from logger import log
from outlook_service import enviar_correo


class RecordatoriosService:
    """
    Revisa la tabla 'presupuestos_seguimiento' y, para cada presupuesto cuya
    fecha de recordatorio ya ha llegado (o pasado) y que todavía no se ha
    marcado como respondido ni como recordatorio ya enviado, manda un email
    al cliente pidiéndole que conteste al presupuesto.
    """

    def __init__(self):
        self.headers = {
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
        }

    def obtener_pendientes(self):

        hoy = date.today().isoformat()

        url = (
            f"{SUPABASE_URL}/rest/v1/presupuestos_seguimiento"
            f"?recordatorio_enviado=eq.false&respondido=eq.false"
            f"&fecha_recordatorio=lte.{hoy}&select=*"
        )

        r = requests.get(url, headers=self.headers)
        r.raise_for_status()

        return r.json()

    def marcar_enviado(self, presupuesto_id):

        url = f"{SUPABASE_URL}/rest/v1/presupuestos_seguimiento?id=eq.{presupuesto_id}"

        r = requests.patch(
            url,
            headers=self.headers,
            json={
                "recordatorio_enviado": True,
                "fecha_recordatorio_enviado": datetime.now(timezone.utc).isoformat(),
            },
        )
        r.raise_for_status()

    def _construir_mensaje(self, presupuesto):

        asunto_presupuesto = presupuesto.get("asunto")
        fecha_envio = presupuesto.get("fecha_envio")

        asunto = "Recordatorio: presupuesto pendiente de respuesta"

        cuerpo = (
            f"Buen día,\n\n"
            f"Nos ponemos en contacto desde Aislamientos Olivilla Tres para saber si tienes alguna duda"
            f"con el presupuesto enviado el pasado {fecha_envio}"
            + (f" ({asunto_presupuesto})" if asunto_presupuesto else "")
            + "Estamos a sus ordenes, no dude en contactarnos para cualquier información.\n\n"
            f"Gracias por su atención.\n\n"
            f"Un saludo,\nAislamientos Olivilla Tres"
        )

        return asunto, cuerpo

    def revisar_y_enviar(self):

        try:
            pendientes = self.obtener_pendientes()
        except Exception as e:
            log(f"ERROR consultando presupuestos pendientes de recordatorio: {e}")
            return

        if not pendientes:
            log("Recordatorios de presupuestos: no hay ninguno pendiente hoy")
            return

        for p in pendientes:

            try:
                email = p.get("email_cliente")

                if not email:
                    log(f"Presupuesto id={p.get('id')} sin email de cliente, se omite")
                    continue

                asunto, cuerpo = self._construir_mensaje(p)

                enviar_correo(
                    email,
                    asunto,
                    cuerpo,
                    cuenta_remitente=CORREO_RECORDATORIOS_PRESUPUESTOS,
                )

                self.marcar_enviado(p["id"])

                log(f"Recordatorio de presupuesto enviado a {email} (id={p['id']})")

            except Exception as e:
                log(f"ERROR enviando recordatorio de presupuesto (id={p.get('id')}): {e}")