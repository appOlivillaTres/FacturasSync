import win32com.client
import os
import time

from logger import log

# DISPID de la propiedad "SendUsingAccount" de MailItem en el modelo de
# objetos de Outlook. La asignación normal por Python
# (mail.SendUsingAccount = cuenta) a veces no se aplica de verdad por un
# problema conocido de pywin32 con esta propiedad concreta (se queda en
# None aunque no dé ningún error). Invocarla directamente por su DISPID
# evita ese problema.
DISPID_SENDUSINGACCOUNT = 64209


def conectar(correo_cuenta):
    """
    Conecta con la bandeja de entrada de una cuenta CONCRETA (no la cuenta
    por defecto del perfil de Outlook), buscándola entre las cuentas que
    ya están dadas de alta en este Outlook por su dirección de email. Así
    se puede leer facturas y albaranes desde dos buzones distintos
    (facturas@olivillatres.com / almacen@olivillatres.com) aunque ambos
    estén configurados en el mismo Outlook del PC.

    Lanza un error si la cuenta no está dada de alta en este Outlook, para
    no acabar leyendo por accidente la bandeja de otra cuenta distinta.
    """

    outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")

    cuenta_encontrada = None

    for cuenta in namespace.Accounts:
        if cuenta.SmtpAddress.lower() == correo_cuenta.lower():
            cuenta_encontrada = cuenta
            break

    if not cuenta_encontrada:
        raise RuntimeError(
            f"La cuenta '{correo_cuenta}' no está añadida en este Outlook. "
            f"Añádela en Archivo > Agregar cuenta para poder leer su bandeja de entrada."
        )

    # 6 = Bandeja de entrada. DeliveryStore es el almacén de correo propio
    # de ESA cuenta (a diferencia de GetDefaultFolder(6) sobre el
    # namespace, que coge la cuenta por defecto del perfil sin más).
    inbox = cuenta_encontrada.DeliveryStore.GetDefaultFolder(6)

    log(f"Bandeja de entrada conectada: '{correo_cuenta}' (carpeta '{inbox.Name}')")

    return inbox


def enviar_correo(destinatario, asunto, cuerpo, cuenta_remitente=None):
    """
    Envía un correo simple (texto plano) usando Outlook, para los
    recordatorios automáticos de presupuestos pendientes de respuesta.

    Si se indica 'cuenta_remitente' (una dirección de email), se busca esa
    cuenta entre las que ya están dadas de alta en Outlook y se usa para
    enviar. Esto permite, por ejemplo, que el buzón que se lee para
    facturas/albaranes sea uno, y que los recordatorios de presupuestos se
    manden desde otro (p.ej. almacen@olivillatres.com), siempre que ambas
    cuentas estén configuradas en el mismo Outlook.

    IMPORTANTE: si se pide una 'cuenta_remitente' concreta y esa cuenta NO
    está dada de alta en este Outlook, el correo NO se envía (se lanza un
    error). Así se evita el problema de que, por no encontrar la cuenta
    correcta, el correo salga igualmente desde la cuenta por defecto (la de
    lectura de facturas), mezclando ambos buzones sin querer.
    """

    outlook_app = win32com.client.Dispatch("Outlook.Application")

    mail = outlook_app.CreateItem(0)  # 0 = olMailItem
    mail.To = destinatario
    mail.Subject = asunto
    mail.Body = cuerpo

    if cuenta_remitente:

        namespace = outlook_app.GetNamespace("MAPI")
        cuenta_encontrada = None

        for cuenta in namespace.Accounts:
            if cuenta.SmtpAddress.lower() == cuenta_remitente.lower():
                cuenta_encontrada = cuenta
                break

        if not cuenta_encontrada:
            raise RuntimeError(
                f"La cuenta '{cuenta_remitente}' no está añadida en este Outlook. "
                f"No se envía el correo para no hacerlo por error desde otra cuenta "
                f"(revisa Archivo > Agregar cuenta en Outlook para añadirla en este PC)."
            )

        log(f"Cuenta encontrada para el envío: '{cuenta_encontrada.SmtpAddress}' (DisplayName: '{cuenta_encontrada.DisplayName}')")

        # Asignación normal (a veces basta con esto)...
        mail.SendUsingAccount = cuenta_encontrada

        # ...pero forzamos también por DISPID, que es el workaround fiable
        # para el bug de pywin32 donde la asignación normal no se aplica.
        mail._oleobj_.Invoke(*(DISPID_SENDUSINGACCOUNT, 0, 8, 0, cuenta_encontrada))

        # Comprobación: ¿se ha quedado realmente asignada la cuenta que queríamos?
        try:
            cuenta_aplicada = mail.SendUsingAccount.SmtpAddress
        except Exception:
            cuenta_aplicada = None
        log(f"mail.SendUsingAccount tras asignar: '{cuenta_aplicada}'")

        if not cuenta_aplicada or cuenta_aplicada.lower() != cuenta_remitente.lower():
            raise RuntimeError(
                f"No se ha podido forzar el envío desde '{cuenta_remitente}' "
                f"(Outlook sigue queriendo usar la cuenta por defecto). No se envía "
                f"el correo para evitar que salga desde la cuenta equivocada."
            )

    mail.Send()
    log(f"Correo enviado a {destinatario}" + (f" usando '{cuenta_remitente}'" if cuenta_remitente else " (cuenta por defecto)"))


from word_service import convertir_word_a_pdf

EXTENSIONES_DOCUMENTO = (".pdf", ".doc", ".docx")


def guardar_adjuntos_pdf(correo, carpeta_destino="temp"):

    import os
    import time

    carpeta_destino = os.path.abspath(carpeta_destino)
    os.makedirs(carpeta_destino, exist_ok=True)

    pdfs = []

    for i in range(1, correo.Attachments.Count + 1):

        adjunto = correo.Attachments.Item(i)
        nombre_lower = adjunto.FileName.lower()

        if not nombre_lower.endswith(EXTENSIONES_DOCUMENTO):
            continue

        ruta = os.path.join(carpeta_destino, adjunto.FileName)

        if os.path.exists(ruta):
            nombre, ext = os.path.splitext(adjunto.FileName)
            ruta = os.path.join(
                carpeta_destino,
                f"{nombre}_{int(time.time())}{ext}"
            )

        try:
            adjunto.SaveAsFile(ruta)

            if nombre_lower.endswith((".doc", ".docx")):

                ruta_pdf = convertir_word_a_pdf(ruta)

                if ruta_pdf:
                    pdfs.append(ruta_pdf)
                else:
                    log(f"AVISO: no se pudo convertir '{ruta}' a PDF, se omite este adjunto")

            else:
                pdfs.append(ruta)

        except Exception as e:
            print(f"ERROR guardando adjunto: {e}")

    return pdfs