import json
import sys
import time
from datetime import datetime, timezone

import msvcrt

from outlook_service import conectar, guardar_adjuntos_pdf
from pdf_service import PDFService
from ia_service import AIService
from supabase_service import SupabaseService
from recordatorios_service import RecordatoriosService

from logger import log
from config import (
    ESTADO_FILE, INTERVALO, INTERVALO_RECORDATORIOS,
    CORREO_FACTURAS, CORREO_ALBARANES,
)


LOCK_FILE = "facturasync.lock"


def asegurar_instancia_unica():
    """
    Evita que se ejecuten dos copias de main.py a la vez (p.ej. la tarea
    programada y una ejecución manual desde VS Code coincidiendo, o dos
    lanzamientos manuales seguidos sin cerrar el anterior). Si ya hay una
    instancia corriendo, esta se cierra inmediatamente en vez de arrancar
    y pisar el estado.json / duplicar la lectura de Outlook.

    Se queda con el archivo de lock abierto durante toda la ejecución (la
    referencia se guarda en una variable global para que no la recoja el
    recolector de basura) — el bloqueo se libera solo al cerrarse el
    proceso, sea de forma normal o por un cuelgue.
    """

    global _lock_handle

    try:
        _lock_handle = open(LOCK_FILE, "w")
        msvcrt.locking(_lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        print("Ya hay una instancia de FacturaSync en ejecución. Cerrando esta.")
        sys.exit(1)


def cargar_estado():

    try:
        with open(ESTADO_FILE, encoding="utf8") as f:
            estado = json.load(f)
    except:
        estado = {}

    # Compatibilidad con el estado.json antiguo, en dos pasos:
    #
    # 1) Formato de un único buzón (clave "ultimo_entryid") → se reutiliza
    #    como punto de partida de FACTURAS.
    # 2) Formato de EntryID por buzón ("ultimo_entryid_facturas" /
    #    "ultimo_entryid_albaranes") → se abandona: el EntryID deja de ser
    #    válido en cuanto el correo se mueve o se borra de la bandeja (fue
    #    la causa de que se releyeran años de correo antiguo varias veces).
    #    Ahora el cursor es la FECHA de recepción, que no depende de que
    #    el correo siga físicamente ahí. Si no hay fecha guardada todavía,
    #    se arranca "desde ahora" en vez de desde el principio del
    #    histórico, para no reprocesar todo de golpe en la migración.
    if "ultima_fecha_facturas" not in estado:
        estado["ultima_fecha_facturas"] = datetime.now(timezone.utc).isoformat()
        log("Migración de estado.json: sin fecha de facturas guardada, se arranca desde ahora (no se reprocesa el histórico).")

    if "ultima_fecha_albaranes" not in estado:
        estado["ultima_fecha_albaranes"] = datetime.now(timezone.utc).isoformat()
        log("Migración de estado.json: sin fecha de albaranes guardada, se arranca desde ahora (no se reprocesa el histórico).")

    return estado


def guardar_estado(estado):

    with open(ESTADO_FILE, "w", encoding="utf8") as f:
        json.dump(estado, f, indent=4)


def obtener_nuevos(carpeta, ultima_fecha_iso):
    """
    Devuelve los correos recibidos DESPUÉS de 'ultima_fecha_iso', del más
    antiguo al más reciente (listos para procesar en orden).

    Antes se comparaba por EntryID del último correo procesado, pero eso
    falla si ese correo deja de existir en la carpeta tal cual (se mueve,
    se archiva, se borra...): al no encontrar coincidencia, el bucle
    recorría TODA la bandeja sin encontrar nunca el punto de corte, y
    trataba años de correo antiguo como "nuevo" en cada ciclo. La fecha de
    recepción no tiene ese problema: no depende de que el correo siga
    físicamente en la carpeta.
    """

    mensajes = carpeta.Items
    mensajes.Sort("[ReceivedTime]", True)  # más reciente primero

    nuevos = []

    for correo in mensajes:

        try:
            recibido = correo.ReceivedTime
        except Exception:
            continue

        # Como está ordenado de más reciente a más antiguo, en cuanto
        # llegamos a un correo recibido antes (o en el mismo instante) que
        # el último ya procesado, todo lo que queda por delante ya se
        # procesó en alguna vuelta anterior.
        if ultima_fecha_iso and recibido.isoformat() <= ultima_fecha_iso:
            break

        nuevos.append(correo)

    nuevos.reverse()

    return nuevos


def procesar_pdf_albaran(ruta_pdf, datos, pdf, ia, supa):
    """
    Procesa un PDF que viene del buzón de albaranes (almacen@olivillatres.com).
    'datos' es el resultado de pdf.extraer_parcial(ruta_pdf), ya calculado
    por el llamador (se reutiliza el texto/imagenes ya extraídos).
    """

    log(f"'{ruta_pdf}' → ALBARÁN (buzón {CORREO_ALBARANES})")

    try:
        imagenes_pag = datos.get("imagenes_paginas") or []

        if imagenes_pag:
            log(
                f"Texto insuficiente incluso tras OCR — usando IA con visión "
                f"directamente sobre la imagen ({len(imagenes_pag)} página/s)"
            )
            datos_albaran = ia.analizar_albaran_imagen(imagenes_pag)
        else:
            try:
                datos_albaran = ia.analizar_albaran_pdf(ruta_pdf)
            except Exception as e_pdf:
                log(f"AVISO: fallo mandando PDF de albarán directo a la IA ({e_pdf}), usando texto extraído como respaldo")
                datos_albaran = ia.analizar_albaran(datos.get("texto", ""))

        log(f"Datos de albarán (IA): {datos_albaran}")
    except Exception as e:
        log(f"ERROR usando IA para albarán: {e}")
        datos_albaran = {}

    datos_albaran["fecha"] = pdf.normalizar_fecha(datos_albaran.get("fecha"))

    # Algunos proveedores (p.ej. Sika) tienen un "Su Pedido (online Shop):
    # alias_cliente" que la IA a veces confunde con el número de pedido
    # real. Un pedido real siempre trae algún dígito.
    if datos_albaran.get("numero_pedido") and not any(
        c.isdigit() for c in str(datos_albaran["numero_pedido"])
    ):
        log(
            f"numero_pedido de albarán descartado por no ser numérico: "
            f"'{datos_albaran['numero_pedido']}'"
        )
        datos_albaran["numero_pedido"] = None

    if not datos_albaran.get("numero_pedido"):
        datos_albaran["numero_pedido"] = pdf.buscar_numero_pedido(datos.get("texto", ""))

    # El NIF que lee la IA directamente del PDF puede venir mal. Si ya
    # tenemos este proveedor dado de alta en 'proveedores', usamos SIEMPRE
    # su NIF registrado en vez del leído por la IA.
    if datos_albaran.get("empresa"):

        try:
            proveedor_encontrado = supa.buscar_proveedor(
                nombre=datos_albaran.get("empresa")
            )

            if proveedor_encontrado and proveedor_encontrado.get("nif"):

                nif_ia = datos_albaran.get("nif")
                nif_registrado = proveedor_encontrado["nif"]

                if nif_ia and nif_ia != nif_registrado:
                    log(
                        f"NIF sobrescrito con el de 'proveedores' "
                        f"('{datos_albaran.get('empresa')}': IA leyó '{nif_ia}', "
                        f"se usa el registrado '{nif_registrado}')"
                    )
                elif not nif_ia:
                    log(
                        f"NIF completado desde 'proveedores' por nombre "
                        f"('{datos_albaran.get('empresa')}' → {nif_registrado})"
                    )

                datos_albaran["nif"] = nif_registrado

        except Exception as e:
            log(f"ERROR buscando NIF por nombre de proveedor: {e}")

    archivo_url_alb = None
    try:
        archivo_url_alb = supa.subir_archivo(ruta_pdf)
    except Exception as e:
        log(f"ERROR subiendo albarán a Storage: {e}")

    supa.insertar_albaran(datos_albaran, archivo_url=archivo_url_alb)
    log("Albarán insertado en tabla albaranes")


def procesar_pdf_factura(ruta_pdf, datos, remitente, asunto, fecha_recibido_correo, pdf, ia, supa):
    """
    Procesa un PDF que viene del buzón de facturas (facturas@olivillatres.com).
    'datos' es el resultado de pdf.extraer_parcial(ruta_pdf) (ya trae
    'fecha_recibido' puesto por el llamador).
    """

    log(f"'{ruta_pdf}' → FACTURA (buzón {CORREO_FACTURAS})")

    # La IA analiza SIEMPRE el texto completo, ya que el regex solo cubre
    # nif/numero_factura/fecha/base/iva/total y deja vacíos el resto de
    # campos del excel (base2, retenciones, etc.)
    try:

        imagenes_pag = datos.get("imagenes_paginas") or []

        if imagenes_pag:
            log(
                f"Texto insuficiente incluso tras OCR — usando IA con visión "
                f"directamente sobre la imagen ({len(imagenes_pag)} página/s)"
            )
            datos_ia = ia.analizar_factura_imagen(imagenes_pag)
        else:
            try:
                datos_ia = ia.analizar_factura_pdf(ruta_pdf)
            except Exception as e_pdf:
                log(f"AVISO: fallo mandando PDF de factura directo a la IA ({e_pdf}), usando texto extraído como respaldo")
                datos_ia = ia.analizar_factura(datos.get("texto", ""))

        log(f"Datos extraídos (IA): {datos_ia}")

        for k, v in datos_ia.items():

            if v not in (None, "", 0):
                datos[k] = v
            elif k not in datos:
                datos[k] = v

    except Exception as e:

        log(f"ERROR usando IA: {e}")

    datos["fecha"] = pdf.normalizar_fecha(datos.get("fecha"))
    fecha_regex = pdf.normalizar_fecha(pdf.buscar_fecha(datos.get("texto", "")))
    if fecha_regex and datos.get("fecha") and fecha_regex != datos["fecha"]:
        log(f"AVISO: fecha de la IA ('{datos['fecha']}') no coincide con la del regex ('{fecha_regex}') — revisar '{ruta_pdf}'")

    numero_regex = pdf.buscar_numero_factura(datos.get("texto", ""))
    if numero_regex and datos.get("numero_factura") and numero_regex != datos["numero_factura"]:
        log(f"AVISO: número de factura de la IA ('{datos['numero_factura']}') no coincide con el del regex ('{numero_regex}') — revisar '{ruta_pdf}'")

    if datos.get("fecha_vencimiento"):
        datos["fecha_vencimiento"] = pdf.normalizar_fecha(datos.get("fecha_vencimiento"))

    if datos.get("numero_pedido") and not any(
        c.isdigit() for c in str(datos["numero_pedido"])
    ):
        log(f"numero_pedido descartado por no ser numérico: '{datos['numero_pedido']}'")
        datos["numero_pedido"] = None

    if not datos.get("numero_pedido"):
        datos["numero_pedido"] = pdf.buscar_numero_pedido(datos.get("texto", ""))

    if datos.get("nif") and not pdf.nif_valido(datos["nif"]):

        nif_regex = pdf.buscar_nif(datos.get("texto", ""))

        if nif_regex and pdf.nif_valido(nif_regex) and nif_regex != datos["nif"]:
            log(f"NIF '{datos['nif']}' (IA) no pasa el dígito de control, se usa '{nif_regex}' (regex) en su lugar")
            datos["nif"] = nif_regex
        else:
            log(f"AVISO: el NIF '{datos['nif']}' no pasa el dígito de control y no hay alternativa mejor — revisar manualmente")

    cuadra, diferencia = pdf.validar_cuadre(datos)
    if not cuadra:
        log(f"AVISO: la factura no cuadra (base+IVA vs total difieren en {diferencia}€) — revisar '{ruta_pdf}'")

    numeros_albaran_detectados = (
        datos.get("numeros_albaran") or pdf.buscar_numeros_albaran(datos.get("texto", ""))
    )

    datos["numeros_albaran"] = numeros_albaran_detectados

    archivo_url = None

    try:
        archivo_url = supa.subir_archivo(ruta_pdf)
    except Exception as e:
        log(f"ERROR subiendo PDF a Storage: {e}")

    proveedor = supa.buscar_proveedor(
        nif=datos.get("nif"),
        nombre=datos.get("proveedor")
    )

    log(
        f"Comprobando si es CSG → nif='{datos.get('nif')}', "
        f"proveedor='{datos.get('proveedor')}'"
    )

    if datos.get("total") and supa.es_csg(
        nif=datos.get("nif"),
        nombre=datos.get("proveedor"),
        texto=datos.get("texto")
    ):

        resultado_ins = supa.insertar_factura(
            datos,
            "facturas_csg",
            archivo_url=archivo_url
        )

        log("Factura insertada en CSG")

        supa.marcar_albaranes_facturados(
            numeros_albaran_detectados,
            datos.get("nif"),
            "facturas_csg",
            resultado_ins["id"],
            numero_pedido=datos.get("numero_pedido")
        )

    elif proveedor and datos.get("total"):

        destino = proveedor["destino"]

        if destino == "obra":

            resultado_ins = supa.insertar_factura(
                datos,
                "facturas_obra",
                archivo_url=archivo_url
            )

            log("Factura insertada en OBRA")

            supa.marcar_albaranes_facturados(
                numeros_albaran_detectados,
                datos.get("nif"),
                "facturas_obra",
                resultado_ins["id"],
                numero_pedido=datos.get("numero_pedido")
            )

        else:

            resultado_ins = supa.insertar_factura(
                datos,
                "facturas_almacen",
                archivo_url=archivo_url
            )

            log("Factura insertada en ALMACÉN")

            supa.marcar_albaranes_facturados(
                numeros_albaran_detectados,
                datos.get("nif"),
                "facturas_almacen",
                resultado_ins["id"],
                numero_pedido=datos.get("numero_pedido")
            )

    else:

        log("Factura incompleta o proveedor no encontrado → PENDIENTES")

        supa.insertar_factura(
            {
                "remitente": remitente,
                "asunto": asunto,
                "clasificado": False,
                "fecha_recibido": fecha_recibido_correo
            },
            "facturas_pendientes",
            archivo_url=archivo_url
        )


def procesar_correo(correo, tipo_forzado, pdf, ia, supa):

    # Se leen TODAS las propiedades del correo que hagan falta al principio,
    # antes de cualquier llamada de red (Supabase, IA...). Si se leyeran
    # más tarde, tras varios segundos de espera de una subida a Storage,
    # la referencia COM al correo puede quedar "desconectada"
    # (CO_E_OBJNOTCONNECTED / -2147220995) y la lectura falla.
    try:
        asunto = correo.Subject
        remitente = correo.SenderEmailAddress
        entry_id = correo.EntryID
        fecha_recibido_iso = correo.ReceivedTime.isoformat()
    except Exception as e:
        log(f"ERROR leyendo propiedades básicas del correo, se omite: {e}")
        return None, None

    log("--------------------------------------")
    log(f"Asunto: {asunto}")
    log(f"Remitente: {remitente}")
    log(f"Adjuntos: {correo.Attachments.Count}")

    try:
        pdfs = guardar_adjuntos_pdf(correo)
    except Exception as e:
        log(f"ERROR guardando adjuntos del correo, se omite este correo: {e}")
        return entry_id, fecha_recibido_iso

    try:
        fecha_recibido_correo = correo.ReceivedTime.strftime("%Y-%m-%d")
    except Exception as e:
        log(f"AVISO: no se pudo leer ReceivedTime del correo, se deja fecha_recibido vacía: {e}")
        fecha_recibido_correo = None

    if not pdfs:
        log("No hay PDFs adjuntos.")
        return entry_id, fecha_recibido_iso

    for ruta_pdf in pdfs:

        try:

            log(f"Procesando: {ruta_pdf}")

            datos = pdf.extraer_parcial(ruta_pdf)

            log(f"Datos extraídos (regex): {datos}")

            datos["fecha_recibido"] = fecha_recibido_correo

            tipo_detectado = pdf.tipo_documento(datos.get("texto", ""))

            if tipo_detectado is None:

                log(
                    f"'{ruta_pdf}' no parece factura ni albarán (buzón "
                    f"'{tipo_forzado}') → se manda a PENDIENTES para revisión manual"
                )

                archivo_url_pend = None

                try:
                    archivo_url_pend = supa.subir_archivo(ruta_pdf)
                except Exception as e:
                    log(f"ERROR subiendo PDF no identificado a Storage: {e}")

                supa.insertar_factura(
                    {
                        "remitente": remitente,
                        "asunto": asunto,
                        "clasificado": False,
                        "fecha_recibido": fecha_recibido_correo
                    },
                    "facturas_pendientes",
                    archivo_url=archivo_url_pend
                )

                continue

            if tipo_detectado != tipo_forzado:
                log(
                    f"AVISO: el contenido de '{ruta_pdf}' parece '{tipo_detectado}' "
                    f"pero llegó al buzón de '{tipo_forzado}' — se procesa igualmente "
                    f"como '{tipo_forzado}' (revisar si está en el buzón equivocado)"
                )

            if tipo_forzado == "albaran":
                procesar_pdf_albaran(ruta_pdf, datos, pdf, ia, supa)
            else:
                procesar_pdf_factura(ruta_pdf, datos, remitente, asunto, fecha_recibido_correo, pdf, ia, supa)

        except Exception as e:

            log(f"ERROR procesando PDF: {e}")

    return entry_id, fecha_recibido_iso


def conectar_con_reintentos(correo, intentos=10, espera=10):
    """
    Igual que conectar(), pero reintentando en vez de morir si Outlook
    todavía no está listo (p.ej. justo al iniciar sesión, cuando la tarea
    programada arranca antes de que Outlook haya terminado de abrirse).
    """

    for i in range(intentos):
        try:
            return conectar(correo)
        except Exception as e:
            log(f"No se pudo conectar a '{correo}' (intento {i + 1}/{intentos}): {e}. Reintentando en {espera}s...")
            time.sleep(espera)

    raise RuntimeError(f"No se pudo conectar a '{correo}' tras {intentos} intentos")


_lock_handle = None  # se rellena en asegurar_instancia_unica(); hay que mantener la referencia viva

log("========================================")
log("FacturaSync iniciado")

asegurar_instancia_unica()

carpeta_facturas = conectar_con_reintentos(CORREO_FACTURAS)
log(f"CARPETA FACTURAS CONECTADA ({CORREO_FACTURAS}): {carpeta_facturas.Name}")
log(f"CORREOS EN INBOX FACTURAS: {carpeta_facturas.Items.Count}")

carpeta_albaranes = conectar_con_reintentos(CORREO_ALBARANES)
log(f"CARPETA ALBARANES CONECTADA ({CORREO_ALBARANES}): {carpeta_albaranes.Name}")
log(f"CORREOS EN INBOX ALBARANES: {carpeta_albaranes.Items.Count}")

pdf = PDFService()
ia = AIService()
supa = SupabaseService()
recordatorios = RecordatoriosService()

ultimo_check_recordatorios = 0

while True:

    try:

        ahora = time.time()
        if ahora - ultimo_check_recordatorios >= INTERVALO_RECORDATORIOS:
            recordatorios.revisar_y_enviar()
            ultimo_check_recordatorios = ahora

        estado = cargar_estado()

        # ── FACTURAS (facturas@olivillatres.com) ──
        nuevos_facturas = obtener_nuevos(carpeta_facturas, estado["ultima_fecha_facturas"])

        for correo in nuevos_facturas:

            try:
                entry_id, fecha_recibido_iso = procesar_correo(correo, "factura", pdf, ia, supa)
            except Exception as e:
                log(f"ERROR procesando correo de facturas, se omite y se continúa con el siguiente: {e}")
                entry_id, fecha_recibido_iso = None, None

            # Se avanza el cursor aunque el correo haya fallado (siempre que
            # se haya podido leer su fecha), para que un correo problemático
            # no bloquee para siempre a los que llegan detrás.
            if fecha_recibido_iso:
                estado["ultima_fecha_facturas"] = fecha_recibido_iso
                guardar_estado(estado)

        # ── ALBARANES (almacen@olivillatres.com) ──
        nuevos_albaranes = obtener_nuevos(carpeta_albaranes, estado["ultima_fecha_albaranes"])

        for correo in nuevos_albaranes:

            try:
                entry_id, fecha_recibido_iso = procesar_correo(correo, "albaran", pdf, ia, supa)
            except Exception as e:
                log(f"ERROR procesando correo de albaranes, se omite y se continúa con el siguiente: {e}")
                entry_id, fecha_recibido_iso = None, None

            if fecha_recibido_iso:
                estado["ultima_fecha_albaranes"] = fecha_recibido_iso
                guardar_estado(estado)

    except Exception as e:

        log(f"ERROR GENERAL: {e}")

        codigo = e.args[0] if getattr(e, "args", None) else None

        if codigo in (-2147023174, -2147352567, -2147417848, -2147220995):

            log("Conexión con Outlook perdida. Reintentando reconectar...")

            reconectado = False

            while not reconectado:

                time.sleep(INTERVALO)

                try:
                    carpeta_facturas = conectar_con_reintentos(CORREO_FACTURAS, intentos=1)
                    carpeta_albaranes = conectar_con_reintentos(CORREO_ALBARANES, intentos=1)
                    log("Reconectado a Outlook (facturas y albaranes).")
                    reconectado = True

                except Exception as e2:
                    log(f"Outlook sigue sin estar disponible: {e2}")

    time.sleep(INTERVALO)