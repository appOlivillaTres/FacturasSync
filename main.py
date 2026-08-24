import json
import time

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


def cargar_estado():

    try:
        with open(ESTADO_FILE, encoding="utf8") as f:
            estado = json.load(f)
    except:
        estado = {}

    # Compatibilidad con el estado.json antiguo (un único buzón, clave
    # "ultimo_entryid"): si ya existe pero todavía no hay claves separadas
    # por buzón, lo reutilizamos como punto de partida de FACTURAS, para
    # no reprocesar de golpe todo el histórico al pasar a dos buzones.
    if "ultimo_entryid_facturas" not in estado:
        estado["ultimo_entryid_facturas"] = estado.get("ultimo_entryid", "")

    if "ultimo_entryid_albaranes" not in estado:
        estado["ultimo_entryid_albaranes"] = ""

    return estado


def guardar_estado(estado):

    with open(ESTADO_FILE, "w", encoding="utf8") as f:
        json.dump(estado, f, indent=4)


def obtener_nuevos(carpeta, ultimo_entryid):

    mensajes = carpeta.Items
    mensajes.Sort("[ReceivedTime]", True)

    nuevos = []

    for correo in mensajes:

        try:

            if correo.EntryID == ultimo_entryid:
                break

            nuevos.append(correo)

        except:
            pass

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


def procesar_pdf_factura(ruta_pdf, datos, correo, fecha_recibido_correo, pdf, ia, supa):
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
                "remitente": correo.SenderEmailAddress,
                "asunto": correo.Subject,
                "clasificado": False,
                "fecha_recibido": fecha_recibido_correo
            },
            "facturas_pendientes",
            archivo_url=archivo_url
        )


def procesar_correo(correo, tipo_forzado, pdf, ia, supa):
    """
    tipo_forzado: 'factura' o 'albaran', según el buzón del que viene el
    correo. Si el contenido del PDF no parece ni factura ni albarán de
    verdad (p.ej. un presupuesto/proforma colado en el buzón equivocado),
    se manda igualmente a PENDIENTES en vez de forzarlo — es la única red
    de seguridad que queda del antiguo tipo_documento().
    """

    log("--------------------------------------")
    log(f"Asunto: {correo.Subject}")
    log(f"Remitente: {correo.SenderEmailAddress}")
    log(f"Adjuntos: {correo.Attachments.Count}")

    pdfs = guardar_adjuntos_pdf(correo)

    try:
        fecha_recibido_correo = correo.ReceivedTime.strftime("%Y-%m-%d")
    except Exception as e:
        log(f"AVISO: no se pudo leer ReceivedTime del correo, se deja fecha_recibido vacía: {e}")
        fecha_recibido_correo = None

    if not pdfs:
        log("No hay PDFs adjuntos.")
        return

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
                        "remitente": correo.SenderEmailAddress,
                        "asunto": correo.Subject,
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
                procesar_pdf_factura(ruta_pdf, datos, correo, fecha_recibido_correo, pdf, ia, supa)

        except Exception as e:

            log(f"ERROR procesando PDF: {e}")


log("========================================")
log("FacturaSync iniciado")

carpeta_facturas = conectar(CORREO_FACTURAS)
log(f"CARPETA FACTURAS CONECTADA ({CORREO_FACTURAS}): {carpeta_facturas.Name}")
log(f"CORREOS EN INBOX FACTURAS: {carpeta_facturas.Items.Count}")

carpeta_albaranes = conectar(CORREO_ALBARANES)
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
        nuevos_facturas = obtener_nuevos(carpeta_facturas, estado["ultimo_entryid_facturas"])

        for correo in nuevos_facturas:
            procesar_correo(correo, "factura", pdf, ia, supa)
            estado["ultimo_entryid_facturas"] = correo.EntryID
            guardar_estado(estado)

        # ── ALBARANES (almacen@olivillatres.com) ──
        nuevos_albaranes = obtener_nuevos(carpeta_albaranes, estado["ultimo_entryid_albaranes"])

        for correo in nuevos_albaranes:
            procesar_correo(correo, "albaran", pdf, ia, supa)
            estado["ultimo_entryid_albaranes"] = correo.EntryID
            guardar_estado(estado)

    except Exception as e:

        log(f"ERROR GENERAL: {e}")

        # Errores típicos cuando Outlook se cierra, se reinicia o el PC
        # sale de suspensión y la conexión COM se queda "muerta":
        #   -2147023174  RPC_S_SERVER_UNAVAILABLE ("El servidor RPC no está disponible")
        #   -2147352567  Excepción COM genérica de Outlook caído
        #   -2147417848  "El objeto invocado se ha desconectado de sus clientes"
        codigo = e.args[0] if getattr(e, "args", None) else None

        if codigo in (-2147023174, -2147352567, -2147417848):

            log("Conexión con Outlook perdida. Reintentando reconectar...")

            reconectado = False

            while not reconectado:

                time.sleep(INTERVALO)

                try:
                    carpeta_facturas = conectar(CORREO_FACTURAS)
                    carpeta_albaranes = conectar(CORREO_ALBARANES)
                    log("Reconectado a Outlook (facturas y albaranes).")
                    reconectado = True

                except Exception as e2:
                    log(f"Outlook sigue sin estar disponible: {e2}")

    time.sleep(INTERVALO)