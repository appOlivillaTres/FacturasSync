import json
import time

from outlook_service import conectar, guardar_adjuntos_pdf
from pdf_service import PDFService
from ia_service import AIService
from supabase_service import SupabaseService
from recordatorios_service import RecordatoriosService

from logger import log
from config import ESTADO_FILE, INTERVALO, INTERVALO_RECORDATORIOS


def cargar_estado():

    try:
        with open(ESTADO_FILE, encoding="utf8") as f:
            return json.load(f)

    except:
        return {"ultimo_entryid": ""}


def guardar_estado(entryid):

    with open(ESTADO_FILE, "w", encoding="utf8") as f:
        json.dump({"ultimo_entryid": entryid}, f, indent=4)


log("========================================")
log("FacturaSync iniciado")

carpeta = conectar()
log(f"CARPETA CONECTADA: {carpeta.Name}")
log(f"CORREOS EN INBOX: {carpeta.Items.Count}")

pdf = PDFService()
ia = AIService()
supa = SupabaseService()
recordatorios = RecordatoriosService()

ultimo_check_recordatorios = 0  # epoch en segundos; 0 fuerza la primera comprobación nada más arrancar

while True:

    try:

        # ── RECORDATORIOS DE PRESUPUESTOS (cada INTERVALO_RECORDATORIOS) ──
        ahora = time.time()
        if ahora - ultimo_check_recordatorios >= INTERVALO_RECORDATORIOS:
            recordatorios.revisar_y_enviar()
            ultimo_check_recordatorios = ahora

        estado = cargar_estado()

        mensajes = carpeta.Items
        mensajes.Sort("[ReceivedTime]", True)

        ultimo = estado["ultimo_entryid"]

        nuevos = []

        for correo in mensajes:

            try:

                if correo.EntryID == ultimo:
                    break

                nuevos.append(correo)

            except:
                pass

        nuevos.reverse()

        for correo in nuevos:

            log("--------------------------------------")
            log(f"Asunto: {correo.Subject}")
            log(f"Remitente: {correo.SenderEmailAddress}")
            log(f"Adjuntos: {correo.Attachments.Count}")

            pdfs = guardar_adjuntos_pdf(correo)

            # Fecha real en que llegó el correo a Outlook (se usa para "fecha_recibido")
            try:
                fecha_recibido_correo = correo.ReceivedTime.strftime("%Y-%m-%d")
            except Exception as e:
                log(f"AVISO: no se pudo leer ReceivedTime del correo, se deja fecha_recibido vacía: {e}")
                fecha_recibido_correo = None

            if not pdfs:

                log("No hay PDFs adjuntos.")

                guardar_estado(correo.EntryID)

                continue

            for ruta_pdf in pdfs:

                try:

                    log(f"Procesando: {ruta_pdf}")

                    datos = pdf.extraer_parcial(ruta_pdf)

                    log(f"Datos extraídos (regex): {datos}")

                    datos["fecha_recibido"] = fecha_recibido_correo

                    tipo = pdf.tipo_documento(datos.get("texto", ""))

                    if tipo is None:

                        log(f"'{ruta_pdf}' no parece factura ni albarán → se manda a PENDIENTES para revisión manual")

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
                                "fecha_recibido": fecha_recibido_correo   # ← AÑADIR
                            },
                            "facturas_pendientes",
                            archivo_url=archivo_url_pend
                        )

                        continue

                    if tipo == "albaran":

                        log(f"'{ruta_pdf}' identificado como ALBARÁN")

                        try:
                            imagenes_pag = datos.get("imagenes_paginas") or []

                            if imagenes_pag:
                                log(
                                    f"Texto insuficiente incluso tras OCR — usando IA con visión "
                                    f"directamente sobre la imagen ({len(imagenes_pag)} página/s)"
                                )
                                datos_albaran = ia.analizar_albaran_imagen(imagenes_pag)
                            else:
                                # Vía principal: mandar el PDF real a la IA (ve la
                                # maquetación tal cual, no el texto ya aplanado por
                                # pdfplumber/OCR). Si falla por lo que sea (red,
                                # PDF corrupto, demasiado grande...), caemos al
                                # texto ya extraído como respaldo.
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

                        # Algunos proveedores (p.ej. Sika) tienen un "Su
                        # Pedido (online Shop): alias_cliente" que la IA a
                        # veces confunde con el número de pedido real. Un
                        # pedido real siempre trae algún dígito.
                        if datos_albaran.get("numero_pedido") and not any(
                            c.isdigit() for c in str(datos_albaran["numero_pedido"])
                        ):
                            log(
                                f"numero_pedido de albarán descartado por no ser numérico: "
                                f"'{datos_albaran['numero_pedido']}'"
                            )
                            datos_albaran["numero_pedido"] = None

                        # Respaldo por regex si la IA no detectó el pedido
                        # (o si se acaba de descartar por no ser numérico)
                        if not datos_albaran.get("numero_pedido"):
                            datos_albaran["numero_pedido"] = pdf.buscar_numero_pedido(datos.get("texto", ""))

                        # El NIF que lee la IA directamente del PDF puede
                        # venir mal (OCR confuso, formato distinto, un NIF
                        # de otra empresa mencionado de refilón, etc.). Si
                        # ya tenemos este proveedor dado de alta en la tabla
                        # 'proveedores' (porque nos ha facturado antes),
                        # usamos SIEMPRE su NIF registrado en vez del leído
                        # por la IA, que es más fiable. El NIF de la IA solo
                        # se usa si el proveedor no está dado de alta.
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

                        continue  # este PDF ya está gestionado, no sigue como factura

                    # La IA analiza SIEMPRE el texto completo, ya que el regex
                    # solo cubre nif/numero_factura/fecha/base/iva/total y deja
                    # vacíos el resto de campos del excel (base2, retenciones, etc.)
                    try:

                        imagenes_pag = datos.get("imagenes_paginas") or []

                        if imagenes_pag:
                            log(
                                f"Texto insuficiente incluso tras OCR — usando IA con visión "
                                f"directamente sobre la imagen ({len(imagenes_pag)} página/s)"
                            )
                            datos_ia = ia.analizar_factura_imagen(imagenes_pag)
                        else:
                            # Vía principal: mandar el PDF real a la IA (ve la
                            # maquetación tal cual, no el texto ya aplanado por
                            # pdfplumber/OCR — esto es lo que más precisión da).
                            # Si falla, caemos al texto ya extraído como respaldo.
                            try:
                                datos_ia = ia.analizar_factura_pdf(ruta_pdf)
                            except Exception as e_pdf:
                                log(f"AVISO: fallo mandando PDF de factura directo a la IA ({e_pdf}), usando texto extraído como respaldo")
                                datos_ia = ia.analizar_factura(datos.get("texto", ""))

                        log(f"Datos extraídos (IA): {datos_ia}")

                        # La IA manda siempre que tenga un valor: lee la factura
                        # con contexto completo, mientras que el regex es frágil
                        # con tablas desordenadas o texto rotado (como en la
                        # factura de la jaula, donde el regex confundió el
                        # importe neto con el total). El regex solo se usa como
                        # respaldo si la IA no devuelve nada para ese campo.
                        for k, v in datos_ia.items():

                            if v not in (None, "", 0):
                                datos[k] = v
                            elif k not in datos:
                                datos[k] = v

                    except Exception as e:

                        log(f"ERROR usando IA: {e}")

                    # Postgres/Supabase necesita la fecha en formato ISO
                    # (YYYY-MM-DD); el regex y a veces la IA la devuelven
                    # en formato español (DD/MM/YYYY).
                    datos["fecha"] = pdf.normalizar_fecha(datos.get("fecha"))

                    if datos.get("fecha_vencimiento"):
                        datos["fecha_vencimiento"] = pdf.normalizar_fecha(datos.get("fecha_vencimiento"))

                    # Algunos proveedores (p.ej. Sika) tienen un "Su Pedido
                    # (online Shop): alias_cliente" que la IA a veces
                    # confunde con el número de pedido real. Un pedido real
                    # siempre trae algún dígito.
                    if datos.get("numero_pedido") and not any(
                        c.isdigit() for c in str(datos["numero_pedido"])
                    ):
                        log(f"numero_pedido descartado por no ser numérico: '{datos['numero_pedido']}'")
                        datos["numero_pedido"] = None

                    # Respaldo por regex si la IA no detectó el pedido
                    # (o si se acaba de descartar por no ser numérico)
                    if not datos.get("numero_pedido"):
                        datos["numero_pedido"] = pdf.buscar_numero_pedido(datos.get("texto", ""))

                    # Validación del NIF por dígito de control: si el que ha
                    # devuelto la IA no es válido (típico si el documento
                    # viene borroso o con letra pequeña), probamos con el
                    # que encuentra el regex en el texto extraído — y si ese
                    # sí es válido, lo usamos en su lugar.
                    if datos.get("nif") and not pdf.nif_valido(datos["nif"]):

                        nif_regex = pdf.buscar_nif(datos.get("texto", ""))

                        if nif_regex and pdf.nif_valido(nif_regex) and nif_regex != datos["nif"]:
                            log(f"NIF '{datos['nif']}' (IA) no pasa el dígito de control, se usa '{nif_regex}' (regex) en su lugar")
                            datos["nif"] = nif_regex
                        else:
                            log(f"AVISO: el NIF '{datos['nif']}' no pasa el dígito de control y no hay alternativa mejor — revisar manualmente")

                    # Comprobación de cuadre: la suma de bases + IVA (y demás
                    # conceptos) debería coincidir con el total. Si no cuadra,
                    # es una señal fuerte de que la IA se ha equivocado en
                    # algún importe (p.ej. total vs subtotal). No bloqueamos
                    # la inserción por esto (preferimos tener el dato, aunque
                    # sea sospechoso, a perderlo), pero lo dejamos bien
                    # visible en el log para poder revisarlo.
                    cuadra, diferencia = pdf.validar_cuadre(datos)
                    if not cuadra:
                        log(f"AVISO: la factura no cuadra (base+IVA vs total difieren en {diferencia}€) — revisar '{ruta_pdf}'")

                    # Calculamos los números de albarán UNA sola vez (antes se
                    # recalculaban en cada rama de destino) y los guardamos en
                    # 'numeros_albaran' (mismo nombre y formato de array que
                    # espera el panel index.html) para que también queden
                    # persistidos en la propia factura, no solo usados de
                    # forma interna para marcar la tabla albaranes.
                    numeros_albaran_detectados = (
                        datos.get("numeros_albaran") or pdf.buscar_numeros_albaran(datos.get("texto", ""))
                    )

                    datos["numeros_albaran"] = numeros_albaran_detectados

                    # Subir el PDF a Supabase Storage para que el panel pueda
                    # mostrarlo (una ruta local de este PC no sirve como archivo_url)
                    archivo_url = None

                    try:

                        archivo_url = supa.subir_archivo(ruta_pdf)

                    except Exception as e:

                        log(f"ERROR subiendo PDF a Storage: {e}")

                    # Buscar proveedor SOLO una vez completados los datos
                    proveedor = supa.buscar_proveedor(
                        nif=datos.get("nif"),
                        nombre=datos.get("proveedor")
                    )

                    # Las facturas de CSG Reformas S.L. (por NIF o por nombre)
                    # van SIEMPRE a facturas_csg, sin mirar el destino normal
                    # configurado en la tabla proveedores.
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

                except Exception as e:

                    log(f"ERROR procesando PDF: {e}")

            guardar_estado(correo.EntryID)

    except Exception as e:

        log(f"ERROR GENERAL: {e}")

        # Errores típicos cuando Outlook se cierra, se reinicia o el PC
        # sale de suspensión y la conexión COM se queda "muerta":
        #   -2147023174  RPC_S_SERVER_UNAVAILABLE ("El servidor RPC no está disponible")
        #   -2147352567  Excepción COM genérica de Outlook caído
        #   -2147417848  "El objeto invocado se ha desconectado de sus clientes"
        # En estos casos NO basta con seguir el bucle: hay que reconectar
        # con Outlook, si no este mismo error se repetirá para siempre.
        codigo = e.args[0] if getattr(e, "args", None) else None

        if codigo in (-2147023174, -2147352567, -2147417848):

            log("Conexión con Outlook perdida. Reintentando reconectar...")

            reconectado = False

            while not reconectado:

                time.sleep(INTERVALO)

                try:
                    carpeta = conectar()
                    log(f"Reconectado a Outlook. CARPETA: {carpeta.Name}")
                    reconectado = True

                except Exception as e2:
                    log(f"Outlook sigue sin estar disponible: {e2}")

    time.sleep(INTERVALO)

# print("FacturaSync iniciado correctamente")


# import win32com.client

# from logger import log
# from config import CARPETA_OUTLOOK


# class OutlookService:

#     def conectar(self):

#         outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#         for i in range(1, outlook.Folders.Count + 1):

#             cuenta = outlook.Folders.Item(i)

#             for carpeta in cuenta.Folders:

#                 if carpeta.Name.upper() == CARPETA_OUTLOOK.upper():

#                     log(f"Carpeta encontrada: {carpeta.Name}")

#                     return carpeta

#         raise Exception(f"No existe la carpeta {CARPETA_OUTLOOK}")

# print("FacturaSync iniciado correctamente")


# import win32com.client

# from logger import log
# from config import CARPETA_OUTLOOK


# class OutlookService:

#     def conectar(self):

#         outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#         for i in range(1, outlook.Folders.Count + 1):

#             cuenta = outlook.Folders.Item(i)

#             for carpeta in cuenta.Folders:

#                 if carpeta.Name.upper() == CARPETA_OUTLOOK.upper():

#                     log(f"Carpeta encontrada: {carpeta.Name}")

#                     return carpeta

#         raise Exception(f"No existe la carpeta {CARPETA_OUTLOOK}")

# print("FacturaSync iniciado correctamente")


# import win32com.client

# from logger import log
# from config import CARPETA_OUTLOOK


# class OutlookService:

#     def conectar(self):

#         outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#         for i in range(1, outlook.Folders.Count + 1):

#             cuenta = outlook.Folders.Item(i)

#             for carpeta in cuenta.Folders:

#                 if carpeta.Name.upper() == CARPETA_OUTLOOK.upper():

#                     log(f"Carpeta encontrada: {carpeta.Name}")

#                     return carpeta

#         raise Exception(f"No existe la carpeta {CARPETA_OUTLOOK}")

# print("FacturaSync iniciado correctamente")


# import win32com.client

# from logger import log
# from config import CARPETA_OUTLOOK


# class OutlookService:

#     def conectar(self):

#         outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#         for i in range(1, outlook.Folders.Count + 1):

#             cuenta = outlook.Folders.Item(i)

#             for carpeta in cuenta.Folders:

#                 if carpeta.Name.upper() == CARPETA_OUTLOOK.upper():

#                     log(f"Carpeta encontrada: {carpeta.Name}")

#                     return carpeta

#         raise Exception(f"No existe la carpeta {CARPETA_OUTLOOK}")