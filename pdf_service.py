import pdfplumber
import pymupdf as fitz  # PyMuPDF
import re
import io
import base64

import pytesseract
from PIL import Image, ImageOps

from config import NIF_PROPIO, TESSERACT_CMD

if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


class PDFService:

    def __init__(self):
        # Páginas del último documento leído, en base64 (PNG). Solo se
        # rellena cuando ni el texto normal ni el OCR (con preprocesado)
        # consiguen sacar apenas dígitos: es la señal de que hace falta
        # mandar la imagen directamente a un modelo con visión como
        # último recurso (ver ia_service.analizar_factura_imagen /
        # analizar_albaran_imagen, y main.py).
        self.ultimas_imagenes_base64 = []

    def leer_texto(self, ruta_pdf):

        self.ultimas_imagenes_base64 = []

        texto = ""

        try:

            with pdfplumber.open(ruta_pdf) as pdf:

                for pagina in pdf.pages:

                    contenido = pagina.extract_text()

                    if contenido:
                        texto += contenido + "\n"

        except Exception as e:

            print("Error leyendo PDF:", e)

        # Algunos PDFs (p.ej. generados con Adobe LiveCycle/XFA, típico en
        # facturas de Sika) tienen los datos reales en una capa que
        # pdfplumber no lee bien: solo saca las etiquetas de la plantilla
        # vacía (que puede ser "larga" en caracteres pero sin apenas
        # números reales). Por eso no basta con mirar la longitud del
        # texto: contamos cuántos dígitos contiene, ya que una factura de
        # verdad está llena de números (importes, fechas, cantidades) y
        # una plantilla vacía casi no tiene ninguno.
        if sum(c.isdigit() for c in texto) < 10:

            try:

                texto_fitz = ""

                with fitz.open(ruta_pdf) as doc:

                    for pagina in doc:
                        texto_fitz += pagina.get_text() + "\n"

                if sum(c.isdigit() for c in texto_fitz) > sum(c.isdigit() for c in texto):
                    texto = texto_fitz

            except Exception as e:

                print("Error leyendo PDF con PyMuPDF:", e)

        # Último recurso: si seguimos sin apenas dígitos, es que el PDF no
        # tiene ninguna capa de texto real — es una foto o un escaneo del
        # papel (imagen incrustada), como pasa con facturas reenviadas por
        # correo hechas con el móvil. Ni pdfplumber ni PyMuPDF.get_text()
        # pueden leer texto de una imagen, así que aquí hace falta OCR de
        # verdad: renderizamos cada página como imagen y se la pasamos a
        # Tesseract.
        if sum(c.isdigit() for c in texto) < 10:

            try:

                texto_ocr = ""
                paginas_b64 = []

                with fitz.open(ruta_pdf) as doc:

                    for pagina in doc:

                        # zoom x4 (antes x3) para mejorar la resolución antes del OCR
                        pix = pagina.get_pixmap(matrix=fitz.Matrix(4, 4))
                        png_bytes = pix.tobytes("png")
                        imagen = Image.open(io.BytesIO(png_bytes))

                        # Preprocesado para Tesseract: escala de grises +
                        # autocontraste + binarizado. Ayuda bastante con
                        # fotos de móvil con sombras, mala iluminación o
                        # poco contraste (típico en escaneos reenviados).
                        imagen_gris = imagen.convert("L")
                        imagen_contraste = ImageOps.autocontrast(imagen_gris)
                        imagen_bin = imagen_contraste.point(lambda p: 255 if p > 160 else 0)

                        # --oem 1: motor LSTM (más preciso que el legacy).
                        # --psm 4: columnas de texto de ancho variable, va
                        # mejor que el psm 3 por defecto en tablas de factura.
                        texto_pagina = pytesseract.image_to_string(
                            imagen_bin, lang="spa", config="--oem 1 --psm 4"
                        )

                        # Si esta página en concreto sale casi sin dígitos,
                        # puede que --psm 4 (columnas) no encaje con este
                        # documento (p.ej. una factura de una sola columna
                        # ancha, o con el bloque de importes centrado).
                        # Probamos --psm 6 (bloque uniforme de texto) como
                        # segundo intento SOLO para esta página, y nos
                        # quedamos con el que saque más dígitos.
                        if sum(c.isdigit() for c in texto_pagina) < 5:

                            texto_pagina_alt = pytesseract.image_to_string(
                                imagen_bin, lang="spa", config="--oem 1 --psm 6"
                            )

                            if sum(c.isdigit() for c in texto_pagina_alt) > sum(c.isdigit() for c in texto_pagina):
                                texto_pagina = texto_pagina_alt

                        texto_ocr += texto_pagina + "\n"

                        # Guardamos la imagen ORIGINAL (sin binarizar) por si
                        # hace falta el respaldo de visión más abajo: un
                        # modelo de IA con visión interpreta mejor los tonos
                        # de gris/color que una imagen ya blanco y negro.
                        paginas_b64.append(base64.b64encode(png_bytes).decode("utf-8"))

                if sum(c.isdigit() for c in texto_ocr) > sum(c.isdigit() for c in texto):
                    texto = texto_ocr
                    print(f"Texto extraído por OCR (con preprocesado): {ruta_pdf}")

                # Si tras el OCR con preprocesado SEGUIMOS sin apenas
                # dígitos, es un escaneo/foto que Tesseract no puede leer
                # bien. Guardamos las páginas para que main.py intente el
                # respaldo de visión como último recurso.
                if sum(c.isdigit() for c in texto) < 10:
                    self.ultimas_imagenes_base64 = paginas_b64
                    print(f"OCR insuficiente, se guardan {len(paginas_b64)} página(s) para respaldo por visión: {ruta_pdf}")

            except Exception as e:

                print("Error leyendo PDF con OCR:", e)

        return self.limpiar_texto(texto)


    def normalizar_fecha(self, fecha):
        """
        Convierte una fecha en formato español (DD/MM/YYYY o DD-MM-YYYY) a
        ISO (YYYY-MM-DD), que es lo que espera la columna date de Supabase.
        Si ya viene en ISO, la deja igual. Si no reconoce el formato,
        devuelve None para no mandar una fecha inválida a la base de datos.
        """

        if not fecha:
            return None

        fecha = str(fecha).strip()

        if re.match(r"^\d{4}-\d{2}-\d{2}$", fecha):
            return fecha

        m = re.match(r"^(\d{2})[/-](\d{2})[/-](\d{4})$", fecha)

        if m:
            dia, mes, anio = m.groups()
            return f"{anio}-{mes}-{dia}"

        return None

    def tipo_documento(self, texto):
        """
        Devuelve 'factura', 'albaran' o None. Sustituye a parece_factura()
        para poder enrutar cada PDF a su tabla correcta.
        """

        if not texto:
            return None

        # Las condiciones generales de venta (letra pequeña al final de casi
        # cualquier documento comercial de proveedores como Sika) mencionan
        # palabras como "factura" en sentido legal/genérico (p.ej. "coste
        # incluido en la factura"), sin que eso signifique que el propio
        # documento sea una factura. Para no confundirnos, solo miramos el
        # encabezado real del documento, cortando el texto justo donde
        # empieza ese bloque de condiciones legales.
        marcadores_boilerplate = [
            "CONDICIONES GENERALES DE VENTA",
            "CONDICIONES GENERALES",
            "TÉRMINOS Y CONDICIONES",
            "TERMINOS Y CONDICIONES",
        ]

        t = texto.upper()

        for marcador in marcadores_boilerplate:
            idx = t.find(marcador)
            if idx != -1:
                t = t[:idx]
                break

        palabras_factura = [
            "FACTURA", "INVOICE", "FACTURE", "RECHNUNG",
            "NOTA DE ABONO", "NOTA DE CRÉDITO", "NOTA DE CREDITO"
        ]

        palabras_albaran = [
            "ALBARÁN", "ALBARAN", "DELIVERY NOTE", "PACKING LIST",
            "NOTA DE ENTREGA", "LIEFERSCHEIN"
        ]

        palabras_excluyentes = [
            "PRESUPUESTO", "PROFORMA", "CATÁLOGO", "CATALOGO", "CONTRATO"
        ]

        def _contiene_alguna(palabras):
            return any(re.search(rf"\b{re.escape(p)}\b", t) for p in palabras)

        tiene_factura = _contiene_alguna(palabras_factura)
        tiene_albaran = _contiene_alguna(palabras_albaran)

        if not tiene_factura and not tiene_albaran and _contiene_alguna(palabras_excluyentes):
            return None

        # Si menciona ambas cosas (habitual: una factura suele listar sus
        # albaranes de referencia), prevalece "factura", porque es el
        # documento fiscal real.
        if tiene_factura:
            return "factura"

        if tiene_albaran:
            return "albaran"

        return None


    def parece_factura(self, texto):
        return self.tipo_documento(texto) == "factura"


    def buscar_numeros_albaran(self, texto):
        """
        Busca TODOS los números de albarán referenciados dentro de una
        factura (respaldo por si la IA no los detecta). Puede haber varios.
        """

        if not texto:
            return []

        t = texto.upper()

        patrones = [
            r"ALBAR[ÁA]N(?:ES)?\s*N?[ºO°]?\.?\s*:?\s*([A-Z0-9][A-Z0-9\-/]{2,20})",
            r"\bALB\.?\s*N?[ºO°]?\.?\s*:?\s*([A-Z0-9][A-Z0-9\-/]{2,20})",
        ]

        encontrados = []

        for patron in patrones:
            for m in re.finditer(patron, t):

                numero = m.group(1).strip()

                # Un número de albarán real siempre trae algún dígito; si
                # no, probablemente hemos capturado la siguiente etiqueta
                # de una tabla (p.ej. "Albarán:\nTransporte:" sin valores
                # pegados, típico de algunos formatos de Sika), no un
                # número real.
                if numero and any(c.isdigit() for c in numero) and numero not in encontrados:
                    encontrados.append(numero)

        # Respaldo: algunos formatos (p.ej. Sika) listan las etiquetas
        # "Pedido:\nAlbarán:\n..." todas juntas al principio, y sus valores
        # reales aparecen más abajo, en el MISMO orden, como líneas sueltas
        # "NÚMERO de FECHA" (p.ej. "86550739 / de 06.07.2026"). Si no hemos
        # encontrado nada con los patrones normales, probamos esto: la
        # segunda línea de ese tipo suele ser el albarán (la primera es el
        # pedido).
        if not encontrados:

            m_lista = list(re.finditer(r"(\d{5,15})\s*/?\s*DE\s*\d{2}\.\d{2}\.\d{4}", t))

            idx_pedido = t.find("PEDIDO")
            idx_albaran = t.find("ALBAR")

            if len(m_lista) >= 2 and idx_pedido != -1 and idx_albaran != -1 and idx_pedido < idx_albaran:
                encontrados.append(m_lista[1].group(1))

        return encontrados

    def buscar_numero_pedido(self, texto):
        """
        Busca el número de pedido de compra (respaldo por si la IA no lo
        detecta). A diferencia de los albaranes, normalmente solo hay UN
        número de pedido por documento.
        """

        if not texto:
            return None

        t = texto.upper()

        patrones = [
            r"N[ÚU]MERO\s+DE\s+PEDIDO\s*:?\s*([A-Z0-9][A-Z0-9\-/]{2,20})",
            r"SU\s+PEDIDO\s*:?\s*([A-Z0-9][A-Z0-9\-/]{2,20})",
            r"\bPEDIDO\s*N?[ºO°]?\.?\s*:?\s*([A-Z0-9][A-Z0-9\-/]{2,20})",
            r"\bORDER\s*N?[ºO°]?\.?\s*:?\s*([A-Z0-9][A-Z0-9\-/]{2,20})",
        ]

        for patron in patrones:

            for m in re.finditer(patron, t):

                numero = m.group(1).strip()

                # Un número de pedido real siempre trae algún dígito; si no,
                # probablemente hemos capturado el nombre de una empresa u
                # otro texto que casualmente sigue a la palabra "Pedido"
                # (p.ej. "Pedido: Olivilla Tres S.L.U.").
                if numero and any(c.isdigit() for c in numero):
                    return numero

        # Respaldo: mismo formato de Sika descrito en buscar_numeros_albaran
        # (etiquetas "Pedido:\nAlbarán:\n..." juntas, valores reales más
        # abajo como líneas "NÚMERO de FECHA"). La primera de esas líneas
        # suele ser el pedido.
        m_lista = list(re.finditer(r"(\d{5,15})\s*/?\s*DE\s*\d{2}\.\d{2}\.\d{4}", t))

        idx_pedido = t.find("PEDIDO")
        idx_albaran = t.find("ALBAR")

        if m_lista and idx_pedido != -1 and idx_albaran != -1 and idx_pedido < idx_albaran:
            return m_lista[0].group(1)

        return None

    # Letras de control para el DNI/NIF de persona física (últimas 2 cifras
    # del NIE se tratan como si empezaran por X/Y/Z → 0/1/2).
    _LETRAS_DNI = "TRWAGMYFPDXBNJZSQVHLCKE"

    # Letra de control esperada para cada primera letra de CIF (empresas).
    # Las que dan solo número (grupo 1) o solo letra (grupo 2) las tratamos
    # aceptando cualquiera de las dos formas.
    _CIF_SOLO_LETRA = set("PQSW")
    _CIF_SOLO_NUMERO = set("ABEH")

    def nif_valido(self, nif):
        """
        Valida el dígito/letra de control de un NIF (persona física), NIE o
        CIF (empresa) español. Sirve para descartar cadenas que encajan con
        el patrón de un NIF por pura casualidad (ruido de OCR, un código de
        producto, un número de cuenta cortado, etc.) pero no son un NIF real.

        Devuelve True/False. Si el formato no se reconoce, devuelve False.
        """

        if not nif:
            return False

        n = re.sub(r"[^A-Z0-9]", "", str(nif).upper())

        # DNI: 8 dígitos + letra
        if re.fullmatch(r"\d{8}[A-Z]", n):
            return n[8] == self._LETRAS_DNI[int(n[:8]) % 23]

        # NIE: X/Y/Z + 7 dígitos + letra
        if re.fullmatch(r"[XYZ]\d{7}[A-Z]", n):
            numero = {"X": "0", "Y": "1", "Z": "2"}[n[0]] + n[1:8]
            return n[8] == self._LETRAS_DNI[int(numero) % 23]

        # CIF: letra + 7 dígitos + dígito o letra de control
        if re.fullmatch(r"[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-Z]", n):

            letra_inicial = n[0]
            digitos = n[1:8]
            control = n[8]

            suma_par = sum(int(d) for d in digitos[1::2])
            suma_impar = 0
            for d in digitos[0::2]:
                doble = int(d) * 2
                suma_impar += doble // 10 + doble % 10

            total = suma_par + suma_impar
            digito_control = (10 - (total % 10)) % 10

            if letra_inicial in self._CIF_SOLO_NUMERO:
                return control == str(digito_control)

            letra_control = "JABCDEFGHI"[digito_control]

            if letra_inicial in self._CIF_SOLO_LETRA:
                return control == letra_control

            # Resto de letras (empresas normales, tipo B): admiten ambas formas
            return control == str(digito_control) or control == letra_control

        return False

    def buscar_nif(self, texto):

        patron = r"\b[ABCDEFGHJNPQRSUVWXYZ]\d{7}[0-9A-Z]\b|\b\d{8}[A-Z]\b"

        candidatos = re.findall(patron, texto.upper())

        propio_norm = re.sub(r"[^A-Z0-9]", "", NIF_PROPIO.upper()) if NIF_PROPIO else None

        candidato_sin_validar = None

        for c in candidatos:

            # El NIF de nuestra propia empresa aparece en casi todas las
            # facturas porque somos el cliente, no quien la emite. Lo
            # saltamos y seguimos buscando el NIF real del proveedor.
            if propio_norm and re.sub(r"[^A-Z0-9]", "", c) == propio_norm:
                continue

            # Preferimos un candidato que pase el dígito de control (evita
            # confundir el NIF real con un código que casualmente encaja en
            # el patrón, típico en OCR de baja calidad). Si ninguno lo pasa,
            # nos quedamos con el primero como respaldo (mejor algo que nada).
            if self.nif_valido(c):
                return c

            if candidato_sin_validar is None:
                candidato_sin_validar = c

        return candidato_sin_validar


    def buscar_numero_factura(self, texto):

        patrones = [

            r"FACTURA\s*N[º°]?\s*FECHA\s*FACTURA\s*([A-Z0-9\s/\-]+?)\s+\d{2}[/-]\d{2}[/-]\d{4}",
            r"FACTURA\s*N[º°]\.?\s*([A-Z0-9/\-]+)",
            r"FACTURA[^\n]*\n\s*([A-Z0-9/\-]{2,15})\s+\d{2}[/-]\d{2}[/-]\d{4}",
            r"N[ÚU]MERO\s+FACTURA\s*[: ]+\s*([A-Z0-9/\-]+)",
            r"FACTURA\s*[: ]+\s*([A-Z0-9/\-]+)",
            r"INVOICE\s*[: ]+\s*([A-Z0-9/\-]+)"

        ]

        texto = texto.upper()

        invalidas = {"FECHA", "REFERENCIA", "VALOR", "Nº", "N"}

        for patron in patrones:

            for m in re.finditer(patron, texto, re.DOTALL):

                numero = m.group(1)

                numero = numero.replace("\n", " ")
                numero = " ".join(numero.split())
                numero = numero.replace(" / ", "/")

                if numero and numero not in invalidas:
                    return numero

        return None


    MESES = {
        "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
        "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
        "septiembre": "09", "setiembre": "09", "octubre": "10",
        "noviembre": "11", "diciembre": "12"
    }


    def buscar_fecha(self, texto):

        patron = r"\b\d{2}[/-]\d{2}[/-]\d{4}\b"

        m = re.search(patron, texto)

        if m:

            return m.group(0)

        # Fecha en formato textual: "31 de enero 2026" o "31 de enero de 2026"
        patron_textual = r"\b(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+(?:de\s+)?(\d{4})\b"

        m2 = re.search(patron_textual, texto, re.IGNORECASE)

        if m2:

            dia = m2.group(1).zfill(2)
            mes = self.MESES.get(m2.group(2).lower())
            anio = m2.group(3)

            if mes:
                return f"{dia}/{mes}/{anio}"

        return None


    def validar_cuadre(self, datos, tolerancia=0.5):
        """
        Comprueba que la suma de bases + cuotas de IVA/RE (menos retención,
        descuento, y más tasas) cuadra razonablemente con el total. Sirve
        para detectar automáticamente cuando la IA ha leído mal un importe
        (p.ej. ha confundido el subtotal con el total, o se ha dejado un
        tramo de IVA), en vez de insertar el dato en Supabase sin más.

        Devuelve (cuadra: bool, diferencia: float | None). Si faltan datos
        para comprobarlo (no hay total, por ejemplo), devuelve (True, None)
        — preferimos no bloquear nada cuando no hay suficiente información,
        no dar un falso positivo de "no cuadra".
        """

        total = datos.get("total")

        if not total:
            return True, None

        def num(clave):
            valor = datos.get(clave)
            try:
                return float(valor) if valor not in (None, "") else 0.0
            except (TypeError, ValueError):
                return 0.0

        suma = (
            num("base1") + num("total_iva1")
            + num("base2") + num("total_iva2")
            + num("base3") + num("total_iva3")
            + num("exento")
            + num("total_re1")
            + num("tasas")
            - num("total_retencion")
            - num("descuento")
        )

        # Si no hay ningún desglose informado (todo a 0), no hay nada que
        # comparar de forma fiable — no lo marcamos como descuadre.
        if suma == 0:
            return True, None

        diferencia = abs(suma - float(total))

        return diferencia <= tolerancia, round(diferencia, 2)

    def analizar(self, ruta_pdf):

        texto = self.leer_texto(ruta_pdf)

        base, iva, total = self.buscar_importes(texto)

        return {

            "texto": texto,

            "nif": self.buscar_nif(texto),

            "numero_factura": self.buscar_numero_factura(texto),

            "fecha": self.buscar_fecha(texto),

            "base": base,

            "iva": iva,

            "total": total,

            "imagenes_paginas": self.ultimas_imagenes_base64

        }


    def _euros_a_float(self, valor):
        valor = valor.replace(".", "").replace(",", ".")
        return float(valor)


    def buscar_importes(self, texto):

        texto = texto.replace("\n", " ")
        texto_up = texto.upper()

        # Formato 1: "base € (pct%) iva € total €" todo en una línea
        patron1 = r"([\d\.\,]+)\s*€\s*\(\d+[\.,]?\d*%\)\s*([\d\.\,]+)\s*€\s*([\d\.\,]+)\s*€"

        m = re.search(patron1, texto_up)

        if m:

            base = self._euros_a_float(m.group(1))
            iva = self._euros_a_float(m.group(2))
            total = self._euros_a_float(m.group(3))

            return base, iva, total

        base = None
        iva = None
        total = None

        # Formato 2: tabla "... BASE IMPONIBLE IMPORTE IVA IMPORTE R.E. % % 21,00% 2.480,00 520,80"
        patron2 = r"BASE\s+IMPONIBLE\s+IMPORTE\s+IVA\s+IMPORTE\s+R\.?E\.?\s*%?\s*%?\s*[\d\.,]+%\s*([\d\.,]+)\s*([\d\.,]+)"

        m = re.search(patron2, texto_up)

        if m:
            base = self._euros_a_float(m.group(1))
            iva = self._euros_a_float(m.group(2))

        # Formato 3: "SUBTOTAL 357,51€ ... IVA 21% 95,03€" (base e iva por separado)
        if base is None:

            m_base = re.search(r"\bSUBTOTAL\b[^€\d]{0,20}([\d\.,]+)\s*€", texto_up)

            if m_base:
                base = self._euros_a_float(m_base.group(1))

        if iva is None:

            m_iva = re.search(r"\bIVA\s*\d{1,2}[.,]?\d*\s*%[^€\d]{0,20}([\d\.,]+)\s*€", texto_up)

            if m_iva:
                iva = self._euros_a_float(m_iva.group(1))

        # Total: usamos límite de palabra (\b) para no confundir con SUBTOTAL,
        # ventana corta para no saltar hasta un importe lejano (p.ej. una cabecera
        # de columna "Cantidad Total"), y nos quedamos con la ÚLTIMA coincidencia,
        # ya que el total real suele ir al final de la factura.
        patron_total = r"\bTOTAL\b[^€]{0,100}?([\d\.,]+)\s*€"

        coincidencias = list(re.finditer(patron_total, texto_up))

        if coincidencias:
            total = self._euros_a_float(coincidencias[-1].group(1))

        return base, iva, total
    def limpiar_texto(self, texto):

        lineas = texto.splitlines()

        resultado = []

        for linea in lineas:

            linea = linea.strip()

            if not linea:
                continue

            if "TCPDF" in linea.upper():
                continue

            if linea.lower().startswith("pagina"):
                continue

            if len(linea) > 150:
                continue

            resultado.append(linea)

        return "\n".join(resultado)
    def extraer_parcial(self, ruta_pdf):

        texto = self.leer_texto(ruta_pdf)

        base, iva, total = self.buscar_importes(texto)

        return {
            "texto": texto,
            "nif": self.buscar_nif(texto),
            "numero_factura": self.buscar_numero_factura(texto),
            "fecha": self.buscar_fecha(texto),
            "base": base,
            "iva": iva,
            "total": total,
            "imagenes_paginas": self.ultimas_imagenes_base64
        }