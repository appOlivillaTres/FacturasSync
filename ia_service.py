import base64
import json

from openai import OpenAI

from config import OPENAI_API_KEY, NIF_PROPIO, NOMBRE_PROPIO
from logger import log


client = OpenAI(api_key=OPENAI_API_KEY)

# Antes se usaba gpt-4.1-mini. Subimos a gpt-4.1 (mismo formato de API,
# misma forma de leer archivos/imágenes) porque lee mejor tablas
# desalineadas, letra pequeña y facturas con mucho ruido — es la mejora
# de precisión más directa que se puede hacer sin cambiar de API.
MODELO_FACTURAS = "gpt-4.1"

# Tamaño máximo de PDF que se manda directamente a OpenAI como archivo.
# Por encima de esto, nos quedamos con el texto ya extraído (evita mandar
# documentos enormes innecesariamente).
MAX_PDF_BYTES_DIRECTO = 15 * 1024 * 1024  # 15 MB


class AIService:

    def _prompt_albaran(self):

        return f"""
Eres un experto en logística.

Analiza el siguiente ALBARÁN (nota de entrega) y devuelve ÚNICAMENTE un JSON válido.
No escribas explicaciones. No uses Markdown. No pongas ```json.

Reglas:
- Fechas en formato YYYY-MM-DD. Si no existe, usa null.
- Importes como número. Si no hay total, usa 0.
- "empresa" es el proveedor que EMITE el albarán, nunca nuestra empresa (NIF {NIF_PROPIO}, {NOMBRE_PROPIO}).
- "nif" es el NIF del proveedor emisor, nunca {NIF_PROPIO}.
- "numero_pedido" es el número de pedido de compra al que corresponde este albarán (puede aparecer como "Pedido", "Nº de pedido", "Su pedido", "Order No", etc.). Si no existe, usa null.

Estructura EXACTA:
{{
  "numero_albaran": null,
  "empresa": null,
  "nif": null,
  "fecha": null,
  "numero_pedido": null,
  "total": 0
}}
"""

    def analizar_albaran(self, texto):

        prompt = f"{self._prompt_albaran()}\n\nAlbarán:\n\n{texto}"

        respuesta = client.chat.completions.create(
            model=MODELO_FACTURAS,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": prompt}]
        )

        return self._parsear_json(respuesta)

    def analizar_albaran_pdf(self, ruta_pdf):
        """
        Manda el PDF ORIGINAL (no el texto ya extraído por pdfplumber/OCR)
        directamente al modelo. Es más preciso que analizar_albaran(texto)
        porque el modelo ve la maquetación real (tablas, columnas,
        posición de cada dato) en vez de una versión aplanada donde esa
        estructura ya se ha perdido — mismo enfoque que ya usa el panel web
        (swift-api.ts) con las facturas de clientes.
        """

        contenido = self._contenido_pdf(ruta_pdf, self._prompt_albaran(), "albaran.pdf")

        respuesta = client.chat.completions.create(
            model=MODELO_FACTURAS,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": contenido}]
        )

        return self._parsear_json(respuesta)

    def analizar_albaran_imagen(self, imagenes_base64):
        """
        Como analizar_albaran(), pero en vez de texto ya extraído (por
        regex/OCR, que en escaneos/fotos puede venir con errores), manda
        las páginas del documento directamente como imágenes a un modelo
        con visión — igual que ya hace el panel web (swift-api.ts). Se usa
        como último recurso cuando ni el texto normal ni el OCR con
        preprocesado consiguen sacar apenas dígitos (ver pdf_service).
        """

        contenido = [{"type": "text", "text": self._prompt_albaran()}]

        for img_b64 in imagenes_base64:
            contenido.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img_b64}"}
            })

        respuesta = client.chat.completions.create(
            model=MODELO_FACTURAS,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": contenido}]
        )

        return self._parsear_json(respuesta)

    def _prompt_factura(self):

        return f"""
Eres un experto en contabilidad española.

Analiza la siguiente FACTURA y devuelve ÚNICAMENTE un JSON válido.

No escribas explicaciones.
No utilices Markdown.
No pongas ```json.
Devuelve solo el objeto JSON.

Reglas:

- Fechas en formato YYYY-MM-DD.
- Importes como números (ejemplo: 1234.56).
- Si un dato no existe usa null.
- Si un importe no existe usa 0.
- No inventes información.
- El NIF/CIF "{NIF_PROPIO}" es el de nuestra propia empresa ({NOMBRE_PROPIO}). Casi todas las facturas lo mencionan porque somos el cliente/destinatario, NO el proveedor. Nunca devuelvas ese NIF como "nif": busca en el texto el NIF real de la empresa que EMITE la factura (el proveedor/vendedor), que es un NIF distinto.
- Para "metodo_pago", usa un valor corto y normalizado según lo que indique la factura: "Transferencia", "Domiciliación", "Giro", "Pagaré", "Cheque", "Tarjeta", "Contado" o "Confirming". Si no se especifica, usa null.
- Para "iban", extrae el número de cuenta bancaria del PROVEEDOR (para hacerle la transferencia), normalizado sin espacios (ejemplo: "ES6321007255500200078438"). Si la factura no incluye ningún IBAN/número de cuenta, usa null.
- Para "numeros_albaran", devuelve un array con TODOS los números de albarán (delivery note / packing list / nota de entrega) que se mencionen en la factura. Puede haber varios (a veces 5 o más). Si no hay ninguno, usa [].
- Para "numero_pedido", devuelve el número de pedido de compra al que corresponde la factura (puede aparecer como "Pedido", "Nº de pedido", "Su pedido", "Order No", etc.). A veces varias facturas distintas de un mismo pedido corresponden todas al mismo albarán, así que este dato es clave para relacionarlas aunque el número de albarán no coincida exactamente. Si no existe, usa null.

El JSON debe tener EXACTAMENTE esta estructura:

{{
  "proveedor": null,
  "nif": null,
  "numero_factura": null,
  "numero_interno": null,
  "fecha": null,
  "fecha_vencimiento": null,
  "concepto": null,
  "metodo_pago": null,
  "numero_pedido": null,

  "base1": 0,
  "pct_iva1": 0,
  "total_iva1": 0,

  "base2": 0,
  "pct_iva2": 0,
  "total_iva2": 0,

  "base3": 0,
  "pct_iva3": 0,
  "total_iva3": 0,

  "exento": 0,

  "pct_re1": 0,
  "total_re1": 0,

  "base_retencion": 0,
  "pct_retencion": 0,
  "total_retencion": 0,

  "tasas": 0,
  "descuento": 0,

  "total": 0,

  "iban": null,
  "numeros_albaran": []
}}
"""

    def analizar_factura(self, texto):

        prompt = f"{self._prompt_factura()}\n\nFactura:\n\n{texto}"

        respuesta = client.chat.completions.create(
            model=MODELO_FACTURAS,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": prompt}]
        )

        return self._parsear_json(respuesta)

    def analizar_factura_pdf(self, ruta_pdf):
        """
        Manda el PDF ORIGINAL (no el texto ya extraído por pdfplumber/OCR)
        directamente al modelo. Es más preciso que analizar_factura(texto)
        porque el modelo ve la maquetación real (tablas, columnas,
        posición de cada dato) en vez de una versión aplanada donde esa
        estructura ya se ha perdido — mismo enfoque que ya usa el panel web
        (swift-api.ts) con las facturas de clientes. Esta es ahora la vía
        PRINCIPAL de lectura (ver main.py); el texto extraído por regex
        queda solo como respaldo si esta llamada falla.
        """

        contenido = self._contenido_pdf(ruta_pdf, self._prompt_factura(), "factura.pdf")

        respuesta = client.chat.completions.create(
            model=MODELO_FACTURAS,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": contenido}]
        )

        return self._parsear_json(respuesta)

    def analizar_factura_imagen(self, imagenes_base64):
        """
        Como analizar_factura(), pero en vez de texto ya extraído (por
        regex/OCR, que en escaneos/fotos puede venir con errores), manda
        las páginas del documento directamente como imágenes a un modelo
        con visión — igual que ya hace el panel web (swift-api.ts). Se usa
        como último recurso cuando ni el texto normal ni el OCR con
        preprocesado consiguen sacar apenas dígitos (ver pdf_service).
        """

        contenido = [{"type": "text", "text": self._prompt_factura()}]

        for img_b64 in imagenes_base64:
            contenido.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img_b64}"}
            })

        respuesta = client.chat.completions.create(
            model=MODELO_FACTURAS,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": contenido}]
        )

        return self._parsear_json(respuesta)

    def _contenido_pdf(self, ruta_pdf, prompt, nombre_archivo):
        """
        Construye el bloque de contenido "file" (PDF en base64) + el texto
        del prompt, tal y como espera la Chat Completions API de OpenAI
        para modelos con soporte de documentos. Si el PDF pesa más de
        MAX_PDF_BYTES_DIRECTO, lanza un error para que la llamada caiga al
        respaldo por texto en main.py.
        """

        with open(ruta_pdf, "rb") as f:
            datos_pdf = f.read()

        if len(datos_pdf) > MAX_PDF_BYTES_DIRECTO:
            raise ValueError(
                f"PDF de {len(datos_pdf) / 1024 / 1024:.1f} MB supera el límite de "
                f"{MAX_PDF_BYTES_DIRECTO / 1024 / 1024:.0f} MB para envío directo a la IA"
            )

        pdf_base64 = base64.b64encode(datos_pdf).decode("utf-8")

        return [
            {
                "type": "file",
                "file": {
                    "filename": nombre_archivo,
                    "file_data": f"data:application/pdf;base64,{pdf_base64}"
                }
            },
            {"type": "text", "text": prompt}
        ]

    def _parsear_json(self, respuesta):

        contenido = respuesta.choices[0].message.content.strip()

        if contenido.startswith("```"):
            contenido = contenido.strip("`")
            contenido = contenido.replace("json\n", "", 1).strip()

        try:
            return json.loads(contenido)
        except json.JSONDecodeError as e:
            log(f"ERROR: la IA devolvió JSON no válido: {e} — contenido: {contenido[:300]}")
            raise
