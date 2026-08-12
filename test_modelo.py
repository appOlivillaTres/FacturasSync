"""
Script de diagnóstico: compara qué lee la IA de una factura real con el
modelo actual (gpt-5) frente al modelo que se usaba antes (gpt-4.1),
para confirmar si el cambio de modelo es la causa de las lecturas erróneas
(por ejemplo, la fecha metida como número de factura).

Uso (desde el PC donde corre FacturaSync, con el entorno virtual activado):
    python test_ia_diagnostico.py "temp/nombre_de_la_factura.pdf"
"""

import sys
import json

from openai import OpenAI
from config import OPENAI_API_KEY
from pdf_service import PDFService
from ia_service import AIService

if len(sys.argv) < 2:
    print('Uso: python test_ia_diagnostico.py "ruta/a/la_factura.pdf"')
    sys.exit(1)

ruta_pdf = sys.argv[1]

pdf_service = PDFService()
ai = AIService()

client = OpenAI(api_key=OPENAI_API_KEY)

# Extraemos el texto solo para tener contexto legible en consola (el envío
# real a la IA usa el PDF directo, igual que hace main.py)
texto = pdf_service.leer_texto(ruta_pdf)
print("=== Primeras líneas del texto extraído (para contexto) ===")
print("\n".join(texto.splitlines()[:15]))
print()

prompt = ai._prompt_factura()
contenido = ai._contenido_pdf(ruta_pdf, prompt, "factura.pdf")

for modelo in ["gpt-5", "gpt-4.1"]:
    print(f"\n=== Resultado con modelo: {modelo} ===")
    try:
        respuesta = client.chat.completions.create(
            model=modelo,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": contenido}],
        )
        datos = ai._parsear_json(respuesta)
        print(json.dumps(datos, indent=4, ensure_ascii=False))
        print(f"\n  numero_factura -> {datos.get('numero_factura')!r}")
        print(f"  fecha          -> {datos.get('fecha')!r}")
        print(f"  proveedor      -> {datos.get('proveedor')!r}")
        print(f"  nif            -> {datos.get('nif')!r}")
        print(f"  total          -> {datos.get('total')!r}")
    except Exception as e:
        print(f"  ERROR con el modelo '{modelo}': {e}")