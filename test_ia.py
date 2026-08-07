from pdf_service import PDFService
from ia_service import AIService
import json

pdf = PDFService()
ai = AIService()

texto = pdf.leer_texto("temp/_7CB0YZLNL.pdf")

datos = ai.analizar_factura(texto)

print(json.dumps(datos, indent=4, ensure_ascii=False))