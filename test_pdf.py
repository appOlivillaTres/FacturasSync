# test_pdf.py
from pdf_service import PDFService
import json

pdf = PDFService()
datos = pdf.extraer_parcial("temp/FACTURA  NÂº F    567-25  OLIVILLA TRES SLU  JAULA.pdf")  # usa una factura real
print(json.dumps(datos, indent=4, ensure_ascii=False))