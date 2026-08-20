import os

import win32com.client
import pythoncom

from logger import log

# Valor de FileFormat de Word para exportar a PDF (wdFormatPDF).
WD_FORMAT_PDF = 17


def convertir_word_a_pdf(ruta_word):
    """
    Convierte un documento Word (.doc o .docx) a PDF usando el propio
    Microsoft Word instalado en este PC, por COM (mismo enfoque que ya se
    usa con Outlook en outlook_service.py). Requiere que Word esté
    instalado — no basta con tener solo Outlook, aunque normalmente vienen
    juntos en la suite Office.

    Devuelve la ruta del PDF generado (mismo sitio, mismo nombre, con
    extensión .pdf), o None si la conversión falla por cualquier motivo
    (para que quien llame pueda decidir qué hacer: omitir el adjunto,
    avisar en el log, etc.).
    """

    ruta_pdf = os.path.splitext(ruta_word)[0] + ".pdf"

    # Si ya existe un PDF con ese mismo nombre (poco probable, pero por si
    # acaso), no lo pisamos: generamos uno con sufijo para no perder nada.
    if os.path.exists(ruta_pdf):
        base, ext = os.path.splitext(ruta_pdf)
        ruta_pdf = f"{base}_conv{ext}"

    word = None
    doc = None

    # Necesario porque este código puede ejecutarse en un hilo/contexto
    # distinto al que inicializó COM para Outlook.
    pythoncom.CoInitialize()

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(ruta_word, ReadOnly=True)
        doc.SaveAs(ruta_pdf, FileFormat=WD_FORMAT_PDF)

        log(f"Word convertido a PDF: '{ruta_word}' -> '{ruta_pdf}'")

        return ruta_pdf

    except Exception as e:
        log(f"ERROR convirtiendo Word a PDF ('{ruta_word}'): {e}")
        return None

    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()