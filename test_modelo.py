from openai import OpenAI
from config import OPENAI_API_KEY
from ia_service import MODELO_FACTURAS

client = OpenAI(api_key=OPENAI_API_KEY)

print(f"Probando modelo configurado: {MODELO_FACTURAS!r}\n")

try:
    respuesta = client.chat.completions.create(
        model=MODELO_FACTURAS,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[{"role": "user", "content": 'Devuelve exactamente este JSON: {"ok": true}'}]
    )
    print("El modelo respondio correctamente:")
    print(respuesta.choices[0].message.content)
except Exception as e:
    print("ERROR con el modelo configurado:")
    print(repr(e))

print("\n" + "-" * 50)
print("Probando tambien con gpt-4.1 para comparar...\n")

try:
    respuesta = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0,
        response_format={"type": "json_object"},
        messages=[{"role": "user", "content": 'Devuelve exactamente este JSON: {"ok": true}'}]
    )
    print("gpt-4.1 respondio correctamente:")
    print(respuesta.choices[0].message.content)
except Exception as e:
    print("ERROR con gpt-4.1:")
    print(repr(e))
