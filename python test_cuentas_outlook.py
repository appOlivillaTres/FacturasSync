"""
Script de diagnostico. Ejecutalo en el PC donde corre FacturaSync
(con Outlook ya abierto) para ver exactamente que cuentas detecta
el codigo por COM y con que direccion de email exacta.

Uso:
    python test_cuentas_outlook.py
"""

import win32com.client

outlook_app = win32com.client.Dispatch("Outlook.Application")
namespace = outlook_app.GetNamespace("MAPI")

print("Cuentas que ve Outlook por codigo (win32com):\n")

if namespace.Accounts.Count == 0:
    print("  (ninguna cuenta detectada)")

for cuenta in namespace.Accounts:
    print(f"  - DisplayName: {cuenta.DisplayName!r}")
    print(f"    SmtpAddress:  {cuenta.SmtpAddress!r}")
    print(f"    AccountType:  {cuenta.AccountType}")
    print()

print("-" * 50)
buscada = "almacen@olivillatres.com"
encontrada = any(c.SmtpAddress.lower() == buscada.lower() for c in namespace.Accounts)
print(f"Buscando '{buscada}'... {'ENCONTRADA' if encontrada else 'NO ENCONTRADA'}")