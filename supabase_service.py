import requests
import re
import os
import time
import mimetypes
from datetime import datetime, timezone

from config import SUPABASE_URL
from config import SUPABASE_KEY
from config import NIF_CSG
from config import NOMBRE_CSG
from logger import log


# Quién debe revisar cada tabla y recibir el aviso por correo.
# TODO: falta el email de Juan David, rellenar aquí.
REVISORES = {
    "facturas_obra": [
        {"nombre": "Angel", "email": "angelperez@olivillatres.com"},
        {"nombre": "Juan David", "email": "TODO_EMAIL_JUAN_DAVID"},
        {"nombre": "Emilio", "email": "efernandez@olivillatres.com"},
        {"nombre": "Cynthia", "email": "cynthiagomez@olivillatres.com"},
    ],
    "facturas_almacen": [
        {"nombre": "Cynthia", "email": "cynthiagomez@olivillatres.com"},
        {"nombre": "Emilio", "email": "efernandez@olivillatres.com"},
    ],
    "facturas_csg": [
        {"nombre": "Angel", "email": "angelperez@olivillatres.com"},
        {"nombre": "Emilio", "email": "efernandez@olivillatres.com"},
        {"nombre": "Cynthia", "email": "cynthiagomez@olivillatres.com"},
    ],
    "ingresos":[
        {"nombre":"Emilio", "email": "efernandez@olivillatres.com"},
        {"nombre": "Angel", "email": "angelperez@olivillatres.com"},
    ]  
}


class SupabaseService:

    def __init__(self):

        self.headers = {
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json"
        }

    def es_csg(self, nif=None, nombre=None, texto=None):
        """
        Comprueba si una factura pertenece a CSG Reformas S.L., por NIF o por
        nombre de proveedor, para forzar su clasificación en facturas_csg
        sin pasar por la lógica normal de destino (obra/almacén).

        Además del NIF/nombre del proveedor detectado, se comprueba también
        en TODO el texto de la factura: hay casos en los que CSG Reformas no
        es quien emite la factura, sino el cliente/obra para la que se hizo
        el trabajo (p.ej. subcontratas facturando a nombre de la obra de
        CSG), y en esos casos el NIF/nombre de CSG no aparece como
        "proveedor" sino en otra parte del documento.
        """

        nif_csg_normalizado = re.sub(r"[^A-Z0-9]", "", NIF_CSG.upper())

        def _coincide_nif(valor):
            if not valor:
                return False
            return re.sub(r"[^A-Z0-9]", "", valor.upper()) == nif_csg_normalizado

        if _coincide_nif(nif):
            log(f"CSG detectado por NIF de proveedor ({nif})")
            return True

        if nombre and NOMBRE_CSG.upper() in self.limpiar_nombre(nombre):
            log(f"CSG detectado por nombre de proveedor ({nombre})")
            return True

        if texto:

            texto_norm = re.sub(r"[^A-Z0-9]", "", texto.upper())

            if nif_csg_normalizado and nif_csg_normalizado in texto_norm:
                log(f"CSG detectado por NIF dentro del texto de la factura ({NIF_CSG})")
                return True

            if NOMBRE_CSG.upper().replace(" ", "") in texto_norm:
                log(f"CSG detectado por nombre dentro del texto de la factura ({NOMBRE_CSG})")
                return True

        return False

    def limpiar_nombre(self, nombre):

        if not nombre:
            return ""

        nombre = nombre.upper()

        quitar = [
            "S.L.",
            "SL",
            "S.L",
            "S A",
            "S.A.",
            "S.A",
            "CB",
            "C.B.",
            "C.B",
            ".",
            ","
        ]

        for q in quitar:
            nombre = nombre.replace(q, "")

        nombre = re.sub(r"\s+", " ", nombre)

        return nombre.strip()

    def buscar_proveedor(self, nif=None, nombre=None):

        # 1º Buscar por NIF
        if nif:

            url = f"{SUPABASE_URL}/rest/v1/proveedores?nif=eq.{nif}&select=*"

            r = requests.get(url, headers=self.headers)

            r.raise_for_status()

            datos = r.json()

            if datos:
                return datos[0]

        # 2º Buscar por nombre
        if nombre:

            url = f"{SUPABASE_URL}/rest/v1/proveedores?select=*"

            r = requests.get(url, headers=self.headers)

            r.raise_for_status()

            proveedores = r.json()

            nombre = self.limpiar_nombre(nombre)

            nombre = self.normalizar_proveedor(nombre)

            for p in proveedores:

                if self.normalizar_proveedor(p["nombre"]) == nombre:
                    return p

        return None

    def subir_archivo(self, ruta_local, bucket="facturas-entrantes"):
        """
        Sube un fichero local al bucket de Supabase Storage y devuelve la
        URL pública, igual que hace subirArchivoStorage() en el panel web.
        """

        nombre_archivo = os.path.basename(ruta_local)
        nombre_limpio = re.sub(r"[^a-zA-Z0-9.\-_]", "_", nombre_archivo)
        ruta_storage = f"{int(time.time() * 1000)}_{nombre_limpio}"

        content_type = mimetypes.guess_type(ruta_local)[0] or "application/octet-stream"

        url = f"{SUPABASE_URL}/storage/v1/object/{bucket}/{ruta_storage}"

        with open(ruta_local, "rb") as f:
            r = requests.post(
                url,
                headers={
                    "apikey": SUPABASE_KEY,
                    "Authorization": f"Bearer {SUPABASE_KEY}",
                    "Content-Type": content_type
                },
                data=f
            )

        r.raise_for_status()

        return f"{SUPABASE_URL}/storage/v1/object/public/{bucket}/{ruta_storage}"

    def insertar_factura(self, datos, tabla, archivo_url=None):

        permitidos_facturas = {
            "numero_interno",
            "numero_factura",
            "fecha",
            "fecha_recibido", 
            "proveedor",
            "nif",
            "concepto",
            "metodo_pago",
            "iban",
            "total",
            "fecha_vencimiento",
            "numeros_albaran",

            "base1",
            "pct_iva1",
            "total_iva1",

            "pct_re1",
            "total_re1",

            "tasas",
            "descuento",

            "base2",
            "pct_iva2",
            "total_iva2",

            "base3",
            "pct_iva3",
            "total_iva3",

            "exento",

            "base_retencion",
            "pct_retencion",
            "total_retencion"
        }

        permitidos_pendientes = {
            "remitente",
            "asunto",
            "clasificado",
            "fecha_recibido"
        }

        permitidos = permitidos_pendientes if tabla == "facturas_pendientes" else permitidos_facturas

        factura = {}

        for k, v in datos.items():

            if k in permitidos:
                factura[k] = v

        if archivo_url:
            factura["archivo_url"] = archivo_url

        if tabla == "facturas_obra":
            factura["nombre_obra"] = None

        # ── COMPROBACIÓN DE DUPLICADOS (solo facturas, no facturas_pendientes) ──
        if tabla in ("facturas_almacen", "facturas_obra", "facturas_csg"):
            duplicado = self.buscar_duplicado(
                numero_factura=factura.get("numero_factura"),
                nif=factura.get("nif"),
                fecha=factura.get("fecha"),
                total=factura.get("total"),
            )
            if duplicado:
                log(
                    f"Factura duplicada detectada (ya existe en {duplicado['tabla']}, "
                    f"id={duplicado['id']}). No se inserta de nuevo. "
                    f"numero_factura={factura.get('numero_factura')}, nif={factura.get('nif')}"
                )

                self.registrar_duplicado(duplicado, factura, origen="python")

                raise Exception(
                    f"Factura duplicada: ya existe en {duplicado['tabla']} (id={duplicado['id']})"
                )

        url = f"{SUPABASE_URL}/rest/v1/{tabla}"

        r = requests.post(
            url,
            headers={
                **self.headers,
                "Prefer": "return=representation"
            },
            json=factura
        )

        if not r.ok:
            # Red de seguridad final contra duplicados: si dos facturas
            # idénticas llegan casi a la vez (dos ejecuciones de
            # FacturaSync, o este backend guardando en el mismo instante
            # en que alguien la mete a mano desde el panel web), el
            # buscar_duplicado() de arriba puede no haber visto todavía
            # la otra factura porque hace un SELECT antes de este INSERT.
            # La restricción única de la base de datos sí las bloquea
            # (código 23505 = unique_violation de Postgres), y aquí lo
            # tratamos igual que un duplicado detectado por el SELECT.
            if tabla in ("facturas_almacen", "facturas_obra", "facturas_csg") and "23505" in r.text:
                log(f"Factura duplicada detectada por la base de datos al insertar en {tabla} (carrera con otro proceso)")
                self.registrar_duplicado({"tabla": tabla, "id": None}, factura, origen="python")
                raise Exception(f"Factura duplicada: la base de datos rechazó la inserción en {tabla} (ya existía)")
            raise Exception(f"Supabase {r.status_code} en {tabla}: {r.text}")

        resultado = r.json()[0]

        if tabla in ("facturas_obra", "facturas_almacen", "facturas_csg"):

            try:
                self.notificar_revision(tabla, resultado)
            except Exception as e:
                log(f"ERROR notificando revisores de {tabla}: {e}")

        return resultado

    def notificar_revision(self, tabla, factura):
        """
        Avisa por correo a quienes tienen que revisar la factura recién
        insertada, usando la función 'clever-worker' ya existente.
        """

        revisores = REVISORES.get(tabla, [])

        if not revisores:
            return

        if tabla == "facturas_obra":
            tipo = "revision_obra"
        elif tabla == "facturas_csg":
            tipo = "revision_csg"
        else:
            tipo = "revision_almacen"

        datos_factura = {
            "proveedor": factura.get("proveedor"),
            "numero_factura": factura.get("numero_factura"),
            "fecha": factura.get("fecha"),
            "fecha_vencimiento": factura.get("fecha_vencimiento"),
            "metodo_pago": factura.get("metodo_pago"),
            "total": factura.get("total"),
            "nombre_obra": factura.get("nombre_obra"),
        }

        for revisor in revisores:

            if not revisor.get("email") or revisor["email"].startswith("TODO_"):
                log(f"Aviso NO enviado a {revisor.get('nombre')}: falta su email en REVISORES")
                continue

            try:

                resp = requests.post(
                    f"{SUPABASE_URL}/functions/v1/clever-worker",
                    headers={
                        "Content-Type": "application/json",
                        "Authorization": f"Bearer {SUPABASE_KEY}"
                    },
                    json={
                        "tipo": tipo,
                        "email": revisor["email"],
                        "nombre": revisor["nombre"],
                        "factura": datos_factura
                    },
                    timeout=15
                )

                if not resp.ok:
                    raise Exception(f"{resp.status_code}: {resp.text}")

                log(f"Aviso de revisión enviado a {revisor['nombre']} ({revisor['email']})")

            except Exception as e:
                log(f"ERROR enviando aviso a {revisor['email']}: {e}")

    def normalizar_proveedor(self, nombre):

        if not nombre:
            return ""

        nombre = nombre.upper()

        quitar = [
            "S.L.", "SL", "S.A.", "SA",
            "C.B.", "CB",
            ".", ","
        ]

        for q in quitar:
            nombre = nombre.replace(q, "")

        nombre = " ".join(nombre.split())

        return nombre.strip()

    def buscar_duplicado(self, numero_factura=None, nif=None, fecha=None, total=None):
        """
        Comprueba si una factura ya existe en facturas_almacen o facturas_obra,
        para no insertarla dos veces si llega reenviada por correo.
        """

        tablas = ["facturas_almacen", "facturas_obra", "facturas_csg"]

        # 1º intento: por número de factura (más fiable si existe)
        if numero_factura:
            for tabla in tablas:
                url = f"{SUPABASE_URL}/rest/v1/{tabla}?numero_factura=eq.{numero_factura}&select=id,numero_factura"
                r = requests.get(url, headers=self.headers)
                r.raise_for_status()
                datos = r.json()
                if datos:
                    return {"tabla": tabla, "id": datos[0]["id"]}

            # Si tenemos número de factura pero no ha coincidido con ninguna
            # ya guardada, NO caemos al respaldo por NIF+fecha+total: daría
            # falsos positivos con dos facturas distintas del mismo
            # proveedor el mismo día (p.ej. mismo importe recurrente).
            return None

        # 2º intento (respaldo): NIF + fecha + total, solo para facturas
        # que no traen numero_factura
        if nif and fecha and total is not None:
            for tabla in tablas:
                url = (
                    f"{SUPABASE_URL}/rest/v1/{tabla}"
                    f"?nif=eq.{nif}&fecha=eq.{fecha}&total=eq.{total}&select=id"
                )
                r = requests.get(url, headers=self.headers)
                r.raise_for_status()
                datos = r.json()
                if datos:
                    return {"tabla": tabla, "id": datos[0]["id"]}

        return None

    def registrar_duplicado(self, duplicado, datos_intento, origen="python"):
        """
        Deja constancia en 'duplicados_detectados' de que ha llegado una
        factura repetida, y marca la factura original como reenviada
        (veces_reenviada / ultimo_reenvio). Replica en el backend de Python
        lo que ya hacía registrarDuplicado() en el panel (index.html), para
        que las facturas duplicadas detectadas por el pipeline de Outlook
        también aparezcan en la pantalla "Duplicados" del panel.
        """

        try:
            payload = {
                "tabla_original": duplicado["tabla"],
                "id_original": duplicado.get("id"),
                "numero_factura": datos_intento.get("numero_factura"),
                "nif": datos_intento.get("nif"),
                "proveedor": datos_intento.get("proveedor"),
                "total": datos_intento.get("total"),
                "origen": origen,
                "archivo_url": datos_intento.get("archivo_url"),
            }

            r = requests.post(
                f"{SUPABASE_URL}/rest/v1/duplicados_detectados",
                headers={**self.headers, "Prefer": "return=representation"},
                json=payload
            )
            r.raise_for_status()

            # Si no sabemos qué fila ganó la carrera (duplicado detectado
            # por la propia base de datos al rechazar el INSERT, no por
            # nuestro SELECT previo), no hay id_original al que sumarle
            # veces_reenviada.
            if duplicado.get("id") is None:
                log("Duplicado registrado en 'duplicados_detectados' (origen=python, sin id de la fila original)")
                return

            r2 = requests.get(
                f"{SUPABASE_URL}/rest/v1/{duplicado['tabla']}"
                f"?id=eq.{duplicado['id']}&select=veces_reenviada",
                headers=self.headers
            )
            r2.raise_for_status()
            original = r2.json()
            veces_actual = (original[0].get("veces_reenviada") or 0) if original else 0

            r3 = requests.patch(
                f"{SUPABASE_URL}/rest/v1/{duplicado['tabla']}?id=eq.{duplicado['id']}",
                headers=self.headers,
                json={
                    "veces_reenviada": veces_actual + 1,
                    "ultimo_reenvio": datetime.now(timezone.utc).isoformat()
                }
            )
            r3.raise_for_status()

            log(f"Duplicado registrado en 'duplicados_detectados' (origen=python)")

        except Exception as e:
            log(f"ERROR registrando duplicado en 'duplicados_detectados': {e}")

    def insertar_albaran(self, datos, archivo_url=None):

        permitidos = {"numero_albaran", "fecha", "empresa", "nif", "total", "numero_pedido"}

        albaran = {k: v for k, v in (datos or {}).items() if k in permitidos}
        albaran.setdefault("facturado", False)

        if archivo_url:
            albaran["archivo_url"] = archivo_url

        # ── COMPROBACIÓN DE DUPLICADOS (mismo criterio que las facturas) ──
        duplicado = self.buscar_duplicado_albaran(
            numero_albaran=albaran.get("numero_albaran"),
            nif=albaran.get("nif"),
            fecha=albaran.get("fecha"),
            total=albaran.get("total"),
        )
        if duplicado:
            log(
                f"Albarán duplicado detectado (ya existe en {duplicado['tabla']}, "
                f"id={duplicado['id']}). No se inserta de nuevo. "
                f"numero_albaran={albaran.get('numero_albaran')}, nif={albaran.get('nif')}"
            )

            self.registrar_duplicado(
                duplicado,
                {
                    "numero_factura": albaran.get("numero_albaran"),
                    "nif": albaran.get("nif"),
                    "proveedor": albaran.get("empresa"),
                    "total": albaran.get("total"),
                    "archivo_url": archivo_url,
                },
                origen="python"
            )

            raise Exception(
                f"Albarán duplicado: ya existe en {duplicado['tabla']} (id={duplicado['id']})"
            )

        url = f"{SUPABASE_URL}/rest/v1/albaranes"

        r = requests.post(
            url,
            headers={**self.headers, "Prefer": "return=representation"},
            json=albaran
        )

        if not r.ok:
            if "23505" in r.text:
                log("Albarán duplicado detectado por la base de datos al insertar (carrera con otro proceso)")
                self.registrar_duplicado({"tabla": "albaranes", "id": None}, {
                    "numero_factura": albaran.get("numero_albaran"),
                    "nif": albaran.get("nif"),
                    "proveedor": albaran.get("empresa"),
                    "total": albaran.get("total"),
                    "archivo_url": archivo_url,
                }, origen="python")
                raise Exception("Albarán duplicado: la base de datos rechazó la inserción (ya existía)")
            raise Exception(f"Supabase {r.status_code} en albaranes: {r.text}")

        return r.json()[0]

    def buscar_duplicado_albaran(self, numero_albaran=None, nif=None, fecha=None, total=None):
        """
        Comprueba si un albarán ya existe en la tabla 'albaranes', para no
        insertarlo dos veces si llega reenviado por correo. Mismo criterio
        que buscar_duplicado() para facturas: primero por número, y si no
        hay número, por NIF + fecha + total.
        """

        # 1º intento: por número de albarán (más fiable si existe)
        if numero_albaran:
            url = (
                f"{SUPABASE_URL}/rest/v1/albaranes"
                f"?numero_albaran=eq.{numero_albaran}&select=id,numero_albaran"
            )
            r = requests.get(url, headers=self.headers)
            r.raise_for_status()
            datos = r.json()
            if datos:
                return {"tabla": "albaranes", "id": datos[0]["id"]}

            # Si tenemos número de albarán pero no ha coincidido con ninguno,
            # NO caemos al respaldo por NIF+fecha+total: eso daría falsos
            # positivos cuando el mismo proveedor manda dos albaranes
            # distintos el mismo día (mismo NIF, misma fecha, a veces mismo
            # total). El respaldo solo tiene sentido cuando no hay número
            # de albarán con el que comparar.
            return None

        # 2º intento (respaldo): NIF + fecha + total, solo para albaranes
        # que no traen numero_albaran
        if nif and fecha and total is not None:
            url = (
                f"{SUPABASE_URL}/rest/v1/albaranes"
                f"?nif=eq.{nif}&fecha=eq.{fecha}&total=eq.{total}&select=id"
            )
            r = requests.get(url, headers=self.headers)
            r.raise_for_status()
            datos = r.json()
            if datos:
                return {"tabla": "albaranes", "id": datos[0]["id"]}

        return None

    def marcar_albaranes_facturados(self, numeros_albaran, nif, tabla_factura, factura_id, numero_pedido=None):
        """
        Por cada número de albarán que aparece en una factura, busca el
        albarán correspondiente en la tabla 'albaranes' (priorizando el
        que coincida también en NIF, por si dos proveedores comparten
        numeración) y lo marca como facturado, enlazándolo con la factura.

        Además, si se indica un número de pedido, también se marcan como
        facturados los albaranes pendientes de ese mismo pedido. Esto cubre
        el caso de proveedores que envían varias facturas (parciales) de
        un mismo albarán, donde el número de albarán no siempre coincide
        exactamente entre factura y factura pero el número de pedido sí.
        """

        ids_ya_marcados = set()

        def _marcar_por_id(albaran_id, numero_mostrado):

            if albaran_id in ids_ya_marcados:
                return

            try:

                patch_url = f"{SUPABASE_URL}/rest/v1/albaranes?id=eq.{albaran_id}"

                rp = requests.patch(
                    patch_url,
                    headers=self.headers,
                    json={
                        "facturado": True,
                        "factura_tabla": tabla_factura,
                        "factura_id": factura_id
                    }
                )
                rp.raise_for_status()

                ids_ya_marcados.add(albaran_id)

                log(f"Albarán '{numero_mostrado}' marcado FACTURADO (factura {tabla_factura} id={factura_id})")

            except Exception as e:
                log(f"ERROR marcando albarán '{numero_mostrado}' como facturado: {e}")

        # 1º: por número de albarán, como hasta ahora
        for numero in (numeros_albaran or []):

            if not numero:
                continue

            try:

                url = (
                    f"{SUPABASE_URL}/rest/v1/albaranes"
                    f"?numero_albaran=eq.{requests.utils.quote(str(numero))}&select=id,nif"
                )

                r = requests.get(url, headers=self.headers)
                r.raise_for_status()
                candidatos = r.json()

                if not candidatos:
                    log(f"Albarán '{numero}' no encontrado en 'albaranes' (no se marca)")
                    continue

                elegido = candidatos[0]

                if nif:
                    for c in candidatos:
                        if c.get("nif") == nif:
                            elegido = c
                            break

                _marcar_por_id(elegido["id"], numero)

            except Exception as e:
                log(f"ERROR marcando albarán '{numero}' como facturado: {e}")

        # 2º: por número de pedido, para cubrir facturas parciales de un
        # mismo albarán que no repiten el número de albarán exacto. Solo
        # tocamos albaranes que sigan pendientes (facturado=false), para no
        # sobrescribir un albarán ya vinculado a otra factura del pedido.
        if numero_pedido:

            try:

                url = (
                    f"{SUPABASE_URL}/rest/v1/albaranes"
                    f"?numero_pedido=eq.{requests.utils.quote(str(numero_pedido))}"
                    f"&facturado=eq.false&select=id,nif,numero_albaran"
                )

                r = requests.get(url, headers=self.headers)
                r.raise_for_status()
                candidatos = r.json()

                if not candidatos:
                    log(f"Sin albaranes pendientes para el pedido '{numero_pedido}'")

                for c in candidatos:

                    if nif and c.get("nif") and c.get("nif") != nif:
                        continue

                    _marcar_por_id(c["id"], c.get("numero_albaran") or f"pedido {numero_pedido}")

            except Exception as e:
                log(f"ERROR marcando albaranes por número de pedido '{numero_pedido}': {e}")