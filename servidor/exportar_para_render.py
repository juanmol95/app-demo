"""
PREPARAR LOS DATOS SECRETOS PARA RENDER
=======================================
Uso (en este computador):
    python exportar_para_render.py

1. Se abre el navegador: inicia sesión con tu cuenta de Microsoft (Uniandes).
2. Se crean dos archivos en la carpeta  para_render  (junto a la carpeta servidor):
     DATAVERSE_SESION.txt  -> tu sesión de Dataverse (es una credencial: NO la compartas)
     PLANTILLA_B64.txt     -> la plantilla de Excel, en texto
3. En Render, en Environment, crea las variables DATAVERSE_SESION y PLANTILLA_B64
   y pega en cada una el contenido de su archivo.

La sesión caduca si no se usa por un tiempo (normalmente semanas). Si en Render aparece
"la sesión de Dataverse del servidor caducó", vuelve a ejecutar este archivo y actualiza
DATAVERSE_SESION en Render.
"""

import base64
from pathlib import Path

import msal
import requests

import dataverse_conexion as dv

CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "para_render"
PLANTILLA = Path(__file__).resolve().parent / "IMPLANT REPORTS - PLANTILLA UNIANDES_0685.xlsx"


def main():
    CARPETA_SALIDA.mkdir(exist_ok=True)

    print("Se abrirá el navegador: inicia sesión con tu cuenta de Microsoft...")
    cache = msal.SerializableTokenCache()
    app = msal.PublicClientApplication(dv.CLIENTE, authority=f"https://login.microsoftonline.com/{dv.TENANT_ID}",
                                       token_cache=cache)
    r = app.acquire_token_interactive([dv.ALCANCE], prompt="select_account")
    if "access_token" not in r:
        raise SystemExit("No se pudo iniciar sesión: " + str(r.get("error_description", r)))

    # Comprobar que la sesión sirve para Dataverse
    yo = requests.get(dv.API + "WhoAmI", headers={"Authorization": "Bearer " + r["access_token"],
                                                  "Accept": "application/json"}, timeout=60)
    yo.raise_for_status()
    print("Sesión de Dataverse correcta para:", r.get("id_token_claims", {}).get("preferred_username", "(cuenta)"))

    (CARPETA_SALIDA / "DATAVERSE_SESION.txt").write_text(
        base64.b64encode(cache.serialize().encode("utf-8")).decode("ascii"), encoding="ascii")
    print("Creado:", CARPETA_SALIDA / "DATAVERSE_SESION.txt", "(credencial: no la compartas)")

    if PLANTILLA.is_file():
        (CARPETA_SALIDA / "PLANTILLA_B64.txt").write_text(base64.b64encode(PLANTILLA.read_bytes()).decode("ascii"),
                                                          encoding="ascii")
        print("Creado:", CARPETA_SALIDA / "PLANTILLA_B64.txt")
    else:
        print("No encontré la plantilla en la carpeta servidor; el Excel no funcionará en Render.")


if __name__ == "__main__":
    main()
