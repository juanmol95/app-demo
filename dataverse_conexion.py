"""
CONEXIÓN CON MICROSOFT DATAVERSE
================================
El servidor se conecta a Dataverse con la cuenta de Microsoft de quien inicia sesión
(modo prototipo: no usa una aplicación registrada propia).

La primera vez:
    python dataverse_conexion.py
Se abre el navegador, la persona inicia sesión con su cuenta de Microsoft y la sesión
queda guardada cifrada en este computador. Después el servidor la usa sin pedirla
de nuevo hasta que caduque; entonces hay que volver a ejecutar este archivo.

Se puede cambiar el entorno con las variables de entorno DATAVERSE_URL y DATAVERSE_TENANT_ID.
"""

import os
from pathlib import Path

import requests
from azure.identity import AuthenticationRecord, InteractiveBrowserCredential, TokenCachePersistenceOptions

DATAVERSE_URL = os.environ.get("DATAVERSE_URL", "https://orgc950bd5e.crm2.dynamics.com").rstrip("/")
TENANT_ID = os.environ.get("DATAVERSE_TENANT_ID", "fabd047c-ff48-492a-8bbb-8f98b9fb9cca")   # Universidad de los Andes
API = DATAVERSE_URL + "/api/data/v9.2/"
ALCANCE = DATAVERSE_URL + "/.default"

CARPETA = Path(__file__).resolve().parent
ARCHIVO_SESION = CARPETA / "dataverse_sesion.json"     # quién inició sesión (no contiene contraseñas ni tokens)
CACHE = TokenCachePersistenceOptions(name="registro_implantes_dataverse")   # tokens cifrados por Windows

_credencial = None


def credencial():
    global _credencial
    if _credencial is None:
        registro = None
        if ARCHIVO_SESION.is_file():
            registro = AuthenticationRecord.deserialize(ARCHIVO_SESION.read_text(encoding="utf-8"))
        _credencial = InteractiveBrowserCredential(tenant_id=TENANT_ID, cache_persistence_options=CACHE,
                                                   authentication_record=registro)
    return _credencial


def iniciar_sesion():
    """Abre el navegador para iniciar sesión y guarda la sesión para las siguientes veces."""
    global _credencial
    _credencial = InteractiveBrowserCredential(tenant_id=TENANT_ID, cache_persistence_options=CACHE)
    registro = _credencial.authenticate(scopes=[ALCANCE])
    ARCHIVO_SESION.write_text(registro.serialize(), encoding="utf-8")
    return registro


def api(metodo, ruta, **kwargs):
    """Llama a la API web de Dataverse. Devuelve la respuesta de requests."""
    token = credencial().get_token(ALCANCE).token
    encabezados = {
        "Authorization": "Bearer " + token,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    encabezados.update(kwargs.pop("headers", {}))
    return requests.request(metodo, API + ruta, headers=encabezados, timeout=60, **kwargs)


if __name__ == "__main__":
    print("Entorno:", DATAVERSE_URL)
    print("Se abrirá el navegador: inicia sesión con tu cuenta de Microsoft...")
    registro = iniciar_sesion()
    print("Sesión iniciada como:", registro.username)
    r = api("GET", "WhoAmI")
    r.raise_for_status()
    datos = r.json()
    usuario = api("GET", f"systemusers({datos['UserId']})?$select=fullname").json()
    print("Conexión con Dataverse correcta. Usuario en el entorno:", usuario.get("fullname"))
