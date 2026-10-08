from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import cv2
import numpy as np
import zxingcpp
import pytesseract
import uvicorn
import re
import datetime

# --- Agregado en la unificación ---
import json
import os
import sqlite3
import threading
from pathlib import Path
import openpyxl
from fastapi import Body
from fastapi.responses import FileResponse, Response
import io
import openpyxl.styles
import openpyxl.utils
import generador_reporte as generador   # el módulo que genera el Implant Report
import dataverse_registros as dvr       # guardar y consultar pacientes en Microsoft Dataverse

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def limpiar_texto(texto: str) -> str:
    """Limpia caracteres nulos o no alfabéticos dejando solo letras, tildes, Ñ y espacios."""
    if not texto:
        return ""
    # Permite letras de la A-Z, a-z, acentos, Ñ/ñ y espacios
    texto_limpio = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', texto)
    return texto_limpio.strip()

# -------------------------------------------------------------
# 1. MÉTODO CÉDULA TRADICIONAL (PDF417)
# -------------------------------------------------------------
def extraer_datos_pdf417(raw_text: str):
    """Extrae datos del código de barras PDF417 """
    fecha_raw = raw_text[152:160].strip() # AAAAMMDD
    
    # Formatear la fecha a AAAA-MM-DD
    fecha_formatted = fecha_raw
    if len(fecha_raw) == 8 and fecha_raw.isdigit():
        fecha_formatted = f"{fecha_raw[:4]}-{fecha_raw[4:6]}-{fecha_raw[6:]}"

    return {
        "documento": raw_text[48:58].lstrip("0"),
        "primer_apellido": limpiar_texto(raw_text[58:81]),
        "segundo_apellido": limpiar_texto(raw_text[81:104]),
        "primer_nombre": limpiar_texto(raw_text[104:127]),
        "segundo_nombre": limpiar_texto(raw_text[127:150]),
        "sexo": raw_text[151:152].strip(),
        "fecha_nacimiento": fecha_formatted,  # AAAA-MM-DD
        "rh": raw_text[166:168].strip()
    }

# -------------------------------------------------------------
# 2. MÉTODO CÉDULA DIGITAL (MRZ OCR)
# -------------------------------------------------------------
def parsear_mrz(lineas):
    """
    Parsea las 3 líneas ICAO 9303 (MRZ) del reverso de la cédula digital.
    """
    linea1 = lineas[0]
    linea2 = lineas[1]
    linea3 = lineas[2]

    # --- A. FECHA DE NACIMIENTO Y SEXO (Línea 2: pos 0-6 YYMMDD) ---
    fecha_nac_raw = linea2[0:6] 
    sexo = linea2[7:8] if len(linea2) > 7 else ""

    fecha_nac = ""
    if len(fecha_nac_raw) == 6 and fecha_nac_raw.isdigit():
        yy = int(fecha_nac_raw[:2])
        mm = fecha_nac_raw[2:4]
        dd = fecha_nac_raw[4:6]

        anio_actual_2d = datetime.date.today().year % 100
        prefix = "20" if yy <= anio_actual_2d else "19"

        # Formato AAAA-MM-DD
        fecha_nac = f"{prefix}{yy:02d}-{mm}-{dd}"

    # --- B. NÚMERO DE DOCUMENTO (Línea 2) ---
    mrz_sub = linea2[15:] if len(linea2) > 15 else linea2
    numeros_encontrados = re.findall(r'\d{7,10}', mrz_sub)
    
    if numeros_encontrados:
        documento = numeros_encontrados[0]
    else:
        documento = re.sub(r'[^0-9]', '', mrz_sub).lstrip("0")

    # --- C. NOMBRES Y APELLIDOS (Línea 3) ---
    partes = linea3.split("<<")
    apellidos_raw = partes[0] if len(partes) > 0 else ""
    nombres_raw = partes[1] if len(partes) > 1 else ""

    apellidos = [a.replace("<", " ").strip() for a in apellidos_raw.split("<") if a]
    nombres = [n.replace("<", " ").strip() for n in nombres_raw.split("<") if n]

    primer_apellido = apellidos[0] if len(apellidos) > 0 else ""
    segundo_apellido = apellidos[1] if len(apellidos) > 1 else ""
    primer_nombre = nombres[0] if len(nombres) > 0 else ""
    segundo_nombre = " ".join(nombres[1:]) if len(nombres) > 1 else ""

    return {
        "documento": documento,
        "primer_apellido": limpiar_texto(primer_apellido),
        "segundo_apellido": limpiar_texto(segundo_apellido),
        "primer_nombre": limpiar_texto(primer_nombre),
        "segundo_nombre": limpiar_texto(segundo_nombre),
        "sexo": sexo if sexo in ["M", "F"] else "",
        "fecha_nacimiento": fecha_nac,
        "rh": ""  # El RH no forma parte del estándar ICAO MRZ
    }

def intentar_ocr_mrz(img):
    """Ejecuta Tesseract OCR buscando la franja MRZ"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    config_mrz = r'--oem 3 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<'
    texto = pytesseract.image_to_string(gray, config=config_mrz)

    lineas_mrz = [l.strip().replace(" ", "") for l in texto.split("\n") if "<" in l and len(l.strip()) >= 15]

    if len(lineas_mrz) >= 3:
        return parsear_mrz(lineas_mrz[-3:])
    return None

# -------------------------------------------------------------
# 3. PROCESADOR UNIFICADO (PDF417 O MRZ)
# -------------------------------------------------------------
def procesar_imagen_cedula(img_bytes):
    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None:
        return None

    h, w = img.shape[:2]
    max_dim = 1600
    if max(h, w) > max_dim:
        scale = max_dim / float(max(h, w))
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)

    # --- PASO A: Wscanear código de barras PDF417 ---
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    contrast_img = clahe.apply(gray)
    _, thresh_img = cv2.threshold(contrast_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    variantes = [("Original", img), ("Grises", gray), ("Contraste", contrast_img), ("Binarizada", thresh_img)]

    for _, imagen_proc in variantes:
        for angulo in [0, 90, 180, 270]:
            if angulo == 90:
                img_rot = cv2.rotate(imagen_proc, cv2.ROTATE_90_CLOCKWISE)
            elif angulo == 180:
                img_rot = cv2.rotate(imagen_proc, cv2.ROTATE_180)
            elif angulo == 270:
                img_rot = cv2.rotate(imagen_proc, cv2.ROTATE_90_COUNTERCLOCKWISE)
            else:
                img_rot = imagen_proc

            resultados = zxingcpp.read_barcodes(img_rot, formats=[zxingcpp.BarcodeFormat.PDF417])

            if resultados:
                res = resultados[0]
                raw_text = res.bytes.decode("iso-8859-1", errors="ignore")
                if len(raw_text) >= 160:
                    datos = extraer_datos_pdf417(raw_text)
                    return {"tipo_codigo": "Lectura PDF417", "datos": datos}

    # --- PASO B: Si no hay PDF417, leer MRZ con OCR ---
    mrz_datos = intentar_ocr_mrz(img)
    if mrz_datos:
        return {"tipo_codigo": "Lectura MRZ", "datos": mrz_datos}

    return None

@app.post("/api/extraer-paciente")
async def extraer_paciente(file: UploadFile = File(...)):
    contents = await file.read()
    resultado = procesar_imagen_cedula(contents)

    if not resultado:
        raise HTTPException(
            status_code=400,
            detail="No se pudo detectar un código PDF417 ni la franja de texto MRZ."
        )

    return {
        "exito": True,
        "tipo_codigo": resultado["tipo_codigo"],
        "paciente": resultado["datos"]
    }

# =============================================================
# 4. AGREGADO EN LA UNIFICACIÓN: base de datos, catálogos y reporte
# =============================================================
# Hasta aquí todo es el código de Tatiana, sin cambios (lectura de la cédula).
# De aquí en adelante está lo que se agregó para unir las partes:
#   - Guardar / consultar / eliminar pacientes en una base de datos (SQLite)
#   - Entregar al formulario las listas oficiales que están en el Excel de la empresa
#   - Generar el Implant Report oficial (módulo generador_reporte_simple.py)
#   - Entregar la página registro_p.html desde este mismo servidor
#
# PENDIENTE de seguridad (antes de usarlo con datos reales):
#   usuario y contraseña, HTTPS, cifrar la base de datos y restringir el CORS de arriba.
#   La foto de la cédula NO se guarda: se procesa en memoria y se descarta.

CARPETA = Path(__file__).resolve().parent
BASE_DATOS = CARPETA / "registros.db"        # aquí quedan los pacientes (un solo archivo)
PAGINA = CARPETA / "registro_p.html"
candado_reporte = threading.Lock()           # el reporte se genera de a uno (usa Excel)


# ---------- Base de datos ----------
def conectar():
    """Abre la base de datos. Las filas se leen como diccionarios."""
    conexion = sqlite3.connect(BASE_DATOS)
    conexion.row_factory = sqlite3.Row
    return conexion


def crear_tablas():
    """Crea las tablas si todavía no existen."""
    with conectar() as db:
        # Un registro por paciente (por número de documento). Los datos van como texto JSON.
        db.execute("""CREATE TABLE IF NOT EXISTS registros (
                          documento   TEXT PRIMARY KEY,
                          datos       TEXT NOT NULL,
                          creado      TEXT NOT NULL,
                          actualizado TEXT NOT NULL)""")
        # Bitácora: quién hizo qué y cuándo (trazabilidad)
        db.execute("""CREATE TABLE IF NOT EXISTS bitacora (
                          id           INTEGER PRIMARY KEY AUTOINCREMENT,
                          fecha        TEXT NOT NULL,
                          accion       TEXT NOT NULL,
                          documento    TEXT,
                          especialista TEXT,
                          detalle      TEXT)""")


def anotar_bitacora(accion, documento, especialista="", detalle=""):
    with conectar() as db:
        db.execute("INSERT INTO bitacora (fecha, accion, documento, especialista, detalle) VALUES (?,?,?,?,?)",
                   (datetime.datetime.now().isoformat(timespec="seconds"), accion, documento, especialista, detalle))


def nombre_especialista(registro):
    esp = registro.get("especialista", {})
    return (str(esp.get("nombre", "")) + " " + str(esp.get("apellido", ""))).strip()


crear_tablas()


# ---------- Catálogos (las listas oficiales salen del Excel de la empresa) ----------
# Estas listas cortas son las mismas que tiene el Excel en sus listas desplegables de la hoja "Datos".
LISTAS_FIJAS = {
    "modos": ["DDD", "DDDR", "VVI", "VVIR", "AAI", "AAIR", "VDD", "VDDR", "VVT", "DDI", "DDIR", "Off"],
    "ubicaciones": ["Left Subcutaneous", "Right Subcutaneous", "Left Subpectoral", "Right Subpectoral",
                    "Left Abdominal", "Right Abdominal", "Left Submammary", "Right Submammary"],
    "posiciones": ["RA", "RV pace/sence", "LV", "RV dual coil", "RV single coil", "SVC", "SQ", "SQLL"],
    "polaridades": ["Bipolar", "'+ Bipolar", "'- Bipolar", "Unipolar", "'+ Unipolar", "'- Unipolar", "N/A"],
    "fabricantes": ["Boston Scientific", "Medtronic", "St. Jude", "ELA", "Oscor", "Biotronik"],
}
_catalogos = {}      # se llena la primera vez que alguien lo pide


def leer_catalogos():
    """Lee del Excel los hospitales (con sus sedes), los municipios y las indicaciones de terapia."""
    if _catalogos:
        return _catalogos

    libro = openpyxl.load_workbook(generador.PLANTILLA, data_only=True)

    # Municipio -> departamento (hoja Municipios: D = municipio, B = departamento)
    municipios = []
    depto_de = {}
    for fila in libro["Municipios"].iter_rows(min_row=2, values_only=True):
        if fila[3] and fila[1]:
            municipios.append([str(fila[3]).strip(), str(fila[1]).strip()])
            depto_de[str(fila[3]).strip()] = str(fila[1]).strip()

    # Nombre del hospital -> ciudad (hoja Cuentas: E = nombre, F = ciudad)
    ciudad_de = {}
    for fila in libro["Cuentas"].iter_rows(min_row=2, values_only=True):
        if fila[4] and fila[5]:
            ciudad_de[str(fila[4]).strip().upper()] = str(fila[5]).strip()

    # Hospitales y sedes (hoja Inst_: B = nombre, D = dirección, E = teléfono)
    hospitales = []
    for fila in libro["Inst_"].iter_rows(min_row=2, max_row=73, values_only=True):
        if fila[1]:
            nombre = str(fila[1]).strip()
            ciudad = ciudad_de.get(nombre.upper(), "")
            hospitales.append({"nombre": nombre,
                               "direccion": str(fila[3] or "").strip(),
                               "telefono": str(fila[4] or "").strip(),
                               "ciudad": ciudad,
                               "departamento": depto_de.get(ciudad, "")})

    # Indicaciones de terapia (hoja Implant Report, columna AF, filas 2 a 27)
    indicaciones = [str(c[0].value).strip() for c in libro["Implant Report"]["AF2:AF27"] if c[0].value]

    _catalogos.update(LISTAS_FIJAS)
    _catalogos.update({"hospitales": hospitales, "municipios": municipios, "indicaciones": indicaciones})
    return _catalogos


@app.get("/api/catalogos")
def catalogos():
    try:
        return leer_catalogos()
    except Exception as error:
        raise HTTPException(status_code=500, detail="No se pudo leer el Excel de la empresa: " + str(error))


# ---------- Dónde se guardan los pacientes: Dataverse o registros.db ----------
# Si alguien inició sesión con dataverse_conexion.py en este computador, se usa Dataverse.
# Para forzar la base local:  $env:ALMACEN="sqlite"
USAR_DATAVERSE = os.environ.get("ALMACEN", "").lower() != "sqlite" and dvr.disponible()


def almacen_guardar(registro, documento):
    if USAR_DATAVERSE:
        return dvr.guardar(registro)
    ahora = datetime.datetime.now().isoformat(timespec="seconds")
    with conectar() as db:
        existe = db.execute("SELECT 1 FROM registros WHERE documento = ?", (documento,)).fetchone()
        if existe:
            db.execute("UPDATE registros SET datos = ?, actualizado = ? WHERE documento = ?",
                       (json.dumps(registro, ensure_ascii=False), ahora, documento))
        else:
            db.execute("INSERT INTO registros (documento, datos, creado, actualizado) VALUES (?,?,?,?)",
                       (documento, json.dumps(registro, ensure_ascii=False), ahora, ahora))
    return "actualizar" if existe else "crear"


def almacen_consultar(documento):
    if USAR_DATAVERSE:
        return dvr.consultar(documento)
    with conectar() as db:
        fila = db.execute("SELECT datos FROM registros WHERE documento = ?", (documento,)).fetchone()
    return json.loads(fila["datos"]) if fila else None


def almacen_eliminar(documento):
    if USAR_DATAVERSE:
        return dvr.eliminar(documento)
    with conectar() as db:
        return bool(db.execute("DELETE FROM registros WHERE documento = ?", (documento,)).rowcount)


def almacen_listar():
    """[(documento, creado, actualizado, registro), ...]"""
    if USAR_DATAVERSE:
        return dvr.listar()
    with conectar() as db:
        filas = db.execute("SELECT documento, datos, creado, actualizado FROM registros ORDER BY creado").fetchall()
    return [(f["documento"], f["creado"], f["actualizado"], json.loads(f["datos"])) for f in filas]


def error_dataverse(err):
    return HTTPException(status_code=502, detail="No se pudo comunicar con Dataverse: " + str(err)
                         + " (si la sesión caducó, ejecuta: python dataverse_conexion.py)")


# ---------- Guardar, consultar y eliminar pacientes ----------
@app.post("/api/registros")
def guardar_registro(registro: dict = Body(...)):
    documento = str(registro.get("paciente", {}).get("documento", "")).strip()
    if not documento:
        raise HTTPException(status_code=400, detail="El número de documento del paciente es obligatorio.")

    try:
        accion = almacen_guardar(registro, documento)
    except Exception as err:
        raise error_dataverse(err)
    anotar_bitacora(accion, documento, nombre_especialista(registro))
    return {"exito": True, "documento": documento, "actualizado": datetime.datetime.now().isoformat(timespec="seconds")}


@app.get("/api/registros/{documento}")
def consultar_registro(documento: str):
    try:
        registro = almacen_consultar(documento.strip())
    except Exception as err:
        raise error_dataverse(err)
    if not registro:
        raise HTTPException(status_code=404, detail="No se encontró ningún paciente con ese documento.")
    return registro


@app.delete("/api/registros/{documento}")
def eliminar_registro(documento: str):
    documento = documento.strip()
    try:
        borrado = almacen_eliminar(documento)
    except Exception as err:
        raise error_dataverse(err)
    if borrado:
        anotar_bitacora("eliminar", documento)
    return {"exito": True, "eliminado": borrado}


# ---------- Exportar todos los registros a Excel ----------
def aplanar(valor, prefijo, fila):
    """Convierte el registro (diccionarios y listas anidadas) en columnas: 'paciente · documento', 'medicos · 1 · nombre'..."""
    if isinstance(valor, dict):
        for clave, v in valor.items():
            aplanar(v, f"{prefijo} · {clave}" if prefijo else clave, fila)
    elif isinstance(valor, list):
        for i, v in enumerate(valor, start=1):
            aplanar(v, f"{prefijo} · {i}", fila)
    else:
        fila[prefijo] = "" if valor is None else valor


@app.get("/api/exportar-excel")
def exportar_excel():
    try:
        registros = almacen_listar()
    except Exception as err:
        raise error_dataverse(err)

    filas, columnas = [], ["documento", "creado", "actualizado"]
    for documento, creado, actualizado, registro in registros:
        fila = {"documento": documento, "creado": creado, "actualizado": actualizado}
        aplanar(registro, "", fila)
        for c in fila:
            if c not in columnas:
                columnas.append(c)
        filas.append(fila)

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Registros"
    hoja.append(columnas)
    for fila in filas:
        hoja.append([fila.get(c, "") for c in columnas])
    for celda in hoja[1]:
        celda.font = openpyxl.styles.Font(bold=True)
    hoja.freeze_panes = "B2"
    hoja.auto_filter.ref = hoja.dimensions
    for i, c in enumerate(columnas, start=1):
        hoja.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(max(len(c), 12), 40)

    contenido = io.BytesIO()
    libro.save(contenido)
    nombre = f"registros_{datetime.date.today().isoformat()}.xlsx"
    return Response(content=contenido.getvalue(),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{nombre}"'})


# ---------- Implant Report oficial ----------
def registro_a_reporte(registro):
    """
    Convierte el registro del formulario (con los nombres de campo de Tatiana)
    en el diccionario que entiende generador_reporte_simple.py.
    """
    paciente = dict(registro.get("paciente", {}))
    diagnostico = registro.get("diagnostico", {})
    implante = registro.get("implante", {})
    medicos = registro.get("medicos", [])

    # El Excel solo tiene un médico: usamos el implantador (o el primero que haya)
    implantador = next((m for m in medicos if m.get("rol") == "Implanting"), medicos[0] if medicos else {})

    # Fecha del procedimiento: la de implante; si no la dieron, la fecha de diligenciamiento
    fecha = implante.get("fecha_implante") or registro.get("fecha_diligenciamiento")

    paciente["diagnostico"] = diagnostico.get("indicacion_principal", "")
    paciente["fe"] = diagnostico.get("fraccion_eyeccion", "")

    dispositivo = dict(implante.get("dispositivo", {}))
    if dispositivo.get("modelo") and not dispositivo.get("fabricante"):
        dispositivo["fabricante"] = "Boston Scientific"

    electrodos = {}
    for nombre, electrodo in implante.get("electrodos", {}).items():
        if electrodo.get("modelo"):                      # solo los electrodos que se llenaron
            electrodo = dict(electrodo)
            electrodo["fecha_implante"] = fecha
            electrodo.setdefault("fabricante", "Boston Scientific")
            electrodos[nombre] = electrodo

    return {
        "especialista": nombre_especialista(registro).upper(),
        "fecha_implante": fecha,
        "paciente": paciente,
        "medico": {"nombre": implantador.get("nombre", ""), "apellido": implantador.get("apellido", "")},
        "hospital": registro.get("hospital", {}),
        "dispositivo": dispositivo,
        "parametros": implante.get("parametros", {}),
        "electrodos": electrodos,
    }


@app.post("/api/generar-reporte/{documento}")
def generar_reporte_paciente(documento: str):
    try:
        registro = almacen_consultar(documento.strip())
    except Exception as err:
        raise error_dataverse(err)
    if not registro:
        raise HTTPException(status_code=404, detail="Primero guarda el registro de este paciente.")

    with candado_reporte:                       # de a un reporte a la vez
        resultado = generador.generar_reporte(registro_a_reporte(registro))

    archivo = resultado["pdf"] or resultado["excel"]
    anotar_bitacora("generar_reporte", documento.strip(), nombre_especialista(registro),
                    os.path.basename(archivo) if archivo else "sin archivo")
    return {
        "exito": archivo is not None,
        "tipo": "pdf" if resultado["pdf"] else "xlsx",
        "url": "/api/reportes/" + os.path.basename(archivo) if archivo else None,
        "avisos": resultado["avisos"],
    }


@app.get("/api/reportes/{nombre}")
def descargar_reporte(nombre: str):
    nombre = os.path.basename(nombre)           # evita rutas raras como ../
    ruta = Path(generador.CARPETA_SALIDA) / nombre
    if not ruta.is_file() or ruta.suffix.lower() not in (".pdf", ".xlsx"):
        raise HTTPException(status_code=404, detail="No existe ese reporte.")
    return FileResponse(ruta)


# ---------- La página ----------
@app.get("/")
def pagina():
    return FileResponse(PAGINA)


if __name__ == "__main__":
    # 127.0.0.1 = solo este computador. Para que otros equipos de la misma red entren,
    # ejecutar con:  set HOST=0.0.0.0   (en PowerShell:  $env:HOST="0.0.0.0")  y luego  py App.py
    HOST = os.environ.get("HOST", "127.0.0.1")
    print("Abre en el navegador:  http://127.0.0.1:8001")
    print("Los pacientes se guardan en:", ("Dataverse " + dvr.dv.DATAVERSE_URL) if USAR_DATAVERSE else BASE_DATOS)
    uvicorn.run(app, host=HOST, port=8001)