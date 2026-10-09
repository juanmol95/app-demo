"""
GUARDAR Y CONSULTAR REGISTROS EN DATAVERSE
==========================================
Mismo registro que envía el formulario (JSON), repartido en las tablas de Dataverse:
Paciente, Implante (con una copia completa del formulario en JSON), Electrodo y
Médico del implante. Por ahora hay un implante por paciente, igual que antes:
guardar de nuevo el mismo documento actualiza su registro.
"""

import datetime
import json

import dataverse_conexion as dv

_conjuntos = {}     # nombre lógico de la tabla -> nombre en la API (p. ej. bsc_paciente -> bsc_pacientes)
_navegacion = {}    # (tabla, columna de búsqueda) -> propiedad de navegación para @odata.bind


def disponible():
    """Hay Dataverse si alguien inició sesión en este computador o si hay sesión en Render."""
    return dv.hay_sesion()


def conjunto(tabla):
    if tabla not in _conjuntos:
        r = dv.api("GET", f"EntityDefinitions(LogicalName='{tabla}')?$select=EntitySetName")
        _revisar(r, "leer tabla " + tabla)
        _conjuntos[tabla] = r.json()["EntitySetName"]
    return _conjuntos[tabla]


def navegacion(tabla, columna):
    if (tabla, columna) not in _navegacion:
        r = dv.api("GET", f"EntityDefinitions(LogicalName='{tabla}')/ManyToOneRelationships"
                          "?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName")
        _revisar(r, "leer relaciones de " + tabla)
        for rel in r.json()["value"]:
            _navegacion[(tabla, rel["ReferencingAttribute"])] = rel["ReferencingEntityNavigationPropertyName"]
    return _navegacion[(tabla, columna)]


def _revisar(r, que):
    if r.status_code >= 300:
        try:
            detalle = r.json()["error"]["message"]
        except Exception:
            detalle = r.text[:300]
        raise RuntimeError(f"Dataverse ({que}): {detalle}")


def _texto(v):
    v = "" if v is None else str(v).strip()
    return v or None


def _fecha(v, nombre="Fecha", minimo="1900-01-01"):
    """AAAA-MM-DD entre 'minimo' y hoy; una fecha imposible (año 1111, 9999999...) se rechaza con un mensaje claro."""
    v = _texto(v)
    if not v:
        return None
    hoy = datetime.date.today().isoformat()
    if len(v) != 10 or not (minimo <= v <= hoy):
        raise ValueError(f"{nombre} no válida ({v}): debe estar entre {minimo[8:]}/{minimo[5:7]}/{minimo[:4]} y hoy.")
    return v


def _numero(v, entero=False):
    v = _texto(v)
    if v is None:
        return None
    try:
        return int(float(v)) if entero else float(v)
    except ValueError:
        return None


def _comilla(v):
    return str(v).replace("'", "''")


def _buscar(tabla, filtro, campos, orden=None):
    ruta = f"{conjunto(tabla)}?$select={campos}&$filter={filtro}"
    if orden:
        ruta += f"&$orderby={orden}"
    r = dv.api("GET", ruta)
    _revisar(r, "buscar en " + tabla)
    return r.json()["value"]


def _crear(tabla, datos):
    r = dv.api("POST", conjunto(tabla), json=datos, headers={"Prefer": "return=representation"})
    _revisar(r, "crear en " + tabla)
    return r.json()[tabla + "id"]


def _actualizar(tabla, id_, datos):
    _revisar(dv.api("PATCH", f"{conjunto(tabla)}({id_})", json=datos), "actualizar " + tabla)


def _borrar(tabla, id_):
    _revisar(dv.api("DELETE", f"{conjunto(tabla)}({id_})"), "borrar en " + tabla)


def _vincular(tabla, columna, tabla_destino, id_destino):
    return {f"{navegacion(tabla, columna)}@odata.bind": f"/{conjunto(tabla_destino)}({id_destino})"}


def _id_paciente(documento):
    v = _buscar("bsc_paciente", f"bsc_name eq '{_comilla(documento)}'", "bsc_pacienteid")
    return v[0]["bsc_pacienteid"] if v else None


def _implantes(id_paciente):
    return _buscar("bsc_implante", f"_bsc_pacienteid_value eq {id_paciente}",
                   "bsc_implanteid,bsc_datosformulario", "modifiedon desc")


def _borrar_hijos(id_implante):
    for tabla in ("bsc_electrodo", "bsc_medicoimplante"):
        for fila in _buscar(tabla, f"_bsc_implanteid_value eq {id_implante}", tabla + "id"):
            _borrar(tabla, fila[tabla + "id"])


def guardar(registro):
    """Crea o actualiza el paciente y su implante. Devuelve 'crear' o 'actualizar'."""
    p = registro.get("paciente", {})
    documento = _texto(p.get("documento"))
    if not documento:
        raise ValueError("El número de documento del paciente es obligatorio.")
    # Validar todas las fechas antes de escribir nada
    _fecha(p.get("fecha_nacimiento"), "Fecha de nacimiento")
    _fecha(registro.get("fecha_diligenciamiento"), "Fecha de diligenciamiento", "2000-01-01")
    _fecha((registro.get("implante") or {}).get("fecha_implante"), "Fecha de implante", "1960-01-01")

    datos_paciente = {
        "bsc_name": documento, "bsc_historiaclinica": _texto(p.get("medical_record")),
        "bsc_primernombre": _texto(p.get("primer_nombre")), "bsc_segundonombre": _texto(p.get("segundo_nombre")),
        "bsc_primerapellido": _texto(p.get("primer_apellido")), "bsc_segundoapellido": _texto(p.get("segundo_apellido")),
        "bsc_sufijo": _texto(p.get("sufijo")), "bsc_fechanacimiento": _fecha(p.get("fecha_nacimiento"), "Fecha de nacimiento"),
        "bsc_sexo": _texto(p.get("sexo")), "bsc_rh": _texto(p.get("rh")), "bsc_eps": _texto(p.get("eps")),
        "bsc_direccion": _texto(p.get("direccion")), "bsc_telefono": _texto(p.get("telefono")),
        "bsc_ciudad": _texto(p.get("ciudad")), "bsc_departamento": _texto(p.get("departamento")),
        "bsc_pais": _texto(p.get("pais")),
    }
    id_paciente = _id_paciente(documento)
    accion = "actualizar" if id_paciente else "crear"
    if id_paciente:
        _actualizar("bsc_paciente", id_paciente, datos_paciente)
    else:
        id_paciente = _crear("bsc_paciente", datos_paciente)

    esp = registro.get("especialista", {})
    diag = registro.get("diagnostico", {})
    hosp = registro.get("hospital", {})
    imp = registro.get("implante", {})
    disp = imp.get("dispositivo", {})
    par = imp.get("parametros", {})
    datos_implante = {
        "bsc_name": f"{documento} · {imp.get('fecha_implante') or registro.get('fecha_diligenciamiento') or ''}".strip(" ·"),
        "bsc_fechadiligenciamiento": _fecha(registro.get("fecha_diligenciamiento"), "Fecha de diligenciamiento", "2000-01-01"),
        "bsc_especialistanombre": _texto(esp.get("nombre")), "bsc_especialistaapellido": _texto(esp.get("apellido")),
        "bsc_diagnostico": _texto(diag.get("indicacion_principal")),
        "bsc_fraccioneyeccion": _numero(diag.get("fraccion_eyeccion")),
        "bsc_indicacionesotras": _texto(diag.get("indicaciones_otras")),
        "bsc_hospitalnombre": _texto(hosp.get("nombre")), "bsc_hospitaldireccion": _texto(hosp.get("direccion")),
        "bsc_hospitalciudad": _texto(hosp.get("ciudad")), "bsc_hospitaldepartamento": _texto(hosp.get("departamento")),
        "bsc_hospitaltelefono": _texto(hosp.get("telefono")),
        "bsc_fechaimplante": _fecha(imp.get("fecha_implante"), "Fecha de implante", "1960-01-01"), "bsc_modelo": _texto(disp.get("modelo")),
        "bsc_serial": _texto(disp.get("serial")), "bsc_fabricante": _texto(disp.get("fabricante")),
        "bsc_ubicacion": _texto(disp.get("ubicacion")), "bsc_modo": _texto(par.get("modo")),
        "bsc_lrl": _numero(par.get("lrl"), True), "bsc_url": _numero(par.get("url"), True),
        "bsc_avdelay": _numero(par.get("av_delay"), True), "bsc_pvarp": _numero(par.get("pvarp"), True),
        "bsc_vrefractory": _numero(par.get("v_ref"), True),
        "bsc_avsearch": _texto(par.get("av_search")), "bsc_vrr": _texto(par.get("vrr")),
        "bsc_atr": _texto(par.get("atr")), "bsc_sbr": _texto(par.get("sbr")),
        "bsc_datosformulario": json.dumps(registro, ensure_ascii=False),
    }
    datos_implante.update(_vincular("bsc_implante", "bsc_pacienteid", "bsc_paciente", id_paciente))
    if _texto(hosp.get("nombre")):
        h = _buscar("bsc_hospital", f"bsc_name eq '{_comilla(hosp['nombre'].strip())}'", "bsc_hospitalid")
        if h:
            datos_implante.update(_vincular("bsc_implante", "bsc_hospitalid", "bsc_hospital", h[0]["bsc_hospitalid"]))

    existentes = _implantes(id_paciente)
    if existentes:
        id_implante = existentes[0]["bsc_implanteid"]
        _actualizar("bsc_implante", id_implante, datos_implante)
        _borrar_hijos(id_implante)
    else:
        id_implante = _crear("bsc_implante", datos_implante)

    titulos = {"auricular_derecho": "RA", "ventricular_derecho": "RV", "ventricular_izquierdo": "LV"}
    for clave, e in (imp.get("electrodos") or {}).items():
        if not _texto(e.get("modelo")):
            continue                       # igual que en el reporte: solo los electrodos con modelo
        tipo = titulos.get(clave, clave)
        fila = {"bsc_name": f"{tipo} · {e.get('modelo')}", "bsc_tipo": tipo,
                "bsc_modelo": _texto(e.get("modelo")), "bsc_serial": _texto(e.get("serial")),
                "bsc_polaridad": _texto(e.get("polaridad")), "bsc_posicion": _texto(e.get("posicion")),
                "bsc_onda": _numero(e.get("onda")), "bsc_impedancia": _numero(e.get("impedancia")),
                "bsc_anchopulso": _numero(e.get("ancho")), "bsc_umbral": _numero(e.get("umbral"))}
        fila.update(_vincular("bsc_electrodo", "bsc_implanteid", "bsc_implante", id_implante))
        _crear("bsc_electrodo", fila)

    for m in registro.get("medicos") or []:
        nombre = " ".join(x for x in (_texto(m.get("nombre")), _texto(m.get("apellido"))) if x)
        if not nombre:
            continue
        fila = {"bsc_name": nombre, "bsc_rol": _texto(m.get("rol")), "bsc_especialidad": _texto(m.get("especialidad")),
                "bsc_nombre": _texto(m.get("nombre")), "bsc_apellido": _texto(m.get("apellido")),
                "bsc_sufijo": _texto(m.get("sufijo")), "bsc_direccion": _texto(m.get("direccion")),
                "bsc_ciudad": _texto(m.get("ciudad")), "bsc_telefono": _texto(m.get("telefono"))}
        fila.update(_vincular("bsc_medicoimplante", "bsc_implanteid", "bsc_implante", id_implante))
        _crear("bsc_medicoimplante", fila)

    return accion


def consultar(documento):
    """Devuelve el registro del formulario (JSON) del paciente, o None si no existe."""
    id_paciente = _id_paciente(documento.strip())
    if not id_paciente:
        return None
    implantes = _implantes(id_paciente)
    if not implantes or not implantes[0].get("bsc_datosformulario"):
        return None
    return json.loads(implantes[0]["bsc_datosformulario"])


def eliminar(documento):
    """Borra el paciente con sus implantes, electrodos y médicos. Devuelve True si existía."""
    id_paciente = _id_paciente(documento.strip())
    if not id_paciente:
        return False
    for imp in _implantes(id_paciente):
        _borrar_hijos(imp["bsc_implanteid"])
        _borrar("bsc_implante", imp["bsc_implanteid"])
    _borrar("bsc_paciente", id_paciente)
    return True


def listar():
    """Todos los registros, para exportar: [(documento, creado, actualizado, registro), ...]"""
    salida = []
    ruta = f"{conjunto('bsc_implante')}?$select=bsc_datosformulario,createdon,modifiedon&$orderby=createdon"
    while ruta:
        r = dv.api("GET", ruta)
        _revisar(r, "listar implantes")
        cuerpo = r.json()
        for f in cuerpo["value"]:
            if not f.get("bsc_datosformulario"):
                continue
            registro = json.loads(f["bsc_datosformulario"])
            salida.append((registro.get("paciente", {}).get("documento", ""), f["createdon"], f["modifiedon"], registro))
        siguiente = cuerpo.get("@odata.nextLink")
        ruta = siguiente.split("/api/data/v9.2/", 1)[1] if siguiente else None
    return salida
