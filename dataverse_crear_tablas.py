"""
CREAR LAS TABLAS DEL REGISTRO DE IMPLANTES EN DATAVERSE
=======================================================
Crea (si no existen) el publicador "BSC" (prefijo bsc), la solución "Registro de Implantes"
y estas tablas con sus columnas y relaciones:

    Hospital            (bsc_hospital)
    Paciente            (bsc_paciente)
    Implante            (bsc_implante)        -> Paciente, Hospital
    Electrodo           (bsc_electrodo)       -> Implante
    Médico del implante (bsc_medicoimplante)  -> Implante

También carga los hospitales de la tabla de instituciones.
Se puede ejecutar varias veces: lo que ya existe no se vuelve a crear.

Uso (después de iniciar sesión con dataverse_conexion.py):
    python dataverse_crear_tablas.py
"""

import json
import sys

import dataverse_conexion as dv

IDIOMA = 3082            # español, idioma base del entorno
SOLUCION = "RegistroImplantes"
SOLO_ENCABEZADO = {"MSCRM.SolutionUniqueName": SOLUCION}


def etiqueta(texto):
    return {"@odata.type": "Microsoft.Dynamics.CRM.Label",
            "LocalizedLabels": [{"@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                 "Label": texto, "LanguageCode": IDIOMA}]}


NO_REQUERIDO = {"Value": "None", "CanBeChanged": True,
                "ManagedPropertyLogicalName": "canmodifyrequirementlevelsettings"}


def texto(nombre, visible, largo=200):
    return {"@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata", "SchemaName": nombre,
            "AttributeType": "String", "FormatName": {"Value": "Text"}, "MaxLength": largo,
            "RequiredLevel": NO_REQUERIDO, "DisplayName": etiqueta(visible)}


def memo(nombre, visible):
    return {"@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata", "SchemaName": nombre,
            "AttributeType": "Memo", "Format": "TextArea", "MaxLength": 1048576,
            "RequiredLevel": NO_REQUERIDO, "DisplayName": etiqueta(visible)}


def fecha(nombre, visible):
    return {"@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata", "SchemaName": nombre,
            "AttributeType": "DateTime", "Format": "DateOnly", "DateTimeBehavior": {"Value": "DateOnly"},
            "RequiredLevel": NO_REQUERIDO, "DisplayName": etiqueta(visible)}


def decimal(nombre, visible):
    return {"@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata", "SchemaName": nombre,
            "AttributeType": "Decimal", "Precision": 2, "MinValue": -1000000, "MaxValue": 1000000,
            "RequiredLevel": NO_REQUERIDO, "DisplayName": etiqueta(visible)}


def entero(nombre, visible):
    return {"@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata", "SchemaName": nombre,
            "AttributeType": "Integer", "Format": "None", "MinValue": -100000, "MaxValue": 100000,
            "RequiredLevel": NO_REQUERIDO, "DisplayName": etiqueta(visible)}


# Tabla -> (nombre visible, plural, columna principal, otras columnas)
TABLAS = {
    "bsc_Hospital": ("Hospital", "Hospitales", texto("bsc_Name", "Nombre de la institución"), [
        texto("bsc_RazonSocial", "Razón social"), texto("bsc_CodSAP", "Cod SAP", 50),
        texto("bsc_Direccion", "Dirección"), texto("bsc_Telefono", "Teléfono", 50),
        texto("bsc_Ciudad", "Ciudad", 100), texto("bsc_Departamento", "Departamento", 100),
        texto("bsc_Zona", "Zona", 50)]),
    "bsc_Paciente": ("Paciente", "Pacientes", texto("bsc_Name", "Documento", 50), [
        texto("bsc_HistoriaClinica", "Historia clínica", 50),
        texto("bsc_PrimerNombre", "Primer nombre", 100), texto("bsc_SegundoNombre", "Segundo nombre", 100),
        texto("bsc_PrimerApellido", "Primer apellido", 100), texto("bsc_SegundoApellido", "Segundo apellido", 100),
        texto("bsc_Sufijo", "Sufijo", 20), fecha("bsc_FechaNacimiento", "Fecha de nacimiento"),
        texto("bsc_Sexo", "Sexo", 20), texto("bsc_RH", "RH", 10), texto("bsc_EPS", "EPS", 100),
        texto("bsc_Direccion", "Dirección"), texto("bsc_Telefono", "Teléfono", 50),
        texto("bsc_Ciudad", "Ciudad", 100), texto("bsc_Departamento", "Departamento", 100),
        texto("bsc_Pais", "País", 100)]),
    "bsc_Implante": ("Implante", "Implantes", texto("bsc_Name", "Registro"), [
        fecha("bsc_FechaDiligenciamiento", "Fecha de diligenciamiento"),
        texto("bsc_EspecialistaNombre", "Especialista - nombre", 100),
        texto("bsc_EspecialistaApellido", "Especialista - apellido", 100),
        texto("bsc_Diagnostico", "Indicación principal", 500), decimal("bsc_FraccionEyeccion", "Fracción de eyección (%)"),
        texto("bsc_IndicacionesOtras", "Otras indicaciones", 1000),
        texto("bsc_HospitalNombre", "Hospital - nombre"), texto("bsc_HospitalDireccion", "Hospital - dirección"),
        texto("bsc_HospitalCiudad", "Hospital - ciudad", 100), texto("bsc_HospitalDepartamento", "Hospital - departamento", 100),
        texto("bsc_HospitalTelefono", "Hospital - teléfono", 50),
        fecha("bsc_FechaImplante", "Fecha de implante"), texto("bsc_Modelo", "Modelo", 100),
        texto("bsc_Serial", "Serial", 100), texto("bsc_Fabricante", "Fabricante", 100),
        texto("bsc_Ubicacion", "Ubicación del implante", 100), texto("bsc_Modo", "Modo de estimulación", 20),
        entero("bsc_LRL", "LRL (ppm)"), entero("bsc_URL", "URL (ppm)"), entero("bsc_AVDelay", "AV Delay (ms)"),
        entero("bsc_PVARP", "PVARP (ms)"), entero("bsc_VRefractory", "V Refractory (ms)"),
        texto("bsc_AVSearch", "AV Search", 10), texto("bsc_VRR", "VRR", 10),
        texto("bsc_ATR", "Atrial Tachy Response", 10), texto("bsc_SBR", "Sudden Brady Response", 10),
        memo("bsc_DatosFormulario", "Datos completos del formulario (JSON)")]),
    "bsc_Electrodo": ("Electrodo", "Electrodos", texto("bsc_Name", "Electrodo"), [
        texto("bsc_Tipo", "Tipo (RA/RV/LV)", 50), texto("bsc_Modelo", "Modelo", 100), texto("bsc_Serial", "Serial", 100),
        texto("bsc_Polaridad", "Polaridad", 50), texto("bsc_Posicion", "Posición", 50),
        decimal("bsc_Onda", "Onda (mV)"), decimal("bsc_Impedancia", "Impedancia (ohms)"),
        decimal("bsc_AnchoPulso", "Ancho de pulso (ms)"), decimal("bsc_Umbral", "Umbral (V)")]),
    "bsc_MedicoImplante": ("Médico del implante", "Médicos del implante", texto("bsc_Name", "Médico"), [
        texto("bsc_Rol", "Rol", 100), texto("bsc_Especialidad", "Especialidad", 100),
        texto("bsc_Nombre", "Nombres", 100), texto("bsc_Apellido", "Apellidos", 100), texto("bsc_Sufijo", "Sufijo", 20),
        texto("bsc_Direccion", "Dirección"), texto("bsc_Ciudad", "Ciudad", 100), texto("bsc_Telefono", "Teléfono", 50)]),
}

# (tabla padre, tabla hija, columna de búsqueda en la hija, nombre visible)
RELACIONES = [
    ("bsc_paciente", "bsc_implante", "bsc_PacienteId", "Paciente"),
    ("bsc_hospital", "bsc_implante", "bsc_HospitalId", "Hospital"),
    ("bsc_implante", "bsc_electrodo", "bsc_ImplanteId", "Implante"),
    ("bsc_implante", "bsc_medicoimplante", "bsc_ImplanteId", "Implante"),
]

# Tabla de instituciones (tal cual la entregó el equipo)
HOSPITALES = [
    {"bsc_name": "HOSPITAL TE CUIDAMOS-SEDE TESORO", "bsc_codsap": "119356", "bsc_razonsocial": "CENTRO CUIDADO TE CUIDAMOS",
     "bsc_direccion": "Avenida Carrera 30 #53b-50", "bsc_telefono": "57-12345678", "bsc_ciudad": "BOGOTA",
     "bsc_departamento": "DC", "bsc_zona": "POOL BOG"},
    {"bsc_name": "HOSPITAL TE CUIDAMOS-SEDE CENTRO", "bsc_codsap": "95153", "bsc_razonsocial": "",
     "bsc_direccion": "Carrera 1 #18a 12", "bsc_telefono": "", "bsc_ciudad": "BOGOTA",
     "bsc_departamento": "DC", "bsc_zona": "POOL BOG"},
    {"bsc_name": "CLINICA CORAZON PAISA", "bsc_codsap": "116182", "bsc_razonsocial": "CLINICA CORA PAISA S.A.S",
     "bsc_direccion": "Carrera 65 #23a-05", "bsc_telefono": "57-87654321", "bsc_ciudad": "Medellin",
     "bsc_departamento": "Antioquia", "bsc_zona": "POOL MDE"},
    {"bsc_name": "CLINICA SANANDO CORAZONES", "bsc_codsap": "123177", "bsc_razonsocial": "IPS SANANDO TU CORAZON S.A",
     "bsc_direccion": "Avenida Calle 26 103 09", "bsc_telefono": "57-23456789", "bsc_ciudad": "BOGOTA",
     "bsc_departamento": "DC", "bsc_zona": "POOL BOG"},
]


def revisar(r, que):
    if r.status_code >= 300:
        print(f"ERROR en {que}: {r.status_code} {r.text[:800]}")
        sys.exit(1)


def obtener_o_crear_publicador():
    v = dv.api("GET", "publishers?$select=publisherid&$filter=uniquename eq 'bsc'").json()["value"]
    if v:
        return v[0]["publisherid"]
    r = dv.api("POST", "publishers", json={"uniquename": "bsc", "friendlyname": "BSC", "customizationprefix": "bsc",
                                          "customizationoptionvalueprefix": 72700},
               headers={"Prefer": "return=representation"})
    revisar(r, "crear publicador")
    print("Publicador BSC creado")
    return r.json()["publisherid"]


def obtener_o_crear_solucion(publicador):
    v = dv.api("GET", f"solutions?$select=solutionid&$filter=uniquename eq '{SOLUCION}'").json()["value"]
    if v:
        return
    r = dv.api("POST", "solutions", json={"uniquename": SOLUCION, "friendlyname": "Registro de Implantes",
                                         "version": "1.0.0.0", "publisherid@odata.bind": f"/publishers({publicador})"})
    revisar(r, "crear solución")
    print("Solución 'Registro de Implantes' creada")


def existe_tabla(logico):
    return dv.api("GET", f"EntityDefinitions(LogicalName='{logico}')?$select=LogicalName").status_code == 200


def columnas_existentes(logico):
    r = dv.api("GET", f"EntityDefinitions(LogicalName='{logico}')/Attributes?$select=LogicalName")
    return {a["LogicalName"] for a in r.json()["value"]}


def crear_tablas():
    for esquema, (visible, plural, principal, columnas) in TABLAS.items():
        logico = esquema.lower()
        if not existe_tabla(logico):
            cuerpo = {"@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata", "SchemaName": esquema,
                      "DisplayName": etiqueta(visible), "DisplayCollectionName": etiqueta(plural),
                      "Description": etiqueta(f"{plural} del registro de implantes"),
                      "OwnershipType": "UserOwned", "IsActivity": False, "HasActivities": False, "HasNotes": False,
                      "Attributes": [dict(principal, IsPrimaryName=True)]}
            r = dv.api("POST", "EntityDefinitions", json=cuerpo, headers=SOLO_ENCABEZADO)
            revisar(r, f"crear tabla {visible}")
            print(f"Tabla '{plural}' creada")
        existentes = columnas_existentes(logico)
        for col in columnas:
            if col["SchemaName"].lower() in existentes:
                continue
            r = dv.api("POST", f"EntityDefinitions(LogicalName='{logico}')/Attributes", json=col, headers=SOLO_ENCABEZADO)
            revisar(r, f"crear columna {col['SchemaName']}")
        print(f"  columnas de '{plural}' listas")


def crear_relaciones():
    for padre, hija, busqueda, visible in RELACIONES:
        if busqueda.lower() in columnas_existentes(hija):
            continue
        cuerpo = {"@odata.type": "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata",
                  "SchemaName": f"{padre}_{hija}_{busqueda.lower()}",
                  "ReferencedEntity": padre, "ReferencedAttribute": f"{padre}id", "ReferencingEntity": hija,
                  "CascadeConfiguration": {"Assign": "NoCascade", "Delete": "RemoveLink", "Merge": "NoCascade",
                                           "Reparent": "NoCascade", "Share": "NoCascade", "Unshare": "NoCascade"},
                  "Lookup": {"@odata.type": "Microsoft.Dynamics.CRM.LookupAttributeMetadata", "SchemaName": busqueda,
                             "AttributeType": "Lookup", "AttributeTypeName": {"Value": "LookupType"},
                             "DisplayName": etiqueta(visible), "RequiredLevel": NO_REQUERIDO}}
        r = dv.api("POST", "RelationshipDefinitions", json=cuerpo, headers=SOLO_ENCABEZADO)
        revisar(r, f"crear relación {hija} -> {padre}")
        print(f"Relación {hija} -> {padre} creada")


def publicar():
    r = dv.api("POST", "PublishAllXml")
    revisar(r, "publicar")
    print("Cambios publicados")


def nombres_de_conjunto():
    salida = {}
    for esquema in TABLAS:
        r = dv.api("GET", f"EntityDefinitions(LogicalName='{esquema.lower()}')?$select=EntitySetName")
        salida[esquema.lower()] = r.json()["EntitySetName"]
    return salida


def cargar_hospitales(conjunto):
    for h in HOSPITALES:
        nombre = h["bsc_name"].replace("'", "''")
        v = dv.api("GET", f"{conjunto}?$select=bsc_name&$filter=bsc_name eq '{nombre}'").json()["value"]
        if v:
            continue
        r = dv.api("POST", conjunto, json={k: val for k, val in h.items() if val})
        revisar(r, f"cargar hospital {h['bsc_name']}")
        print("Hospital cargado:", h["bsc_name"])


if __name__ == "__main__":
    print("Entorno:", dv.DATAVERSE_URL)
    obtener_o_crear_solucion(obtener_o_crear_publicador())
    crear_tablas()
    crear_relaciones()
    publicar()
    conjuntos = nombres_de_conjunto()
    print("Nombres en la API:", json.dumps(conjuntos, ensure_ascii=False))
    cargar_hospitales(conjuntos["bsc_hospital"])
    print("Listo.")
