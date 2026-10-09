"""
CERTIFICADO PARA ABRIR LA APP CON HTTPS (MODO CELULAR)
======================================================
Los navegadores solo permiten la cámara en vivo en páginas https (o en el mismo
computador). Para usar la cámara desde el celular en la red local, el servidor se
arranca con un certificado propio ("autofirmado"), que se crea aquí la primera vez.

El celular mostrará un aviso ("La conexión no es privada") porque el certificado no es
de una entidad reconocida: se acepta una vez con "Configuración avanzada" -> "Continuar".
"""

import datetime
import ipaddress
import socket
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID

CARPETA = Path(__file__).resolve().parent / "certificado"
CERTIFICADO = CARPETA / "certificado.pem"
CLAVE = CARPETA / "clave.pem"


def direcciones_locales():
    """Direcciones IPv4 de este computador en las redes a las que está conectado."""
    ips = {"127.0.0.1"}
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    return sorted(ip for ip in ips if not ip.startswith("169.254."))


def asegurar_certificado():
    """Crea el certificado si no existe (válido 2 años) y devuelve (certificado, clave)."""
    if CERTIFICADO.is_file() and CLAVE.is_file():
        return str(CERTIFICADO), str(CLAVE)
    CARPETA.mkdir(exist_ok=True)
    clave = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    nombre = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "Registro de Implantes (local)")])
    ahora = datetime.datetime.now(datetime.timezone.utc)
    alternativos = [x509.DNSName("localhost")] + [x509.IPAddress(ipaddress.ip_address(ip)) for ip in direcciones_locales()]
    certificado = (
        x509.CertificateBuilder()
        .subject_name(nombre).issuer_name(nombre)
        .public_key(clave.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(ahora - datetime.timedelta(days=1))
        .not_valid_after(ahora + datetime.timedelta(days=730))
        .add_extension(x509.SubjectAlternativeName(alternativos), critical=False)
        .sign(clave, hashes.SHA256())
    )
    CLAVE.write_bytes(clave.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL,
                                          serialization.NoEncryption()))
    CERTIFICADO.write_bytes(certificado.public_bytes(serialization.Encoding.PEM))
    return str(CERTIFICADO), str(CLAVE)
