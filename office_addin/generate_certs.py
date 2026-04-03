"""
Genera certificados SSL autofirmados para el Office Add-in.
Los Office Add-ins requieren HTTPS obligatoriamente.

Uso: python generate_certs.py
"""

import subprocess
import os
import sys

CERT_DIR = os.path.join(os.path.dirname(__file__), 'certs')
CERT_FILE = os.path.join(CERT_DIR, 'server.crt')
KEY_FILE = os.path.join(CERT_DIR, 'server.key')


def generate_with_openssl():
    """Genera certificados usando OpenSSL"""
    os.makedirs(CERT_DIR, exist_ok=True)

    cmd = [
        'openssl', 'req', '-x509', '-newkey', 'rsa:2048',
        '-keyout', KEY_FILE,
        '-out', CERT_FILE,
        '-days', '365',
        '-nodes',
        '-subj', '/CN=localhost/O=ExcelIAChat/C=ES',
        '-addext', 'subjectAltName=DNS:localhost,IP:127.0.0.1'
    ]

    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


def generate_with_python():
    """Genera certificados usando Python (cryptography library)"""
    try:
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        import datetime
    except ImportError:
        print("  Instalando cryptography...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'cryptography'])
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        import datetime
        import ipaddress

    os.makedirs(CERT_DIR, exist_ok=True)

    # Generar clave privada
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    # Generar certificado
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "ES"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Excel IA Chat"),
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
    ])

    import ipaddress as ipmod

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.utcnow())
        .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=365))
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(ipmod.IPv4Address("127.0.0.1")),
            ]),
            critical=False,
        )
        .sign(key, hashes.SHA256())
    )

    # Guardar
    with open(KEY_FILE, "wb") as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption()
        ))

    with open(CERT_FILE, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

    return True


def install_cert_windows():
    """Instala el certificado en Windows para que sea confiable"""
    print()
    print("  Para que Excel confíe en el certificado, instálalo:")
    print()
    print("  OPCIÓN A - Automático (requiere admin):")
    print(f'    certutil -addstore -user "Root" "{CERT_FILE}"')
    print()
    print("  OPCIÓN B - Manual:")
    print(f"    1. Haz doble clic en: {CERT_FILE}")
    print("    2. 'Instalar certificado' > 'Usuario actual'")
    print("    3. 'Colocar todos los certificados en el siguiente almacén'")
    print("    4. Examinar > 'Entidades de certificación raíz de confianza'")
    print("    5. Siguiente > Finalizar")
    print()


if __name__ == '__main__':
    print("=" * 55)
    print("  Generador de certificados SSL - Office Add-in")
    print("=" * 55)
    print()

    if os.path.exists(CERT_FILE):
        print(f"  Ya existen certificados en: {CERT_DIR}")
        resp = input("  ¿Regenerar? (s/N): ").strip().lower()
        if resp != 's':
            print("  Cancelado.")
            sys.exit(0)

    print("  Generando certificados SSL...")

    success = generate_with_openssl()
    if not success:
        print("  OpenSSL no disponible, usando Python...")
        success = generate_with_python()

    if success:
        print()
        print(f"  Certificado: {CERT_FILE}")
        print(f"  Clave:       {KEY_FILE}")
        print()
        print("  Certificados generados correctamente!")

        if sys.platform == 'win32':
            install_cert_windows()
    else:
        print("  ERROR: No se pudieron generar los certificados.")
        sys.exit(1)
