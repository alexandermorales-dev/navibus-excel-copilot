"""
Dev server for Excel AI Copilot.
Starts an HTTPS server on https://localhost:3000 with a self-signed certificate.
Required for Office Add-in sideloading (Office rejects HTTP).

Usage:
    python dev_server.py

First run: generates a self-signed certificate and installs it to the
Windows Trusted Root store so Office trusts it.
"""

import http.server
import ssl
import os
import sys
import subprocess
import datetime

PORT = 3000
DIR = os.path.dirname(os.path.abspath(__file__))
CERT_DIR = os.path.join(DIR, ".certs")
CERT_FILE = os.path.join(CERT_DIR, "localhost.crt")
KEY_FILE = os.path.join(CERT_DIR, "localhost.key")

def generate_cert():
    """Generate a self-signed certificate using Python's cryptography library."""
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa

    os.makedirs(CERT_DIR, exist_ok=True)

    print("Generating self-signed certificate...")

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
    ])

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.now(datetime.timezone.utc))
        .not_valid_after(datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=365))
        .add_extension(
            x509.SubjectAlternativeName([x509.DNSName("localhost")]),
            critical=False,
        )
        .sign(key, hashes.SHA256())
    )

    # Write cert as PEM
    with open(CERT_FILE, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

    # Write key as PEM
    with open(KEY_FILE, "wb") as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))

    print("Certificate generated.")

    # Install to Windows Trusted Root store so Office trusts it
    install_to_trusted_root(CERT_FILE)
    return True

def install_to_trusted_root(cert_path):
    """Add the cert to the Windows CurrentUser Trusted Root store."""
    print("Installing certificate to Trusted Root store...")
    ps_script = f'''
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("{cert_path}")
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()
Write-Output "Cert installed to Trusted Root."
'''
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0:
            print(result.stdout.strip())
        else:
            print(f"Warning: Could not auto-install to Trusted Root: {result.stderr.strip()}")
            print("You may see a security warning when loading the add-in. This is normal for dev certs.")
    except Exception as e:
        print(f"Warning: Could not install cert to Trusted Root: {e}")
        print("You may see a security warning when loading the add-in. This is normal for dev certs.")

def main():
    if not os.path.exists(CERT_FILE) or not os.path.exists(KEY_FILE):
        if not generate_cert():
            print("\nFailed to generate certificates. Cannot start HTTPS server.")
            print("Alternative: Host on GitHub Pages — see README.md.")
            sys.exit(1)

    print(f"\nStarting HTTPS server on https://localhost:{PORT}")
    print(f"Serving files from: {DIR}")
    print(f"\nTo sideload in Excel:")
    print(f"  1. Open Excel")
    print(f"  2. Insert > Add-ins > Manage My Add-ins > Upload My Add-in")
    print(f"  3. Select: {os.path.join(DIR, 'manifest-localhost.xml')}")
    print(f"\nPress Ctrl+C to stop.\n")

    # Change to project directory so SimpleHTTPRequestHandler serves from there
    os.chdir(DIR)

    # Custom handler with no-cache headers (prevents Office task pane from caching JS)
    class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
        def end_headers(self):
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
            super().end_headers()

    handler = NoCacheHandler
    httpd = http.server.HTTPServer(("localhost", PORT), handler)

    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=CERT_FILE, keyfile=KEY_FILE)
    httpd.socket = context.wrap_socket(httpd.socket, server_side=True)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
        httpd.shutdown()

if __name__ == "__main__":
    main()
