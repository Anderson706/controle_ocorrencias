#!/usr/bin/env python3
"""
CCTV Control Panel — modo SERVIDOR (acesso via navegador na rede).

Roda o app Flask com waitress (servidor WSGI de produção) ouvindo em TODAS as
interfaces de rede, para que outras máquinas acessem por
http://<ip-do-servidor>:<porta> pelo navegador — sem instalar nada no cliente.

IMPORTANTE:
  - Só ESTA máquina (o servidor) precisa de acesso à VPN/Oracle.
  - É preciso liberar a porta no Firewall do Windows (ver instruções abaixo).

Uso:
    python serve.py                  # porta padrão 5000, HTTP
    python serve.py --port 8080
    python serve.py --threads 8
    python serve.py --https          # HTTPS com certificado autoassinado
                                      # (necessário p/ câmera funcionar em
                                      # tablets/celulares acessando por IP —
                                      # navegador só libera getUserMedia()
                                      # em HTTPS ou localhost)

Liberar a porta no Firewall (PowerShell como Administrador, uma vez):
    New-NetFirewallRule -DisplayName "CCTV Control Panel" `
        -Direction Inbound -Protocol TCP -LocalPort 5000 -Action Allow

Com --https, o navegador do tablet vai mostrar um aviso de "conexão não
seguro" na primeira vez (o certificado é autoassinado, não vem de uma CA
confiável) — é só avançar/aceitar uma vez por dispositivo. A partir daí a
câmera funciona normalmente.
"""
import argparse
import datetime
import ipaddress
import os
import socket
import ssl
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _gerar_certificado_self_signed(cert_path, key_path, ips):
    """Gera um certificado TLS autoassinado (válido ~2 anos) cobrindo localhost
    e os IPs de rede local desta máquina, e grava em disco (reaproveitado nas
    próximas execuções — não gera de novo a cada start)."""
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    nome = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "CCTV Control Panel")])

    alt_names = [x509.DNSName("localhost"), x509.IPAddress(ipaddress.ip_address("127.0.0.1"))]
    for ip in ips:
        try:
            alt_names.append(x509.IPAddress(ipaddress.ip_address(ip)))
        except ValueError:
            pass

    agora = datetime.datetime.now(datetime.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(nome)
        .issuer_name(nome)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(agora - datetime.timedelta(days=1))
        .not_valid_after(agora + datetime.timedelta(days=825))
        .add_extension(x509.SubjectAlternativeName(alt_names), critical=False)
        .sign(key, hashes.SHA256())
    )

    with open(key_path, "wb") as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))
    with open(cert_path, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))


def _real_ssl_context(protocol):
    """Cria um ssl.SSLContext genuíno da stdlib para uso de SERVIDOR.

    Em redes com proxy de inspeção TLS (caso da DHL), o pacote pip-system-certs
    substitui ssl.SSLContext globalmente por um wrapper (truststore) pensado só
    para verificação do lado CLIENTE — usado num socket servidor, ele quebra em
    wrap_socket(server_side=True). Pegamos a classe original guardada pelo
    próprio truststore antes do patch, sem desfazer o patch pro resto do processo
    (outras partes do app, ex. conexão Oracle, podem depender dele)."""
    try:
        from pip._vendor.truststore._api import _original_SSLContext
        return _original_SSLContext(protocol)
    except Exception:
        return ssl.SSLContext(protocol)


def _lan_ips():
    """Descobre os IPs de rede local desta máquina (para imprimir a URL de acesso)."""
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None):
            ip = info[4][0]
            if ":" not in ip and not ip.startswith("127."):
                ips.add(ip)
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))   # não envia nada — só resolve a interface de saída
        ips.add(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    return sorted(ips)


def main():
    parser = argparse.ArgumentParser(description="CCTV Control Panel — servidor web")
    parser.add_argument("--host", default="0.0.0.0", help="interface de escuta (padrão: todas)")
    parser.add_argument("--port", type=int, default=5000, help="porta (padrão: 5000)")
    parser.add_argument("--threads", type=int, default=8,
                        help="threads do waitress (padrão: 8; mantenha <= pool do Oracle)")
    parser.add_argument("--https", action="store_true",
                        help="serve via HTTPS com certificado autoassinado "
                             "(necessário p/ câmera funcionar em tablets/celulares)")
    args = parser.parse_args()

    # Dimensiona o pool Oracle ANTES de importar o app: uma conexão "quente" por
    # thread + folga, evitando abrir conexões sob demanda (caro via VPN) nos picos.
    os.environ.setdefault("CCTV_POOL_SIZE",    str(args.threads))
    os.environ.setdefault("CCTV_MAX_OVERFLOW", str(max(4, args.threads // 2)))

    print("Carregando aplicação...", flush=True)
    from app import app, _init_db

    print("Aplicando migrações de schema (rápido após a 1ª vez)...", flush=True)
    try:
        _init_db()
    except Exception as exc:
        print(f"  Aviso: _init_db falhou ({exc}). Seguindo mesmo assim.", flush=True)

    ips    = _lan_ips()
    scheme = "https" if args.https else "http"

    print("\n" + "=" * 60)
    print("  CCTV Control Panel — SERVIDOR ATIVO" + (" (HTTPS)" if args.https else ""))
    print("=" * 60)
    print(f"  Porta: {args.port}   Threads: {args.threads}")
    print("  Acesse de outras máquinas da rede pelo navegador:")
    if ips:
        for ip in ips:
            print(f"     ->  {scheme}://{ip}:{args.port}")
    else:
        print(f"     ->  {scheme}://<ip-desta-maquina>:{args.port}")
    print(f"  Nesta máquina:  {scheme}://localhost:{args.port}")
    if args.https:
        print("  (certificado autoassinado — aceite o aviso do navegador na 1ª visita)")
    print("=" * 60)
    print("  (Ctrl+C para parar o servidor)\n", flush=True)

    if args.https:
        # waitress não tem suporte nativo a TLS (seu loop não-bloqueante não
        # conduz o handshake corretamente quando a gente embrulha o socket à
        # força — foi isso que causava o reset de conexão). O servidor de dev
        # do Flask/werkzeug trata SSL nativamente (1 thread bloqueante por
        # conexão), então usamos ele só para o modo --https.
        cert_path = os.path.join(BASE_DIR, "_server_cert.pem")
        key_path  = os.path.join(BASE_DIR, "_server_key.pem")
        if not (os.path.exists(cert_path) and os.path.exists(key_path)):
            print("Gerando certificado autoassinado (1ª vez nesta máquina)...", flush=True)
            _gerar_certificado_self_signed(cert_path, key_path, ips)
        ctx = _real_ssl_context(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(cert_path, key_path)
        app.run(host=args.host, port=args.port, threaded=True,
                debug=False, use_reloader=False, ssl_context=ctx)
    else:
        try:
            from waitress import serve
        except ImportError:
            print("ERRO: waitress não instalado. Rode:  pip install waitress", file=sys.stderr)
            sys.exit(1)
        serve(
            app,
            host=args.host,
            port=args.port,
            threads=args.threads,
            ident=None,           # não anuncia "Server: waitress" nos cabeçalhos
            channel_timeout=120,  # encerra conexões ociosas/lentas após 120s
            cleanup_interval=30,
        )


if __name__ == "__main__":
    main()
