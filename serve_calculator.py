from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


HOST = "127.0.0.1"
PORT = 8000
SITE_DIR = Path(__file__).resolve().parent / "calculator-site"


def main():
    if not SITE_DIR.exists():
        raise SystemExit(f"Missing site folder: {SITE_DIR}")

    handler = lambda *args, **kwargs: SimpleHTTPRequestHandler(
        *args,
        directory=str(SITE_DIR),
        **kwargs,
    )

    server = ThreadingHTTPServer((HOST, PORT), handler)
    print(f"Serving calculator at http://{HOST}:{PORT}/")
    print("Press Ctrl+C to stop.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping calculator server.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
