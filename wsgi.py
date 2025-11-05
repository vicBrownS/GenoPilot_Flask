"""
WSGI entrypoint for production (Gunicorn).
- Exposes `app` for Gunicorn: `wsgi:app`
- Allows local `python wsgi.py` for quick tests (port configurable via $PORT)
"""
import os

try:
    # Preferred: application factory
    from app import create_app  # ajusta si tu módulo/fábrica tiene otro nombre
    app = create_app()
except ImportError:
    # Fallback: app global (por si tu proyecto no usa fábrica)
    from app import app  # type: ignore

if __name__ == "__main__":
    # Permite ejecutar localmente: `python wsgi.py`
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
