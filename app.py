import os

from app import create_app


def _env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return str(value).lower() in {'1', 'true', 'yes', 'on'}


app = create_app()


if __name__ == '__main__':
    host = os.environ.get('FLASK_RUN_HOST', '0.0.0.0')
    port = int(os.environ.get('FLASK_RUN_PORT', 5000))
    debug = _env_flag('FLASK_DEBUG', default=True)
    use_waitress = _env_flag('USE_WAITRESS', default=False) or os.environ.get('FLASK_ENV') == 'production'

    print(f"\n{'='*60}")
    print(f"  KP Generator Flask Application")
    print(f"{'='*60}")
    print(f"  Host: {host}")
    print(f"  Port: {port}")
    print(f"  Debug: {debug}")
    print(f"  Server: {'Waitress' if use_waitress else 'Flask Development'}")
    print(f"{'='*60}")
    print(f"\n  Open in browser: http://localhost:{port}")
    print(f"  AI Agent: http://localhost:{port}/ai-agent")
    print(f"  Admin Panel: http://localhost:{port}/admin/ai-agent")
    print(f"\n{'='*60}\n")

    if use_waitress:
        try:
            from waitress import serve
        except ModuleNotFoundError as exc:  # pragma: no cover - runtime fallback
            raise RuntimeError('waitress is required for production runs. Install via pip install waitress') from exc

        print("Starting Waitress server...")
        serve(app, host=host, port=port)
    else:
        print("Starting Flask development server...")
        app.run(debug=debug, host=host, port=port)
