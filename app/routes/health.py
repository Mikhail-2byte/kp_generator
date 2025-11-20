from flask import Blueprint, jsonify

from app.services.healthcheck import run_health_checks

health_bp = Blueprint('health', __name__)


@health_bp.route('/health', methods=['GET'])
@health_bp.route('/healthz', methods=['GET'])
def health():
    """Возвращает состояние ключевых зависимостей приложения."""
    report = run_health_checks()
    status_code = 200 if report['status'] == 'ok' else 503
    return jsonify(report), status_code

