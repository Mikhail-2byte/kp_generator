from flask import render_template

from app.ui import build_context


def register_error_handlers(app):
    @app.errorhandler(404)
    def not_found_error(_error):
        return render_template(
            '404.html',
            **build_context('index', 'Страница не найдена')
        ), 404

    @app.errorhandler(500)
    def internal_error(_error):
        return render_template(
            '500.html',
            **build_context('index', 'Внутренняя ошибка')
        ), 500
