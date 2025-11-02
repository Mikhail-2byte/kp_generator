import os
import json
import secrets
import logging
from logging.handlers import RotatingFileHandler
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()

def load_config(app):
    """Загружает конфигурацию из settings.json и переменных окружения"""
    try:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config', 'settings.json')
        
        config_data = {}
        
        # Загружаем из JSON если файл существует
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                app.logger.info(f'Config loaded successfully: {config_path}')
        else:
            app.logger.warning(f'Config file not found: {config_path}')
            
        # Переопределяем чувствительные данные из переменных окружения
        debug_flag = os.environ.get('DEBUG')
        if debug_flag is None:
            debug_flag = os.environ.get('FLASK_DEBUG', 'False')

        env_config = {
            'secret_key': os.environ.get('SECRET_KEY'),
            'database_url': os.environ.get('DATABASE_URL', 'sqlite:///kp_generator.db'),
            'debug': str(debug_flag).lower() == 'true',
            'log_level': os.environ.get('LOG_LEVEL', 'INFO'),
        }
        
        config_data.update(env_config)
        return config_data
        
    except Exception as e:
        app.logger.error(f'Error loading config: {e}')
        return {}

def setup_app_security(app, config):
    """Настраивает безопасность приложения"""
    secret_key = config.get('secret_key') or secrets.token_hex(32)
    app.secret_key = secret_key
    
    # Настройка Flask на основе конфигурации
    app.config['DEBUG'] = config.get('debug', False)
    app.config['TESTING'] = False
    
    return secret_key

def setup_logging(app, config):
    """Настраивает логирование"""
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    log_level = getattr(logging, config.get('log_level', 'INFO'))
    
    file_handler = RotatingFileHandler('logs/kp_generator.log', maxBytes=10240, backupCount=10)
    file_handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'))
    file_handler.setLevel(log_level)
    
    error_handler = RotatingFileHandler('logs/kp_generator_errors.log', maxBytes=10240, backupCount=5)
    error_handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'))
    error_handler.setLevel(logging.ERROR)
    
    app.logger.addHandler(file_handler)
    app.logger.addHandler(error_handler)
    app.logger.setLevel(log_level)
