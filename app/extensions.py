from flask_wtf import CSRFProtect
from flask_login import LoginManager


csrf = CSRFProtect()

login_manager = LoginManager()
login_manager.login_view = 'auth.profile'
login_manager.login_message_category = 'info'
