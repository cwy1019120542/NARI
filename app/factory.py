from flask import Flask
from .config import config
from .extention import init_extention
from .views.user import user_blueprint
from .views.error import error_blueprint
from .views.main_config import main_config_blueprint
from .views.send_config import send_config_blueprint
from .views.receive_config import receive_config_blueprint
from .views.static import static_blueprint

def create_app(config_name):
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    app.register_blueprint(user_blueprint, url_prefix='/nari/user')
    app.register_blueprint(error_blueprint)
    app.register_blueprint(main_config_blueprint, url_prefix='/nari/user/<int:user_id>/main_config')
    app.register_blueprint(send_config_blueprint, url_prefix='/nari/user/<int:user_id>/send_config')
    app.register_blueprint(receive_config_blueprint, url_prefix='/nari/user/<int:user_id>/receive_config')
    app.register_blueprint(static_blueprint, url_prefix='/nari')
    init_extention(app)
    return app
