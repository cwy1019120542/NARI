import os
from app.extention import celery, init_celery
from app.factory import create_app
config_name = os.getenv('FLASK_CONFIG', 'default')
app = create_app(config_name)
init_celery(celery, app)