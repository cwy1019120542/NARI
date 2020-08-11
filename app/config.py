import os
import sys

class BaseConfig:
    BASE_DIR = '/home/cwy'
    CONFIG_FILES_DIR = os.path.join(BASE_DIR, 'config_files')
    STATIC_FILES_DIR = os.path.join(BASE_DIR, 'static_files')
    SQLALCHEMY_DATABASE_URI = os.environ.get('SQLALCHEMY_DATABASE_URI', 'mysql://cwy:never1019120542,@localhost:3306/NARI')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    BROKER_URL = os.environ.get('BROKER_URL', 'redis://localhost:6379/0')
    CELERY_RESULT_BACKEND = os.environ.get('CELERY_RESULT_BACKEND', 'redis://localhost:6379/0')
    CELERYD_CONCURRENCY = 20
    # SQLALCHEMY_ECHO = True
    CELERY_TIMEZONE = 'Asia/Shanghai'
    CELERY_ACCEPT_CONTENT = ['pickle', 'json', 'msgpack', 'yaml']
    CELERY_IMPORTS = (
        'app.tasks.receive_mail', 'app.tasks.send_mail'
    )
    SECRET_KEY = 'ddc81ea8-ac80-11ea-b922-507b9d122005'
    SERVER_MAIL_IP = os.environ.get('SERVER_MAIL_IP', 'smtp.qq.com')
    SERVER_MAIL_USER = os.environ.get('SERVER_MAIL_USER', '1019120542@qq.com')
    SERVER_MAIL_PASSWORD = os.environ.get('SERVER_MAIL_PASSWORD', 'aswhfohkwtaebfcc')

class DefaultConfig(BaseConfig):
    DEBUG = False

class DevelopConfig(BaseConfig):
    DEBUG = True

class ProductConfig(BaseConfig):
    DEBUG = False

config = {
    "default": DefaultConfig,
    "develop": DevelopConfig,
    "product": ProductConfig
}