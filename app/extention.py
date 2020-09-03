import os
import logging
import logging.handlers
import datetime
from celery import Celery
from redis import Redis
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine

db = SQLAlchemy()
bcrypt = Bcrypt()
redis = Redis(db=0, host=os.environ.get('REDIS_HOST', 'localhost'), port=6379)
BROKER_URL = os.getenv('BROKER_URL', 'redis://localhost:6379/0')
CELERY_RESULT_BACKEND = os.getenv('CELERY_RESULT_BACKEND', 'redis://localhost:6379/0')
celery = Celery(__name__, backend=CELERY_RESULT_BACKEND, broker=BROKER_URL)
engine = create_engine(os.environ.get('SQLALCHEMY_DATABASE_URI', 'mysql://cwy:never1019120542,@localhost:3306/NARI'), pool_timeout=3600, pool_recycle=60)
Session = sessionmaker(bind=engine)

send_logger = logging.getLogger('send')
send_logger.setLevel(logging.DEBUG)
receive_logger = logging.getLogger('receive')
receive_logger.setLevel(logging.DEBUG)
# rf_handler = logging.handlers.TimedRotatingFileHandler('all.log', when='midnight', interval=1, backupCount=7, atTime=datetime.time(0, 0, 0, 0))
# rf_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

send_handler = logging.FileHandler('/var/log/send.log')
send_handler.setLevel(logging.DEBUG)
send_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(filename)s[:%(lineno)d] - %(message)s"))

receive_handler = logging.FileHandler('/var/log/receive.log')
receive_handler.setLevel(logging.DEBUG)
receive_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(filename)s[:%(lineno)d] - %(message)s"))

send_logger.addHandler(send_handler)
receive_logger.addHandler(receive_handler)
def init_extention(app):
    db.init_app(app)
    bcrypt.init_app(app)

def init_celery(celery, app):
    celery.conf.update(app.config)
    TaskBase = celery.Task
    class ContextTask(TaskBase):
        def __call__(self, *args, **kwargs):
            with app.app_context():
                return TaskBase.__call__(self, *args, **kwargs)
    celery.Task = ContextTask