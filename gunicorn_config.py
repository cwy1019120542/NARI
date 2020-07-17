import os
from gevent import monkey
monkey.patch_all()
debug = True
bind = os.environ.get('GUNICORN_BIND', '127.0.0.1')
loglevel = 'debug'
pidfile = 'app/logs/gunicorn_pid.txt'
logfile = 'app/logs/gunicorn_log.txt'
workers = 4
worker_class = 'gevent'