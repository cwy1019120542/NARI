#!/bin/bash
source /etc/profile
service vsftpd restart
service ssh restart
service mysql restart
service redis-server restart
rm /var/log/celery_nari.log /var/log/flask_nari.log /var/log/send.log /var/log/receive.log
touch /var/log/celery_nari.log /var/log/flask_nari.log /var/log/send.log /var/log/receive.log
#nohup celery worker -A /home/cwy/NARI/celery_worker.celery >> /var/log/celery_nari.log 2>&1 &
#nohup gunicorn /home/cwy/PycharmProjects/NARIProjects/NARI/manage:app -c /home/cwy/PycharmProjects/NARIProjects/NARI/gunicorn_config.py >> /var/log/flask_nari.log 2>&1 &