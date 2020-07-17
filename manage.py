import os
from flask_cors import CORS
from flask_script import Manager
from flask_migrate import Migrate, MigrateCommand
from app.factory import create_app
from app.extention import db
from app.func_tools import response

config_name = os.getenv('FLASK_CONFIG', 'default')
app = create_app(config_name)
CORS(app, supports_credentials=True)
migrate = Migrate(app, db)
manage = Manager(app)
manage.add_command('db', MigrateCommand)

if __name__ == '__main__':
    manage.run()