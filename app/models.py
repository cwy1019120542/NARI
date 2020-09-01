import os
from flask import current_app
from .extention import db, redis

class User(db.Model):
    __tablename__ = 'user'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(100))
    password = db.Column(db.String(100))
    name =  db.Column(db.String(10))
    department =  db.Column(db.String(10))
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)

    def get_info(self):
        return {
            "id": self.id,
            "email": self.email,
            "name": self.name,
            "department": self.department,
            "status": self.status,
            "create_timestamp": self.create_timestamp,
            "change_timestamp": self.change_timestamp
        }

class MainConfig(db.Model):
    __tablename__ = 'main_config'
    id = db.Column(db.Integer, primary_key=True)
    config_name = db.Column(db.String(50))
    function_type = db.Column(db.Integer)
    send_config_id = db.Column(db.Integer)
    receive_config_id = db.Column(db.Integer)
    user_id = db.Column(db.Integer)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    email = db.Column(db.String(100))
    password = db.Column(db.String(100))
    status = db.Column(db.Integer, default=1)
    run_timestamp = db.Column(db.Integer)

    def get_son_info(self, model, config_id):
        if config_id:
            config = db.session.query(model).filter_by(id=config_id).first()
            if config:
                config_info = config.get_info()
            else:
                config_info = {}
        else:
            config_info = {}
        return config_info

    def get_info(self):
        from .func_tools import get_task_info
        send_task_id = redis.get(f'{self.id}_send_task_id')
        receive_task_id = redis.get(f'{self.id}_receive_task_id')
        send_task_info = get_task_info(send_task_id)
        receive_task_info = get_task_info(receive_task_id)
        send_config_info = self.get_son_info(SendConfig, self.send_config_id)
        receive_config_info = self.get_son_info(ReceiveConfig, self.receive_config_id)
        config_files_dir = current_app.config['CONFIG_FILES_DIR']
        from .func_tools import get_file_name
        match_excel = get_file_name(config_files_dir, self.id, 'match_excel')
        return {
            "id": self.id,
            "config_name": self.config_name,
            "function_type": self.function_type,
            "send_config_info": send_config_info,
            "receive_config_info": receive_config_info,
            "user_id": self.user_id,
            "create_timestamp": self.create_timestamp,
            "change_timestamp": self.change_timestamp,
            "email": self.email,
            "password": self.password,
            "status": self.status,
            "run_timestamp": self.run_timestamp,
            "send_task_info": send_task_info,
            "receive_task_info": receive_task_info,
            "match_excel": match_excel
        }

class SendConfig(db.Model):
    __tablename__ = 'send_config'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer)
    subject = db.Column(db.String(50))
    content = db.Column(db.Text)
    sheet = db.Column(db.Text)
    field_row = db.Column(db.String(50))
    split_field = db.Column(db.String(50))
    is_split = db.Column(db.String(50))
    is_timing = db.Column(db.Boolean)
    start_timestamp = db.Column(db.Integer)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    ip = db.Column(db.String(50))
    port = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)
    send_excel_name = db.Column(db.String(50))
    send_excel_uuid_name = db.Column(db.String(100))
    send_excel_field = db.Column(db.Text)

    def get_info(self):
        return {
                "id": self.id,
                "user_id": self.user_id,
                "subject": self.subject,
                "content": self.content,
                "sheet": self.sheet,
                "field_row": self.field_row,
                "split_field": self.split_field,
                "is_split": self.is_split,
                "is_timing": self.is_timing,
                "start_timestamp": self.start_timestamp,
                "create_timestamp": self.create_timestamp,
                "change_timestamp": self.change_timestamp,
                "ip": self.ip,
                "port": self.port,
                "status": self.status,
                "send_excel_name": self.send_excel_name,
                "send_excel_uuid_name": self.send_excel_uuid_name,
                "send_excel_field": self.send_excel_field
            }

class ReceiveConfig(db.Model):
    __tablename__ = 'receive_config'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer)
    subject = db.Column(db.String(50))
    sheet_info = db.Column(db.Text)
    is_timing = db.Column(db.Boolean)
    start_timestamp = db.Column(db.Integer)
    is_remind = db.Column(db.Boolean)
    remind_ip = db.Column(db.String(50))
    remind_port = db.Column(db.Integer)
    remind_subject = db.Column(db.String(50))
    remind_content = db.Column(db.Text)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    ip = db.Column(db.String(50))
    port = db.Column(db.Integer)
    read_start_timestamp = db.Column(db.Integer)
    read_end_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)
    template_excel_name = db.Column(db.String(50))
    template_excel_uuid_name = db.Column(db.String(100))
    template_excel_field = db.Column(db.Text)

    def get_info(self):
        return {
                "id": self.id,
                "user_id": self.user_id,
                "subject": self.subject,
                "sheet_info": self.sheet_info,
                "is_timing": self.is_timing,
                "start_timestamp": self.start_timestamp,
                "is_remind": self.is_remind,
                "remind_ip": self.remind_ip,
                "remind_port": self.remind_port,
                "remind_subject": self.remind_subject,
                "remind_content": self.remind_content,
                "create_timestamp": self.create_timestamp,
                "change_timestamp": self.change_timestamp,
                "ip": self.ip,
                "port": self.port,
                "read_start_timestamp": self.read_start_timestamp,
                "read_end_timestamp": self.read_end_timestamp,
                "status": self.status,
                "template_excel_name": self.template_excel_name,
                "template_excel_uuid_name": self.template_excel_uuid_name,
                "template_excel_field": self.template_excel_field
            }

class UpdateMessage(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)

    def get_info(self):
        return {
            "id": self.id,
            "content": self.content,
            "create_timestamp": self.create_timestamp,
            "change_timestamp": self.change_timestamp,
            "status": self.status
        }

class SendHistory(db.Model):
    __tablename__ = 'send_history'
    id = db.Column(db.Integer, primary_key=True)
    target = db.Column(db.String(50))
    email = db.Column(db.String(100))
    main_config_id = db.Column(db.Integer)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)
    is_success = db.Column(db.Boolean)
    message = db.Column(db.String(50))

    def get_info(self):
        return {
            "id": self.id,
            "target": self.target,
            "email": self.email,
            "main_config_id": self.main_config_id,
            "create_timestamp": self.create_timestamp,
            "is_success": self.is_success,
            "message": self.message,
            "status": self.status
        }

class ReceiveHistory(db.Model):
    __tablename__ = 'receive_history'
    id = db.Column(db.Integer, primary_key=True)
    target = db.Column(db.String(50))
    email = db.Column(db.String(100))
    main_config_id = db.Column(db.Integer)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)
    is_success = db.Column(db.Boolean)
    message = db.Column(db.String(50))

    def get_info(self):
        return {
            "id": self.id,
            "target": self.target,
            "email": self.email,
            "main_config_id": self.main_config_id,
            "create_timestamp": self.create_timestamp,
            "is_success": self.is_success,
            "message": self.message,
            "status": self.status
        }

class SapConfig(db.Model):
    __tablename__ = "sap_config"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer)
    account = db.Column(db.String(50))
    password = db.Column(db.String(100))
    main_body = db.Column(db.String(100))
    subject = db.Column(db.String(100))
    start_date = db.Column(db.String(20))
    end_date = db.Column(db.String(20))
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)

    def get_info(self):
        return {
            "id": self.id,
            "user_id": self.user_id,
            "account": self.account,
            "password": self.password,
            "main_body": self.main_body,
            "subject": self.subject,
            "start_date": self.start_date,
            "end_date": self.end_date,
            "create_timestamp": self.create_timestamp,
            "change_timestamp": self.change_timestamp,
            "status": self.status
        }
