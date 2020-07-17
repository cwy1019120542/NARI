import os
from flask import current_app
from .extention import db, redis

class User(db.Model):
    __tablename__ = 'user'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(30))
    password = db.Column(db.String(100))
    name =  db.Column(db.String(10))
    department =  db.Column(db.String(10))
    status = db.Column(db.Integer, default=1)

    def get_info(self):
        return {
            "id": self.id,
            "email": self.email,
            "name": self.name,
            "department": self.department,
            "status": self.status
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
    email = db.Column(db.String(50))
    password = db.Column(db.String(100))
    status = db.Column(db.Integer, default=1)
    no_response = db.Column(db.Text)
    no_attachment = db.Column(db.Text)
    run_timestamp = db.Column(db.Integer)
    send_number = db.Column(db.Integer)
    receive_number = db.Column(db.Integer)

    def get_info(self):
        from .func_tools import get_task_info
        send_task_id = redis.get(f'main_config_{self.id}_send')
        receive_task_id = redis.get(f'main_config_{self.id}_receive')
        send_task_info = get_task_info(send_task_id)
        receive_task_info = get_task_info(receive_task_id)
        send_config_info = db.session.query(SendConfig).get(self.send_config_id).get_info() if self.send_config_id else {}
        receive_config_info = db.session.query(ReceiveConfig).get(self.receive_config_id).get_info() if self.receive_config_id else {}
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
            "no_response": self.no_response,
            "no_attachment": self.no_attachment,
            "run_timestamp": self.run_timestamp,
            "send_number": self.send_number,
            "receive_number": self.receive_number,
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
    agreement = db.Column(db.String(10))
    status = db.Column(db.Integer, default=1)

    def get_info(self):
        return {
                "id": self.id,
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
                "agreement": self.agreement,
                "status": self.status
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
    remind_agreement = db.Column(db.String(10))
    remind_subject = db.Column(db.String(50))
    remind_content = db.Column(db.Text)
    create_timestamp = db.Column(db.Integer)
    change_timestamp = db.Column(db.Integer)
    ip = db.Column(db.String(50))
    port = db.Column(db.Integer)
    agreement = db.Column(db.String(10))
    read_start_timestamp = db.Column(db.Integer)
    read_end_timestamp = db.Column(db.Integer)
    status = db.Column(db.Integer, default=1)

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
                "remind_agreement": self.remind_agreement,
                "remind_subject": self.remind_subject,
                "remind_content": self.remind_content,
                "create_timestamp": self.create_timestamp,
                "change_timestamp": self.change_timestamp,
                "ip": self.ip,
                "port": self.port,
                "agreement": self.agreement,
                "read_start_timestamp": self.read_start_timestamp,
                "read_end_timestamp": self.read_end_timestamp,
                "status": self.status
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

