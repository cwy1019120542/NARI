import os
import json
import time
import zipfile
import smtplib
from flask import Blueprint, request, current_app
from ..func_tools import is_login, response,file_resource, resource_manage, get_run_config, create_task_id, resource_limit, parameter_check, return_file, get_file_path, send_multi_mail, return_zip
from ..models import User, MainConfig, SendConfig, ReceiveConfig, SendHistory, ReceiveHistory
from ..parameter_config import main_config_parameter, remind_parameter, send_history_parameter, receive_history_parameter
from ..extention import db, redis, celery
from ..tasks.send_mail import send_mail
from ..tasks.receive_mail import receive_mail

main_config_blueprint = Blueprint('Main_config', __name__)

@main_config_blueprint.route('', methods=['GET', 'POST'])
@main_config_blueprint.route('/<int:main_config_id>', methods=['GET', 'PUT', 'DELETE'])
@is_login
def main_config(user_id, main_config_id=None):
    request_args = request.args
    request_json = request.json
    request_method = request.method
    if request_method in ["POST", "PUT"]:
        is_right, clean_data = parameter_check(request_json, [('send_config_id', int, True, 11), ('receive_config_id', int, True, 11)], False)
        if not is_right:
            return clean_data
        send_config_id = clean_data.get('send_config_id')
        if send_config_id:
            is_success, return_data = resource_limit([(User, user_id, None), (SendConfig, send_config_id, "user_id")])
            if not is_success:
                return return_data
        receive_config_id = clean_data.get('receive_config_id')
        if receive_config_id:
            is_success, return_data = resource_limit([(User, user_id, None), (ReceiveConfig, receive_config_id, "user_id")])
            if not is_success:
                return return_data
    return resource_manage([(User, user_id, None), (MainConfig, main_config_id, "user_id")], request_method, request_args, request_json, main_config_parameter)

@main_config_blueprint.route('/<int:main_config_id>/match_excel', methods=["GET", "POST", "PUT", "DELETE"])
@is_login
def match_excel(user_id, main_config_id):
    config_files_dir = current_app.config['CONFIG_FILES_DIR']
    excel_dir = os.path.join(config_files_dir, str(main_config_id), 'match_excel')
    return file_resource([(User, user_id, None), (MainConfig, main_config_id, "user_id")], excel_dir, request.method, "match_excel", request.files)

@main_config_blueprint.route('/<int:main_config_id>/send_excel', methods=["GET", "POST", "PUT", "DELETE"])
@is_login
def send_excel(user_id, main_config_id):
    config_files_dir = current_app.config['CONFIG_FILES_DIR']
    excel_dir = os.path.join(config_files_dir, str(main_config_id), 'send_excel')
    return file_resource([(User, user_id, None), (MainConfig, main_config_id, "user_id")], excel_dir, request.method, "send_excel", request.files)

@main_config_blueprint.route('/<int:main_config_id>/start', methods=["POST"])
@is_login
def start(user_id, main_config_id):
    is_success, return_data = resource_limit([(User, user_id, None), (MainConfig, main_config_id, "user_id")])
    if not is_success:
        return return_data
    active_task_id_list, scheduled_task_id_list = get_run_config(user_id)
    if main_config_id in active_task_id_list or main_config_id in scheduled_task_id_list:
        return response(False, 403, "该配置已在运行")
    main_config_query, main_config = return_data[1:3]
    function_type = main_config.function_type
    is_send = False
    is_receive = False
    if function_type == 1 or function_type == 3:
        send_config_id = main_config.send_config_id
        if send_config_id:
            is_success, return_data = resource_limit([(User, user_id, None), (SendConfig, send_config_id, "user_id")])
            if is_success:
                send_config_query, send_config = return_data[1:3]
                is_send = True
            else:
                return return_data
        else:
            return response(False, 400, "发邮件配置未设置")
    if function_type == 2 or function_type == 3:
        receive_config_id = main_config.receive_config_id
        if receive_config_id:
            is_success, return_data = resource_limit([(User, user_id, None), (ReceiveConfig, receive_config_id, "user_id")])
            if is_success:
                receive_config_query, receive_config = return_data[1:3]
                is_receive = True
            else:
                return return_data
        else:
            return response(False, 400, "收邮件配置未设置")
    app_config = current_app.config
    config = {
        "CONFIG_FILES_DIR": app_config["CONFIG_FILES_DIR"]
    }
    main_config_info = main_config.get_info()
    now_time = int(time.time())
    if is_send:
        send_task_id = create_task_id('send_mail', user_id=user_id, send_config_id=send_config_id, main_config_id=main_config_id)
        send_config_info = send_config.get_info()
        countdown = 0
        if send_config_info['is_timing']:
            start_time = send_config_info['start_timestamp']
            if not start_time or start_time < now_time:
                return response(False, 400, "定时时间错误")
            countdown = start_time - int(time.time())
        redis.set(f'main_config_{main_config_id}_send', send_task_id)
        send_mail.apply_async(kwargs={"config": config, "main_config_info": main_config_info, "send_config_info": send_config_info}, task_id=send_task_id, countdown=countdown)
    if is_receive:
        receive_task_id = create_task_id('receive_mail', user_id=user_id, receive_config_id=receive_config_id, main_config_id=main_config_id)
        receive_config_info = receive_config.get_info()
        countdown = 0
        if receive_config_info['is_timing']:
            start_time = receive_config_info['start_timestamp']
            if not start_time or start_time < now_time:
                return response(False, 400, "定时时间错误")
            countdown = start_time - int(time.time())
        redis.set(f'main_config_{main_config_id}_receive', receive_task_id)
        receive_mail.apply_async(kwargs={"config":config, "main_config": main_config_info, "receive_config": receive_config_info}, task_id=receive_task_id, countdown=countdown)
    return response(True, 202, "成功")

@main_config_blueprint.route('/<int:main_config_id>/stop', methods=['POST'])
@is_login
def stop(user_id, main_config_id):
    is_success, return_data = resource_limit([(User, user_id, None), (MainConfig, main_config_id, "user_id")])
    if not is_success:
        return return_data
    send_task_id = redis.get(f'main_config_{main_config_id}_send')
    receive_task_id = redis.get(f'main_config_{main_config_id}_receive')
    print(send_task_id, receive_task_id)
    if send_task_id:
        celery.control.revoke(send_task_id.decode(), terminate=True)
    if receive_task_id:
        celery.control.revoke(receive_task_id.decode(), terminate=True)
    return response(True, 200, "成功")

@main_config_blueprint.route('/<int:main_config_id>/template_excel', methods=["GET", "POST", "PUT"])
@is_login
def template_excel(user_id, main_config_id):
    config_files_dir = current_app.config['CONFIG_FILES_DIR']
    excel_dir = os.path.join(config_files_dir, str(main_config_id), 'template_excel')
    return file_resource([(User, user_id, None), (MainConfig, main_config_id, "user_id")], excel_dir, request.method, "template_excel", request.files)

@main_config_blueprint.route('/<int:main_config_id>/result_excel', methods=["GET"])
@is_login
def result_excel(user_id, main_config_id):
    is_success, return_data = resource_limit([(User, user_id, None), (MainConfig, main_config_id, "user_id")])
    if not is_success:
        return return_data
    config_files_dir = current_app.config['CONFIG_FILES_DIR']
    is_success, return_data = get_file_path(config_files_dir, main_config_id, "result_excel")
    if not is_success:
        return response(False, 404, return_data)
    else:
        return return_file(return_data)

@main_config_blueprint.route('/<int:main_config_id>/remind', methods=["POST"])
@is_login
def remind(user_id, main_config_id):
    request_json = request.json
    is_success, return_data = parameter_check(request_json, remind_parameter["POST"])
    if not is_success:
        return return_data
    clean_data = return_data
    is_success, return_data = resource_limit([(User, user_id, None), (MainConfig, main_config_id, "user_id")])
    if not is_success:
        return return_data
    main_config_query, main_config = return_data[1:3]
    receive_config_id = main_config.receive_config_id
    if not receive_config_id:
        return response(False, 403, "无催办配置")
    else:
        is_sucess, return_data = resource_limit([(User, user_id, None), (MainConfig, main_config_id, "user_id")])
        if not is_success:
            return return_data
        receive_config_query, receive_config = return_data[1:3]
        username = main_config.email
        password = main_config.password
        remind_subject = receive_config.remind_subject
        remind_content = receive_config.remind_content
        remind_ip = receive_config.remind_ip
        remind_port = receive_config.remind_port
        email_list = clean_data["email"].split('|')
        is_success, return_data = send_multi_mail(remind_ip, remind_port, username, password, email_list, remind_subject, remind_content)
        if not is_success:
            return response(False, 400, return_data)
        return response(True, 200, "成功")

@main_config_blueprint.route('/active', methods=["get"])
@is_login
def active(user_id):
    active_task_id_list, scheduled_task_id_list = get_run_config(user_id)
    return response(True, 200, '成功', {
        "active_task": active_task_id_list,
        "scheduled_task": scheduled_task_id_list
    })

@main_config_blueprint.route('/<int:main_config_id>/log', methods=['GET'])
@is_login
def task_log(user_id, main_config_id):
    is_success, return_data = resource_limit([(User, user_id, None), (MainConfig, main_config_id, "user_id")])
    if not is_success:
        return return_data
    send_key = f"{main_config_id}_send_log"
    receive_key = f"{main_config_id}_receive_log"
    send_log = [i.decode() for i in redis.lrange(send_key, 0, -1)] if redis.exists(send_key) else []
    receive_log = [i.decode() for i in redis.lrange(receive_key, 0, -1)] if redis.exists(receive_key) else []
    return response(True, 200, "成功", {
        "send_log": send_log,
        "receive_log": receive_log
    })

@main_config_blueprint.route('/<int:main_config_id>/send_history', methods=['GET'])
@main_config_blueprint.route('/<int:main_config_id>/send_history/<int:history_id>', methods=['GET', 'DELETE'])
@is_login
def send_history(user_id, main_config_id, history_id=None):
    request_method = request.method
    request_args = request.args
    request_json = request.json
    return resource_manage([(User, user_id, None), (MainConfig, main_config_id, "user_id"), (SendHistory, history_id, "main_config_id")], request_method, request_args, request_json, send_history_parameter)

@main_config_blueprint.route('/<int:main_config_id>/receive_history', methods=['GET'])
@main_config_blueprint.route('/<int:main_config_id>/receive_history/<int:history_id>', methods=['GET', 'DELETE'])
@is_login
def receive_history(user_id, main_config_id, history_id=None):
    request_method = request.method
    request_args = request.args
    request_json = request.json
    return resource_manage([(User, user_id, None), (MainConfig, main_config_id, "user_id"), (ReceiveHistory, history_id, "main_config_id")], request_method, request_args, request_json, receive_history_parameter)

@main_config_blueprint.route('/<int:main_config_id>/split_excel', methods=['GET'])
@is_login
def split_file(user_id, main_config_id):
    zip_path = os.path.join(current_app.config['CONFIG_FILES_DIR'], str(main_config_id), f'拆分表格.zip')
    split_file_dir = os.path.join(current_app.config['CONFIG_FILES_DIR'], str(main_config_id), "split_excel")
    return return_zip([(User, user_id, None), (MainConfig, main_config_id, "user_id")], zip_path, split_file_dir)

@main_config_blueprint.route('/<int:main_config_id>/receive_excel', methods=['GET'])
@is_login
def receive_excel(user_id, main_config_id):
    zip_path = os.path.join(current_app.config['CONFIG_FILES_DIR'], str(main_config_id), f'收取表格.zip')
    receive_file_dir = os.path.join(current_app.config['CONFIG_FILES_DIR'], str(main_config_id), "receive_excel")
    return return_zip([(User, user_id, None), (MainConfig, main_config_id, "user_id")], zip_path, receive_file_dir)