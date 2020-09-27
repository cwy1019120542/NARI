import os
import shutil
from datetime import datetime
from flask import Blueprint, request, current_app
from ..func_tools import is_login, to_xlsx, response, get_run_config, create_task_id, save_file, parameter_check, return_zip
from ..parameter_config import check_file_dict
from ..tasks.generate_report import generate_report
from ..models import User

report_blueprint = Blueprint("report", __name__)

@report_blueprint.route("/excel", methods=["POST"])
@is_login
def report_excel(user_id):
    request_files = request.files
    REPORT_FILES_DIR = current_app.config["REPORT_FILES_DIR"]
    now_date_str = request.form.get("date")
    if not now_date_str:
        return response(False, 400, "参数错误")
    try:
        datetime.strptime(now_date_str, "%Y-%m")
    except:
        return response(False, 400, "参数错误")
    file_dir = os.path.join(REPORT_FILES_DIR, str(user_id), now_date_str, "origin")
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    save_list = []
    for key, value in check_file_dict.items():
        is_success, return_data = save_file(key, request_files, False, file_dir, value)
        if not is_success:
            continue
        save_list.append(value)
    return response(True, 200, "成功", save_list)

@report_blueprint.route("/start", methods=["POST"])
@is_login
def start(user_id):
    request_json = request.json
    is_success, return_data = parameter_check(request_json, [("date", str, False, 10)])
    if not is_success:
        return return_data
    now_date_str = return_data["date"]
    try:
        datetime.strptime(now_date_str, "%Y-%m")
    except:
        return response(False, 400, "参数错误")
    active_task_id_list, scheduled_task_id_list = get_run_config(user_id, "user_id")
    if active_task_id_list or scheduled_task_id_list:
        return response(False, 403, "该配置已在运行")
    task_id = create_task_id("report", user_id=user_id)
    config = current_app.config
    REPORT_FILES_DIR = config["REPORT_FILES_DIR"]
    STATIC_FILES_DIR = config["STATIC_FILES_DIR"]
    code_file_path = os.path.join(STATIC_FILES_DIR, "公司代码.xlsx")
    file_dir = os.path.join(REPORT_FILES_DIR, str(user_id), now_date_str)
    year, month = [int(i) for i in now_date_str.split("-")]
    last_month = month -1 if month != 1 else 12
    last_year = year if month != 1 else year - 1
    last_date_str = "-".join([str(last_year), str(last_month)])
    last_year_date_str = f"{year-1}-12"
    last_year_file_dir = os.path.join(REPORT_FILES_DIR, str(user_id), last_year_date_str, "result")
    last_file_dir = os.path.join(REPORT_FILES_DIR, str(user_id), last_date_str, "result")
    generate_report.apply_async(kwargs={"file_dir": file_dir, "code_file_path": code_file_path, "last_file_dir": last_file_dir, "last_year_file_dir": last_year_file_dir},
                              task_id=task_id)
    return response(True, 202, "成功")

@report_blueprint.route("/stop", methods=["POST"])
@is_login
def stop(user_id):
    pass

@report_blueprint.route('/file/<date>', methods=['GET'])
@is_login
def report_file(user_id, date):
    zip_path = os.path.join(current_app.config['REPORT_FILES_DIR'], str(user_id), date, '稽核处文件.zip')
    file_dir = os.path.join(current_app.config['REPORT_FILES_DIR'], str(user_id), date, "result")
    return return_zip([(User, user_id, None)], zip_path, file_dir)


