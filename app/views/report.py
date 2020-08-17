import os
import shutil
from flask import Blueprint, request, current_app
from ..func_tools import is_login, to_xlsx, response
from ..parameter_config import accept_file_type
from ..tasks.generate_report import generate_report

report_blueprint = Blueprint("report", __name__)

@report_blueprint.route("excel", methods=["POST"])
@is_login
def report_excel(user_id):
    request_files = request.files
    REPORT_FILES_DIR = current_app.config["REPORT_FILES_DIR"]
    now_date_str = request.form.get("date")
    if not now_date_str:
        return response(False, 400, "参数错误")
    file_dir = os.path.join(REPORT_FILES_DIR, str(user_id), now_date_str)
    if os.path.exists(file_dir):
        shutil.rmtree(file_dir)
    os.makedirs(file_dir)
    for i in range(1, 13):
        parameter = f"excel{i}"
        if parameter in request_files:
            file = request_files[parameter]
            file_name = file.filename
            if file_name.endswith(accept_file_type):
                file_suffix = file_name.split(".")[1]
                save_file_name = f"{i}.{file_suffix}"
                file_path = os.path.join(file_dir, save_file_name)
                file.save(file_path)
                if not save_file_name.endswith(".xlsx"):
                    to_xlsx(file_path)
    return response(True, 200, "成功")

@report_blueprint.route("start", methods=["POST"])
@is_login
def start(user_id):
    pass

@report_blueprint.route("stop", methods=["POST"])
@is_login
def stop(user_id):
    pass


