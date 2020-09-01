import os
import openpyxl
import uuid
from urllib.parse import quote
from flask import Blueprint, current_app, send_from_directory, make_response, request
from ..func_tools import return_file, db, response, resource_manage, save_file
from ..models import UpdateMessage
from ..parameter_config import update_message_premeter, accept_file_type

public_blueprint = Blueprint('public', __name__)

@public_blueprint.route('/match_excel_template', methods=["GET"])
def match_excel_template():
    static_files_dir = current_app.config['STATIC_FILES_DIR']
    filename = "邮箱对应模板表.xlsx"
    file_path = os.path.join(static_files_dir, filename)
    return return_file(file_path)

@public_blueprint.route('/update_message', methods=['GET', 'POST'])
@public_blueprint.route('/update_message/<int:message_id>', methods=['GET', 'PUT', 'DELETE'])
def update_message(message_id=None):
    request_method = request.method
    request_args = request.args
    request_json = request.json
    if request.method in ['POST', 'PUT', 'DELETE']:
        if request.json.pop('password', None) != 'cwy1019120542':
            return response(False, 403, "没有权限设置更新日志")
    return resource_manage([(UpdateMessage, message_id, None)], request_method, request_args, request_json, update_message_premeter)

@public_blueprint.route("/excel", methods=["POST"])
@public_blueprint.route("/excel/<name>", methods=["GET", "DELETE"])
def excel(name=None):
    temp_files_dir = current_app.config["TEMP_FILES_DIR"]
    if request.method == "GET":
        request_args = request.args
        if not name.endswith(accept_file_type):
            return response(False, 400, "文件格式不合法")
        file_path = os.path.join(temp_files_dir, name)
        if not os.path.exists(file_path):
            return response(False, 404, "请求的资源不存在")
        excel = openpyxl.load_workbook(file_path, data_only=True)
        sheet_name_list = excel.sheetnames
        header_data = None
        sheet_name = None
        header_row = None
        if "sheet" in request_args and "header_row" in request_args:
            sheet_name = request_args["sheet"]
            header_row = request_args["header_row"]
            if sheet_name not in sheet_name_list:
                return response(False, 404, "sheet不存在")
            sheet = excel[sheet_name]
            header_data = [i.value for i in sheet[header_row] if i.value]
        excel.close()
        return response(True, 200, "成功", {
            "sheet_name_list": sheet_name_list,
            "header_data": header_data,
            "sheet_name": sheet_name,
            "header_row": header_row
        })
    elif request.method == "POST":
        request_file = request.files
        return save_file("excel", request_file, False, temp_files_dir, str(uuid.uuid1()))[1]
    elif request.method == "DELETE":
        temp_file_path = os.path.join(temp_files_dir, name)
        if not os.path.exists(temp_file_path):
            return response(False, 404, "资源不存在")
        os.remove(temp_file_path)
        return response(True, 204, "成功")


