import os
from urllib.parse import quote
from flask import Blueprint, current_app, send_from_directory, make_response, request
from ..func_tools import return_file, db, response, resource_manage
from ..models import UpdateMessage
from ..parameter_config import update_message_premeter

static_blueprint = Blueprint('static', __name__)

@static_blueprint.route('/match_excel_template', methods=["GET"])
def match_excel_template():
    static_files_dir = current_app.config['STATIC_FILES_DIR']
    filename = "邮箱对应模板表.xlsx"
    file_path = os.path.join(static_files_dir, filename)
    return return_file(file_path)

@static_blueprint.route('/update_message', methods=['GET', 'POST'])
@static_blueprint.route('/update_message/<int:message_id>', methods=['GET', 'PUT', 'DELETE'])
def update_message(message_id=None):
    request_method = request.method
    request_args = request.args
    request_json = request.json
    if request.method in ['POST', 'PUT', 'DELETE']:
        if request.json.pop('password', None) != 'cwy1019120542':
            return response(False, 403, "没有权限设置更新日志")
    return resource_manage([(UpdateMessage, message_id, None)], request_method, request_args, request_json, update_message_premeter)


