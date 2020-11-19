import os
import shutil
from flask import Blueprint, request, current_app
from ..models import SapConfig, User
from ..parameter_config import sap_config_parameter, sap_log_parameter
from ..func_tools import resource_manage, is_login, response, resource_limit, parameter_check, return_file, file_resource
from ..extention import redis

sap_config_blueprint = Blueprint('sap_config', __name__)

@sap_config_blueprint.route("/", methods=["GET", "POST"])
@sap_config_blueprint.route("/<int:config_id>", methods=["GET", "PUT"])
@is_login
def sap_config(user_id, config_id=None):
    request_method = request.method
    request_args = request.args
    request_json = request.json
    return resource_manage([(User, user_id, None), (SapConfig, config_id, "user_id")], request_method, request_args, request_json, sap_config_parameter)

@sap_config_blueprint.route("/<int:config_id>/log", methods=["GET", "POST", "DELETE"])
@is_login
def log(user_id, config_id):
    is_success, return_data = resource_limit([(User, user_id, None), (SapConfig, config_id, "user_id")])
    if not is_success:
        return return_data
    key = f"{config_id}_sap_log"
    if request.method == "GET":
        sap_log = [i.decode() for i in redis.lrange(key, 0, -1)] if redis.exists(key) else []
        return response(True, 200, "成功", sap_log)
    elif request.method == "POST":
        request_json = request.json
        is_success, return_data = parameter_check(request_json, sap_log_parameter["POST"])
        if not is_success:
            return return_data
        log_content = return_data["log"]
        redis.rpush(key, log_content)
        sap_log = [i.decode() for i in redis.lrange(key, 0, -1)]
        return response(True, 200, "成功", sap_log)
    elif request.method == "DELETE":
        redis.delete(key)
        return response(True, 204, "成功")

@sap_config_blueprint.route("/<int:config_id>/should_file", methods=["GET", "POST"])
@is_login
def should_file(user_id, config_id):
    config = current_app.config
    sap_files_dir = os.path.join(config["SAP_FILES_DIR"], str(user_id), "should")
    request_method = request.method
    request_file = request.files
    return file_resource([(User, user_id, None), (SapConfig, config_id, "user_id")], sap_files_dir, request_method, "file", request_file, file_type="*")

@sap_config_blueprint.route("/<int:config_id>/pre_file", methods=["GET", "POST"])
@is_login
def pre_file(user_id, config_id):
    config = current_app.config
    sap_files_dir = os.path.join(config["SAP_FILES_DIR"], str(user_id), "pre")
    request_method = request.method
    request_file = request.files
    return file_resource([(User, user_id, None), (SapConfig, config_id, "user_id")], sap_files_dir, request_method, "file", request_file, file_type="*")

@sap_config_blueprint.route("/<int:config_id>/other_file", methods=["GET", "POST"])
@is_login
def other_file(user_id, config_id):
    config = current_app.config
    sap_files_dir = os.path.join(config["SAP_FILES_DIR"], str(user_id), "other")
    request_method = request.method
    request_file = request.files
    return file_resource([(User, user_id, None), (SapConfig, config_id, "user_id")], sap_files_dir, request_method, "file", request_file, file_type="*")

@sap_config_blueprint.route('/code_excel', methods=["POST"])
@is_login
def code_excel(user_id):
    file_dir = current_app.config["STATIC_FILES_DIR"]
    return file_resource([(User, user_id, None)], file_dir, request.method, "file",
                  request.files, file_type="excel", is_reset=False, new_file_name="公司代码")

@sap_config_blueprint.route('/name_excel', methods=["POST"])
@is_login
def name_excel(user_id):
    file_dir = current_app.config["STATIC_FILES_DIR"]
    return file_resource([(User, user_id, None)], file_dir, request.method, "file",
                  request.files, file_type="excel", is_reset=False, new_file_name="公司全称与简称对照表")