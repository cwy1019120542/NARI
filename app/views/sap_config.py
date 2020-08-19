from flask import Blueprint, request
from ..models import SapConfig, User
from ..parameter_config import sap_config_parameter, sap_log_parameter
from ..func_tools import resource_manage, is_login, response, resource_limit, parameter_check
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