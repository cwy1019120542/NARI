from flask import Blueprint, request
from ..func_tools import is_login, resource_manage
from ..models import SendConfig, User
from ..parameter_config import send_config_parameter

send_config_blueprint = Blueprint('send_config', __name__)

@send_config_blueprint.route('', methods=['GET', 'POST'])
@send_config_blueprint.route('/<int:send_config_id>', methods=['GET', 'PUT'])
@is_login
def send_config(user_id, send_config_id=None):
    request_args = request.args
    request_json = request.json
    request_method = request.method
    return resource_manage([(User, user_id, None), (SendConfig, send_config_id, "user_id")], request_method, request_args, request_json, send_config_parameter)