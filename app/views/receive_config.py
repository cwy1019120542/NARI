from flask import Blueprint, request
from ..func_tools import is_login, resource_manage
from ..models import ReceiveConfig, User
from ..parameter_config import receive_config_parameter

receive_config_blueprint = Blueprint('receive_config', __name__)

@receive_config_blueprint.route('', methods=['GET', 'POST'])
@receive_config_blueprint.route('/<int:receive_config_id>', methods=['GET', 'PUT'])
@is_login
def send_config(user_id, receive_config_id=None):
    request_args = request.args
    request_json = request.json
    request_method = request.method
    return resource_manage([(User, user_id, None), (ReceiveConfig, receive_config_id, "user_id")], request_method, request_args, request_json, receive_config_parameter)