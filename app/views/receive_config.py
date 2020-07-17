from flask import Blueprint, request
from ..func_tools import is_login, config_resource
from ..models import ReceiveConfig
from ..api_config import receive_config_parameter

receive_config_blueprint = Blueprint('receive_config', __name__)

@receive_config_blueprint.route('', methods=['GET', 'POST'])
@receive_config_blueprint.route('/<int:receive_config_id>', methods=['GET', 'PUT', 'DELETE'])
@is_login
def send_config(user_id, receive_config_id=None):
    request_args = request.args
    request_json = request.json
    request_method = request.method
    return config_resource(user_id, ReceiveConfig, receive_config_id, request_method, request_args, request_json, receive_config_parameter)