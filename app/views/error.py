from flask import Blueprint
from ..func_tools import response

error_blueprint = Blueprint('error', __name__)

@error_blueprint.app_errorhandler(404)
def error_404(e):
    return response(False, 404, "请求的资源不存在")

@error_blueprint.app_errorhandler(405)
def error_405(e):
    return response(False, 405, "不支持该请求方法")

@error_blueprint.app_errorhandler(500)
def error_500(e):
    return response(False, 500, "服务端出错")
