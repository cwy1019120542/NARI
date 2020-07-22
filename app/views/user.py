import random
import smtplib
from flask import Blueprint, request, current_app, session
from ..extention import bcrypt, db, redis
from ..parameter_config import user_parameter
from ..func_tools import response, parameter_check, page_filter, captcha_check, is_login, resource_limit, smtp_send_mail
from ..models import User

user_blueprint = Blueprint('user', __name__)

@user_blueprint.route('/<int:user_id>', methods=['GET', 'PUT', 'DELETE'])
@user_blueprint.route('', methods=['POST', 'GET'])
@is_login
def user(user_id=None):
    if user_id:
        is_success, return_data = resource_limit([User, user_id, None])
        if not is_success:
            return return_data
        else:
            user_query, user = return_data[1:3]
    if request.method == 'GET':
        if user_id:
            user_info = user.get_info()
            return response(True, 200, "成功", user_info)
        request_args = request.args
        is_right, clean_data = parameter_check(request_args, user_parameter["GET"], False)
        if not is_right:
            return clean_data
        user_list, page_info = page_filter(User, clean_data, user_parameter["fuzzy_field"])
        result = [i.get_info() for i in user_list]
        return response(True, 200, "成功", result, **page_info)
    elif request.method == 'POST':
        request_json = request.json
        is_right, clean_data = parameter_check(request_json, user_parameter["POST"])
        if not is_right:
            return clean_data
        email = clean_data['email']
        user = db.session.query(User).filter_by(email=email).first()
        if user:
            return response(False, 403, f"用户{email}已存在")
        captcha = clean_data.pop('captcha')
        check_result = captcha_check(email, captcha)
        if check_result:
            return response(False, 403,  check_result)
        password = clean_data['password']
        password_hash = bcrypt.generate_password_hash(password)
        clean_data['password'] = password_hash
        new_user = User(**clean_data)
        db.session.add(new_user)
        db.session.commit()
        result = new_user.get_info()
        return response(True, 201, "成功", result)
    elif request.method == 'PUT':
        request_json = request.json
        is_right, clean_data = parameter_check(request_json, user_parameter["PUT"], False)
        if not is_right:
            return clean_data
        if 'password' in clean_data:
            password = clean_data['password']
            password_hash = bcrypt.generate_password_hash(password)
            clean_data['password'] = password_hash
        user_query.update(clean_data)
        db.session.commit()
        result = db.session.query(User).get(user_id).get_info()
        return response(True, 201, "成功", result)
    elif request.method == 'DELETE':
        user_query.update({"status": 0})
        db.session.commit()
        return response(True, 204, "成功")

@user_blueprint.route('/login', methods=['POST'])
def login():
    request_json = request.json
    parameter_group_list = (('email', str, False, 50), ('password', str, False, 100))
    is_right, clean_data = parameter_check(request_json, parameter_group_list)
    if not is_right:
        return clean_data
    email = clean_data['email']
    password = clean_data['password']
    user= db.session.query(User).filter_by(email=email).first()
    if not user:
        return response(False, 401, f"用户{email}未注册")
    password_hash = user.password
    if not bcrypt.check_password_hash(password_hash, password):
        return response(False, 401, f"用户{email}密码错误")
    user_id = user.id
    session['user_id'] = user_id
    result = user.get_info()
    return response(True, 200, "成功", result)

@user_blueprint.route('/send_captcha', methods=['POST'])
def send_captcha():
    request_json = request.json
    is_right, clean_data = parameter_check(request_json, (('email', str, False, 50), ))
    if not is_right:
        return clean_data
    captcha = random.randrange(10000, 99999)
    email = clean_data['email']
    config = current_app.config
    SERVER_MAIL_IP = config['SERVER_MAIL_IP']
    SERVER_MAIL_USER = config['SERVER_MAIL_USER']
    SERVER_MAIL_PASSWORD = config['SERVER_MAIL_PASSWORD']
    try:
        smtp_obj = smtplib.SMTP(SERVER_MAIL_IP)
        smtp_obj.login(SERVER_MAIL_USER, SERVER_MAIL_PASSWORD)
        smtp_send_mail(smtp_obj, SERVER_MAIL_USER, email, "邮件收发系统注册验证码", str(captcha))
    except Exception as error:
        print(error)
        return response(False, 403, "发邮件失败")
    redis.set(email, captcha)
    redis.expire(email, 600)
    return response(True, 200, "成功")


@user_blueprint.route('/<int:user_id>/logout', methods=['POST'])
@is_login
def logout(user_id):
    session.pop('user_id')
    return response(True, 200, "成功")


