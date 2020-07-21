import poplib
import chardet
import email
import smtplib
import openpyxl
import datetime
import os
import time
import shutil
import json
from traceback import print_exc
from ..models import MainConfig, ReceiveHistory
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from ..extention import celery, Session, receive_logger, redis
from ..func_tools import get_match_dict, get_file_path, response, get_header_row, get_column_number, smtp_send_mail, send_multi_mail

def generate_log(main_config_id, level, message):
    getattr(receive_logger, level)(message)
    redis.rpush(f"{main_config_id}_receive_log", message)

def decode_str(s):  # 字符编码转换
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def getTimeStamp(utcstr):
    format_str_list = [('%a, %d %b %Y %H:%M:%S +0800', 0), ('%a, %d %b %Y %H:%M:%S +0800 (CST)', 0), ('%a, %d %b %Y %H:%M:%S -0900', 17), ('%a, %d %b %Y %H:%M:%S -0900 (CST)', 17)]
    for format_str, add_hour in format_str_list:
        try:
            utcdatetime = datetime.datetime.strptime(utcstr, format_str)
        except Exception as error:
            continue
        utcdatetime += datetime.timedelta(hours=add_hour)
        localtimestamp = utcdatetime.timestamp()
        return localtimestamp

def pop3_receive_mail():
    pass

def get_column_data_list(sheet, field_data_list, header_row, insert_none=True):
    column_number_list = get_column_number()
    max_row = sheet.max_row - header_row
    field_data_list_test = [i.value for i in sheet[header_row] if i.value]
    column_data_list = []
    for field in field_data_list:
        if field in field_data_list_test:
            field_index = field_data_list_test.index(field)
            column_number = column_number_list[field_index]
            column_data = [i.value for i in sheet[column_number][header_row:]]
            column_data_list.append(column_data)
        else:
            if insert_none:
                column_data_list.append([None for i in range(max_row)])
    return column_data_list

def decode_group(data):
    if data[-1][1]:
        decode_data = data[-1][0].decode(data[-1][1])
    else:
        decode_data = data[-1][0]
        if isinstance(decode_data, bytes):
            decode_data = decode_data.decode()
        if ' ' in decode_data:
            decode_data = decode_data.split(' ')[1][1:-1]
    decode_data = decode_data.strip().strip('<').strip('>')
    return decode_data

def get_target_group(target, match_dict):
    target_email = match_dict[target]
    target_group = [target, '|'.join(target_email)]
    return target_group

@celery.task(track_started=True)
def receive_mail(config, main_config, receive_config):
    try:
        main_config_id = main_config['id']
        redis.delete(f"{main_config_id}_receive_log")
        generate_log(main_config_id, "info", f"main_config_id {main_config_id} 收邮件任务开始")
        config_files_dir = config['CONFIG_FILES_DIR']
        username = main_config['email']
        password = main_config['password']
        host = receive_config['ip']
        port = receive_config['port']
        is_success, match_dict = get_match_dict(config_files_dir, main_config_id)
        if not is_success:
            generate_log(main_config_id, "critical", f"main_config_id {main_config_id} 缺失邮箱对应表")
            return "运行失败,缺失邮箱对应表"
        target_list = list(match_dict.keys())
        new_match_dict = {}
        for key, value in match_dict.items():
            for from_email in value:
                if from_email in new_match_dict:
                    new_match_dict[from_email].append(key)
                else:
                    new_match_dict[from_email] = [key]
        target_subject = receive_config['subject']
        sheet_info = receive_config['sheet_info']
        read_start_timestamp = receive_config['read_start_timestamp']
        read_end_timestamp = receive_config['read_end_timestamp']
        is_remind = receive_config['is_remind']
        remind_subject = receive_config['remind_subject']
        remind_content = receive_config['remind_content']
        remind_ip = receive_config['remind_ip']
        remind_port = receive_config['remind_port']
        remind_agreement = receive_config['remind_agreement']
        generate_log(main_config_id, "info", f"main_config_id {main_config_id} 开始收邮件")
        try:
            pop_obj = poplib.POP3(host)
            pop_obj.user(username)
            pop_obj.pass_(password)
        except Exception as error:
            generate_log(main_config_id, "critical", f"main_config_id {main_config_id} 无法登陆收件邮箱 {error}")
            return "运行失败,无法登陆收件邮箱"
        resp, mail_list, octets = pop_obj.list()
        mail_count = len(mail_list)
        no_response_list = []
        no_attachment_list = []
        no_attachment_target_list = set()
        generate_log(main_config_id, "info", f"main_config_id {main_config_id} mail_count {mail_count}")
        file_dir = os.path.join(config_files_dir, str(main_config_id), "receive_excel")
        if os.path.exists(file_dir):
            shutil.rmtree(file_dir)
        os.makedirs(file_dir)
        history_list = []
        for i in range(mail_count, 0, -1):
            try:
                resp, lines, octets = pop_obj.retr(i)
            except Exception as error:
                generate_log(main_config_id, "error", f"main_config_id {main_config_id} 邮件解析失败 {error}")
                continue
            msg_content = b'\r\n'.join(lines).decode("utf8", "ignore")
            msg = Parser().parsestr(msg_content)
            if not msg['from'] or not msg['subject'] or not msg['date']:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} 信息缺失")
                continue
            from_email = decode_header(msg['from'])
            from_email = decode_group(from_email)
            subject_group = decode_header(msg['subject'])
            subject = decode_group(subject_group)
            msg_timestamp = getTimeStamp(msg['date'])
            if not msg_timestamp:
                generate_log(main_config_id, "error", f"{from_email} {subject} {msg['date']} 时间无法解析")
                continue
            if msg_timestamp < read_start_timestamp:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} {from_email} {subject} {msg_timestamp} 小于开始时间")
                break
            if msg_timestamp > read_end_timestamp:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} {from_email} {subject} {msg_timestamp} 大于结束时间")
                continue
            if from_email not in new_match_dict:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} {from_email} is not useful")
                continue
            target_group = new_match_dict[from_email]
            target_group_str = "|".join(target_group)
            check_list = [i for i in target_group if i not in target_list]
            if check_list:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} {from_email} {new_match_dict.get(from_email)} was handled")
                continue
            generate_log(main_config_id, "info", f"main_config_id {main_config_id} useful {from_email} {subject} {msg_timestamp}")
            if target_subject in subject:
                for part in msg.walk():
                    file_name = part.get_filename()
                    if file_name:
                        h = email.header.Header(file_name)
                        dh = email.header.decode_header(h)
                        filename = dh[0][0]
                        if dh[0][1]:
                            filename = decode_str(str(filename, dh[0][1]))
                            generate_log(main_config_id, "info", f"main_config_id {main_config_id} filename {filename}")
                            data = part.get_payload(decode=True)
                            file_path = os.path.join(file_dir, f'{target_group_str}_{filename}')
                            with open(file_path, 'wb') as f:
                                f.write(data)
                            for target in target_group:
                                if target in target_list:
                                    target_list.remove(target)
                                if target in no_attachment_target_list:
                                    no_attachment_target_list.remove(target)
                            history_list.append(ReceiveHistory(email=from_email, target=target_group_str, timestamp=time.time(), main_config_id=main_config_id, status=True, message="success"))
                    else:
                        for target in target_group:
                            if target in target_list:
                                no_attachment_target_list.add(target)
        pop_obj.quit()
        for no_attachment_target in no_attachment_target_list:
            email_group = '|'.join(match_dict[no_attachment_target])
            history_list.append(ReceiveHistory(email=email_group, target=no_attachment_target, timestamp=time.time(), main_config_id=main_config_id, status=False, message="缺失附件"))
        for target in target_list:
            if target not in no_attachment_target_list:
                email_group = '|'.join(match_dict[target])
                history_list.append(ReceiveHistory(email=email_group, target=target, timestamp=time.time(), main_config_id=main_config_id, status=False, message="未回复"))
        generate_log(main_config_id, "info", f"main_config_id {main_config_id} 邮件收取完毕 no_attachment_list {no_attachment_list} no_response_list {no_response_list}")
        if is_remind:
            remind_target_list = []
            remind_target_list.extend(no_response_list)
            remind_target_list.extend(no_attachment_list)
            if remind_target_list:
                email_list = []
                for remind_target_group  in remind_target_list:
                    email_list.extend(remind_target_group[1].split('|'))
                is_success, return_data = send_multi_mail(remind_ip, remind_port, username, password, email_list,
                                                          remind_subject, remind_content)
                if not is_success:
                    return return_data
        sheet_info = json.loads(sheet_info)
        is_success, return_data = get_file_path(config_files_dir, main_config_id, "template_excel")
        template_path = False
        if is_success:
            template_path = return_data
        is_success, receive_excel_path_list = get_file_path(config_files_dir, main_config_id, "receive_excel", True)
        if not is_success or not receive_excel_path_list:
            generate_log(main_config_id, "error", f"main_config_id {main_config_id} 未收取到指定邮件")
        else:
            if not template_path:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} 没有提供模板")
                result_list = []
                receive_excel_check = openpyxl.load_workbook(receive_excel_path_list[0], data_only=True)
                sheet_name_list = receive_excel_check.sheetnames
                receive_excel_check.close()
                for sheet_name in sheet_name_list:
                    field_data_list = []
                    for receive_excel_path in receive_excel_path_list[:]:
                        receive_excel = openpyxl.load_workbook(receive_excel_path, data_only=True)
                        receive_excel_sheet_name_list = receive_excel.sheetnames
                        if sheet_name not in receive_excel_sheet_name_list:
                            generate_log(main_config_id, "error", f"main_config_id {main_config_id} {receive_excel_path} 没有sheet页 {sheet_name}")
                            continue
                        sheet = receive_excel[sheet_name]
                        header_row = get_header_row(sheet, None)
                        field_data_list_test = [i.value for i in sheet[header_row] if i.value]
                        for field in field_data_list_test:
                            if field not in field_data_list:
                                field_data_list.append(field)
                        receive_excel.close()
                    all_sheet_data = []
                    for path_index, receive_excel_path in enumerate(receive_excel_path_list):
                        receive_excel = openpyxl.load_workbook(receive_excel_path, data_only=True)
                        receive_excel_sheet_name_list = receive_excel.sheetnames
                        if sheet_name not in receive_excel_sheet_name_list:
                            generate_log(main_config_id, "error", f"main_config_id {main_config_id} {receive_excel_path} 没有sheet页 {sheet_name}")
                            continue
                        sheet = receive_excel[sheet_name]
                        header_row = get_header_row(sheet, None)
                        if path_index == 0:
                            header_data = [[i.value for i in sheet[i]] for i in range(1, header_row)]
                            all_sheet_data.extend(header_data)
                            all_sheet_data.append(field_data_list)
                        column_data_list = get_column_data_list(sheet, field_data_list, header_row)
                        all_sheet_data.extend(list(zip(*column_data_list)))
                        receive_excel.close()
                    result_list.append(all_sheet_data)
                    generate_log(main_config_id, "info", f"main_config_id {main_config_id} 表格数据生成完毕")
                work_book = openpyxl.Workbook('结果表.xlsx')
                result_dir = os.path.join(config_files_dir, str(main_config_id), "result_excel")
                if not os.path.exists(result_dir):
                    os.makedirs(result_dir)
                result_path = os.path.join(result_dir, '结果表.xlsx')
                for sheet_name, result in reversed(list(zip(sheet_name_list, result_list))):
                    sheet = work_book.create_sheet(sheet_name, 0)
                    for data in result:
                        sheet.append(data)
                work_book.save(result_path)
                work_book.close()
            else:
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} 提供模板表")
                result_dir = os.path.join(config_files_dir, str(main_config_id), "result_excel")
                if not os.path.exists(result_dir):
                    os.makedirs(result_dir)
                result_path = os.path.join(result_dir, '结果表.xlsx')
                shutil.copy(template_path, result_path)
                result_list = []
                result_excel = openpyxl.load_workbook(result_path, data_only=True)
                result_excel_sheet_name_list = result_excel.sheetnames
                for sheet_name, header_row, merge_field, fill_field in sheet_info:
                    if len(sheet_info) == 1:
                        sheet = result_excel[result_excel_sheet_name_list[0]]
                    else:
                        if sheet_name not in result_excel_sheet_name_list:
                            generate_log(main_config_id, "error", f"main_config_id {main_config_id} result_excel 没有sheet页 {sheet_name}")
                            continue
                        else:
                            sheet = result_excel[sheet_name]
                    header_row = get_header_row(sheet, header_row)
                    field_data_list = [i.value for i in sheet[header_row]]
                    check_row_list = [i.value for i in sheet[header_row+1] if i.value]
                    if len(check_row_list) > 2:
                        generate_log(main_config_id, "info", f"main_config_id {main_config_id} 模板表有数据")
                        merge_field_list = merge_field.split('|')
                        fill_field_list = fill_field.split('|')
                        merge_fill_dict = {}
                        for receive_excel_path in receive_excel_path_list:
                            generate_log(main_config_id, "info", f"main_config_id {main_config_id} 聚合 {os.path.split(receive_excel_path)[1]}")
                            receive_excel = openpyxl.load_workbook(receive_excel_path, data_only=True)
                            receive_excel_sheet_name_list = receive_excel.sheetnames
                            if len(sheet_info) == 1:
                                receive_sheet = receive_excel[receive_excel_sheet_name_list[0]]
                            else:
                                if sheet_name not in receive_excel_sheet_name_list:
                                    generate_log(main_config_id, "error", f"main_config_id {main_config_id} receive_excel {receive_excel_path} 没有sheet页 {sheet_name}")
                                    continue
                                else:
                                    receive_sheet = receive_excel[sheet_name]
                            merge_data_list = get_column_data_list(receive_sheet, merge_field_list, header_row, False)
                            merge_data_group = list(zip(*merge_data_list))
                            fill_data_list = get_column_data_list(receive_sheet, fill_field_list, header_row, False)
                            fill_data_group = list(zip(*fill_data_list))
                            for merge_group, fill_group in zip(merge_data_group, fill_data_group):
                                fill_data = merge_fill_dict.get(tuple(merge_group))
                                if not fill_data:
                                    merge_fill_dict[tuple(merge_group)] = fill_group
                                else:
                                    if fill_data.count(None) > fill_group.count(None):
                                        merge_fill_dict[tuple(merge_group)] = fill_group
                            receive_excel.close()
                        result_list.append(merge_fill_dict)
                    else:
                        generate_log(main_config_id, "info", f"main_config_id {main_config_id} 模板表无数据")
                        all_sheet_data = []
                        for receive_excel_path in receive_excel_path_list:
                            receive_excel = openpyxl.load_workbook(receive_excel_path, data_only=True)
                            receive_excel_sheet_name_list = receive_excel.sheetnames
                            if len(sheet_info) == 1:
                                receive_sheet = receive_excel[receive_excel_sheet_name_list[0]]
                            else:
                                if sheet_name not in receive_excel_sheet_name_list:
                                    generate_log(main_config_id, "error", f"main_config_id {main_config_id} receive_excel {receive_excel_path} 没有sheet页 {sheet_name}")
                                    continue
                                else:
                                    receive_sheet = receive_excel[sheet_name]
                            column_data_list = get_column_data_list(receive_sheet, field_data_list, header_row)
                            all_row_data = [i for i in zip(*column_data_list) if i.count(None) < 2]
                            all_sheet_data.extend(all_row_data)
                            receive_excel.close()
                        result_list.append(all_sheet_data)
                generate_log(main_config_id, "info", f"main_config_id {main_config_id} 数据聚合完毕")
                result_excel.close()
                for result_index, result in enumerate(result_list):
                    sheet_info[result_index].append(result)
                result_excel = openpyxl.load_workbook(result_path, data_only=True)
                for sheet_name, header_row, merge_field, fill_field, result in sheet_info:
                    if len(sheet_info) == 1:
                        sheet = result_excel[result_excel_sheet_name_list[0]]
                    else:
                        if sheet_name not in result_excel_sheet_name_list:
                            generate_log(main_config_id, "error", f"main_config_id {main_config_id} result_excel 没有sheet页 {sheet_name}")
                            continue
                        else:
                            sheet = result_excel[sheet_name]
                    header_row = get_header_row(sheet, header_row)
                    field_data_list = [i.value for i in sheet[header_row]]
                    column_number_list = get_column_number()
                    if isinstance(result, list):
                        for data_index, data in enumerate(result, start=1):
                            for column_index, single_data in enumerate(data):
                                column_number = column_number_list[column_index]
                                sheet[f"{column_number}{header_row+data_index}"].value = single_data
                    else:
                        merge_field_dict = {}
                        merge_field_list = merge_field.split('|')
                        fill_field_list = fill_field.split('|')
                        for row_index, row_data in enumerate(list(sheet.values)[header_row:], start=header_row+1):
                            key_list = []
                            for merge_field in merge_field_list:
                                merge_field_index = field_data_list.index(merge_field)
                                key_list.append(row_data[merge_field_index])
                            merge_field_dict[tuple(key_list)] = row_index
                        for fill_field_in_index, fill_field in enumerate(fill_field_list):
                            fill_field_index = field_data_list.index(fill_field)
                            column_number = column_number_list[fill_field_index]
                            for merge_field_group, row_index in merge_field_dict.items():
                                fill_value_list = result.get(merge_field_group)
                                if fill_value_list:
                                    fill_value = fill_value_list[fill_field_in_index]
                                else:
                                    fill_value = None
                                sheet[f"{column_number}{row_index}"].value = fill_value
                result_excel.save(result_path)
                result_excel.close()
        session = Session()
        session.add_all(history_list)
        session.commit()
        session.close()
        generate_log(main_config_id, "info", f"main_config_id {main_config_id} 运行成功")
        return "运行成功"
    except Exception as error:
        generate_log(main_config_id, "critical", f"main_config_id {main_config_id} 运行失败 {print_exc()}")
        return "运行失败,未知错误"


if __name__ == '__main__':
    config = {'CONFIG_FILES_DIR': r'C:\Users\cwy\PycharmProjects\untitled1\NARI\app\config_files'}
    main_config_info = {'id': 2, 'email': '1019120542@qq.com', 'password': 'aswhfohkwtaebfcc'}
    receive_config_info = {'subject': "测试邮件", 'content': "内容", "sheet": "Sheet1", "read_number": 100, 'split_field': "用途",
                        "ip": "pop.qq.com", "port": 110, "sheet_info": "[]", "is_remind": False, "remind_subject": None, "remind_content": None,
                           "remind_ip": None, "remind_port": None, "remind_agreement": None}
    return_data = receive_mail(config, main_config_info, receive_config_info)
