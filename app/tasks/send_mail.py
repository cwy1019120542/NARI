import os
import shutil
import re
import smtplib
import openpyxl
import time
from traceback import print_exc
from ..models import MainConfig
from ..extention import celery, Session, send_logger
from ..func_tools import response, get_file_path, get_match_dict, check_dir, smtp_send_mail, get_header_row

def clean_file_name(file_name):
    sub_list = r"[/\\\:\*\?\"\<\>\|]"
    new_file_name = re.sub(sub_list, '_', str(file_name))
    return new_file_name


@celery.task(track_started=True)
def send_mail(config, main_config_info, send_config_info):
    try:
        main_config_id = main_config_info['id']
        send_logger.info(f"main_config_id {main_config_id} 发邮件任务开始")
        config_files_dir = config['CONFIG_FILES_DIR']
        is_success, return_data = get_match_dict(config_files_dir, main_config_id)
        if not is_success:
            send_logger.critical(f"main_config_id {main_config_id} 缺失邮箱对应表")
            return "运行失败,缺失邮箱对应表"
        match_dict = return_data
        total_target_list = list(match_dict.keys())
        subject = send_config_info['subject']
        content = send_config_info['content']
        sheet_list = send_config_info['sheet'].split('|')
        field_row_list = [int(i) if i else None for i in send_config_info['field_row'].split('|')]
        is_split_list = [int(i) if i else 0 for i in send_config_info['is_split'].split('|')]
        split_field = send_config_info['split_field']
        ip = send_config_info['ip']
        port = send_config_info['port']
        is_success, return_data = get_file_path(config_files_dir, main_config_id, "send_excel")
        if not is_success:
            send_logger.critical(f"main_config_id {main_config_id} 未上传要发送的邮件")
            return "运行失败,未上传要发送的邮件"
        send_excel_path = return_data
        send_excel_name = os.path.splitext(os.path.split(send_excel_path)[1])[0]
        split_dir = os.path.join(config_files_dir, str(main_config_id), "split_excel")
        if os.path.exists(split_dir):
            shutil.rmtree(split_dir)
        os.makedirs(split_dir)
        empty_template_path = os.path.join(split_dir, "empty_template.xlsx")
        shutil.copy(send_excel_path, empty_template_path)
        send_logger.info(f"main_config_id {main_config_id} 开始生成空模板")
        empty_excel = openpyxl.open(empty_template_path)
        empty_excel_sheet_name_list = empty_excel.sheetnames
        for empty_excel_sheet_name in empty_excel_sheet_name_list:
            if empty_excel_sheet_name not in sheet_list:
                empty_excel.remove(empty_excel[empty_excel_sheet_name])
        for sheet_name, field_row, is_split in zip(sheet_list, field_row_list, is_split_list):
            empty_sheet = empty_excel[sheet_name]
            header_row = get_header_row(empty_sheet, field_row)
            max_row = int(empty_sheet.max_row)
            send_logger.info(f"main_config_id {main_config_id} sheet_name {sheet_name} max_row {max_row}")
            if is_split:
                empty_sheet.delete_rows(header_row+1, max_row+1)
        empty_excel.save(empty_template_path)
        empty_excel.close()
        send_logger.info(f"main_config_id {main_config_id} 空模板表生成完毕")
        send_excel = openpyxl.load_workbook(send_excel_path, data_only=True)
        result_dict = {}
        target_list = []
        if split_field:
            for sheet, field_row, is_split in zip(sheet_list, field_row_list, is_split_list):
                send_sheet = send_excel[sheet]
                if is_split:
                    field_row = get_header_row(send_sheet, field_row)
                    field_row = int(field_row) - 1
                    total_data_list = list(send_sheet.values)
                    data_list = total_data_list[field_row+1:]
                    field_data_list = total_data_list[field_row]
                    split_field_index = field_data_list.index(split_field)
                    for data in data_list:
                        split_value = data[split_field_index]
                        if split_value not in target_list:
                            target_list.append(split_value)
        else:
            target_list = list(match_dict.keys())
        for target in target_list:
            result_dict[target] = {}
            for sheet_name in sheet_list:
                result_dict[target][sheet_name] = None
        send_logger.info(f"main_config_id {main_config_id} 数据模板生成完毕")
        sheet_split_dict = dict(zip(sheet_list, is_split_list))
        for sheet_name, field_row, is_split in zip(sheet_list, field_row_list, is_split_list):
            send_sheet = send_excel[sheet_name]
            new_data_dict = {}
            if is_split:
                field_row = get_header_row(send_sheet, field_row)
                field_row = int(field_row) - 1
                total_data_list = list(send_sheet.values)
                data_list = total_data_list[field_row+1:]
                field_data_list = total_data_list[field_row]
                split_field_index = field_data_list.index(split_field)
                for data in data_list:
                    split_value = data[split_field_index]
                    if split_value not in new_data_dict:
                        new_data_dict[split_value] = []
                    new_data_dict[split_value].append(data)
                for key, value in new_data_dict.items():
                    result_dict[key][sheet_name] = list(value)
            else:
                for target in target_list:
                    result_dict[target][sheet_name] = []
        send_excel.close()
        send_logger.info(f"main_config_id {main_config_id} 数据生成完毕")
        for key, value in result_dict.items():
            if key not in total_target_list:
                send_logger.error(f"main_config_id {main_config_id} {key} 未配置邮箱")
                continue
            single_excel_path = os.path.join(split_dir, f'{clean_file_name(key)}_{send_excel_name}_.xlsx')
            shutil.copy(empty_template_path, single_excel_path)
            work_book = openpyxl.load_workbook(single_excel_path, data_only=True)
            for sheet_name, sheet_value in value.items():
                is_split = sheet_split_dict[sheet_name]
                if not is_split:
                    continue
                work_sheet = work_book[sheet_name]
                if sheet_value == None:
                    work_book.remove(work_sheet)
                    continue
                for row_value in sheet_value:
                    work_sheet.append(row_value)
            work_book.save(single_excel_path)
            work_book.close()
            send_logger.info(f"main_config_id {main_config_id} {key} 表生成完毕")
        sender = main_config_info['email']
        password = main_config_info['password']
        try:
            smtp_obj = smtplib.SMTP(ip, port)
            smtp_obj.login(sender, password)
        except Exception as error:
            send_logger.critical(f"main_config_id {main_config_id} 无法登陆发件邮箱 {error}")
            return "运行失败,无法登陆发件邮箱"
        send_number = 0
        for key in result_dict:
            if key in match_dict:
                email_list = match_dict[key]
                split_path = os.path.join(split_dir, f'{clean_file_name(key)}_{send_excel_name}_.xlsx')
                for email in email_list:
                    try:
                        smtp_send_mail(smtp_obj, sender, email, subject, content, [split_path])
                    except Exception as error:
                        send_logger.error(f"main_config_id {main_config_id} {email} 发送失败 {error}")
                        continue
                    else:
                        send_number += 1
                        send_logger.info(f"main_config_id {main_config_id} {email} 发送成功")
        smtp_obj.quit()
        session = Session()
        update_dict = {
            "run_timestamp": time.time(),
            "send_number": send_number
        }
        session.query(MainConfig).filter_by(id=main_config_id).update(update_dict)
        session.commit()
        session.close()
        send_logger.info(f"main_config_id {main_config_id}  运行成功")
        return "运行成功"
    except Exception as error:
        send_logger.critical(f"main_config_id {main_config_id} 运行失败 {print_exc()}")
        return "运行失败,未知错误"


