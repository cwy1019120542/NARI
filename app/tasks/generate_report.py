import os
import openpyxl
from datetime import datetime, timedelta
from ..extention import celery

def generate_key_value(column_value_list, field_list, target_field_list, header_row):
    target_list = []
    for target_field in target_field_list:
        target_index = field_list.index(target_field)
        target_list.append([i.value for i in column_value_list[target_index][header_row:]])
    if len(target_list) == 1:
        target = target_list[0]
    else:
        target = list(zip(*target_list))
    return target

def generate_dict(sheet, header_row, key_field_list, value_field_list):
    field_list = [i.value for i in sheet[header_row]]
    column_value_list = list(sheet.columns)
    key_list = generate_key_value(column_value_list, field_list, key_field_list, header_row)
    value_list = generate_key_value(column_value_list, field_list, value_field_list, header_row)
    data_dict = dict(zip(key_list, value_list))
    return data_dict

def compare_with_last(data_dict, count_dict, file_name, last_file_dir, result_amount_field):
    result_dict = {}
    if file_name in os.listdir(last_file_dir):
        last_excel = openpyxl.load_workbook(os.path.join(last_file_dir, file_name))
        last_sheet = last_excel[last_excel.sheetnames[0]]
        last_data_dict = generate_dict(last_sheet, 1, [], ["公司名称", result_amount_field])
        last_excel.close()
    else:
        last_data_dict = {}
    for company, amount in data_dict.items():
        last_amount = float(last_data_dict.get(company, 0))
        change_amount = amount - last_amount
        result_dict[company] = [amount, change_amount, count_dict[company]]
    return result_dict

def generate_result_dict(file_path, init_amount_field, code_dict, total_file_name, inner_file_name, last_file_dir, result_amount_field):
    excel = openpyxl.load_workbook(file_path, read_only=True)
    sheet = excel[excel.sheetnames[0]]
    column_value_list = list(sheet.columns)
    field_list = [i.value for i in sheet[1]]
    data_list = generate_key_value(column_value_list, field_list, ["公司代码", "公司名称", "利润中心", init_amount_field, "客户属性"], 1)
    for data in data_list[:]:
        if data[0] in ["4600", "4606", "4608", "4609"]:
            data_list[1] = code_dict[data[2]]
    total_data_dict = {}
    inner_data_dict = {}
    total_count_dict = {}
    inner_count_dict = {}
    for data in data_list:
        company = data[1]
        amount = float(data[3])
        if data[4] == "国网系统内-集团内":
            inner_data_dict[company] = inner_data_dict[company] + amount if company in inner_data_dict else amount
            inner_count_dict[company] = inner_count_dict[company] + 1 if company in inner_count_dict else 1
        total_data_dict[company] = total_data_dict[company] + amount if company in total_data_dict else amount
        total_count_dict[company] = total_count_dict[company] + 1 if company in total_count_dict else 1
    total_result_dict = compare_with_last(total_data_dict, total_count_dict, total_file_name, last_file_dir, result_amount_field)
    inner_result_dict = compare_with_last(inner_data_dict, inner_count_dict, inner_file_name, last_file_dir, result_amount_field)
    return total_result_dict, inner_result_dict

def handle_1_2(file_dir, last_file_dir, code_dict):
    total_file_name = "内部关联交易-收入确认与开票不同步_总体情况.xlsx"
    inner_file_name = "内部关联交易-收入确认与开票不同步_系统内情况.xlsx"
    file1_path = os.path.join(file_dir, "1.xlsx")
    file2_path = os.path.join(file_dir, "2.xlsx")
    generate_result_dict(file1_path, "预开票金额", code_dict, total_file_name, inner_file_name, last_file_dir, "预开票(金额)")
    generate_result_dict(file2_path, "滞后开票金额", code_dict, total_file_name, inner_file_name, last_file_dir, "滞后开票(金额)")

@celery.task
def generate_report(user_id, file_dir, code_file_path, last_file_dir):
    code_excel = openpyxl.load_workbook(code_file_path)
    code_sheet = code_excel[code_excel.sheetnames[0]]
    code_dict = generate_dict(code_sheet, 1, ["利润中心"], ["分公司及事业部名称"])
    file_name_list = os.listdir(file_dir)
    if "1.xlsx" in file_name_list and "2.xlsx" in file_name_list:
        handle_1_2(file_dir, last_file_dir, code_dict)







