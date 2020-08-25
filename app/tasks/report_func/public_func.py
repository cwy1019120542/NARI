import os
from app.parameter_config import check_file_dict

def generate_key_value(column_value_list, field_list, target_field_list, header_row):
    target_list = []
    for target_field in target_field_list:
        target_index = field_list.index(target_field)
        target_list.append([i.value for i in column_value_list[target_index][header_row:]])
    if len(target_list) == 1:
        target = [str(i) if i else i for i in target_list[0]]
    else:
        target = [[str(j) if j else j for j in i] for i in zip(*target_list)]
    return target

def generate_dict(sheet, header_row, key_field_list, value_field_list):
    field_list = [i.value for i in sheet[header_row]]
    column_value_list = list(sheet.columns)
    key_list = generate_key_value(column_value_list, field_list, key_field_list, header_row)
    value_list = generate_key_value(column_value_list, field_list, value_field_list, header_row)
    data_dict = dict(zip(key_list, value_list))
    data_dict.pop(None, None)
    return data_dict

def handle_num(num):
    return round(float(num) / 10000, 2)

def replace_company(data_list, centre_company_dict):
    remove_data_list = []
    for data_index, data in enumerate(data_list[:]):
        if not data[1]:
            remove_data_list.append(data)
            continue
        if data[0] in ["4600", "4606", "4608", "4609"]:
            if data[2] not in centre_company_dict:
                print(f"{data[2]}缺失")
                remove_data_list.append(data)
                continue
            data_list[data_index][1] = centre_company_dict[data[2]]
    for remove_data in remove_data_list:
        data_list.remove(remove_data)

def split_total_inner(data_list):
    total_data_dict = {}
    inner_data_dict = {}
    for data in data_list:
        company = data[1]
        amount = float(data[3])
        if data[4] == "国网系统内-集团内":
            inner_data_dict[company] = inner_data_dict[company] + amount if company in inner_data_dict else amount
        total_data_dict[company] = total_data_dict[company] + amount if company in total_data_dict else amount
    return total_data_dict, inner_data_dict

def generate_percent_rate(num1, num2):
    return f'{round((float(num1) / float(num2)) * 100, 2)}%' if float(num2) else None

def generate_change_info(amount, last_amount):
    handle_amount = handle_num(amount)
    change_amount = handle_amount - float(last_amount)
    rate = generate_percent_rate(change_amount, last_amount)
    return handle_amount, change_amount, rate
