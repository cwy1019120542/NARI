import os, openpyxl
from app.parameter_config import check_file_dict
from .public_func import generate_key_value, handle_num

def generate_result_list(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[workbook.sheetnames[0]]
    column_value_list = list(sheet.columns)
    field_list = [i.value for i in sheet[2]]
    data_list = generate_key_value(column_value_list, field_list,
                                   ["单位名称", "销售确认收入金额", "采购收货金额"], 3)
    workbook.close()
    data_dict = {}
    for data in data_list:
        company = data[0]
        amount1 = float(data[1]) if data[1] else 0
        amount2 = float(data[2]) if data[2] else 0
        data_dict[company] = (data_dict[company][0] + amount1, data_dict[company][1] + amount2) if company in data_dict else (amount1, amount2)
    result_list = []
    for company, data in data_dict.items():
        amount1, amount2 = data
        diff_amount = amount1 - amount2
        result_list.append([company, handle_num(amount1), handle_num(amount2), handle_num(diff_amount)])
    return result_list

def generate_file(result_list, result_file_path):
    workbook = openpyxl.Workbook()
    sheet = workbook[workbook.sheetnames[0]]
    result_list.sort(key=lambda x:x[3], reverse=True)
    sheet.append(["序号", "单位名称", "销售确认收入金额", "采购收货金额", "差异金额"])
    sum_amount1 = 0
    sum_amount2 = 0
    for result_index, result in enumerate(result_list, start=1):
        sum_amount1 += result[1]
        sum_amount2 += result[2]
        sheet.append([result_index]+result)
    sheet.append([None, "合计", sum_amount1, sum_amount2, sum_amount1-sum_amount2])
    workbook.save(result_file_path)
    workbook.close()

def handle_12(file_dir):
    file_name = "内部关联交易-收入确认与收货不同步.xlsx"
    origin_file_path = os.path.join(file_dir, "origin", f"{check_file_dict['receive_excel']}.xlsx")
    if not os.path.exists(origin_file_path):
        print("12文件丢失")
        return
    result_list = generate_result_list(origin_file_path)
    result_file_dir = os.path.join(file_dir, "result")
    if not os.path.exists(result_file_dir):
        os.makedirs(result_file_dir)
    result_file_path = os.path.join(result_file_dir, file_name)
    generate_file(result_list, result_file_path)
    print("12end")