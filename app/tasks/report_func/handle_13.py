import os, openpyxl
from .public_func import generate_key_value, replace_company, handle_num
from app.parameter_config import check_file_dict

def generate_result_list(file_path, centre_company_dict):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[workbook.sheetnames[0]]
    column_value_list = list(sheet.columns)
    field_list = [i.value for i in sheet[1]]
    origin_data_list = generate_key_value(column_value_list, field_list,
                                   ["公司代码", "公司名称", "利润中心", "客户名称", "销售方开票金额", "确认收入金额", "是否一致"], 1)
    workbook.close()
    data_list = [i for i in origin_data_list if i[-1]=="不一致" and i[1] not in i[3]]
    replace_company(data_list, centre_company_dict)
    data_dict = {}
    for data in data_list:
        company = data[1]
        amount1 = float(data[4]) if data[4] else 0
        amount2 = float(data[5]) if data[5] else 0
        if company not in data_dict:
            data_dict[company] = [0, 0, 0, 0]
        if amount1 > amount2:
            data_dict[company][0] += amount1 - amount2
            data_dict[company][1] += 1
        else:
            data_dict[company][2] += amount2 - amount1
            data_dict[company][3] += 1
    result_list = []
    for company, data in data_dict.items():
        pre_amount, pre_count, suf_amount, suf_count = data
        result_list.append([company, handle_num(pre_amount), pre_count, handle_num(suf_amount), suf_count])
    return result_list

def generate_file(result_list, result_file_path):
    workbook = openpyxl.Workbook()
    sheet = workbook[workbook.sheetnames[0]]
    result_list.sort(key=lambda x:x[1], reverse=True)
    sheet.append(["序号", "公司名称", "预开票(金额)", "预开票(个数)", "滞后开票(金额)", "滞后开票(个数)"])
    sum_pre_amount = 0
    sum_pre_count = 0
    sum_suf_amount = 0
    sum_suf_count = 0
    for result_index, result in enumerate(result_list, start=1):
        sum_pre_amount += result[1]
        sum_pre_count += result[2]
        sum_suf_amount += result[3]
        sum_suf_count += result[4]
        sheet.append([result_index]+result)
    sheet.append([None, "合计", sum_pre_amount, sum_pre_count, sum_suf_amount, sum_suf_count])
    workbook.save(result_file_path)
    workbook.close()

def handle_13(file_dir, centre_company_dict):
    file_name = "内部关联交易-收入确认与开票不同步.xlsx"
    origin_file_path = os.path.join(file_dir, "origin", f"{check_file_dict['open_excel']}.xlsx")
    if not os.path.exists(origin_file_path):
        print("13文件丢失")
        return
    result_list = generate_result_list(origin_file_path, centre_company_dict)
    result_file_dir = os.path.join(file_dir, "result")
    if not os.path.exists(result_file_dir):
        os.makedirs(result_file_dir)
    result_file_path = os.path.join(result_file_dir, file_name)
    generate_file(result_list, result_file_path)
    print("13end")