import os, openpyxl
from .public_func import generate_percent_rate, get_data_list, generate_dict, generate_change_info
from app.parameter_config import check_file_dict

def compare_with_last(data_dict, count_dict, file_path):
    if os.path.exists(file_path):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[workbook.sheetnames[0]]
        last_data_dict = generate_dict(sheet, 1, ["公司名称"], ["本月异常生产成本暂估金额 万元", "本月异常的项目数量 个"])
    else:
        last_data_dict = {}
    result_list = []
    for company, amount in data_dict.items():
        count = count_dict[company]
        last_data = [float(i) for i in last_data_dict.get(company, (0, 0))]
        last_amount, last_count = last_data
        handle_amount, change_amount, rate = generate_change_info(amount, last_amount)
        result_list.append([company, handle_amount, count, last_amount, last_count, change_amount, rate])
    return result_list

def generate_result_list(file_path, last_file_path):
    data_list = get_data_list(file_path, 4, 5, ["公司名称", "累计：生产成本暂估", "暂估占结转成本比(%)"])
    data_dict = {}
    count_dict = {}
    for data in data_list:
        company, amount, rate = data
        if rate and float(rate.strip("%")) > 20:
            amount = float(amount) if amount else 0
            data_dict[company] = data_dict[company] + amount if company in data_dict else amount
            count_dict[company] = count_dict[company] + 1 if company in count_dict else 1
    result_list = compare_with_last(data_dict, count_dict, last_file_path)
    return result_list

def generate_file(result_list, result_file_path):
    workbook = openpyxl.Workbook()
    sheet = workbook[workbook.sheetnames[0]]
    result_list.sort(key=lambda x:x[1], reverse=True)
    sheet.append(["序号", "公司名称", "本月异常生产成本暂估金额 万元", "本月异常的项目数量 个",
                  "年初异常生产成本暂估金额 万元", "年初异常的项目数量 个", "异常生产成本暂估较年初金额增减变化 万元", "较年初增降幅度 %"])
    sum_amount = 0
    sum_count = 0
    sum_last_amount = 0
    sum_last_count = 0
    for result_index, result in enumerate(result_list, start=1):
        sum_amount += result[1]
        sum_count += result[2]
        sum_last_amount += result[3]
        sum_last_count += result[4]
        sheet.append([result_index]+result)
    sum_change_amount = sum_amount - sum_last_amount
    rate = generate_percent_rate(sum_change_amount, sum_last_amount)
    sheet.append([None, "合计", sum_amount, sum_count, sum_last_amount, sum_last_count, sum_change_amount, rate])
    workbook.save(result_file_path)
    workbook.close()

def handle_14(file_dir, last_file_dir):
    file_name = "本月项目生产成本暂估比例异常(达20%以上)情况统计表.xlsx"
    origin_file_path = os.path.join(file_dir, "origin", f"{check_file_dict['error_excel']}.xlsx")
    if not os.path.exists(origin_file_path):
        print("14文件缺失")
        return
    last_file_path = os.path.join(last_file_dir, file_name)
    result_list = generate_result_list(origin_file_path, last_file_path)
    result_file_dir = os.path.join(file_dir, "result")
    if not os.path.join(result_file_dir):
        os.makedirs(result_file_dir)
    result_file_path = os.path.join(result_file_dir, file_name)
    generate_file(result_list, result_file_path)
    print("14end")