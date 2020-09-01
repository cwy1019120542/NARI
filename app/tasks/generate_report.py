import os
import openpyxl
from ..extention import celery
from .report_func.handle_1_2 import handle_1_2
from .report_func.handle_3_4 import  handle_3_4
from .report_func.handle_5 import handle_5
from .report_func.handle_6 import handle_6
from .report_func.handle_7 import handle_7
from .report_func.handle_8 import handle_8
from .report_func.handle_9 import handle_9
from .report_func.handle_10 import handle_10
from .report_func.handle_11 import handle_11
from .report_func.handle_12 import handle_12
from .report_func.handle_13 import handle_13
from .report_func.handle_14 import handle_14
from .report_func.public_func import generate_dict

@celery.task
def generate_report(file_dir, code_file_path, last_file_dir, last_year_file_dir):
    workbook = openpyxl.load_workbook(code_file_path, data_only=True)
    code_sheet = workbook[workbook.sheetnames[0]]
    code_department_dict = generate_dict(code_sheet, 1, ["公司代码"], ["单位名称"])
    centre_sheet = workbook[workbook.sheetnames[1]]
    centre_company_dict = generate_dict(centre_sheet, 1, ["利润中心"], ["分公司及事业部名称"])
    # print(centre_company_dict)
    workbook.close()
    handle_1_2(file_dir, last_file_dir, centre_company_dict)
    handle_3_4(file_dir, last_year_file_dir, centre_company_dict)
    handle_5(file_dir, centre_company_dict)
    handle_6(file_dir, last_year_file_dir, centre_company_dict, code_department_dict)
    handle_7(file_dir, last_year_file_dir, centre_company_dict, code_department_dict)
    handle_8(file_dir, last_year_file_dir, centre_company_dict, code_department_dict)
    handle_9(file_dir, last_year_file_dir, centre_company_dict, code_department_dict)
    handle_10(file_dir, last_year_file_dir, centre_company_dict, code_department_dict)
    handle_11(file_dir, last_year_file_dir, centre_company_dict, code_department_dict)
    handle_12(file_dir)
    handle_13(file_dir, centre_company_dict)
    handle_14(file_dir, last_year_file_dir)






