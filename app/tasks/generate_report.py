import os
import openpyxl
from ..extention import celery
from .report_func.handle_1_2 import handle_1_2
from .report_func.handle_3_4 import  handle_3_4
from .report_func.public_func import generate_dict

@celery.task
def generate_report(file_dir, code_file_path, last_file_dir, last_year_file_dir):
    code_excel = openpyxl.load_workbook(code_file_path, data_only=True)
    code_sheet = code_excel[code_excel.sheetnames[0]]
    code_dict = generate_dict(code_sheet, 1, ["利润中心"], ["分公司及事业部名称"])
    code_excel.close()
    handle_1_2(file_dir, last_file_dir, code_dict)
    handle_3_4(file_dir, last_file_dir, code_dict)







