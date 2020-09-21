import os
import openpyxl
from datetime import datetime, timedelta
from .report import Report

class Report_1(Report):

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            data_dict[company] = data_dict[company] + amount if company in data_dict else amount
        return data_dict, {}

    def get_result_list(self, data_dict, count_dict):
        last_data_dict = self.get_last_data()
        result_list = []
        for company, amount in data_dict.items():
            last_amount = float(last_data_dict.get(company, 0))
            handle_amount, change_amount, rate = Report.generate_change_info(amount, last_amount)
            result_list.append([company, handle_amount, last_amount, change_amount, rate])
            if change_amount > 0:
                self.save_company_list.append(company)
        result_list.sort(key=lambda x: x[3], reverse=True)
        sum_amount = 0
        sum_last_amount = 0
        sum_change_amount = 0
        for result_index, result in enumerate(result_list[:]):
            sum_amount += result[1]
            sum_last_amount += result[2]
            sum_change_amount += result[3]
            result_list[result_index].insert(0, result_index+1)
        sum_rate = Report.generate_percent_rate(sum_change_amount, sum_last_amount)
        sum_list = [None, "合计", sum_amount, sum_last_amount, sum_change_amount, sum_rate]
        result_list.append(sum_list)
        return result_list

class Report_2(Report_1):

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            if data[4] == "国网系统内-集团内":
                data_dict[company] = data_dict[company] + amount if company in data_dict else amount
        return data_dict, {}

class Report_3(Report):

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            year_amount = sum(float(i) if i else 0 for i in data[5:8])
            data_dict[company] = (data_dict[company][0] + amount, data_dict[company][1] + year_amount) if company in data_dict else (amount, year_amount)
        return data_dict, {}

    def get_result_list(self, data_dict, count_dict):
        last_data_dict = self.get_last_data()
        result_list = []
        for company, amount_group in data_dict.items():
            amount, year_amount = amount_group
            self.save_company_list.append(company)
            last_amount_group = last_data_dict.get(company, (0, 0))
            last_amount, last_year_amount = [float(i) for i in last_amount_group]
            handle_amount, change_amount, rate = Report.generate_change_info(amount, last_amount)
            year_handle_amount, year_change_amount, year_rate = Report.generate_change_info(year_amount, last_year_amount)
            result_list.append(
                [company, handle_amount, last_amount, change_amount, rate, year_handle_amount, last_year_amount,
                 year_change_amount, year_rate])
        result_list.sort(key=lambda x: x[1], reverse=True)
        sum_amount = 0
        sum_last_amount = 0
        sum_change_amount = 0
        sum_year_amount = 0
        sum_last_year_amount = 0
        sum_year_change_amount = 0
        for result_index, result in enumerate(result_list):
            sum_amount += result[1]
            sum_last_amount += result[2]
            sum_change_amount += result[3]
            sum_year_amount += result[5]
            sum_last_year_amount += result[6]
            sum_year_change_amount += result[7]
            result_list[result_index].insert(0, result_index+1)
        sum_rate = Report.generate_percent_rate(sum_change_amount, sum_last_amount)
        sum_year_rate = Report.generate_percent_rate(sum_year_change_amount, sum_last_year_amount)
        sum_list = [None, "合计", sum_amount, sum_last_amount, sum_change_amount, sum_rate, sum_year_amount,
                    sum_last_year_amount, sum_year_change_amount, sum_year_rate]
        result_list.append(sum_list)
        return result_list

class Report_4(Report_3):

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            year_amount = sum(float(i) if i else 0 for i in data[5:8])
            if data[4] == "国网系统内-集团内":
                data_dict[company] = (data_dict[company][0] + amount, data_dict[company][1] + year_amount) if company in data_dict else (amount, year_amount)
        return data_dict, {}

class Report_5(Report):

    def fix_data_list(self, data_list, order_field_list):
        now_date_str = os.path.split(self.file_dir)[1]
        year, month = [int(i) for i in now_date_str.split('-')]
        last_date = datetime(year=year, month=month, day=31)
        remove_data_list = []
        for data_index, data in enumerate(data_list[:]):
            data_date_str = data[4]
            if not data_date_str:
                remove_data_list.append(data)
                continue
            data_date = datetime.strptime(str(data_date_str), "%Y%m%d")
            if data_date > last_date:
                print(data_date)
                remove_data_list.append(data)
                continue
        for remove_data in remove_data_list:
            data_list.remove(remove_data)

    def merge_data(self, data_list):
        data_dict = {}
        count_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            count = int(data[5]) if data[5] else 0
            data_dict[company] = data_dict[company] + amount if company in data_dict else amount
            count_dict[company] = count_dict[company] + count if company in count_dict else count
        return data_dict, count_dict

    def get_result_list(self, data_dict, count_dict):
        result_list = []
        for company, amount in data_dict.items():
            self.save_company_list.append(company)
            count = count_dict[company]
            handle_amount = Report.handle_num(amount)
            result_list.append([company, handle_amount, count])
        result_list.sort(key=lambda x:x[1], reverse=True)
        sum_amount = 0
        sum_count = 0
        for result_index, result in enumerate(result_list):
            sum_amount += result[1]
            sum_count += result[2]
            result_list[result_index].insert(0, result_index+1)
        result_list.append([None, "合计", sum_amount, sum_count])
        return result_list

class Report_6(Report):

    def fix_data_list(self, data_list, order_field_list):
        order_field_list.insert(1, "公司名称")
        remove_data_list = []
        for data_index, data in enumerate(data_list[:]):
            year_amount = sum(float(i) if i else 0 for i in data[3:7])
            code = data[0]
            department = self.code_department_dict.get(code)
            data_list[data_index].insert(1, department)
            if year_amount <= 0:
                remove_data_list.append(data_list[data_index])
        filter_data_list = data_list[:]
        for remove_data in remove_data_list:
            filter_data_list.remove(remove_data)
        return filter_data_list


    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            year_amount = sum(float(i) if i else 0 for i in data[4:8])
            data_dict[company] = [data_dict[company][0] + amount,
                                  data_dict[company][1] + year_amount] if company in data_dict else [amount,
                                                                                                     year_amount]
        return data_dict, {}

    def get_result_list(self, data_dict, count_dict):
        last_data_dict = self.get_last_data()
        result_list = []
        for company, data in data_dict.items():
            last_data = last_data_dict.get(company, (0, 0))
            amount, year_amount = data
            self.save_company_list.append(company)
            last_amount, last_year_amount = [float(i) for i in last_data]
            handle_amount, change_amount, rate = Report.generate_change_info(amount, last_amount)
            handle_year_amount, change_year_amount, year_rate = Report.generate_change_info(year_amount, last_year_amount)
            result_list.append(
                [company, handle_amount, last_amount, change_amount, rate, handle_year_amount, last_year_amount,
                 change_year_amount, year_rate])
        result_list.sort(key=lambda x: x[5], reverse=True)
        sum_amount = 0
        sum_last_amount = 0
        sum_year_amount = 0
        sum_last_year_amount = 0
        for result_index, result in enumerate(result_list):
            sum_amount += result[1]
            sum_last_amount += result[2]
            sum_year_amount += result[5]
            sum_last_year_amount += result[6]
            result_list[result_index].insert(0, result_index+1)
        sum_change_amount = sum_amount - sum_last_amount
        sum_change_year_amount = sum_year_amount - sum_last_year_amount
        sum_rate = Report.generate_percent_rate(sum_change_amount, sum_last_amount)
        sum_year_rate = Report.generate_percent_rate(sum_change_year_amount, sum_last_year_amount)
        result_list.append([None, "合计", sum_amount, sum_last_amount, sum_change_amount, sum_rate, sum_year_amount,
                      sum_last_year_amount, sum_change_year_amount, sum_year_rate])
        return result_list

class Report_7(Report_6):

    def fix_data_list(self, data_list, order_field_list):
        order_field_list.insert(0, "公司名称")
        order_field_list.insert(0, "公司代码")
        for data_index, data in enumerate(data_list[:]):
            centre = data[0]
            if centre == "0000009999":
                code = None
                department = "虚拟"
            else:
                code = centre[1:5]
                department = self.code_department_dict.get(code)
            data_list[data_index].insert(0, department)
            data_list[data_index].insert(0, code)

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            year_amount = sum(float(i) if i else 0 for i in data[4:7])
            data_dict[company] = [data_dict[company][0] + amount,
                                  data_dict[company][1] + year_amount] if company in data_dict else [amount,
                                                                                                     year_amount]
        return data_dict, {}

class Report_8(Report_7):
    pass

class Report_9(Report_7):

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            amount = float(data[3]) if data[3] else 0
            year_amount = sum(float(i) if i else 0 for i in data[4:9])
            data_dict[company] = [data_dict[company][0] + amount,
                                  data_dict[company][1] + year_amount] if company in data_dict else [amount,
                                                                                                     year_amount]
        return data_dict, {}

class Report_10(Report_7):
    pass

class Report_11(Report_7):
    pass

class Report_12(Report):

    def fix_data_list(self, data_list, order_field_list):
        print(order_field_list)
        order_field_list[0] = "公司名称"
        print(order_field_list)

    def replace_company(self, data_list):
        pass

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[0]
            amount1 = float(data[1]) if data[1] else 0
            amount2 = float(data[2]) if data[2] else 0
            data_dict[company] = (
            data_dict[company][0] + amount1, data_dict[company][1] + amount2) if company in data_dict else (
            amount1, amount2)
        return data_dict, {}

    def get_result_list(self, data_dict, count_dict):
        result_list = []
        for company, data in data_dict.items():
            self.save_company_list.append(company)
            amount1, amount2 = data
            diff_amount = amount1 - amount2
            result_list.append([company, Report.handle_num(amount1), Report.handle_num(amount2), Report.handle_num(diff_amount)])
        result_list.sort(key=lambda x: x[3], reverse=True)
        sum_amount1 = 0
        sum_amount2 = 0
        for result_index, result in enumerate(result_list):
            sum_amount1 += result[1]
            sum_amount2 += result[2]
            result_list[result_index].insert(0, result_index+1)
        result_list.append([None, "合计", sum_amount1, sum_amount2, sum_amount1 - sum_amount2])
        return result_list

class Report_13(Report):

    def fix_data_list(self, data_list, order_field_list):
        insert_field_list = ["滞后开票(个数)", "滞后开票(金额)", "预开票(个数)", "预开票(金额)"]
        for insert_field in insert_field_list:
            order_field_list.insert(6, insert_field)
        remove_data_list = []
        for data_index, data in enumerate(data_list[:]):
            if data[6] == "一致" or data[1] in data[3]:
                remove_data_list.append(data)
                continue
            amount1 = float(data[4]) if data[4] else 0
            amount2 = float(data[5]) if data[5] else 0
            insert_data_list = [0, 0, 0, 0]
            if amount1 > amount2:
                insert_data_list = [amount1-amount2, 1, 0, 0]
            elif amount1 < amount2:
                insert_data_list = [0, 0, amount2-amount1, 1]
            insert_data_list.reverse()
            for insert_data in insert_data_list:
                data_list[data_index].insert(6, insert_data)
        for remove_data in remove_data_list:
            data_list.remove(remove_data)

    def merge_data(self, data_list):
        data_dict = {}
        for data in data_list:
            company = data[1]
            if company not in data_dict:
                data_dict[company] = [0, 0, 0, 0]
            data_dict[company] = [sum(i) for i in zip(data_dict[company], data[6:10])]
        return data_dict, {}

    def get_result_list(self, data_dict, count_dict):
        result_list = []
        for company, data in data_dict.items():
            pre_amount, pre_count, suf_amount, suf_count = data
            if pre_amount != 0 or suf_amount != 0:
                self.save_company_list.append(company)
            result_list.append([company, Report.handle_num(pre_amount), pre_count, Report.handle_num(suf_amount), suf_count])
        result_list.sort(key=lambda x: x[1], reverse=True)
        sum_pre_amount = 0
        sum_pre_count = 0
        sum_suf_amount = 0
        sum_suf_count = 0
        for result_index, result in enumerate(result_list):
            sum_pre_amount += result[1]
            sum_pre_count += result[2]
            sum_suf_amount += result[3]
            sum_suf_count += result[4]
            result_list[result_index].insert(0, result_index+1)
        result_list.append([None, "合计", sum_pre_amount, sum_pre_count, sum_suf_amount, sum_suf_count])
        return result_list

class Report_14(Report):

    def fix_data_list(self, data_list, order_field_list):
        for data in data_list[:]:
            rate = data[4]
            if not rate or float(rate.strip("%")) < 20 or float(rate.strip("%")) > 100:
                data_list.remove(data)

    def merge_data(self, data_list):
        data_dict = {}
        count_dict = {}
        for data in data_list:
            company = data[1]
            amount = data[3]
            amount = float(amount) if amount else 0
            data_dict[company] = data_dict[company] + amount if company in data_dict else amount
            count_dict[company] = count_dict[company] + 1 if company in count_dict else 1
        return data_dict, count_dict

    def get_result_list(self, data_dict, count_dict):
        last_data_dict = self.get_last_data()
        result_list = []
        for company, amount in data_dict.items():
            self.save_company_list.append(company)
            count = count_dict[company]
            last_data = [float(i) for i in last_data_dict.get(company, (0, 0))]
            last_amount, last_count = last_data
            handle_amount, change_amount, rate = Report.generate_change_info(amount, last_amount)
            result_list.append([company, handle_amount, count, last_amount, last_count, change_amount, rate])
        result_list.sort(key=lambda x: x[1], reverse=True)
        sum_amount = 0
        sum_count = 0
        sum_last_amount = 0
        sum_last_count = 0
        for result_index, result in enumerate(result_list):
            sum_amount += result[1]
            sum_count += result[2]
            sum_last_amount += result[3]
            sum_last_count += result[4]
            result_list[result_index].insert(0, result_index+1)
        sum_change_amount = sum_amount - sum_last_amount
        rate = Report.generate_percent_rate(sum_change_amount, sum_last_amount)
        result_list.append([None, "合计", sum_amount, sum_count, sum_last_amount, sum_last_count, sum_change_amount, rate])
        return result_list







