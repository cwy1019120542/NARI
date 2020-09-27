main_config_parameter = {
    "POST": [('config_name', str, False, 50), ('function_type', int, False, 11), ('send_config_id', int, True, 11), ('receive_config_id', int, True, 11), ('email', str, False, 100), ('password', str, False, 100)],
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('config_name', str, False, 50), ('function_type', int, False, 11), ('send_config_id', int, True, 11), ('receive_config_id', int, True, 11), ('email', str, False, 50)],
    "fuzzy_field": ["config_name", "email"]
}
user_parameter = {
    "POST": [('email', str, False, 50), ('captcha', int, False, 11), ('password', str, False, 100), ('name', str, False, 10), ('department', str, False, 10)],
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('email', str, False, 100), ('name', str, False, 10), ('department', str, False, 10)],
    "PUT": [('password', str, False, 100), ('name', str, False, 10), ('department', str, False, 10)],
    "fuzzy_field": ["email", "name", 'department']
}

send_config_parameter = {
    "POST": [('subject', str, False, 50), ('content', str, False, 10000), ('split_type', int, False, 11), ('sheet', str, False, 10000),
             ('field_row', str, False, 50), ('split_field', str, False, 50), ("is_split", str, False, 50), ('is_timing', bool, False, 10),
            ('start_timestamp', int, True, 11), ('ip', str, False, 50), ('port', int, False, 11),
             ("send_excel_uuid_name", str, True, 100), ("send_excel_field", str, True, 10000)],
    "GET":[('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11)],
    "fuzzy_field": []
}
receive_config_parameter = {
    "POST": [('subject', str, False, 50), ('sheet_info', str, False, 10000), ('is_timing', bool, False, 10),
                         ('start_timestamp', int, True, 11), ('ip', str, False, 50), ('port', int, False, 11), ('is_remind', bool, False, 10), ('remind_subject', str, False, 50),
                            ('remind_ip', str, False, 50), ('remind_port', int, False, 11), ('remind_content', str, False, 10000), ('read_start_timestamp', int, False, 11), ('read_end_timestamp', int, False, 11), ("template_excel_uuid_name", str, True, 100), ("template_excel_field", str, True, 10000)],
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11)],
    "fuzzy_field": []
}
remind_parameter = {
    "POST": [('email', str, False, 10000)]
}
update_message_premeter = {
    "POST": [('content', str, False, 10000)]
}
send_history_parameter = {
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('target', str, False, 50), ('email', str, False, 100), ('main_config_id', int, False, 11), ('is_success', bool, False, 10), ('message', str, False, 50)],
    "fuzzy_field": ["target", "email", "message"]
}
receive_history_parameter = {
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('target', str, False, 50), ('email', str, False, 100), ('main_config_id', int, False, 11), ('is_success', bool, False, 10), ('message', str, False, 50)],
    "fuzzy_field": ["target", "email", "message"]
}
sap_config_parameter = {
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ("config_type", int, False, 11)],
    "POST": [("config_type", int, False, 11), ("account", str, True, 50), ("password", str, True, 100), ("main_body", str, True, 100), ("subject", str, True, 100), ("start_date", str, True, 20), ("end_date", str, True, 20)],
    "fuzzy_field": []
}
sap_log_parameter = {
    "POST": [("log", str, False, 100)]
}
accept_file_type = (".xlsx", ".xls", ".XLSX", ".XLS")

check_file_dict = {
    "pre_excel": "预开票账龄表",
    "suf_excel": "滞后开票账龄表",
    "balance_analyse_excel": "投运完成在制品余额分析",
    "product_cost_excel": "项目生产成本挂账",
    "receive_excel": "内部关联交易一致性报表",
    "open_excel": "收入确认与开票不同步",
    "error_excel": "生产成本暂估异常情况分析",
    "pay_cost_excel": "应付暂估成本账龄表",
    "pay_receive_excel": "应付暂估收货账龄表",
    "pre_pay_excel": "预付款项账龄表",
    "other_receive_excel": "其他应收账龄表",
    "other_pay_excel": "其他应付账龄表",
}