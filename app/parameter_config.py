main_config_parameter = {
    "POST": [('config_name', str, False, 50), ('function_type', int, False, 11), ('send_config_id', int, True, 11), ('receive_config_id', int, True, 11), ('email', str, False, 100), ('password', str, False, 100)],
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('config_name', str, False, 50), ('function_type', int, False, 11), ('send_config_id', int, True, 11), ('receive_config_id', int, True, 11), ('email', str, False, 50), ('status', int, False, 11)],
    "fuzzy_field": ["config_name", "email"]
}
user_parameter = {
    "POST": [('email', str, False, 50), ('captcha', int, False, 11), ('password', str, False, 100), ('name', str, False, 10), ('department', str, False, 10)],
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('email', str, False, 100), ('name', str, False, 10), ('department', str, False, 10), ('status', int, False, 11)],
    "PUT": [('password', str, False, 100), ('name', str, False, 10), ('department', str, False, 10)],
    "fuzzy_field": ["email", "name", 'department']
}

send_config_parameter = {
    "POST": [('subject', str, False, 50), ('content', str, False, 10000), ('sheet', str, False, 10000), ('field_row', str, False, 50), ('split_field', str, False, 50), ("is_split", str, False, 50), ('is_timing', bool, False, 10),
                         ('start_timestamp', int, True, 11), ('ip', str, False, 50), ('port', int, False, 11), ("send_excel_name", str, True, 50), ("send_excel_field", str, True, 10000)],
    "GET":[('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('status', int, False, 11)],
    "fuzzy_field": []
}
receive_config_parameter = {
    "POST": [('subject', str, False, 50), ('sheet_info', str, False, 10000), ('is_timing', bool, False, 10),
                         ('start_timestamp', int, True, 11), ('ip', str, False, 50), ('port', int, False, 11), ('is_remind', bool, False, 10), ('remind_subject', str, False, 50),
                            ('remind_ip', str, False, 50), ('remind_port', int, False, 11), ('remind_content', str, False, 10000), ('read_start_timestamp', int, False, 11), ('read_end_timestamp', int, False, 11), ("template_excel_name", str, True, 50), ("template_excel_field", str, True, 10000)],
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('status', int, False, 11)],
    "fuzzy_field": []
}
remind_parameter = {
    "POST": [('email', str, False, 10000)]
}
update_message_premeter = {
    "POST": [('content', str, False, 10000)]
}
send_history_parameter = {
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('target', str, False, 50), ('email', str, False, 100), ('main_config_id', int, False, 11), ('status', int, False, 11), ('is_success', bool, False, 10), ('message', str, False, 50)],
    "fuzzy_field": ["target", "email", "message"]
}
receive_history_parameter = {
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('target', str, False, 50), ('email', str, False, 100), ('main_config_id', int, False, 11), ('status', int, False, 11), ('is_success', bool, False, 10), ('message', str, False, 50)],
    "fuzzy_field": ["target", "email", "message"]
}
sap_config_parameter = {
    "GET": [('id', int, False, 11), ('offset', int, False, 11), ('limit', int, False, 11), ('status', int, False, 11)],
    "POST": [("account", str, False, 50), ("password", str, False, 100), ("main_body", str, False, 100), ("subject", str, False, 100), ("start_date", str, False, 20), ("end_date", str, False, 20)],
    "fuzzy_field": []
}
sap_log_parameter = {
    "POST": [("log", str, False, 100)]
}
accept_file_type = (".xlsx", ".xls", ".XLSX", ".XLS")