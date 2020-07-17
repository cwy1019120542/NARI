main_config_parameter = [('config_name', str, False), ('function_type', int, False), ('send_config_id', int, True), ('receive_config_id', int, True), ('email', str, False), ('password', str, False)]
user_parameter = [('email', str, False), ('captcha', int, False), ('password', str, False), ('name', str, False), ('department', str, False)]
send_config_parameter = [('subject', str, False), ('content', str, False), ('sheet', str, False), ('field_row', str, False), ('split_field', str, False), ("is_split", str, False), ('is_timing', bool, False),
                         ('start_timestamp', int, True), ('ip', str, False), ('port', int, False), ('agreement', str, False)]
receive_config_parameter = [('subject', str, False), ('sheet_info', str, False), ('is_timing', bool, False),
                         ('start_timestamp', int, True), ('ip', str, False), ('port', int, False), ('agreement', str, False), ('is_remind', bool, False), ('remind_subject', str, False),
                            ('remind_ip', str, False), ('remind_port', int, False), ('remind_agreement', str, False), ('remind_content', str, False), ('read_start_timestamp', int, False), ('read_end_timestamp', int, False)]
remind_parameter = [('is_all', bool, False), ('email', str, False)]
update_message_premeter = [('content', str, False)]
