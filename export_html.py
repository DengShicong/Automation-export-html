import os
import pandas as pd
from datetime import datetime
from netmiko import ConnectHandler
import html
import re
import threading


def parse_show_interfaces_alias(output):
    down_lines = [line for line in output.split('\n') if 'down' in line.lower()]
    return "\n".join(down_lines)

def parse_show_interfaces_status(output):
    down_lines = [line for line in output.split('\n') if 'full' in line.lower()]
    return "\n".join(down_lines)

def connect_and_execute_commands(device_info, cmds, log_dir, html_template):
    all_output = ""
    try:
        with ConnectHandler(**device_info) as net_connect:
            net_connect.enable()
            for cmd in cmds:
                if cmd.strip().lower() == 'show running-directory':
                    continue  # Skip 'show running-directory' command

                output = net_connect.send_command(cmd)

                if cmd.strip().lower() =='show interfaces status':
                    output = parse_show_interfaces_status(output)

                if cmd.strip().lower() == 'show interfaces alias':
                    output = parse_show_interfaces_alias(output)

                all_output += f"{device_info['host']}\tCommand: {cmd}\n{output}\n\n"

            if all_output.strip():
                # Using template and replace placeholders
                formatted_output = html_template.replace('{{command_outputs}}', f"<pre>{html.escape(all_output)}</pre>")
                formatted_output = formatted_output.replace('{{generation_time}}',
                                                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"{timestamp}_{device_info['host']}.html"
                filepath = os.path.join(log_dir, filename)

                with open(filepath, "w") as file:
                    file.write(formatted_output)

    except Exception as e:
        print(f"Connection failed for {device_info['host']}: {e}")


def execute_commands_and_save_to_html(template_path):
    template_data = pd.ExcelFile(template_path)
    assets_data = template_data.parse('assets')

    log_dir = "LOG-HTML"
    os.makedirs(log_dir, exist_ok=True)

    # Read the HTML template file
    with open('template.html', 'r') as file:
        html_template = file.read()

    threads = []

    for index, row in assets_data.iterrows():
        device_info = {
            "device_type": row["device_type"],
            "host": row["IP"],
            "username": row["username"],
            "password": row["password"],
            "port": int(row["port"]) if pd.notna(row["port"]) else 22,
            "secret": row["secret"]
        }

        cmds = template_data.parse(row["device_type"]).iloc[:, 1].dropna().tolist()

        thread = threading.Thread(target=connect_and_execute_commands, args=(device_info, cmds, log_dir, html_template))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()


template_path = 'template.xlsx'  # Replace with the actual path to your template file
execute_commands_and_save_to_html(template_path)