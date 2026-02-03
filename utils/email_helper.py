import smtplib
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def parse_contact_list(path):
    if not os.path.exists(path): return {}
    try:
        df = pd.read_excel(path)
        # 统一转字符串并去空格
        df.columns = [str(c).strip() for c in df.columns]

        # 简单的列名适配，防止用户 ContactList 表头不完全一致
        line_col = next((c for c in df.columns if "条线" in c or "部门" in c or "Line" in c), None)
        email_col = next((c for c in df.columns if "邮箱" in c or "Email" in c or "mail" in c), None)

        if not line_col or not email_col:
            return {}

        contacts = {}
        for _, row in df.iterrows():
            line = str(row[line_col]).strip()
            emails = [e.strip() for e in str(row[email_col]).replace(';', ',').replace('，', ',').split(',') if e.strip()]
            contacts[line] = list(set(contacts.get(line, []) + emails))
        return contacts
    except:
        return {}


def generate_preview(output_dir, contact_path):
    contacts = parse_contact_list(contact_path)
    preview = []
    if not os.path.exists(output_dir): return preview
    for filename in os.listdir(output_dir):
        if filename.endswith(".xlsx"):
            line_name = os.path.splitext(filename)[0]
            preview.append({"line": line_name, "file": filename, "recipients": contacts.get(line_name, [])})
    return preview


def send_emails_batch(smtp_config, dispatch_list, output_dir, yyyymm="", body_template=""):
    """
    执行批量邮件发送
    :param yyyymm: 目标年月 (e.g. "2023-10")
    :param body_template: 用户自定义的邮件正文模版，支持 {line_name}, {month}
    """
    results = {"success": [], "failed": []}

    try:
        port = int(smtp_config['port'])
    except ValueError:
        return {"error": "SMTP 端口必须是数字"}

    # 1. 建立 SMTP 连接
    try:
        if port == 465:
            server = smtplib.SMTP_SSL(smtp_config['server'], port, timeout=15)
        else:
            server = smtplib.SMTP(smtp_config['server'], port, timeout=15)
            try:
                server.starttls()  # 尝试 STARTTLS 加密
            except:
                pass # 有些服务器25端口不需要TLS
        server.login(smtp_config['user'], smtp_config['password'])
    except Exception as e:
        return {"error": f"SMTP 登录失败: {str(e)}"}

    # 默认模版（兜底）
    if not body_template:
        body_template = """老师们好,
    {month}【{line_name}】业务数据出炉, 请查阅附件。
    
    谢谢!"""

    # 2. 遍历分发列表发送
    for item in dispatch_list:
        department = item['line']  # 获取业务条线名称

        if not item['recipients']:
            results['failed'].append({"line": department, "reason": "名单无收件人"})
            continue

        try:
            msg = MIMEMultipart()

            # 构造标题
            msg['Subject'] = f"{yyyymm} {department} 数据报表"
            msg['From'] = smtp_config['user']
            msg['To'] = ", ".join(item['recipients'])

            # === 模版变量替换 ===
            # 将用户输入的模版变量替换为实际值
            current_body = body_template.replace("{line_name}", department) \
                .replace("{month}", yyyymm)

            # 添加邮件正文 (使用 plain 格式，保留换行符)
            msg.attach(MIMEText(current_body, 'plain', 'utf-8'))

            # 附加 Excel 文件
            file_path = os.path.join(output_dir, item['file'])
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=item['file'])
                part['Content-Disposition'] = f'attachment; filename="{item["file"]}"'
                msg.attach(part)
            else:
                raise Exception(f"附件文件 {item['file']} 不存在")

            # 发送
            server.sendmail(smtp_config['user'], item['recipients'], msg.as_string())
            results['success'].append(department)

        except Exception as e:
            results['failed'].append({"line": department, "reason": str(e)})

    try:
        server.quit()
    except:
        pass

    return results