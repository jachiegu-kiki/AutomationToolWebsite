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
        df.columns = [str(c).strip() for c in df.columns]
        contacts = {}
        for _, row in df.iterrows():
            line = str(row['条线']).strip()
            emails = [e.strip() for e in str(row['邮箱']).replace(';', ',').split(',') if e.strip()]
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


def send_emails_batch(smtp_config, dispatch_list, output_dir, yyyymm=""):
    """
    执行批量邮件发送
    :param yyyymm: 目标年月 (e.g. "2023-10")
    """
    results = {"success": [], "failed": []}
    port = int(smtp_config['port'])

    # 1. 建立 SMTP 连接
    try:
        if port == 465:
            server = smtplib.SMTP_SSL(smtp_config['server'], port, timeout=10)
        else:
            server = smtplib.SMTP(smtp_config['server'], port, timeout=10)
            try:
                server.starttls()  # 尝试 STARTTLS 加密
            except:
                pass
        server.login(smtp_config['user'], smtp_config['password'])
    except Exception as e:
        return {"error": f"SMTP 登录失败: {str(e)}"}

    # 2. 遍历分发列表发送
    for item in dispatch_list:
        department = item['line']  # 获取业务条线名称

        if not item['recipients']:
            results['failed'].append({"line": department, "reason": "名单无收件人"})
            continue

        try:
            msg = MIMEMultipart()

            # [核心修改] 构造包含月份和条线的标题
            # 例如: "2023-10 - 华北项目部 - 数据报告"
            date_prefix = f"{yyyymm} " if yyyymm else ""
            msg['Subject'] = f"{date_prefix}{department} - 业务数据报告"

            msg['From'] = smtp_config['user']
            msg['To'] = ", ".join(item['recipients'])

            # 邮件正文
            msg.attach(MIMEText(f"""老师们好,
                    {yyyymm}【{department}】业务部门数据出炉啦,请查阅;
                         ***附件资料包含当财年及同期的明细数据(毛签、退费、报完成、佣金以及截止当月的预收存量)***
                    此邮件由系统自动发送,如有疑问请联系相关人员。
                谢谢!
                广州前途财务部""", 'plain'))

            # 附加 Excel 文件
            file_path = os.path.join(output_dir, item['file'])
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=item['file'])
                part['Content-Disposition'] = f'attachment; filename="{item["file"]}"'
                msg.attach(part)
            else:
                raise Exception("附件文件不存在")

            # 发送
            server.sendmail(smtp_config['user'], item['recipients'], msg.as_string())
            results['success'].append(department)

        except Exception as e:
            results['failed'].append({"line": department, "reason": str(e)})

    server.quit()
    return results