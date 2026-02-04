import json
import os
import re
from datetime import datetime

import requests
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

# === 引入 Utils ===
# [修改] 移除了 n8n_helper
from utils.excel_helper import get_raw_files, process_and_split
from utils.email_helper import generate_preview, send_emails_batch

app = Flask(__name__)
CORS(app)

# ==================== 配置參數 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_DIR = os.path.join(BASE_DIR, 'RawData')
OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
CONTACT_LIST_PATH = os.path.join(BASE_DIR, 'ContactList.xlsx')
RECIPIENT_CACHE_FILE = os.path.join(BASE_DIR, 'recipients_cache.json')

os.makedirs(RAW_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# [新增] 内置 SMTP 服务商配置 (架构师建议：将其硬编码在后端比前端更安全、更可控)
SMTP_PROVIDERS = {
    'gmail': {'host': 'smtp.gmail.com', 'port': 587},
    'office365': {'host': 'smtp.office365.com', 'port': 587},
    'qq': {'host': 'smtp.qq.com', 'port': 465},
    '163': {'host': 'smtp.163.com', 'port': 465},
    'aliyun': {'host': 'smtp.mxhichina.com', 'port': 465}
}

def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def safe_chinese_filename(filename):
    filename = os.path.basename(filename)
    clean_name = re.sub(r'[^\w\u4e00-\u9fa5\.\-]', '', filename)
    return clean_name

def load_workflows():
    try:
        with open(os.path.join(BASE_DIR, 'workflows.json'), 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []

# --- 服務端收件人緩存讀寫 ---
def load_recipient_cache():
    if not os.path.exists(RECIPIENT_CACHE_FILE):
        return {}
    try:
        with open(RECIPIENT_CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_recipient_cache(data):
    try:
        with open(RECIPIENT_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except:
        return False

# ==================== 頁面路由 ====================
@app.route('/')
def index():
    workflows = load_workflows()
    return render_template('dashboard.html', workflows=workflows)

@app.route('/tool/<tool_id>')
def tool_runner(tool_id):
    if tool_id == 'excel-split':
        return render_template('excel_tool.html')
    workflows = load_workflows()
    workflow = next((w for w in workflows if w['id'] == tool_id), None)
    if not workflow:
        return "找不到該工具配置", 404
    return render_template('tool_runner.html', workflow=workflow)

# ==================== API 路由：Excel 工具專用 ====================
@app.route('/api/excel/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "未檢測到檔案"})
    file = request.files['file']
    force_overwrite = request.form.get('force') == 'true'
    if file.filename == '':
        return jsonify({"status": "error", "message": "未選擇檔案"})
    if file and allowed_file(file.filename):
        try:
            filename = safe_chinese_filename(file.filename)
            save_path = os.path.join(RAW_DATA_DIR, filename)
            if os.path.exists(save_path) and not force_overwrite:
                return jsonify({
                    "status": "exists",
                    "message": f"檔案 [{filename}] 已存在，是否覆蓋？",
                    "filename": filename
                })
            file.save(save_path)
            action_msg = "覆蓋" if force_overwrite else "上傳"
            return jsonify({"status": "success", "message": f"檔案 {filename} {action_msg}成功"})
        except Exception as e:
            return jsonify({"status": "error", "message": f"保存失敗: {str(e)}"})
    else:
        return jsonify({"status": "error", "message": "僅支援 .xlsx 或 .xls 格式"})

@app.route('/api/excel/delete', methods=['POST'])
def delete_file():
    data = request.json
    filename = data.get('filename')
    if not filename:
        return jsonify({"status": "error", "message": "未指定檔案名稱"})
    safe_name = safe_chinese_filename(filename)
    file_path = os.path.join(RAW_DATA_DIR, safe_name)
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            return jsonify({"status": "success", "message": f"已刪除檔案: {safe_name}"})
        except Exception as e:
            return jsonify({"status": "error", "message": f"刪除失敗: {str(e)}"})
    else:
        return jsonify({"status": "error", "message": "檔案不存在或已被刪除"})

@app.route('/api/excel/files')
def list_files():
    keyword = request.args.get('keyword', '条线')
    return jsonify(get_raw_files(RAW_DATA_DIR, header_keyword=keyword))

# [修改] 移除了 sync_data 路由

@app.route('/api/excel/split', methods=['POST'])
def split_files():
    selection = request.json.get('selection')
    if not selection:
        return jsonify({"status": "error", "message": "未接收到有效的配置信息"})
    result = process_and_split(selection, RAW_DATA_DIR, OUTPUT_DIR)
    return jsonify(result)

@app.route('/api/excel/preview')
def preview_emails():
    original_list = generate_preview(OUTPUT_DIR, CONTACT_LIST_PATH)
    cache = load_recipient_cache()
    for item in original_list:
        line_name = item['line']
        if line_name in cache:
            item['recipients'] = cache[line_name]
    return jsonify(original_list)

@app.route('/api/excel/save_recipient', methods=['POST'])
def save_recipient():
    data = request.json
    line = data.get('line')
    recipients = data.get('recipients', [])
    if not line:
        return jsonify({"status": "error", "message": "無效的業務條線"})
    cache = load_recipient_cache()
    cache[line] = recipients
    if save_recipient_cache(cache):
        return jsonify({"status": "success", "message": "已保存"})
    else:
        return jsonify({"status": "error", "message": "寫入文件失敗"})

# [修改] 重构发送逻辑：支持 Provider 预设
@app.route('/api/excel/send', methods=['POST'])
def send_emails():
    data = request.json
    target_date = data.get('yyyymm', '')
    email_body = data.get('email_body', '')

    # 提取 SMTP 信息
    provider_key = data.get('smtp_provider', 'gmail')
    smtp_user = data.get('smtp_user')
    smtp_password = data.get('smtp_password')

    if not smtp_user or not smtp_password:
        return jsonify({"error": "请输入完整的发件账号与密码"})

    # 组装配置
    provider_config = SMTP_PROVIDERS.get(provider_key)
    if not provider_config:
        return jsonify({"error": "不支持的邮件服务商"})

    full_smtp_config = {
        "server": provider_config['host'],
        "port": provider_config['port'],
        "user": smtp_user,
        "password": smtp_password
    }

    return jsonify(send_emails_batch(
        full_smtp_config,
        data['dispatch_list'],
        OUTPUT_DIR,
        target_date,
        body_template=email_body
    ))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)