import json
import os
from datetime import datetime

import requests
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

# === 引入新模組的 Utils (請確保 utils 資料夾已建立) ===
from utils.excel_helper import get_raw_files, process_and_split
from utils.n8n_helper import trigger_n8n_sync
from utils.email_helper import generate_preview, send_emails_batch

app = Flask(__name__)
CORS(app)

# ==================== 配置參數 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_DIR = os.path.join(BASE_DIR, 'RawData')
OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
CONTACT_LIST_PATH = os.path.join(BASE_DIR, 'ContactList.xlsx')

# 確保目錄存在
os.makedirs(RAW_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

N8N_BASE_URL = os.getenv('N8N_URL', 'http://192.168.1.9:5678')
N8N_EXCEL_SYNC_URL = f"{N8N_BASE_URL}/webhook/SplitData"

def load_workflows():
    """讀取設定檔，用於生成首頁的選單"""
    try:
        with open(os.path.join(BASE_DIR, 'workflows.json'), 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []

# ==================== 頁面路由 (View Routes) ====================

@app.route('/')
def index():
    """
    [核心路由] 系統首頁
    說明：此處強制顯示 Dashboard 儀表板，不會直接進入任何工具。
    """
    workflows = load_workflows()
    # 渲染儀表板頁面，傳入工具列表供選單顯示
    return render_template('dashboard.html', workflows=workflows)


@app.route('/tool/<tool_id>')
def tool_runner(tool_id):
    """
    [工具路由] 根據 ID 決定進入哪個子系統
    """
    # 1. 判斷是否為「Excel 智能拆分系統」
    if tool_id == 'excel-split':
        # 只有當網址是 /tool/excel-split 時，才顯示 Excel 工具頁面
        return render_template('excel_tool.html')

    # 2. 其他通用工具 (既有的 N8N 工具)
    workflows = load_workflows()
    workflow = next((w for w in workflows if w['id'] == tool_id), None)

    if not workflow:
        return "找不到該工具配置", 404

    return render_template('tool_runner.html', workflow=workflow)

# ==================== API 路由：通用 (維持不變) ====================

@app.route('/api/n8n/health')
def n8n_health_check():
    try:
        response = requests.get(f'{N8N_BASE_URL}/webhook/healthz', timeout=2)
        return jsonify({'success': response.status_code == 200})
    except:
        return jsonify({'success': False})

@app.route('/api/proxy', methods=['POST'])
def proxy_request():
    try:
        data = request.json
        webhook_suffix = data.get('webhook_suffix')
        payload = data.get('payload', {})
        payload['timestamp'] = datetime.now().isoformat()
        payload['source'] = 'generic-web-interface'

        full_url = f"{N8N_BASE_URL}{webhook_suffix}"
        response = requests.post(full_url, json=payload, timeout=30)
        try:
            resp_data = response.json()
        except:
            resp_data = {}

        if response.status_code == 200:
            return jsonify({'success': True, 'message': '執行成功', 'data': resp_data})
        else:
            return jsonify({'success': False, 'message': f'N8N 錯誤: {response.status_code}'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ==================== API 路由：Excel 工具專用 ====================

@app.route('/api/excel/files')
def list_files():
    return jsonify(get_raw_files(RAW_DATA_DIR))

@app.route('/api/excel/sync', methods=['POST'])
def sync_data():
    date_val = request.json.get('date')
    return jsonify(trigger_n8n_sync(N8N_EXCEL_SYNC_URL, date_val))

@app.route('/api/excel/split', methods=['POST'])
def split_files():
    selection = request.json.get('selection')
    if not selection:
        return jsonify({"status": "error", "message": "未接收到有效的配置信息"})
    result = process_and_split(selection, RAW_DATA_DIR, OUTPUT_DIR)
    return jsonify(result)

@app.route('/api/excel/preview')
def preview_emails():
    return jsonify(generate_preview(OUTPUT_DIR, CONTACT_LIST_PATH))

@app.route('/api/excel/send', methods=['POST'])
def send_emails():
    data = request.json
    return jsonify(send_emails_batch(data['smtp_config'], data['dispatch_list'], OUTPUT_DIR))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)