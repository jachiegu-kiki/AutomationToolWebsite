import json
import os
from datetime import datetime

import requests
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# 配置參數
N8N_BASE_URL = os.getenv('N8N_URL', 'http://192.168.1.9:5678')


def load_workflows():
    """讀取設定檔"""
    try:
        with open('workflows.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []


# ==================== 頁面路由 ====================

@app.route('/')
def index():
    """首頁：列出所有可用工具"""
    workflows = load_workflows()
    # 如果只有一個工具，直接跳轉進去（模擬原本的使用體驗），或者顯示列表
    # 這裡我們先顯示列表，方便未來擴充
    return render_template('dashboard.html', workflows=workflows)


@app.route('/tool/<tool_id>')
def tool_runner(tool_id):
    """通用工具執行器：根據 JSON 渲染頁面"""
    workflows = load_workflows()
    workflow = next((w for w in workflows if w['id'] == tool_id), None)

    if not workflow:
        return "找不到該工具配置", 404

    return render_template('tool_runner.html', workflow=workflow)


# ==================== API 路由 ====================

@app.route('/api/n8n/health')
def n8n_health_check():
    """檢查 N8N 健康狀態 (維持原功能)"""
    try:
        response = requests.get(f'{N8N_BASE_URL}/webhook/healthz', timeout=2)
        return jsonify({'success': response.status_code == 200})
    except:
        return jsonify({'success': False})


@app.route('/api/proxy', methods=['POST'])
def proxy_request():
    """
    通用代理接口
    接收前端的數據，轉發給 JSON 設定中指定的 N8N webhook
    """
    try:
        data = request.json
        webhook_suffix = data.get('webhook_suffix')
        payload = data.get('payload', {})

        # 自動注入時間戳與來源
        payload['timestamp'] = datetime.now().isoformat()
        payload['source'] = 'generic-web-interface'

        full_url = f"{N8N_BASE_URL}{webhook_suffix}"

        print(f"正在轉發請求到: {full_url}")
        print(f"數據內容: {payload}")

        response = requests.post(full_url, json=payload, timeout=30)

        # 嘗試解析 N8N 回傳的 JSON，如果不是 JSON 則回傳空字典
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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)