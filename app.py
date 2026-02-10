import json
import os
import re
from datetime import datetime

import requests
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

# === 引入 Utils ===
from utils.excel_helper import (
    get_raw_files, process_and_split,
    clear_directory, build_group_mapping
)
from utils.n8n_helper import trigger_n8n_sync
from utils.email_helper import generate_preview, send_emails_batch

app = Flask(__name__)
CORS(app)

# ==================== 配置參數 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_DIR = os.path.join(BASE_DIR, 'RawData')
OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
CONTACT_LIST_PATH = os.path.join(BASE_DIR, 'ContactList.xlsx')

# 持久化配置文件
RECIPIENT_CACHE_FILE = os.path.join(BASE_DIR, 'recipients_cache.json')
GROUP_CONFIG_FILE = os.path.join(BASE_DIR, 'group_config.json')

os.makedirs(RAW_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

N8N_BASE_URL = os.getenv('N8N_URL', 'http://192.168.1.9:5678')
N8N_EXCEL_SYNC_URL = f"{N8N_BASE_URL}/webhook/SplitData"

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


# ==================== 通用工具函数 ====================

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_chinese_filename(filename):
    """保留中文字符、英數字、點、下劃線、中劃線"""
    filename = os.path.basename(filename)
    clean_name = re.sub(r'[^\w\u4e00-\u9fa5\.\-]', '', filename)
    return clean_name


def load_workflows():
    try:
        with open(os.path.join(BASE_DIR, 'workflows.json'), 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []


# --- 收件人緩存 ---
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


# --- 分组参照配置 ---
def load_group_config():
    """
    加载分组参照路径配置
    返回: { "ref_benqi": "本期路径", "ref_tongqi": "同期路径" }
    """
    default = {"ref_benqi": "", "ref_tongqi": ""}
    if not os.path.exists(GROUP_CONFIG_FILE):
        return default
    try:
        with open(GROUP_CONFIG_FILE, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        # 兼容旧版 key (ref_26/ref_25 → ref_benqi/ref_tongqi)
        if "ref_26" in cfg and "ref_benqi" not in cfg:
            cfg["ref_benqi"] = cfg.pop("ref_26", "")
        if "ref_25" in cfg and "ref_tongqi" not in cfg:
            cfg["ref_tongqi"] = cfg.pop("ref_25", "")
        return {
            "ref_benqi": cfg.get("ref_benqi", ""),
            "ref_tongqi": cfg.get("ref_tongqi", ""),
        }
    except:
        return default


def save_group_config(data):
    try:
        with open(GROUP_CONFIG_FILE, 'w', encoding='utf-8') as f:
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


# ==================== API 路由：通用 ====================

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

# --- 文件管理 ---

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
    return jsonify(get_raw_files(RAW_DATA_DIR))


# --- 同步 API：先清空 RawData 再触发 N8N ---

@app.route('/api/excel/sync', methods=['POST'])
def sync_data():
    date_val = request.json.get('date')

    # 第一步：清空 RawData
    clear_result = clear_directory(RAW_DATA_DIR)
    app.logger.info(
        f"同步前清空 RawData: 删除 {clear_result['removed']} 个, "
        f"失败 {clear_result['failed']} 个"
    )

    # 第二步：触发 N8N
    sync_result = trigger_n8n_sync(N8N_EXCEL_SYNC_URL, date_val)

    if sync_result.get('status') == 'success':
        sync_result['message'] = (
            f"已清空 RawData ({clear_result['removed']} 个旧文件)，"
            f"并成功触发 {date_val} 数据抓取。"
        )
    elif clear_result['removed'] > 0:
        sync_result['message'] = (
            f"已清空 RawData ({clear_result['removed']} 个旧文件)，"
            f"但 N8N 触发失败: {sync_result.get('message', '未知错误')}"
        )

    return jsonify(sync_result)


# --- 分组参照配置 API ---

@app.route('/api/excel/group_config', methods=['GET'])
def get_group_config():
    config = load_group_config()
    for key in ['ref_benqi', 'ref_tongqi']:
        path = config.get(key, '')
        config[f'{key}_exists'] = bool(path and os.path.exists(path))
    return jsonify(config)


@app.route('/api/excel/group_config', methods=['POST'])
def set_group_config():
    data = request.json
    config = {
        "ref_benqi": data.get("ref_benqi", "").strip(),
        "ref_tongqi": data.get("ref_tongqi", "").strip(),
    }

    if save_group_config(config):
        status_info = {}
        for key in ['ref_benqi', 'ref_tongqi']:
            path = config[key]
            status_info[key] = {
                "path": path,
                "exists": bool(path and os.path.exists(path))
            }
        return jsonify({
            "status": "success",
            "message": "配置已保存",
            "files": status_info
        })
    else:
        return jsonify({"status": "error", "message": "配置保存失败"})


@app.route('/api/excel/test_mapping', methods=['POST'])
def test_group_mapping():
    """测试分组映射是否能正常加载"""
    config = load_group_config()
    result = build_group_mapping(config)
    return jsonify({
        "status": "success" if result["stats"]["total"] > 0 else "warning",
        "stats": result["stats"],
        "logs": result["logs"]
    })


# --- 拆分 API：自动加载分组映射后再拆分 ---

@app.route('/api/excel/split', methods=['POST'])
def split_files():
    selection = request.json.get('selection')
    if not selection:
        return jsonify({"status": "error", "message": "未接收到有效的配置信息"})

    # 检查是否有任何 Sheet 选了分组类别
    needs_mapping = False
    for fn_config in selection.values():
        for sheet_info in fn_config.values():
            if sheet_info.get('group_type', ''):
                needs_mapping = True
                break

    # 加载分组映射
    group_mappings = None
    if needs_mapping:
        group_config = load_group_config()
        has_ref = any(group_config.get(k, '').strip() for k in ['ref_benqi', 'ref_tongqi'])
        if has_ref:
            try:
                mapping_result = build_group_mapping(group_config)
                if mapping_result["stats"]["total"] > 0:
                    group_mappings = mapping_result["mappings"]
                    app.logger.info(
                        f"分组映射加载成功: "
                        f"签约 {mapping_result['stats']['签约_total']}, "
                        f"退费 {mapping_result['stats']['退费_total']}"
                    )
                else:
                    app.logger.warning("分组映射为空")
            except Exception as e:
                app.logger.error(f"分组映射加载失败: {e}")
        else:
            app.logger.info("未配置分组参照路径")

    # 执行拆分
    result = process_and_split(selection, RAW_DATA_DIR, OUTPUT_DIR, group_mappings)
    return jsonify(result)


# --- 邮件相关 API ---

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


@app.route('/api/excel/send', methods=['POST'])
def send_emails():
    data = request.json
    target_date = data.get('yyyymm', '')
    return jsonify(send_emails_batch(
        data['smtp_config'],
        data['dispatch_list'],
        OUTPUT_DIR,
        target_date
    ))


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
