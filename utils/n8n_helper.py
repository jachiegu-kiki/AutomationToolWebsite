import requests
import logging

def trigger_n8n_sync(webhook_url, target_date):
    """
    触发 N8N 并回传年月数据
    """
    try:
        # 将用户选择的日期发送至 N8N
        payload = {"trigger": "manual_web_ui", "target_date": target_date}
        response = requests.post(webhook_url, json=payload, timeout=15)
        response.raise_for_status()
        return {"status": "success", "message": f"已成功触发 {target_date} 数据抓取。"}
    except Exception as e:
        logging.error(f"N8N Sync Failed: {e}")
        return {"status": "error", "message": f"连接 N8N 失败: {str(e)}"}