"""
今日工作待办 - 后端数据服务器
支持飞书表格双向同步 + 自动从预计出货时间生成待办提醒
"""
import json
import os
import sys
import uuid
import time
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import urllib.request
import urllib.error

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'reminders_data.json')
CONFIG_FILE = os.path.join(BASE_DIR, 'feishu_config.json')
HTML_FILE = os.path.join(BASE_DIR, 'supplier-reminder.html')

# Serve frontend HTML
@app.route('/')
def serve_html():
    if os.path.exists(HTML_FILE):
        return send_file(HTML_FILE)
    else:
        return '<h1>错误：找不到 supplier-reminder.html</h1><p>请确保该文件与 reminder_server.py 在同一目录</p>', 404

# --- Feishu Config ---
FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")
SPREADSHEET_TOKEN = os.environ.get("SPREADSHEET_TOKEN", "")
SHEET_ID = os.environ.get("SHEET_ID", "aa2707")  # 超时汇总

# Column mapping: website field -> Feishu column index (0-based A=0)
# Columns: A=供应商, B=款号, C=品类, D=款式, E=底胚, F=缺货数量, G=未送汇总, H=更新时间, I=进度, J=预计出货时间
FEISHU_COLUMNS = {
    "supplier": 0,    # A: 供应商
    "style": 1,       # B: 款号
    "category": 2,    # C: 品类
    "color": 3,       # D: 款式
    "fabric": 4,      # E: 底胚
    "qty": 5,         # F: 缺货数量
    "undelivered": 6, # G: 未送汇总
    "update_time": 7, # H: 更新时间
    "progress": 8,    # I: 进度
    "est_delivery": 9 # J: 预计出货时间
}

# --- Token Cache ---
_token_cache = {"token": None, "expires_at": 0}

def get_feishu_token():
    """获取飞书 tenant_access_token（带缓存）"""
    now = time.time()
    if _token_cache["token"] and _token_cache["expires_at"] > now + 60:
        return _token_cache["token"]

    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    data = json.dumps({"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        resp = urllib.request.urlopen(req, timeout=10)
        result = json.loads(resp.read())
        if result.get("code") == 0:
            _token_cache["token"] = result["tenant_access_token"]
            _token_cache["expires_at"] = now + result.get("expire", 7200)
            return _token_cache["token"]
        return None
    except Exception as e:
        print(f"获取飞书Token失败: {e}")
        return None


def feishu_api(method, path, body=None, params=None):
    """调用飞书 Open API"""
    token = get_feishu_token()
    if not token:
        return {"code": -1, "msg": "无法获取飞书Token"}

    url = f"https://open.feishu.cn/open-apis{path}"
    if params:
        qs = "&".join(f"{k}={urllib.parse.quote(str(v))}" for k, v in params.items())
        url += f"?{qs}"

    data = json.dumps(body).encode() if body else None
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        resp = urllib.request.urlopen(req, timeout=15)
        return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        return {"code": -1, "msg": f"HTTP {e.code}: {e.reason}"}
    except Exception as e:
        return {"code": -1, "msg": str(e)}


def read_feishu_sheet(range_str):
    """读取飞书表格指定范围（返回FormattedValue）"""
    import urllib.parse
    params = {
        "valueRenderOption": "FormattedValue",
        "dateTimeRenderOption": "FormattedString"
    }
    qs = "&".join(f"{k}={urllib.parse.quote(v)}" for k, v in params.items())
    path = f"/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{urllib.parse.quote(range_str, safe='')}?{qs}"
    result = feishu_api("GET", path)
    if result.get("code") == 0:
        return result["data"]["valueRange"]["values"]
    return []


def write_feishu_cells(range_str, values):
    """写入飞书表格单元格"""
    body = {"valueRange": {"range": range_str, "values": values}}
    path = f"/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    return feishu_api("PUT", path, body=body)


def parse_feishu_date(raw):
    """解析飞书返回的各种日期格式 → YYYY-MM-DD"""
    import re
    if not raw:
        return ""
    raw = str(raw).strip()

    # 1. 标准日期：2026-07-04 / 2026/07/04
    for pat in [r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})']:
        m = re.match(pat, raw)
        if m:
            return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"

    # 2. 中文日期：7月4日 / 6月27日
    m = re.match(r'(\d{1,2})月(\d{1,2})日?', raw)
    if m:
        return f"2026-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}"

    # 3. 中旬/下旬：7月中旬 / 8月下旬
    m = re.match(r'(\d{1,2})月[中下]旬', raw)
    if m:
        day = '15' if '中' in raw else '25'
        return f"2026-{m.group(1).zfill(2)}-{day}"

    # 4. 月/日格式：7/4
    m = re.match(r'(\d{1,2})/(\d{1,2})$', raw)
    if m:
        return f"2026-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}"

    # 5. Excel serial number (skip if it looks like a small number like 6)
    try:
        serial = int(raw)
        if serial > 45000:  # roughly 2023+
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=serial)).strftime('%Y-%m-%d')
    except:
        pass

    return ""


def excel_serial_to_date(serial):
    return parse_feishu_date(serial)


def data_to_reminder(row, row_index):
    """将飞书行数据映射为提醒格式"""
    def cell(idx):
        return row[idx] if idx < len(row) and row[idx] else ""

    est_raw = cell(FEISHU_COLUMNS["est_delivery"])
    est_date = excel_serial_to_date(est_raw) if est_raw else ""

    return {
        "id": f"fs_{row_index}",
        "supplier": cell(FEISHU_COLUMNS["supplier"]),
        "style": cell(FEISHU_COLUMNS["style"]),
        "color": cell(FEISHU_COLUMNS["color"]),
        "qty": int(cell(FEISHU_COLUMNS["undelivered"])) if cell(FEISHU_COLUMNS["undelivered"]) else 0,
        "date": est_date,
        "time": "",
        "remark": cell(FEISHU_COLUMNS["progress"]),
        "completed": False,
        "source": "feishu",
        "feishu_row": row_index,
        "category": cell(FEISHU_COLUMNS["category"]),
        "fabric": cell(FEISHU_COLUMNS["fabric"]),
        "undelivered": cell(FEISHU_COLUMNS["undelivered"]),
        "feishu_update": cell(FEISHU_COLUMNS["update_time"]),
    }


# --- Data helpers ---
def load_data():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ============ API Routes ============

@app.route('/api/reminders', methods=['GET'])
def get_reminders():
    """获取所有提醒（本地 + 飞书合并）"""
    local = load_data()
    # 读取飞书数据
    feishu_data = []
    try:
        rows = read_feishu_sheet(f"{SHEET_ID}!A3:J500")
        for i, row in enumerate(rows):
            if row and any(c for c in row):
                r = data_to_reminder(row, i + 3)
                # 只保留有预计出货时间的行作为提醒
                if r["date"]:
                    feishu_data.append(r)
    except Exception as e:
        print(f"读取飞书数据失败: {e}")

    # 合并：飞书数据优先，本地补充（以款号为key去重）
    local_styles = {r.get("style", ""): r for r in local}
    merged = {}
    for r in feishu_data:
        key = r.get("style", "")
        merged[key] = r
    for r in local:
        key = r.get("style", "")
        if key and key not in merged:
            merged[key] = r

    return jsonify(list(merged.values()))


@app.route('/api/reminders', methods=['POST'])
def save_reminders():
    """批量保存提醒"""
    data = request.get_json()
    if not isinstance(data, list):
        return jsonify({'error': '数据格式错误'}), 400
    for item in data:
        if 'id' not in item or not item['id']:
            item['id'] = 'r_' + str(int(datetime.now().timestamp() * 1000)) + '_' + uuid.uuid4().hex[:6]
    save_data(data)
    return jsonify({'ok': True, 'count': len(data)})


@app.route('/api/reminders/<reminder_id>', methods=['POST'])
def update_reminder(reminder_id):
    """更新单条提醒"""
    data = load_data()
    body = request.get_json()
    updated = False
    for item in data:
        if item.get('id') == reminder_id:
            item.update(body)
            updated = True
            break
    if not updated:
        return jsonify({'error': '未找到'}), 404
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/reminders/today', methods=['GET'])
def get_today_reminders():
    """获取今日待办（含逾期）"""
    data = load_data()
    today = datetime.now().strftime('%Y-%m-%d')
    result = []
    for item in data:
        if item.get('completed'):
            continue
        item_date = item.get('date', '')
        if not item_date:
            continue
        if item_date <= today:
            result.append(item)
    result.sort(key=lambda x: (x.get('date', '') > today, x.get('date', '')))
    return jsonify(result)


@app.route('/api/reminders/daily-summary', methods=['GET'])
def get_daily_summary():
    data = load_data()
    today = datetime.now().strftime('%Y-%m-%d')
    weekday = ['日','一','二','三','四','五','六'][datetime.now().weekday()]
    overdue, today_items = [], []
    for item in data:
        if item.get('completed'):
            continue
        d = item.get('date', '')
        if not d:
            continue
        if d < today:
            overdue.append(item)
        elif d == today:
            today_items.append(item)

    lines = [f"【今日工作待办 - {today} 星期{weekday}】", ""]
    if overdue:
        lines.append(f"🔴 已逾期 ({len(overdue)}条)：")
        for r in overdue:
            p = [r.get('style','')]
            if r.get('supplier'): p.append(f"[{r['supplier']}]")
            if r.get('color'): p.append(r['color'])
            if r.get('qty'): p.append(f"x{r['qty']}")
            t = f" {r['time']}" if r.get('time') else ''
            rm = f" — {r['remark']}" if r.get('remark') else ''
            lines.append(f"  • {' '.join(p)} — {r['date']}{t}{rm}")
        lines.append("")
    if today_items:
        lines.append(f"🟠 今天到期 ({len(today_items)}条)：")
        for r in today_items:
            p = [r.get('style','')]
            if r.get('supplier'): p.append(f"[{r['supplier']}]")
            if r.get('color'): p.append(r['color'])
            if r.get('qty'): p.append(f"x{r['qty']}")
            t = f" {r['time']}" if r.get('time') else ''
            rm = f" — {r['remark']}" if r.get('remark') else ''
            lines.append(f"  • {' '.join(p)}{t}{rm}")
        lines.append("")
    if not overdue and not today_items:
        lines.append("✅ 今日无待办提醒")
    lines.append("---")
    lines.append("请尽快跟进供应商交货情况")
    return jsonify({
        'text': '\n'.join(lines),
        'overdue_count': len(overdue),
        'today_count': len(today_items),
        'total_active': len([i for i in data if not i.get('completed')])
    })


# ============ 飞书同步 API ============

@app.route('/api/feishu/data', methods=['GET'])
def get_feishu_data():
    """读取飞书表格原始数据"""
    try:
        rows = read_feishu_sheet(f"{SHEET_ID}!A3:J500")
        result = []
        for i, row in enumerate(rows):
            if row and any(c for c in row):
                result.append(data_to_reminder(row, i + 3))
        return jsonify({"code": 0, "data": result, "count": len(result)})
    except Exception as e:
        return jsonify({"code": -1, "msg": str(e)}), 500


@app.route('/api/feishu/sync', methods=['POST'])
def sync_to_feishu():
    """同步本地提醒到飞书表格"""
    local = load_data()
    body = request.get_json() or {}
    items = body.get('items', local)

    # 先读取飞书现有数据（包含款号行号映射）
    rows = read_feishu_sheet(f"{SHEET_ID}!B3:B500")
    style_row_map = {}
    for i, row in enumerate(rows):
        if row:
            style_row_map[row[0]] = i + 3  # row number in sheet

    synced = 0
    errors = []

    for item in items:
        style = item.get('style', '')
        if not style:
            continue
        date = item.get('date', '')
        remark = item.get('remark', '')

        if style in style_row_map:
            row_num = style_row_map[style]
            # 更新预计出货时间(J列)和进度(I列)
            try:
                write_feishu_cells(f"{SHEET_ID}!I{row_num}:J{row_num}", [[remark, date]])
                synced += 1
            except Exception as e:
                errors.append(f"{style}: {e}")

    return jsonify({"ok": True, "synced": synced, "errors": errors})


@app.route('/api/feishu/auto-reminders', methods=['GET'])
def auto_reminders():
    """自动从飞书预计出货时间生成提醒"""
    try:
        rows = read_feishu_sheet(f"{SHEET_ID}!A3:J500")
        auto_reminders = []
        for i, row in enumerate(rows):
            if not row or not any(c for c in row):
                continue
            r = data_to_reminder(row, i + 3)
            if r["date"]:
                auto_reminders.append(r)

        return jsonify({"code": 0, "data": auto_reminders, "count": len(auto_reminders)})
    except Exception as e:
        return jsonify({"code": -1, "msg": str(e)}), 500


@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        'status': 'ok',
        'time': datetime.now().isoformat(),
        'feishu': 'connected' if get_feishu_token() else 'disconnected'
    })


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5199))
    print(f"数据文件: {DATA_FILE}")
    print(f"飞书表格: {SPREADSHEET_TOKEN}")
    print(f"启动于端口: {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
