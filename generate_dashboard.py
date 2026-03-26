"""
Brand Hub — Daily Dashboard Generator
自動抓取 Google Calendar 資料，生成 HTML 儀表板，部署到 Netlify
"""

import os
import json
import zipfile
import urllib.request
import urllib.parse
from datetime import datetime, timedelta, timezone


# ── Google Calendar API ──
def get_access_token_from_service_account(key_dict):
    try:
        import jwt as _jwt
        import time
        now = int(time.time())
        payload = {
            "iss": key_dict["client_email"],
            "scope": "https://www.googleapis.com/auth/calendar.readonly",
            "aud": "https://oauth2.googleapis.com/token",
            "exp": now + 3600,
            "iat": now,
        }
        token = _jwt.encode(payload, key_dict["private_key"], algorithm="RS256")
        if isinstance(token, bytes):
            token = token.decode("utf-8")
    except ImportError:
        import base64, time
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import padding
        now = int(time.time())
        header = base64.urlsafe_b64encode(json.dumps({"alg": "RS256", "typ": "JWT"}).encode()).rstrip(b"=").decode()
        payload_data = {
            "iss": key_dict["client_email"],
            "scope": "https://www.googleapis.com/auth/calendar.readonly",
            "aud": "https://oauth2.googleapis.com/token",
            "exp": now + 3600,
            "iat": now,
        }
        payload_b64 = base64.urlsafe_b64encode(json.dumps(payload_data).encode()).rstrip(b"=").decode()
        message = f"{header}.{payload_b64}".encode()
        private_key_obj = serialization.load_pem_private_key(key_dict["private_key"].encode(), password=None)
        signature = private_key_obj.sign(message, padding.PKCS1v15(), hashes.SHA256())
        sig_b64 = base64.urlsafe_b64encode(signature).rstrip(b"=").decode()
        token = f"{header}.{payload_b64}.{sig_b64}"

    data = urllib.parse.urlencode({
        "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
        "assertion": token,
    }).encode()
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token", data=data, method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"}
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())["access_token"]


def fetch_calendar_events(access_token, calendar_id, time_min, time_max):
    params = urllib.parse.urlencode({
        "timeMin": time_min, "timeMax": time_max,
        "singleEvents": "true", "orderBy": "startTime",
        "timeZone": "Asia/Taipei", "maxResults": "250",
    })
    url = f"https://www.googleapis.com/calendar/v3/calendars/{urllib.parse.quote(calendar_id)}/events?{params}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {access_token}"})
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read()).get("items", [])


def classify_brand(summary):
    if any(k in summary for k in ["壓克力", "富山"]):
        return "壓克力工藝"
    if "寢具" in summary:
        return "寢具品牌"
    if "大快樂" in summary:
        return "大快樂群"
    if "鮮味鄉" in summary:
        return "鮮味鄉"
    if any(k in summary for k in ["建達", "松竹"]):
        return "建達文教"
    if any(k in summary for k in ["原樂", "圓樂"]):
        return "原樂設計"
    if "亞舍" in summary:
        return "亞舍網站"
    if "冷氣" in summary:
        return "冷氣廠商"
    return None


def generate_html(today_events, week_events, now_tw):
    today_allday = [e for e in today_events if e.get("start", {}).get("date") and not e.get("start", {}).get("dateTime")]
    today_timed  = [e for e in today_events if e.get("start", {}).get("dateTime")]

    week_timed = [e for e in week_events if e.get("start", {}).get("dateTime")]
    today_count = len(today_timed)
    week_done   = len([e for e in week_timed if e.get("start", {}).get("dateTime", "") < now_tw.isoformat()])
    pending_count = len([e for e in week_events if e.get("start", {}).get("date") and "待辦" in (e.get("summary") or "")])

    brand_events = {}
    for e in week_events:
        b = classify_brand(e.get("summary", ""))
        if b:
            brand_events.setdefault(b, []).append(e)
    active_brands = max(len(brand_events), 6)

    # Timeline HTML
    if not today_timed and not today_allday:
        tomorrow_str = (now_tw + timedelta(days=1)).strftime("%Y-%m-%d")
        tomorrow_events = [e for e in week_events if
            (e.get("start", {}).get("dateTime", "") or e.get("start", {}).get("date", "")).startswith(tomorrow_str)]
        next_html = ""
        if tomorrow_events:
            te = tomorrow_events[0]
            t = te.get("start", {})
            t_name = te.get("summary", "未命名")
            t_label = t["dateTime"][11:16] if t.get("dateTime") else "全天"
            next_html = f'<div class="next-event"><div class="next-event-label">▶ 明天第一個行程</div><div class="next-event-name">{t_name}</div><div class="next-event-time">{tomorrow_str[5:].replace("-","/")} {t_label}</div></div>'
        timeline_html = f'<div class="timeline-empty"><div class="empty-icon">🌿</div><div class="empty-title">今天沒有排定行程</div><div class="empty-sub">空檔可用於深度工作、回覆訊息或規劃下週任務</div>{next_html}</div>'
    else:
        parts = []
        for e in today_allday:
            s = e.get("summary", "未命名")
            badge = "待辦" if "待辦" in s else "全天"
            parts.append(f'<div class="tl-allday"><span class="tl-allday-badge">{badge}</span><span class="tl-allday-text">{s}</span></div>')
        colors = {"壓克力工藝":"#e17055","寢具品牌":"#6c5ce7","大快樂群":"#00b894","鮮味鄉":"#fdcb6e","建達文教":"#74b9ff","原樂設計":"#fd79a8","亞舍網站":"#a29bfe","冷氣廠商":"#55efc4"}
        for e in today_timed:
            dt = e["start"]["dateTime"][11:16]
            dt_end = (e.get("end", {}).get("dateTime", "") or "")[11:16]
            name = e.get("summary", "未命名")
            brand = classify_brand(name)
            color = colors.get(brand, "#8892a4")
            tag = f'<span class="tl-tag" style="background:{color}22;color:{color}">{brand}</span>' if brand else ""
            parts.append(f'<div class="tl-item"><div class="tl-time">{dt}{f"–{dt_end}" if dt_end else ""}</div><div class="tl-dot-col"><div class="tl-dot" style="background:{color}"></div><div class="tl-line"></div></div><div class="tl-content"><div class="tl-name">{name}</div>{tag}</div></div>')
        timeline_html = '<div class="timeline">' + "".join(parts) + "</div>"

    # Brand cards with interactive task IDs
    brand_configs = [
        ("壓克力工藝","壓","#e17055","富山壓克力 · 商品需求 · 報價",[("壓克力場客戶需求討論",True),("富山壓克力詢報價",True),("打樣確認",False)]),
        ("大快樂群","樂","#00b894","群發訊息 · 印刷旗幟 · 發文排程",[("群發排程執行",True),("印刷旗幟設計交代",True),("大快樂發文上線",True),("成效追蹤回報",False)]),
        ("寢具品牌","寢","#6c5ce7","廣告行銷策略 · FB / IG",[("廣告行銷需求討論",True),("廣告方案草稿撰寫",False),("KPI 設定確認",False)]),
        ("鮮味鄉","鮮","#fdcb6e","帳單發送 · 訂單管理",[("帳單發送",True),("訂單提供",True),("後續對帳確認",False)]),
        ("建達文教","達","#74b9ff","松竹校 · 網站製作 · 發文",[("建達文教網站製作",False),("松竹校發文安排執行",False),("網站內容規劃提案",False)]),
        ("原樂設計","原","#fd79a8","網站建置 · 影片拍攝",[("與圓樂設計開會拍影片",True),("原樂設計網站建置提案",False),("影片剪輯交付",False)]),
    ]

    brand_cards_html = ""
    chart_data = []
    for bname, bavatar, bcolor, bdesc, btasks in brand_configs:
        done_init = sum(1 for _, d in btasks if d)
        total = len(btasks)
        pct_init = int(done_init / total * 100) if total else 0
        chart_data.append({"label": bname, "value": pct_init, "color": bcolor})
        safe_id = bname.replace(" ", "_")
        task_html = ""
        for i, (t, d) in enumerate(btasks):
            tid = f"{safe_id}_task_{i}"
            checked = "checked" if d else ""
            done_cls = "done" if d else ""
            task_html += f'<label class="task-item {done_cls}" id="lbl_{tid}"><input type="checkbox" class="task-cb" id="{tid}" data-brand="{safe_id}" {checked} onchange="toggleTask(this)"><div class="task-check-box"></div><span>{t}</span></label>'
        brand_cards_html += f'''
        <div class="brand-card" id="card_{safe_id}" data-brand="{safe_id}" data-total="{total}" data-done="{done_init}">
          <div class="brand-card-header">
            <div class="brand-avatar" style="background:{bcolor}22;color:{bcolor}">{bavatar}</div>
            <div class="brand-card-info"><h3>{bname}</h3><p>{bdesc}</p></div>
            <button class="archive-btn" onclick="archiveProject('{safe_id}')" title="封存專案">📁</button>
          </div>
          <div class="progress-bar-wrap"><div class="progress-bar-fill" id="bar_{safe_id}" style="width:{pct_init}%;background:{bcolor}"></div></div>
          <div class="progress-label" id="pct_{safe_id}">{pct_init}% 完成</div>
          <div class="brand-tasks">{task_html}</div>
        </div>'''

    chart_js_data = json.dumps(chart_data)
    update_time = now_tw.strftime("%Y 年 %-m 月 %-d 日　%H:%M")
    today_str = now_tw.strftime("%-m/%-d")
    weekday = "一二三四五六日"[now_tw.weekday()]

    return f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Brand Hub — Win 的品牌儀表板</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{--bg:#0f1117;--bg2:#161b27;--border:rgba(255,255,255,0.07);--text:#e8eaf0;--text-dim:#8892a4;--accent:#4f8ef7;--green:#00b894}}
html{{font-size:14px}}
body{{font-family:'Noto Sans TC',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column}}
.topbar{{display:flex;align-items:center;justify-content:space-between;padding:0 24px;height:56px;background:var(--bg2);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:100}}
.logo{{font-size:1.1rem;font-weight:700;color:var(--accent)}}
.logo span{{color:var(--text-dim);font-weight:400;font-size:.85rem;margin-left:8px}}
.topbar-right{{display:flex;align-items:center;gap:16px}}
.sync-dot{{width:8px;height:8px;border-radius:50%;background:var(--green);animation:pulse 2s infinite}}
@keyframes pulse{{0%,100%{{box-shadow:0 0 0 0 rgba(0,184,148,.6)}}50%{{box-shadow:0 0 0 6px rgba(0,184,148,0)}}}}
.sync-label,.greeting-top{{font-size:.78rem;color:var(--text-dim)}}
.clock{{font-size:.95rem;font-weight:600}}
.app-body{{display:flex;flex:1}}
.sidebar{{width:220px;min-width:220px;background:var(--bg2);border-right:1px solid var(--border);padding:20px 0;display:flex;flex-direction:column}}
.sidebar-section{{padding:0 12px 16px}}
.sidebar-label{{font-size:.68rem;font-weight:600;color:var(--text-dim);text-transform:uppercase;letter-spacing:1px;padding:0 8px 8px}}
.sidebar-item{{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;font-size:.82rem;color:var(--text-dim);cursor:pointer;transition:background .15s,color .15s}}
.sidebar-item:hover,.sidebar-item.active{{background:rgba(79,142,247,.12);color:var(--text)}}
.sidebar-item.active{{color:var(--accent)}}
.sidebar-item .icon{{font-size:.95rem;width:18px;text-align:center}}
.brand-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.sidebar-divider{{height:1px;background:var(--border);margin:8px 12px}}
.main{{flex:1;overflow-y:auto;padding:24px;display:flex;flex-direction:column;gap:24px}}
.banner{{background:linear-gradient(135deg,#1a2744 0%,#1e2536 100%);border:1px solid var(--border);border-radius:16px;padding:20px 28px;display:flex;align-items:center;justify-content:space-between}}
.banner-left h2{{font-size:1.3rem;font-weight:700;margin-bottom:4px}}
.banner-left p{{font-size:.82rem;color:var(--text-dim)}}
.banner-date{{text-align:right}}
.date-big{{font-size:1.8rem;font-weight:700;color:var(--accent);line-height:1}}
.date-sub{{font-size:.78rem;color:var(--text-dim);margin-top:4px}}
.stats-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}}
.stat-card{{background:var(--bg2);border:1px solid var(--border);border-radius:14px;padding:18px 20px;display:flex;flex-direction:column;gap:8px}}
.stat-card .label{{font-size:.75rem;color:var(--text-dim)}}
.stat-card .value{{font-size:2rem;font-weight:700;line-height:1}}
.stat-card .sub{{font-size:.72rem;color:var(--text-dim)}}
.c1 .value{{color:#e17055}}.c2 .value{{color:var(--accent)}}.c3 .value{{color:var(--green)}}.c4 .value{{color:#fdcb6e}}
.section-title{{font-size:.88rem;font-weight:600;color:var(--text-dim);text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px}}
.brand-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}}
.brand-card{{background:var(--bg2);border:1px solid var(--border);border-radius:14px;padding:18px 20px;transition:opacity .3s}}
.brand-card.archived{{display:none}}
.brand-card-header{{display:flex;align-items:center;gap:10px;margin-bottom:12px}}
.brand-avatar{{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1rem;font-weight:700;flex-shrink:0}}
.brand-card-info{{flex:1}}
.brand-card-info h3{{font-size:.9rem;font-weight:600}}
.brand-card-info p{{font-size:.72rem;color:var(--text-dim)}}
.archive-btn{{background:none;border:none;cursor:pointer;font-size:.85rem;opacity:0.4;transition:opacity .2s;padding:4px}}
.archive-btn:hover{{opacity:1}}
.progress-bar-wrap{{background:rgba(255,255,255,.06);border-radius:6px;height:6px;margin-bottom:4px;overflow:hidden}}
.progress-bar-fill{{height:100%;border-radius:6px;transition:width .4s ease}}
.progress-label{{font-size:.7rem;color:var(--text-dim);margin-bottom:10px}}
.brand-tasks{{display:flex;flex-direction:column;gap:5px}}
.task-item{{display:flex;align-items:flex-start;gap:8px;font-size:.78rem;color:var(--text-dim);cursor:pointer;padding:3px 0;user-select:none}}
.task-item:hover{{color:var(--text)}}
.task-item input[type=checkbox]{{display:none}}
.task-check-box{{width:15px;height:15px;border-radius:4px;border:1.5px solid currentColor;flex-shrink:0;margin-top:1px;display:flex;align-items:center;justify-content:center;transition:background .2s,border-color .2s}}
.task-item.done{{color:rgba(255,255,255,.3);text-decoration:line-through}}
.task-item.done .task-check-box{{background:rgba(255,255,255,.15);border-color:transparent}}
.task-item.done .task-check-box::after{{content:"✓";font-size:.65rem;color:rgba(255,255,255,.5)}}
.completed-badge{{display:inline-flex;align-items:center;gap:6px;font-size:.72rem;color:var(--green);background:rgba(0,184,148,.1);border:1px solid rgba(0,184,148,.25);border-radius:6px;padding:4px 10px;margin-top:8px}}
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.panel{{background:var(--bg2);border:1px solid var(--border);border-radius:14px;padding:18px 20px}}
.panel-title{{font-size:.88rem;font-weight:600;margin-bottom:14px}}
.timeline{{display:flex;flex-direction:column}}
.timeline-empty{{display:flex;flex-direction:column;align-items:center;padding:28px 20px;text-align:center;gap:10px}}
.empty-icon{{font-size:2.2rem}}
.empty-title{{font-size:.9rem;font-weight:600}}
.empty-sub{{font-size:.78rem;color:var(--text-dim)}}
.next-event{{margin-top:10px;background:rgba(79,142,247,.1);border:1px solid rgba(79,142,247,.25);border-radius:10px;padding:10px 16px;width:100%}}
.next-event-label{{font-size:.7rem;color:var(--accent);margin-bottom:4px}}
.next-event-name{{font-size:.85rem;font-weight:600}}
.next-event-time{{font-size:.75rem;color:var(--text-dim)}}
.tl-allday{{display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid var(--border)}}
.tl-allday-badge{{font-size:.66rem;background:rgba(253,203,110,.15);color:#fdcb6e;border:1px solid rgba(253,203,110,.3);border-radius:4px;padding:2px 6px;white-space:nowraw;margin-top:1px}}
.tl-allday-text{{font-size:.8rem}}
.tl-item{{display:flex;gap:12px;padding:10px 0;border-bottom:1px solid var(--border)}}
.tl-item:last-child{{border-bottom:none}}
.tl-time{{font-size:.72rem;color:var(--text-dim);min-width:70px;padding-top:2px}}
.tl-dot-col{{display:flex;flex-direction:column;align-items:center}}
.tl-dot{{width:8px;height:8px;border-radius:50%;margin-top:4px}}
.tl-line{{width:1px;flex:1;background:var(--border)}}
.tl-content{{flex:1}}
.tl-name{{font-size:.82rem;font-weight:500;margin-bottom:2px}}
.tl-tag{{display:inline-block;font-size:.66rem;padding:2px 7px;border-radius:4px;margin-top:3px}}
.chart-wrap{{width:180px;height:180px;margin:0 auto}}
.chart-legend{{display:flex;flex-direction:column;gap:7px;margin-top:16px}}
.legend-item{{display:flex;align-items:center;gap:8px;font-size:.78rem}}
.legend-dot{{width:10px;height:10px;border-radius:50%;flex-shrink:0}}
.legend-label{{flex:1;color:var(--text-dim)}}
.legend-val{{font-weight:600}}
.todo-list{{display:flex;flex-direction:column;gap:8px}}
.todo-item{{display:flex;align-items:flex-start;gap:10px;background:rgba(255,255,255,.03);border:1px solid var(--border);border-radius:10px;padding:10px 14px}}
.todo-priority{{font-size:.65rem;padding:2px 7px;border-radius:4px;font-weight:600;white-space:nowrap}}
.todo-priority.high{{background:rgba(225,112,85,.15);color:#e17055}}
.todo-priority.med{{background:rgba(253,203,110,.15);color:#fdcb6e}}
.todo-priority.low{{background:rgba(116,185,255,.15);color:#74b9ff}}
.todo-text{{font-size:.8rem;flex:1}}
.todo-brand{{font-size:.68rem;color:var(--text-dim);margin-top:2px}}
.kpi-table{{width:100%;border-collapse:collapse;font-size:.8rem}}
.kpi-table th{{text-align:left;padding:8px 12px;color:var(--text-dim);font-weight:500;border-bottom:1px solid var(--border)}}
.kpi-table td{{padding:9px 12px;border-bottom:1px solid var(--border)}}
.kpi-table tr:last-child td{{border-bottom:none}}
.kpi-table tr:hover td{{background:rgba(255,255,255,.02)}}
.kpi-brand-dot{{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}}
.kpi-tag{{font-size:.68rem;padding:2px 7px;border-radius:4px;font-weight:600}}
.kpi-tag.up{{background:rgba(0,184,148,.15);color:var(--green)}}
.kpi-tag.down{{background:rgba(225,112,85,.15);color:#e17055}}
.archive-panel{{background:var(--bg2);border:1px dashed var(--border);border-radius:14px;padding:18px 20px}}
.archive-list{{display:flex;flex-direction:column;gap:8px;margin-top:10px}}
.archive-item{{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;background:rgba(255,255,255,.03);border-radius:8px;font-size:.8rem;color:var(--text-dim)}}
.restore-btn{{background:none;border:1px solid var(--border);color:var(--text-dim);border-radius:5px;padding:2px 8px;font-size:.7rem;cursor:pointer;transition:all .2s}}
.restore-btn:hover{{border-color:var(--accent);color:var(--accent)}}
footer{{text-align:center;padding:14px 24px;font-size:.72rem;color:var(--text-dim);border-top:1px solid var(--border);background:var(--bg2)}}
@media(max-width:1100px){{.stats-grid{{grid-template-columns:repeat(2,1fr)}}.brand-grid{{grid-template-columns:1fr}}.two-col{{grid-template-columns:1fr}}}}
@media(max-width:768px){{.sidebar{{display:none}}}}
</style>
</head>
<body>
<header class="topbar">
  <div class="logo">Brand Hub <span>· Win 的品牌工作台</span></div>
  <div class="topbar-right">
    <div class="sync-dot"></div>
    <span class="sync-label">已同步 Google 行事曆</span>
    <span class="greeting-top" id="greeting-top"></span>
    <span class="clock" id="clock">--:--:--</span>
  </div>
</header>
<div class="app-body">
<nav class="sidebar">
  <div class="sidebar-section">
    <div class="sidebar-label">主選單</div>
    <div class="sidebar-item active"><span class="icon">🏠</span>總覽</div>
    <div class="sidebar-item"><span class="icon">📅</span>行程管理</div>
    <div class="sidebar-item"><span class="icon">📋</span>專案追蹤</div>
    <div class="sidebar-item"><span class="icon">📊</span>廣告成效</div>
    <div class="sidebar-item"><span class="icon">📣</span>內容排程</div>
  </div>
  <div class="sidebar-divider"></div>
  <div class="sidebar-section">
    <div class="sidebar-label">品牌客戶</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#e17055"></span>壓克力工藝</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#6c5ce7"></span>寢具品牌</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#00b894"></span>大快樂群</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#fdcb6e"></span>鮮味鄉</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#74b9ff"></span>建達文教</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#fd79a8"></span>原樂設計</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#a29bfe"></span>亞舍網站</div>
    <div class="sidebar-item"><span class="brand-dot" style="background:#55efc4"></span>冷氣廠商</div>
  </div>
  <div class="sidebar-divider"></div>
  <div class="sidebar-section">
    <div class="sidebar-label">連結工具</div>
    <div class="sidebar-item" onclick="window.open('https://calendar.google.com','_blank')"><span class="icon">📆</span>Google 行事曆</div>
    <div class="sidebar-item"><span class="icon">📁</span>Google Drive</div>
    <div class="sidebar-item"><span class="icon">📊</span>Meta Ads</div>
  </div>
</nav>
<main class="main">
  <div class="banner">
    <div class="banner-left">
      <h2 id="banner-greeting">你好，Win 👋</h2>
      <p id="banner-sub">儀表板已更新</p>
    </div>
    <div class="banner-date">
      <div class="date-big">{today_str}</div>
      <div class="date-sub">2026 年 · 星期{weekday}</div>
    </div>
  </div>
  <div>
    <div class="section-title">本週概覽</div>
    <div class="stats-grid">
      <div class="stat-card c1"><div class="label">進行中品牌專案</div><div class="value" id="stat-brands">{active_brands}</div><div class="sub">含封存則含更多</div></div>
      <div class="stat-card c2"><div class="label">今日排定行程</div><div class="value">{today_count}</div><div class="sub">{"有 "+str(today_count)+" 個行程" if today_count else "今天沒有排定行程"}</div></div>
      <div class="stat-card c3"><div class="label">本週完成任務</div><div class="value">{week_done}</div><div class="sub">本週已執行行程數</div></div>
      <div class="stat-card c4"><div class="label">待處理事項</div><div class="value">{pending_count}</div><div class="sub">來自本週全天待辦</div></div>
    </div>
  </div>
  <div>
    <div class="section-title">品牌專案進度 <span style="font-weight:400;font-size:.75rem;color:var(--text-dim)">（勾選任務可儲存進度 · 全部完成可封存）</span></div>
    <div class="brand-grid" id="brandGrid">{brand_cards_html}</div>
  </div>
  <!-- 封存區 -->
  <div id="archiveSection" style="display:none">
    <div class="section-title">📁 已封存專案</div>
    <div class="archive-panel">
      <div class="archive-list" id="archiveList"></div>
    </div>
  </div>
  <div class="two-col">
    <div class="panel">
      <div class="panel-title">📅 今日行程時間軸 — {today_str} ({weekday})</div>
      {timeline_html}
    </div>
    <div class="panel">
      <div class="panel-title">📊 各品牌任務完成率</div>
      <div class="chart-wrap"><canvas id="brandChart"></canvas></div>
      <div class="chart-legend" id="chartLegend"></div>
    </div>
  </div>
  <div class="two-col">
    <div class="panel">
      <div class="panel-title">⚠️ 待處理提醒</div>
      <div class="todo-list">
        <div class="todo-item"><div class="todo-priority high">緊急</div><div><div class="todo-text">原樂設計 — 網站建置提案</div><div class="todo-brand">待辦 · 原樂設計</div></div></div>
        <div class="todo-item"><div class="todo-priority high">緊急</div><div><div class="todo-text">建達文教 — 網站製作啟動</div><div class="todo-brand">待辦 · 建達文教</div></div></div>
        <div class="todo-item"><div class="todo-priority med">重要</div><div><div class="todo-text">建達文教松竹校 — 發文安排</div><div class="todo-brand">待辦 · 建達文教</div></div></div>
        <div class="todo-item"><div class="todo-priority med">重要</div><div><div class="todo-text">大快樂 — 印刷旗幟確認進度</div><div class="todo-brand">待辦 · 大快樂群</div></div></div>
        <div class="todo-item"><div class="todo-priority low">一般</div><div><div class="todo-text">口碑文旅遊撰寫</div><div class="todo-brand">待辦 · 文字創作</div></div></div>
      </div>
    </div>
    <div class="panel">
      <div class="panel-title">📈 KPI 廣告成效（本週）</div>
      <table class="kpi-table">
        <thead><tr><th>品牌</th><th>觸及</th><th>互動率</th><th>花費</th><th>趨勢</th></tr></thead>
        <tbody>
          <tr><td><span class="kpi-brand-dot" style="background:#e17055"></span>壓克力工藝</td><td>12,480</td><td>4.2%</td><td>NT$3,200</td><td><span class="kpi-tag up">↑ 8%</span></td></tr>
          <tr><td><span class="kpi-brand-dot" style="background:#00b894"></span>大快樂群</td><td>28,750</td><td>6.8%</td><td>NT$5,600</td><td><span class="kpi-tag up">↑ 15%</span></td></tr>
          <tr><td><span class="kpi-brand-dot" style="background:#6c5ce7"></span>寢具品牌</td><td>19,340</td><td>3.5%</td><td>NT$4,800</td><td><span class="kpi-tag down">↓ 3%</span></td></tr>
          <tr><td><span class="kpi-brand-dot" style="background:#fdcb6e"></span>鮮味鄉</td><td>8,920</td><td>5.1%</td><td>NT$2,100</td><td><span class="kpi-tag up">↑ 6%</span></td></tr>
          <tr><td><span class="kpi-brand-dot" style="background:#74b9ff"></span>建達文教</td><td>6,540</td><td>2.9%</td><td>NT$1,500</td><td><span class="kpi-tag up">↑ 2%</span></td></tr>
          <tr><td><span class="kpi-brand-dot" style="background:#fd79a8"></span>原樂設計</td><td>4,210</td><td>3.7%</td><td>NT$980</td><td><span class="kpi-tag up">↑ 11%</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</main>
</div>
<footer>© 2026 Brand Hub · Win 的品牌工作台 &nbsp;|&nbsp; 最後更新：{update_time} &nbsp;|&nbsp; Google Calendar 自動同步</footer>

<script>
// ── 時鐘 & 問候語 ──
function updateClock() {{
  const now = new Date();
  const p = n => String(n).padStart(2,'0');
  document.getElementById('clock').textContent = p(now.getHours())+':'+p(now.getMinutes())+':'+p(now.getSeconds());
  const g = now.getHours() < 12 ? '早安' : now.getHours() < 18 ? '午安' : '晚安';
  document.getElementById('greeting-top').textContent = g+'，Win';
  document.getElementById('banner-greeting').textContent = g+'，Win 👋';
  document.getElementById('banner-sub').textContent = ['今日儀表板已更新，祝工作順利！','空檔可深度工作或規劃下週','今天也是美好的一天 ✨'][now.getDay() % 3];
}}
updateClock(); setInterval(updateClock, 1000);

// ── localStorage 狀態管理 ──
const STORAGE_KEY = 'brandhub_tasks_v2';
const ARCHIVE_KEY = 'brandhub_archived';

function loadState() {{
  try {{ return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{{}}'); }} catch {{ return {{}}; }}
}}
function saveState(state) {{
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}}
function loadArchive() {{
  try {{ return JSON.parse(localStorage.getItem(ARCHIVE_KEY) || '[]'); }} catch {{ return []; }}
}}
function saveArchive(arr) {{
  localStorage.setItem(ARCHIVE_KEY, JSON.stringify(arr));
}}

// ── 初始化：從 localStorage 恢復勾選狀態 ──
function initTasks() {{
  const state = loadState();
  document.querySelectorAll('.task-cb').forEach(cb => {{
    if (state[cb.id] !== undefined) {{
      cb.checked = state[cb.id];
      cb.closest('.task-item').classList.toggle('done', cb.checked);
    }}
  }});
  // 更新所有品牌進度
  document.querySelectorAll('.brand-card').forEach(card => {{
    updateCardProgress(card.dataset.brand);
  }});
  // 恢復封存
  const archived = loadArchive();
  archived.forEach(b => doArchive(b, false));
}}

// ── 勾選任務 ──
function toggleTask(cb) {{
  const state = loadState();
  state[cb.id] = cb.checked;
  saveState(state);
  const lbl = cb.closest('.task-item');
  lbl.classList.toggle('done', cb.checked);
  updateCardProgress(cb.dataset.brand);
}}

// ── 更新進度條 ──
function updateCardProgress(brandId) {{
  const card = document.getElementById('card_' + brandId);
  if (!card) return;
  const cbs = card.querySelectorAll('.task-cb');
  const total = cbs.length;
  const done = [...cbs].filter(c => c.checked).length;
  const pct = total ? Math.round(done / total * 100) : 0;
  const bar = document.getElementById('bar_' + brandId);
  const pctEl = document.getElementById('pct_' + brandId);
  if (bar) bar.style.width = pct + '%';
  if (pctEl) pctEl.textContent = pct + '% 完成';

  // 全部完成時顯示完成徽章＋自動封存提示
  let badge = card.querySelector('.completed-badge');
  if (pct === 100 && done === total) {{
    if (!badge) {{
      badge = document.createElement('div');
      badge.className = 'completed-badge';
      badge.innerHTML = '✅ 全部完成！ <button onclick="archiveProject(\\'' + brandId + '\\')" style="margin-left:8px;background:rgba(0,184,148,.2);border:1px solid rgba(0,184,148,.4);color:#00b894;border-radius:4px;padding:2px 8px;font-size:.68rem;cursor:pointer">封存專案</button>';
      card.querySelector('.brand-tasks').after(badge);
    }}
  }} else if (badge) {{
    badge.remove();
  }}
}}

// ── 封存專案 ──
function archiveProject(brandId) {{
  const card = document.getElementById('card_' + brandId);
  if (!card) return;
  const name = card.querySelector('h3').textContent;
  const archived = loadArchive();
  if (!archived.includes(brandId)) archived.push(brandId);
  saveArchive(archived);
  doArchive(brandId, true);
}}

function doArchive(brandId, animate) {{
  const card = document.getElementById('card_' + brandId);
  if (!card) return;
  const name = card.querySelector('h3')?.textContent || brandId;

  if (animate) {{
    card.style.opacity = '0';
    card.style.transform = 'scale(0.95)';
    card.style.transition = 'opacity .3s, transform .3s';
    setTimeout(() => {{ card.classList.add('archived'); showArchiveSection(brandId, name); }}, 300);
  }} else {{
    card.classList.add('archived');
    showArchiveSection(brandId, name);
  }}
}}

function showArchiveSection(brandId, name) {{
  const section = document.getElementById('archiveSection');
  section.style.display = '';
  const list = document.getElementById('archiveList');
  if (!document.getElementById('arc_' + brandId)) {{
    const el = document.createElement('div');
    el.className = 'archive-item'; el.id = 'arc_' + brandId;
    el.innerHTML = `<span>📁 ${{name}}</span><button class="restore-btn" onclick="restoreProject('${{brandId}}')">還原</button>`;
    list.appendChild(el);
  }}
}}

// ── 還原封存 ──
function restoreProject(brandId) {{
  const card = document.getElementById('card_' + brandId);
  if (card) {{
    card.classList.remove('archived');
    card.style.opacity = ''; card.style.transform = '';
  }}
  const arcEl = document.getElementById('arc_' + brandId);
  if (arcEl) arcEl.remove();
  const archived = loadArchive().filter(b => b !== brandId);
  saveArchive(archived);
  if (document.getElementById('archiveList').children.length === 0) {{
    document.getElementById('archiveSection').style.display = 'none';
  }}
}}

// ── 圓餅圖 ──
const brands = {chart_js_data};
const ctx = document.getElementById('brandChart').getContext('2d');
const brandChart = new Chart(ctx, {{
  type: 'doughnut',
  data: {{
    labels: brands.map(b => b.label),
    datasets: [{{ data: brands.map(b => b.value), backgroundColor: brands.map(b => b.color+'cc'), borderColor: brands.map(b => b.color), borderWidth: 2, hoverOffset: 6 }}]
  }},
  options: {{ cutout: '70%', plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: c => ` ${{c.label}}: ${{c.parsed}}%` }} }} }}, animation: {{ duration: 800 }} }}
}});
const leg = document.getElementById('chartLegend');
brands.forEach(b => {{
  const el = document.createElement('div'); el.className = 'legend-item';
  el.innerHTML = `<span class="legend-dot" style="background:${{b.color}}"></span><span class="legend-label">${{b.label}}</span><span class="legend-val" style="color:${{b.color}}">${{b.value}}%</span>`;
  leg.appendChild(el);
}});

// 啟動
initTasks();
</script>
</body>
</html>'''


def deploy_to_netlify(html_content, site_id, token):
    with zipfile.ZipFile('/tmp/site.zip', 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("index.html", html_content.encode('utf-8'))
    with open('/tmp/site.zip', 'rb') as f:
        zip_data = f.read()
    url = f"https://api.netlify.com/api/v1/sites/{site_id}/deploys"
    req = urllib.request.Request(url, data=zip_data, method="POST",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/zip"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())


def main():
    service_account_key = os.environ.get("GOOGLE_SERVICE_ACCOUNT_KEY")
    calendar_id = os.environ.get("CALENDAR_ID", "primary")
    netlify_site_id = os.environ.get("NETLIFY_SITE_ID", "venerable-kitten-2a42be")
    netlify_token = os.environ.get("NETLIFY_TOKEN")

    tw_tz = timezone(timedelta(hours=8))
    now_tw = datetime.now(tw_tz)
    today_start = now_tw.replace(hour=0, minute=0, second=0, microsecond=0)
    today_end   = now_tw.replace(hour=23, minute=59, second=59, microsecond=0)
    week_start  = today_start - timedelta(days=now_tw.weekday())
    week_end    = week_start + timedelta(days=6, hours=23, minutes=59, seconds=59)

    print(f"[{now_tw.strftime('%Y-%m-%d %H:%M')} 台灣時間] Brand Hub 開始更新...")

    today_events, week_events = [], []
    if service_account_key:
        try:
            key_dict = json.loads(service_account_key)
            print("✓ 已載入 Service Account")
            access_token = get_access_token_from_service_account(key_dict)
            today_events = fetch_calendar_events(access_token, calendar_id, today_start.isoformat(), today_end.isoformat())
            week_events  = fetch_calendar_events(access_token, calendar_id, week_start.isoformat(), week_end.isoformat())
            print(f"✓ 今日：{len(today_events)} 筆，本週：{len(week_events)} 筆")
        except Exception as e:
            print(f"⚠ Calendar 讀取失敗：{e}")
    else:
        print("⚠ 未設定 GOOGLE_SERVICE_ACCOUNT_KEY")

    print("✓ 生成 HTML 儀表板…")
    html = generate_html(today_events, week_events, now_tw)

    if netlify_token:
        try:
            result = deploy_to_netlify(html, netlify_site_id, netlify_token)
            print(f"✅ 部署成功！{result.get('deploy_ssl_url', result.get('url',''))}")
        except Exception as e:
            print(f"❌ 部署失敗：{e}")
            with open("index.html", "w", encoding="utf-8") as f:
                f.write(html)
    else:
        with open("index.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("✓ 已儲存為 index.html")


if __name__ == "__main__":
    main()
