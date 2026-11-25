"""投产规划填报系统 · Streamlit 原生版（Ant Design 风格）。

复用同一套 Python 后端：注册表元数据、公式引擎、发号器、权限与导入导出服务。
界面按 React 版复刻：左侧折叠菜单 + 顶部用户栏 + 卡片查询区 + 工具栏 + 数据表格。
导航统一走 history.pushState + popstate，避免整页重载白屏。

部署命令：streamlit run streamlit_app.py --server.port 8501
"""
import os
import sys
import shutil
import tempfile
from html import escape
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

RUNTIME_DIR = Path(tempfile.gettempdir()) / "planning_system_runtime"
RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("PLANNING_DB_PATH", str(RUNTIME_DIR / "data.db"))
_db_target = Path(os.environ["PLANNING_DB_PATH"])
if not _db_target.exists():
    _seed_db = ROOT / "backend" / "app" / "data.db"
    if _seed_db.exists():
        shutil.copy(_seed_db, _db_target)

import streamlit as st
import pandas as pd
from openpyxl import load_workbook

from backend.app import auth, registry, services, seed
from backend.app import database as db

st.set_page_config(page_title="投产规划填报系统", page_icon=None, layout="wide",
                   initial_sidebar_state="collapsed")

# 内联 SVG 图标（线性风格，对齐 Ant Design）
SVG = {
    "md": '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/></svg>',
    "planning": '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>',
    "calc": '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>',
    "perm": '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"/><path d="m9 12 2 2 4-4"/></svg>',
    "dash": '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
}

CSS = """
<style>
:root{--primary:#1677ff;--line:#f0f0f0;--text:#262626;--muted:#8c8c8c;}
.stApp{background:#f5f5f5!important;}
[data-testid="stHeader"],[data-testid="stSidebar"]{display:none!important;}
.block-container{padding:68px 24px 40px 224px!important;max-width:none!important;}
.side{position:fixed;left:0;top:0;width:200px;height:100vh;background:#fff;border-right:1px solid var(--line);z-index:1000;overflow-y:auto;}
.brand{height:56px;display:flex;align-items:center;gap:8px;padding:0 14px;border-bottom:1px solid var(--line);font-size:15px;font-weight:600;color:var(--text);}
.brand .logo{width:20px;height:20px;border-radius:4px;background:linear-gradient(135deg,#1677ff,#69c0ff);display:inline-block;flex:none;}
.group{margin:6px 8px;}
.group summary{list-style:none;padding:8px 10px;font-size:13px;font-weight:600;color:var(--text);cursor:pointer;border-radius:6px;display:flex;align-items:center;gap:8px;}
.group summary::-webkit-details-marker{display:none;}
.group summary:hover{background:#f5f5f5;}
.group summary.on{background:#f5f5f5;}
.group summary .ico{width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;color:#595959;flex:none;}
.group summary .arrow{margin-left:auto;color:#bfbfbf;font-size:10px;transition:transform .2s;}
.group[open] summary .arrow{transform:rotate(90deg);}
.group .item{display:flex;align-items:center;gap:6px;padding:6px 10px 6px 34px;font-size:12px;color:#595959;text-decoration:none;border-radius:6px;margin:1px 0;}
.group .item:hover{background:#f5f5f5;color:var(--primary);}
.group .item.active{background:#e6f4ff;color:var(--primary);font-weight:600;}
.side-link{display:flex;align-items:center;gap:8px;margin:6px 8px;padding:8px 10px;font-size:13px;font-weight:600;color:var(--text);text-decoration:none;border-radius:6px;}
.side-link:hover{background:#f5f5f5;}
.side-link .ico{width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;color:#595959;flex:none;}
.topbar{position:fixed;left:200px;right:0;top:0;height:56px;background:#fff;border-bottom:1px solid var(--line);z-index:999;display:flex;align-items:center;justify-content:flex-end;gap:10px;padding:0 24px;}
.role-badge{background:#e6f4ff;color:var(--primary);padding:3px 10px;border-radius:4px;font-size:12px;}
.avatar{width:28px;height:28px;border-radius:50%;background:#f0f0f0;display:flex;align-items:center;justify-content:center;color:#8c8c8c;font-size:13px;}
.username{font-size:13px;color:var(--text);}
.nav{font-size:12px;color:var(--muted);text-decoration:none;margin-left:6px;}
.nav:hover{color:var(--primary);}
.card{background:#fff;border-radius:8px;padding:16px;margin-bottom:16px;box-shadow:0 1px 2px rgba(0,0,0,.03);}
.sec-title{font-size:14px;font-weight:600;color:var(--text);}
.hint{font-size:12px;color:var(--muted);margin-left:8px;font-weight:400;}
.count{font-size:12px;color:var(--muted);padding-left:2px;}
div[data-testid="stHorizontalBlock"]{gap:6px!important;}
div[data-testid="stButton"],div[data-testid="stDownloadButton"]{margin-top:0!important;margin-bottom:0!important;}
.stButton>button,.stDownloadButton>button{height:32px!important;padding:0 12px!important;border-radius:6px!important;border:1px solid #d9d9d9!important;background:#fff!important;color:var(--text)!important;font-size:12px!important;box-shadow:none!important;}
.stButton>button[kind="primary"],.stDownloadButton>button[kind="primary"]{background:var(--primary)!important;border-color:var(--primary)!important;color:#fff!important;}
.stTextInput>div>div>input,.stNumberInput>div>div>input,.stSelectbox>div>div>select,.stDateInput>div>div>input{height:32px!important;border-radius:6px!important;border-color:#d9d9d9!important;font-size:12px!important;}
.stTextInput label,.stNumberInput label,.stSelectbox label,.stDateInput label{font-size:12px!important;color:#595959!important;}
.stFileUploader{border:1px dashed #d9d9d9!important;border-radius:6px!important;padding:8px!important;}
table.grid{border-collapse:collapse;width:100%;font-size:12px;white-space:nowrap;}
table.grid th{background:#fafafa;color:#595959;font-weight:600;text-align:left;padding:10px 12px;border-bottom:1px solid var(--line);}
table.grid td{padding:10px 12px;border-bottom:1px solid #f5f5f5;color:var(--text);}
table.grid tr:hover td{background:#f5f9ff;}
table.grid th .qmark{color:var(--primary);margin-left:3px;font-size:11px;}
.tag{display:inline-block;padding:1px 8px;border-radius:4px;font-size:11px;background:#f5f5f5;color:#595959;border:1px solid #e8e8e8;}
.nav-btn{font-size:12px;text-decoration:none;margin-right:10px;}
.nav-btn-edit{color:var(--primary);}
.nav-btn-del{color:#ff4d4f;}
.pager{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);margin-top:12px;}
.pager a{padding:4px 10px;border:1px solid #d9d9d9;border-radius:6px;color:var(--text);text-decoration:none;background:#fff;}
.pager a.cur{background:var(--primary);border-color:var(--primary);color:#fff;}
.pager a.disabled{opacity:.45;pointer-events:none;}
.metrics{display:flex;gap:16px;flex-wrap:wrap;}
.metric{flex:1;min-width:180px;background:#fff;border-radius:8px;padding:18px;box-shadow:0 1px 2px rgba(0,0,0,.03);}
.metric .mlabel{font-size:12px;color:var(--muted);}
.metric .mval{font-size:26px;font-weight:600;color:var(--text);margin-top:6px;}
.tagchip{display:inline-block;padding:3px 10px;border-radius:4px;font-size:12px;background:#e6f4ff;color:var(--primary);margin:0 6px 6px 0;text-decoration:none;}
.tagchip.grey{background:#f5f5f5;color:#8c8c8c;}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)
QP = st.query_params

# 所有站内导航统一走 pushState + popstate，触发 Streamlit 局部 rerun，避免整页白屏
_NAV_JS = ("event.preventDefault();history.pushState({},'',this.getAttribute('href'));"
           "window.dispatchEvent(new PopStateEvent('popstate'));return false;")


def nav_link(label, href, cls=""):
    return ('<a class="' + cls + '" href="' + href + '" onclick="' + _NAV_JS + '">'
            + label + '</a>')


def goto(**kwargs):
    for k, v in kwargs.items():
        if v is None:
            QP.pop(k, None)
        else:
            QP[k] = str(v)
    st.rerun()


def format_value(value, fmt):
    """与前端 format.js 一致：千分位 / 两位小数 / 百分比（统一两位小数）。"""
    if value is None or value == "":
        return ""
    fmt = fmt or ""
    if not fmt:
        return str(value)
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    if "百分比" in fmt:
        return f"{num * 100:.2f}%"
    if "千分位" in fmt:
        return f"{num:,.2f}" if "两位小数" in fmt else f"{num:,.6f}".rstrip("0").rstrip(".")
    if "两位小数" in fmt:
        return f"{num:.2f}"
    return str(value)


def visible_fields(code, role):
    return [f for f in registry.fields_of(code) if auth.field_visible(code, f["key"], role)]


def render_input(f, current, widget_key):
    name, key, typ = f["name"], f["key"], f["type"]
    if typ in ("formula", "auto_id"):
        st.text_input(name, value="系统计算/发号，只读", disabled=True, key=widget_key)
        return None
    if typ == "integer":
        val = int(current) if current not in (None, "") else None
        return st.number_input(name, value=val, step=1, key=widget_key)
    if typ == "number":
        val = float(current) if current not in (None, "") else None
        return st.number_input(name, value=val, format="%.6f", key=widget_key)
    if f["enum"]:
        opts = [""] + [str(e) for e in f["enum"]]
        idx = opts.index(str(current)) if current is not None and str(current) in opts else 0
        return st.selectbox(name, opts, index=idx, key=widget_key)
    return st.text_input(name, value="" if current is None else str(current), key=widget_key,
                         placeholder=(f.get("desc") or "")[:80])


def html_table(headers, rows):
    """Ant Design 风格表格 HTML。rows 为单元格字符串（已转义）二维数组，自动加序号列。"""
    h = "".join("<th>" + x + "</th>" for x in headers)
    body = []
    for i, row in enumerate(rows, 1):
        cells = "".join("<td>" + c + "</td>" for c in row)
        body.append("<tr><td>" + str(i) + "</td>" + cells + "</tr>")
    return '<table class="grid"><thead><tr><th>序号</th>' + h + '</tr></thead><tbody>' + "".join(body) + '</tbody></table>'


# ---------- 登录 ----------
if QP.get("logout") == "1":
    st.session_state.user = None
    QP.pop("logout", None)
    st.rerun()

if st.session_state.get("user") is None:
    st.markdown('<div class="card" style="max-width:760px;margin:48px auto 8px;">'
                '<div class="brand" style="border:none;padding:0 0 12px 0;">'
                '<span class="logo"></span>投产规划填报系统</div>'
                '<div class="sec-title">登录</div></div>', unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        with st.form("login_form"):
            username = st.text_input("用户名")
            password = st.text_input("密码", type="password")
            if st.form_submit_button("登录", type="primary"):
                u = auth.USER_BY_NAME.get(username.strip())
                if u and u["password"] == password:
                    st.session_state.user = u
                    st.rerun()
                st.error("用户名或密码错误")
        st.caption("体验账号：admin / super / budget / bp / bu / bg / cert，密码均为 123456")
    st.stop()

user = st.session_state.user
role = user["role"]
buttons = auth.BUTTONS.get(role, [])
role_name = auth.ROLES[role]["name"]

# 模块导航（按权限过滤）
nav = []
for top in registry.TOP_MODULES:
    children = []
    for child in top["children"]:
        code = child["code"]
        if code in registry.MODULE_BY_CODE and not auth.can_access_module(role, code):
            continue
        children.append((code, child["name"], top.get("code", "")))
    if children:
        nav.append((top["name"], children, top.get("code", "")))

codes = [c for _, cs, _ in nav for c, _, _ in cs]
module_code = QP.get("module") or codes[0]
if module_code not in codes:
    module_code = codes[0]
view = QP.get("view", "data")
edit_id = QP.get("edit")
delete_id = QP.get("delete")
try:
    page = max(1, int(QP.get("page", "1")))
except ValueError:
    page = 1

# 当前模块所属分组
current_group = ""
for _, cs, g in nav:
    if any(c == module_code for c, _, _ in cs):
        current_group = g
        break
is_special = view in ("dashboard", "perm")

# 左侧菜单（分组可折叠 + 独立项，对齐 React 版）
side = ['<div class="side"><div class="brand"><span class="logo"></span>投产规划填报系统</div>']
for top_name, children, gcode in nav:
    opened = (gcode == current_group and not is_special)
    side.append('<details class="group"' + (' open' if opened else '') + '><summary class="'
                + ('on' if opened else '') + '"><span class="ico">' + SVG.get(gcode, SVG["calc"])
                + '</span>' + escape(top_name) + '<span class="arrow">▶</span></summary>')
    for code, name, _ in children:
        active = (code == module_code and not is_special)
        side.append(nav_link('<span class="item' + (' active' if active else '') + '">'
                             + escape(name) + '</span>',
                             '?module=' + code + '&view=data'))
    side.append('</details>')
side.append(nav_link('<span class="side-link"><span class="ico">' + SVG["dash"]
                     + '</span>工作台</span>', '?view=dashboard&module=' + module_code))
side.append(nav_link('<span class="side-link"><span class="ico">' + SVG["perm"]
                     + '</span>权限管理</span>', '?view=perm&module=' + module_code))
side.append('</div>')
st.markdown("".join(side), unsafe_allow_html=True)

initial = user["name"][:1]
perm = " / ".join(buttons)
st.markdown('<div class="topbar"><span class="role-badge">' + escape(role_name) + '</span>'
            '<span class="avatar">' + escape(initial) + '</span>'
            '<span class="username">' + escape(user["name"]) + '</span>'
            + nav_link('退出', '?logout=1&module=' + module_code, 'nav') + '</div>',
            unsafe_allow_html=True)

# ---------- 工作台 ----------
if view == "dashboard":
    md = sum(len(cs) for _, cs, g in nav if g == "md")
    plan = sum(len(cs) for _, cs, g in nav if g == "planning")
    calc = sum(len(cs) for _, cs, g in nav if g == "calc")
    st.markdown('<div class="card"><div class="sec-title">工作台</div>'
                '<span class="hint">欢迎，' + escape(user["name"]) + '（' + escape(role_name)
                + '）｜ 系统依据飞书《各模块字段定义》构建，当前已落地 ' + str(registry.FIELD_COUNT)
                + ' 个字段、' + str(len(registry.MODULES)) + ' 个模块。</span></div>',
                unsafe_allow_html=True)
    st.markdown('<div class="metrics">'
                '<div class="metric"><div class="mlabel">已落地模块</div><div class="mval">'
                + str(len(registry.MODULES)) + '</div></div>'
                '<div class="metric"><div class="mlabel">字段总数</div><div class="mval">'
                + str(registry.FIELD_COUNT) + '</div></div>'
                '<div class="metric"><div class="mlabel">主数据模块</div><div class="mval">'
                + str(md) + '</div></div>'
                '<div class="metric"><div class="mlabel">投产填报模块</div><div class="mval">'
                + str(plan) + '</div></div></div>', unsafe_allow_html=True)
    st.markdown('<div class="card" style="margin-top:16px;"><div class="sec-title">模块导航</div></div>',
                unsafe_allow_html=True)
    for top_name, children, gcode in nav:
        chips = "".join(nav_link(escape(name), '?module=' + code + '&view=data', 'tagchip')
                        for code, name, _ in children)
        st.markdown('<div style="margin:0 16px 12px;font-size:13px;font-weight:600;">'
                    + SVG.get(gcode, SVG["calc"]) + ' ' + escape(top_name) + '</div>'
                    '<div style="margin:0 16px 12px;">' + chips + '</div>',
                    unsafe_allow_html=True)
    st.stop()

# ---------- 权限管理 ----------
if view == "perm":
    st.markdown('<div class="card"><div class="sec-title">权限管理</div>'
                '<span class="hint">对齐详设 §9 · 功能权限 / 按钮权限 / 行权限 / 字段权限</span></div>',
                unsafe_allow_html=True)
    role_rows = [
        ["系统管理员", "全部模块，全部按钮，全部数据"],
        ["超级管理员", "全部模块，配置与审核"],
        ["预算管理员", "预算填报、参数维护、审核"],
        ["财务 BP", "四算填报、数据查看与分析"],
        ["BU 负责人", "本 BU 数据查看与填报"],
        ["BG 负责人", "本 BG 数据查看与填报"],
        ["核算会计", "字段查看、凭证相关校验"],
    ]
    st.markdown('<div class="card"><div class="sec-title">角色定义</div></div>',
                unsafe_allow_html=True)
    st.markdown('<div class="card" style="padding:0;overflow-x:auto;">'
                + html_table(["角色", "职责"], [[escape(a), escape(b)] for a, b in role_rows])
                + '</div>', unsafe_allow_html=True)
    ctx = [["姓名", escape(user["name"])], ["角色", escape(role_name)],
           ["数据范围(BU)", escape(str(user.get("bu") or "-"))],
           ["数据范围(BG)", escape(str(user.get("bg") or "-"))],
           ["可用按钮", escape(perm)]]
    st.markdown('<div class="card"><div class="sec-title">当前登录上下文</div></div>',
                unsafe_allow_html=True)
    st.markdown('<div class="card" style="padding:0;overflow-x:auto;">'
                + html_table(["项目", "值"], ctx) + '</div>', unsafe_allow_html=True)
    st.stop()

module = registry.MODULE_BY_CODE.get(module_code)
if module is None:
    st.markdown('<div class="card">该模块字段定义尚未在飞书《各模块字段定义》中提供，保留菜单入口（预留空位）。</div>',
                unsafe_allow_html=True)
    st.stop()

fields = visible_fields(module_code, role)

# ---------- 删除确认 ----------
if delete_id:
    st.markdown('<div class="card"><div class="sec-title">删除确认</div>'
                '<p style="font-size:13px;color:#595959;">确定删除记录 ID ' + escape(delete_id)
                + ' 吗？此操作不可撤销。</p></div>', unsafe_allow_html=True)
    c1, c2, _ = st.columns([1, 1, 6])
    with c1:
        if st.button("确认删除", type="primary"):
            res = services.delete_record(module_code, [int(delete_id)])
            if res["ok"]:
                st.success("已删除 " + str(res["deleted"]) + " 行")
            else:
                st.error("删除失败")
            goto(delete=None)
    with c2:
        if st.button("取消"):
            goto(delete=None)

# ---------- 查询区 ----------
queries = {}
query_fields = [f for f in fields if f["query"]]
if view == "data" and not delete_id:
    st.markdown('<div class="card"><span class="sec-title">查询条件</span>'
                '<span class="hint">支持模糊查询，多字段同时填写为「且」关系</span>'
                + nav_link('字段定义', '?module=' + module_code + '&view=fields', 'nav')
                + '</div>', unsafe_allow_html=True)
    qcols = st.columns(min(3, len(query_fields)) + 1)
    for i, f in enumerate(query_fields):
        with qcols[i % len(qcols)]:
            qkey = "q_" + module_code + "_" + f["key"]
            val = st.text_input(f["name"], value=st.session_state.get(qkey, ""),
                                label_visibility="collapsed",
                                placeholder="查询" + f["name"],
                                help=(f.get("desc") or f["name"]), key=qkey)
            if val and val.strip():
                queries[f["key"]] = val.strip()
    with qcols[-1]:
        if st.button("重置查询"):
            for f in query_fields:
                st.session_state["q_" + module_code + "_" + f["key"]] = ""
            goto(page=None)

# ---------- 数据列表 ----------
if view == "data" and not delete_id:
    with st.spinner("加载中…"):
        result = services.list_records(module_code, role, user, page=page, page_size=20,
                                       field_queries=queries or None)
    rows = result["rows"]
    total = result["total"]

    t1, t2, t3, t4, t5, t6 = st.columns([0.9, 0.9, 0.9, 0.9, 1.15, 8])
    with t1:
        if st.button("刷新"):
            st.rerun()
    with t2:
        if st.button("新增", type="primary"):
            if "create" in buttons:
                goto(view="new")
            else:
                st.warning("无新增权限")
    with t3:
        if st.button("导入"):
            if "import" in buttons:
                goto(view="import")
            else:
                st.warning("无导入权限")
    with t4:
        buf = services.export_excel(module_code)
        st.download_button("导出", data=buf.getvalue(), file_name=module["name"] + ".xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           disabled="export" not in buttons)
    with t5:
        tpl = services.download_template(module_code)
        st.download_button("下载模板", data=tpl.getvalue(), file_name=module["name"] + "-导入模板.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           disabled="import" not in buttons)
    with t6:
        st.markdown('<span class="count">共 ' + format(total, ",") + ' 条 · ' + str(len(fields))
                    + ' 个字段（全部展示）</span>', unsafe_allow_html=True)

    headers = []
    for f in fields:
        mark = '<span class="qmark">⌕</span>' if f["query"] else ""
        suffix = " ƒ" if f["type"] == "formula" else ""
        headers.append(escape(f["name"]) + suffix + mark)
    headers += ["状态", "操作"]
    table_rows = []
    for r in rows:
        line = []
        for f in fields:
            if f["type"] == "formula":
                status = r.get("_fs_" + f["key"])
                v = r.get(f["key"])
                val = (format_value(v, f["format"]) if status == "OK" and v not in (None, "")
                       else ("不适用" if status and status != "OK" else ""))
            else:
                val = format_value(r.get(f["key"]), f["format"])
            line.append(escape(str(val)))
        line.append('<span class="tag">' + escape(str(r.get("_status") or "草稿")) + '</span>')
        acts = []
        if "edit" in buttons:
            acts.append(nav_link('编辑', '?module=' + module_code + '&edit=' + str(r["_id"]),
                                 'nav-btn nav-btn-edit'))
        if "delete" in buttons:
            acts.append(nav_link('删除', '?module=' + module_code + '&delete=' + str(r["_id"]),
                                 'nav-btn nav-btn-del'))
        line.append("".join(acts) or "-")
        table_rows.append(line)
    st.markdown('<div class="card" style="padding:0;overflow-x:auto;">'
                + html_table(headers, table_rows) + '</div>', unsafe_allow_html=True)

    total_pages = max(1, (total + 19) // 20)
    prev_cls = "disabled" if page <= 1 else ""
    next_cls = "disabled" if page >= total_pages else ""
    pager = ['<div class="pager">']
    pager.append(nav_link('上一页', '?module=' + module_code + '&page=' + str(page - 1), prev_cls))
    for p in range(max(1, page - 2), min(total_pages, page + 2) + 1):
        cls = "cur" if p == page else ""
        pager.append(nav_link(str(p), '?module=' + module_code + '&page=' + str(p), cls))
    pager.append(nav_link('下一页', '?module=' + module_code + '&page=' + str(page + 1), next_cls))
    pager.append('<span>共 ' + str(total_pages) + ' 页 ｜ ƒ = 公式字段</span></div>')
    st.markdown("".join(pager), unsafe_allow_html=True)

# ---------- 新增 / 编辑 ----------
if view == "new" or edit_id:
    is_edit = bool(edit_id)
    if is_edit and "edit" not in buttons:
        st.warning("无编辑权限")
    else:
        raw = None
        if is_edit:
            raw = next((r for r in db.all_rows(module_code) if r["_id"] == int(edit_id)), None)
            if raw is None:
                st.error("记录不存在")
                st.stop()
        st.markdown('<div class="card"><div class="sec-title">' + ("编辑记录" if is_edit else "新增记录")
                    + '</div><span class="hint">公式字段与自动生成字段由系统维护，不可填写</span>'
                    + nav_link('返回列表', '?module=' + module_code + '&view=data', 'nav')
                    + '</div>', unsafe_allow_html=True)
        with st.form("form_" + module_code):
            values = {}
            chunks = [fields[i:i + 12] for i in range(0, len(fields), 12)]
            tabs = st.tabs([f"字段 {i * 12 + 1}-{min((i + 1) * 12, len(fields))}"
                            for i in range(len(chunks))])
            for tab, chunk in zip(tabs, chunks):
                with tab:
                    for f in chunk:
                        current = raw.get(f["key"]) if raw else None
                        values[f["key"]] = render_input(
                            f, current, ("edit" if is_edit else "new") + "_" + module_code + "_" + f["key"])
            ok = st.form_submit_button("保存", type="primary")
        if ok:
            body = {k: v for k, v in values.items() if v not in (None, "")}
            res = (services.update_record(module_code, int(edit_id), body, role) if is_edit
                   else services.create_record(module_code, body, role))
            if res["ok"]:
                st.success("保存成功")
                goto(view="data", edit=None)
            else:
                st.error("校验未通过：\n- " + "\n- ".join(res["errors"]))

# ---------- 导入 ----------
if view == "import":
    if "import" not in buttons:
        st.warning("无导入权限")
    else:
        st.markdown('<div class="card"><div class="sec-title">Excel 导入</div>'
                    '<span class="hint">模板仅包含「可导入」且非只读字段</span>'
                    + nav_link('返回列表', '?module=' + module_code + '&view=data', 'nav')
                    + '</div>', unsafe_allow_html=True)
        tpl = services.download_template(module_code)
        st.download_button("下载导入模板", data=tpl.getvalue(),
                           file_name=module["name"] + "-导入模板.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        uploaded = st.file_uploader("上传 Excel（.xlsx）", type=["xlsx"])
        if uploaded:
            try:
                wb = load_workbook(BytesIO(uploaded.getvalue()), data_only=True)
                sheet_rows = list(wb.active.iter_rows(values_only=True))
                if not sheet_rows:
                    st.warning("文件为空")
                else:
                    headers = [str(h).replace(" *", "").strip() if h is not None else "" for h in sheet_rows[0]]
                    records = [{headers[i]: r[i] for i in range(min(len(headers), len(r)))}
                               for r in sheet_rows[1:] if any(c is not None for c in r)]
                    preview = services.import_preview(module_code, records)
                    c1, c2, c3 = st.columns(3)
                    c1.metric("总行数", preview["total"])
                    c2.metric("校验通过", preview["valid"])
                    c3.metric("校验失败", preview["invalid"])
                    rows = [[str(x["row"]), "通过" if x["valid"] else "失败",
                             escape("；".join(x["errors"]))] for x in preview["rows"]]
                    st.markdown('<div class="card" style="padding:0;overflow-x:auto;">'
                                + html_table(["行号", "校验", "错误"], rows) + '</div>',
                                unsafe_allow_html=True)
                    if st.button("确认导入通过的行", type="primary"):
                        ok_rows = [x["values"] for x in preview["rows"] if x["valid"]]
                        res = services.import_commit(module_code, ok_rows)
                        if res["ok"]:
                            st.success("已导入 " + str(res["imported"]) + " 行")
                        else:
                            st.error("导入失败")
                        goto(view="data")
            except Exception as exc:
                st.error("解析失败：" + str(exc))

# ---------- 字段定义 ----------
if view == "fields":
    rows = [[escape(f["name"]), escape(f["type"]), "是" if f["required"] else "",
             "是" if f["query"] else "", "是" if f["importable"] else "",
             escape(f["format"] or ""), escape(f["validate"] or ""), escape(f.get("desc") or "")]
            for f in fields]
    st.markdown('<div class="card"><div class="sec-title">字段定义</div>'
                '<span class="hint">来源：飞书《各模块字段定义》</span>'
                + nav_link('返回列表', '?module=' + module_code + '&view=data', 'nav')
                + '</div>', unsafe_allow_html=True)
    st.markdown('<div class="card" style="padding:0;overflow-x:auto;">'
                + html_table(["字段名称", "类型", "必填", "查询字段", "可导入", "格式要求", "校验", "说明"], rows)
                + '</div>', unsafe_allow_html=True)
