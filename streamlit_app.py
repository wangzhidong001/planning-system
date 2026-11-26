"""投产规划填报系统 · Streamlit 原生版（Ant Design 风格）。

复用同一套 Python 后端：注册表元数据、公式引擎、发号器、权限与导入导出服务。
导航全部使用原生 st.button + st.query_params（Python 端局部 rerun），
不使用 HTML <a href> 或 history.pushState，避免整页重载导致 session 丢失回登录页。

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
                   initial_sidebar_state="expanded")

CSS = """
<style>
:root{--primary:#1677ff;--line:#f0f0f0;--text:#262626;--muted:#8c8c8c;}
.stApp{background:#f5f5f5!important;}
[data-testid="stHeader"]{display:none!important;}
/* 原生侧栏：定制为 200px 白色菜单 */
section[data-testid="stSidebar"]{width:200px!important;min-width:200px!important;background:#fff!important;border-right:1px solid var(--line)!important;}
section[data-testid="stSidebar"] > div:first-child{padding:0!important;}
[data-testid="stSidebarCollapseButton"],[data-testid="stSidebarCollapsedControl"]{display:none!important;}
.sb-brand{height:56px;display:flex;align-items:center;gap:8px;padding:0 14px;border-bottom:1px solid var(--line);font-size:15px;font-weight:600;color:var(--text);}
.sb-brand .logo{width:20px;height:20px;border-radius:4px;background:linear-gradient(135deg,#1677ff,#69c0ff);display:inline-block;flex:none;}
.sb-group{padding:8px 10px;font-size:13px;font-weight:600;color:var(--text);margin:6px 8px 0;}
.sb-group .ico{margin-right:6px;}
.block-container{padding:68px 24px 40px 24px!important;max-width:none!important;}
/* 侧栏按钮统一样式 */
section[data-testid="stSidebar"] .stButton{margin:0!important;}
section[data-testid="stSidebar"] .stButton > button{
  width:100%!important;height:auto!important;min-height:32px!important;border:none!important;
  background:transparent!important;color:#595959!important;box-shadow:none!important;
  font-size:12px!important;text-align:left!important;padding:7px 10px 7px 30px!important;
  border-radius:6px!important;font-weight:400!important;
}
section[data-testid="stSidebar"] .stButton > button:hover{background:#f5f5f5!important;color:var(--primary)!important;}
section[data-testid="stSidebar"] .stButton > button[kind="primary"]{background:#e6f4ff!important;color:var(--primary)!important;font-weight:600!important;}
.topbar{position:fixed;left:200px;right:0;top:0;height:56px;background:#fff;border-bottom:1px solid var(--line);z-index:999;display:flex;align-items:center;justify-content:flex-end;gap:10px;padding:0 24px;}
.role-badge{background:#e6f4ff;color:var(--primary);padding:3px 10px;border-radius:4px;font-size:12px;}
.avatar{width:28px;height:28px;border-radius:50%;background:#f0f0f0;display:flex;align-items:center;justify-content:center;color:#8c8c8c;font-size:13px;}
.username{font-size:13px;color:var(--text);}
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
table.grid th[title]{cursor:help;}
table.grid th[title]:hover{background:#e6f4ff;color:var(--primary);}
/* 行内操作表格：表头与数据行用 st.columns 固定列宽，整体横向滚动（表头与数据同步滚动） */
.block-container{overflow-x:auto!important;}
div[data-testid="stHorizontalBlock"]:has(.rt-th),div[data-testid="stHorizontalBlock"]:has(.rt-td){
  gap:0!important;margin-bottom:0!important;flex-wrap:nowrap!important;
  width:max-content!important;min-width:100%!important;
}
div[data-testid="stHorizontalBlock"]:has(.rt-th)>div[data-testid="stColumn"],
div[data-testid="stHorizontalBlock"]:has(.rt-td)>div[data-testid="stColumn"]{
  flex:0 0 108px!important;min-width:108px!important;max-width:108px!important;
}
.rt-th,.rt-td{background:#fff;font-size:12px;padding:9px 8px;border-bottom:1px solid #f0f0f0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.rt-th{background:#fafafa;color:#595959;font-weight:600;cursor:help;}
.rt-th[title]:hover{background:#e6f4ff;color:var(--primary);}
.rt-td{color:var(--text);}
.rt-qmark{color:var(--primary);margin-left:2px;}
.stButton>button[kind="tertiary"]{height:26px!important;min-height:26px!important;padding:0 5px!important;border:none!important;background:transparent!important;color:var(--primary)!important;box-shadow:none!important;font-size:12px!important;}
.stButton>button[kind="tertiary"]:hover{background:#e6f4ff!important;}
.tag{display:inline-block;padding:1px 8px;border-radius:4px;font-size:11px;background:#f5f5f5;color:#595959;border:1px solid #e8e8e8;}
.metrics{display:flex;gap:16px;flex-wrap:wrap;}
.metric{flex:1;min-width:180px;background:#fff;border-radius:8px;padding:18px;box-shadow:0 1px 2px rgba(0,0,0,.03);}
.metric .mlabel{font-size:12px;color:var(--muted);}
.metric .mval{font-size:26px;font-weight:600;color:var(--text);margin-top:6px;}
.tagchip{display:inline-block;padding:3px 10px;border-radius:4px;font-size:12px;background:#e6f4ff;color:var(--primary);margin:0 6px 6px 0;}
.tagchip.grey{background:#f5f5f5;color:#8c8c8c;}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)
QP = st.query_params

# 图标（Streamlit button icon 用 emoji / material 符号）
ICON = {
    "md": "🗄", "planning": "✎", "calc": "▦", "perm": "🛡", "dash": "◷",
}


def goto(**kwargs):
    """用官方 query_params API 跳转：Python 端触发局部 rerun，不操作浏览器 URL。"""
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


def render_input(f, current, widget_key, code=None, row=None, user_name="", is_edit=False):
    """渲染单个字段输入控件。
    - 只读字段（formula/auto_id）：显示最终结果值（公式引擎计算）并置灰，不写入提交体。
    - 数字字段：用 text_input 展示「格式要求」格式化后的值（如 6,208.10），
      支持键盘/粘贴，去掉 number_input 的 +/- 步进按钮；提交时解析回数值。
    """
    name, typ = f["name"], f["type"]
    fmt = f.get("format") or ""
    if typ in ("formula", "auto_id"):
        st.session_state.pop(widget_key, None)  # 强制使用本次计算出的展示值
        if typ == "auto_id":
            shown = "" if current in (None, "") else str(current)
            help_txt = "系统发号，自动生成，不可编辑"
        elif is_edit and code and row is not None:
            from backend.app import formula_engine
            v, status = formula_engine.compute_field(code, f, row, user_name)
            shown = (format_value(v, fmt) if status == "OK" and v not in (None, "")
                     else ("不适用" if status and status != "OK" else ""))
            help_txt = "公式字段，系统自动计算，不可编辑"
        else:
            shown, help_txt = "", "公式字段，保存后系统自动计算"
        st.text_input(name, value=shown, disabled=True, key=widget_key, help=help_txt)
        return None
    if typ in ("integer", "number"):
        shown = format_value(current, fmt) if current not in (None, "") else ""
        txt = st.text_input(name, value=shown, key=widget_key,
                            placeholder=(f.get("desc") or "")[:80])
        if txt in (None, ""):
            return None
        cleaned = str(txt).strip().replace(",", "").replace("，", "")
        pct = "百分比" in fmt
        if pct:
            cleaned = cleaned.rstrip("%")
        try:
            num = float(cleaned)
        except (TypeError, ValueError):
            return txt  # 非数值交由后端校验层提示
        if pct:
            num = num / 100
        return int(num) if typ == "integer" else num
    if f["enum"]:
        opts = [""] + [str(e) for e in f["enum"]]
        idx = opts.index(str(current)) if current is not None and str(current) in opts else 0
        return st.selectbox(name, opts, index=idx, key=widget_key)
    return st.text_input(name, value="" if current is None else str(current), key=widget_key,
                         placeholder=(f.get("desc") or "")[:80])


def html_table(headers, rows, header_titles=None):
    """Ant Design 风格表格 HTML。rows 为单元格字符串（已转义）二维数组，自动加序号列。
    header_titles 与 headers 一一对应，渲染为 <th title="...">，鼠标悬停显示字段说明。"""
    titles = header_titles or []
    ths = []
    for i, x in enumerate(headers):
        tip = str(titles[i]) if i < len(titles) and titles[i] else ""
        attr = ' title="' + escape(tip) + '"' if tip else ""
        ths.append("<th" + attr + ">" + x + "</th>")
    h = "".join(ths)
    body = []
    for i, row in enumerate(rows, 1):
        cells = "".join("<td>" + c + "</td>" for c in row)
        body.append("<tr><td>" + str(i) + "</td>" + cells + "</tr>")
    return '<table class="grid"><thead><tr><th>序号</th>' + h + '</tr></thead><tbody>' + "".join(body) + '</tbody></table>'


# ---------- 登录 ----------
if st.session_state.get("user") is None:
    # 登录页隐藏侧栏
    st.markdown('<style>section[data-testid="stSidebar"]{display:none!important;}</style>',
                unsafe_allow_html=True)
    st.markdown('<div class="card" style="max-width:760px;margin:48px auto 8px;">'
                '<div class="sb-brand" style="border:none;padding:0 0 12px 0;">'
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
        children.append((code, child["name"]))
    if children:
        nav.append((top["name"], children, top.get("code", "")))

codes = [c for _, cs, _ in nav for c, _ in cs]
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

current_group = ""
for _, cs, g in nav:
    if any(c == module_code for c, _ in cs):
        current_group = g
        break
is_special = view in ("dashboard", "perm")

# ---------- 左侧菜单（原生 sidebar + 原生按钮，保证局部 rerun）----------
with st.sidebar:
    st.markdown('<div class="sb-brand"><span class="logo"></span>投产规划填报系统</div>',
                unsafe_allow_html=True)
    for top_name, children, gcode in nav:
        st.markdown('<div class="sb-group">' + ICON.get(gcode, "▦") + ' '
                    + escape(top_name) + '</div>', unsafe_allow_html=True)
        for code, name in children:
            active = (code == module_code and not is_special)
            if st.button(name, key="nav_" + code, type="primary" if active else "secondary",
                         use_container_width=True):
                goto(module=code, view="data", edit=None, delete=None, page=None)
    st.markdown('<div class="sb-group">快捷入口</div>', unsafe_allow_html=True)
    if st.button(ICON["dash"] + " 工作台", key="nav_dash",
                 type="primary" if view == "dashboard" else "secondary", use_container_width=True):
        goto(view="dashboard", edit=None, delete=None)
    if st.button(ICON["perm"] + " 权限管理", key="nav_perm",
                 type="primary" if view == "perm" else "secondary", use_container_width=True):
        goto(view="perm", edit=None, delete=None)
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
    if st.button("退出登录", key="nav_logout", use_container_width=True):
        st.session_state.user = None
        st.rerun()

initial = user["name"][:1]
perm = " / ".join(buttons)
st.markdown('<div class="topbar"><span class="role-badge">' + escape(role_name) + '</span>'
            '<span class="avatar">' + escape(initial) + '</span>'
            '<span class="username">' + escape(user["name"]) + '</span></div>',
            unsafe_allow_html=True)

# ---------- 工作台 ----------
if view == "dashboard":
    md = sum(len(cs) for _, cs, g in nav if g == "md")
    plan = sum(len(cs) for _, cs, g in nav if g == "planning")
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
        st.markdown('<div style="margin:0 16px 6px;font-size:13px;font-weight:600;">'
                    + ICON.get(gcode, "▦") + ' ' + escape(top_name) + '</div>',
                    unsafe_allow_html=True)
        ncols = st.columns(4)
        for i, (code, name) in enumerate(children):
            with ncols[i % 4]:
                if st.button(name, key="db_" + code, use_container_width=True):
                    goto(module=code, view="data", edit=None, delete=None, page=None)
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
    st.markdown('<div class="card"><div class="sec-title">角色定义</div></div>', unsafe_allow_html=True)
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

# ---------- 查询区 ----------
queries = {}
query_fields = [f for f in fields if f["query"]]
if view == "data":
    qcols = st.columns(min(3, len(query_fields)) + 2)
    for i, f in enumerate(query_fields):
        with qcols[i % len(qcols)]:
            qkey = "q_" + module_code + "_" + f["key"]
            val = st.text_input(f["name"], value=st.session_state.get(qkey, ""),
                                label_visibility="collapsed",
                                placeholder="查询" + f["name"],
                                help=(f.get("desc") or f["name"]), key=qkey)
            if val and val.strip():
                queries[f["key"]] = val.strip()
    with qcols[-2]:
        if st.button("重置查询"):
            for f in query_fields:
                st.session_state["q_" + module_code + "_" + f["key"]] = ""
            goto(page=None)
    with qcols[-1]:
        if st.button("字段定义", key="btn_fields"):
            goto(view="fields", edit=None, delete=None)

# ---------- 数据列表 ----------
if view == "data":
    with st.spinner("加载中…"):
        result = services.list_records(module_code, role, user, page=page, page_size=20,
                                       field_queries=queries or None)
    rows = result["rows"]
    total = result["total"]

    b1, b2, b3, b4, b5, b6 = st.columns([0.9, 0.9, 0.9, 0.9, 1.15, 8])
    with b1:
        if st.button("刷新"):
            st.rerun()
    with b2:
        if st.button("新增", type="primary"):
            if "create" in buttons:
                goto(view="new", edit=None, delete=None)
            else:
                st.warning("无新增权限")
    with b3:
        if st.button("导入"):
            if "import" in buttons:
                goto(view="import", edit=None, delete=None)
            else:
                st.warning("无导入权限")
    with b4:
        buf = services.export_excel(module_code)
        st.download_button("导出", data=buf.getvalue(), file_name=module["name"] + ".xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           disabled="export" not in buttons)
    with b5:
        tpl = services.download_template(module_code)
        st.download_button("下载模板", data=tpl.getvalue(), file_name=module["name"] + "-导入模板.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           disabled="import" not in buttons)
    with b6:
        st.markdown('<span class="count">共 ' + format(total, ",") + ' 条 · ' + str(len(fields))
                    + ' 个字段（全部展示）</span>', unsafe_allow_html=True)

    # 行内表格：表头与数据行均用原生 st.columns 保证列对齐；每行末尾为原生 st.button，
    # 点击走 goto() 局部 rerun，绝不使用 HTML 链接（避免整页刷新/跳登录页）。
    weights = [0.45] + [1.0] * len(fields) + [0.8, 1.05]
    header_titles = [(f.get("desc") or "").strip() or (f["name"] + "（暂无字段说明）") for f in fields]

    hcols = st.columns(weights)
    with hcols[0]:
        st.markdown('<div class="rt-th">序号</div>', unsafe_allow_html=True)
    for j, f in enumerate(fields):
        with hcols[j + 1]:
            mark = '<span class="rt-qmark">⌕</span>' if f["query"] else ""
            suffix = " ƒ" if f["type"] == "formula" else ""
            st.markdown('<div class="rt-th" title="' + escape(header_titles[j]) + '">'
                        + escape(f["name"]) + suffix + mark + '</div>', unsafe_allow_html=True)
    with hcols[-2]:
        st.markdown('<div class="rt-th">状态</div>', unsafe_allow_html=True)
    with hcols[-1]:
        st.markdown('<div class="rt-th">操作</div>', unsafe_allow_html=True)

    for i, r in enumerate(rows, 1):
        rid = str(r["_id"])
        rcols = st.columns(weights)
        with rcols[0]:
            st.markdown('<div class="rt-td">' + str(i) + '</div>', unsafe_allow_html=True)
        for j, f in enumerate(fields):
            with rcols[j + 1]:
                if f["type"] == "formula":
                    status = r.get("_fs_" + f["key"])
                    v = r.get(f["key"])
                    val = (format_value(v, f["format"]) if status == "OK" and v not in (None, "")
                           else ("不适用" if status and status != "OK" else ""))
                else:
                    val = format_value(r.get(f["key"]), f["format"])
                st.markdown('<div class="rt-td" title="' + escape(str(val)) + '">'
                            + escape(str(val)) + '</div>', unsafe_allow_html=True)
        with rcols[-2]:
            st.markdown('<div class="rt-td"><span class="tag">'
                        + escape(str(r.get("_status") or "草稿")) + '</span></div>', unsafe_allow_html=True)
        with rcols[-1]:
            ec, dc = st.columns(2)
            with ec:
                if st.button("✎ 编辑", key="row_ed_" + rid, type="tertiary",
                             help=("编辑第 " + str(i) + " 行") if "edit" in buttons else "无编辑权限",
                             disabled="edit" not in buttons, use_container_width=True):
                    goto(edit=rid, view="data", delete=None)
            with dc:
                if st.button("🗑 删除", key="row_del_" + rid, type="tertiary",
                             help=("删除第 " + str(i) + " 行") if "delete" in buttons else "无删除权限",
                             disabled="delete" not in buttons, use_container_width=True):
                    goto(delete=rid, view="data", edit=None)

    # 分页（原生按钮）
    total_pages = max(1, (total + 19) // 20)
    pg = st.columns([1, 1, 1, 1, 1, 4])
    with pg[0]:
        if st.button("上一页", disabled=page <= 1, key="pg_prev"):
            goto(page=page - 1)
    for idx, p in enumerate(range(max(1, page - 1), min(total_pages, page + 1) + 1)):
        with pg[idx + 1]:
            if st.button(("· " if p != page else "") + str(p),
                         type="primary" if p == page else "secondary", key="pg_" + str(p)):
                goto(page=p)
    with pg[4]:
        if st.button("下一页", disabled=page >= total_pages, key="pg_next"):
            goto(page=page + 1)
    with pg[5]:
        st.markdown('<span class="count">共 ' + str(total_pages) + ' 页 ｜ ƒ = 公式字段</span>',
                    unsafe_allow_html=True)

# ---------- 新增（整页，编辑/删除已改为模态对话框） ----------
if view == "new":
    if "create" not in buttons:
        st.warning("无新增权限")
    else:
        st.markdown('<div class="card"><div class="sec-title">新增记录</div>'
                    '<span class="hint">公式字段与自动生成字段由系统维护，不可填写</span></div>',
                    unsafe_allow_html=True)
        if st.button("返回列表", key="back_list"):
            goto(view="data", edit=None, delete=None)
        with st.form("form_" + module_code):
            values = {}
            chunks = [fields[i:i + 12] for i in range(0, len(fields), 12)]
            tabs = st.tabs([f"字段 {i * 12 + 1}-{min((i + 1) * 12, len(fields))}"
                            for i in range(len(chunks))])
            for tab, chunk in zip(tabs, chunks):
                with tab:
                    for f in chunk:
                        current = None
                        values[f["key"]] = render_input(
                            f, current, "new_" + module_code + "_" + f["key"])
            ok = st.form_submit_button("保存", type="primary")
        if ok:
            body = {k: v for k, v in values.items() if v not in (None, "")}
            res = services.create_record(module_code, body, role)
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
                    '<span class="hint">模板仅包含「可导入」且非只读字段</span></div>',
                    unsafe_allow_html=True)
        if st.button("返回列表", key="back_list_imp"):
            goto(view="data", edit=None, delete=None)
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
                '<span class="hint">来源：飞书《各模块字段定义》</span></div>', unsafe_allow_html=True)
    if st.button("返回列表", key="back_list_fields"):
        goto(view="data", edit=None, delete=None)
    st.markdown('<div class="card" style="padding:0;overflow-x:auto;">'
                + html_table(["字段名称", "类型", "必填", "查询字段", "可导入", "格式要求", "校验", "说明"], rows)
                + '</div>', unsafe_allow_html=True)

# ---------- 模态对话框：编辑 / 删除确认 ----------
@st.dialog("编辑记录", width="large")
def _edit_dialog(code, rid, flds, current_role):
    st.caption("只读字段显示系统计算结果，不可编辑；数字字段按格式要求显示，可直接输入或粘贴")
    current_user_name = (st.session_state.get("user") or {}).get("name", "")
    raw = next((r for r in db.all_rows(code) if r["_id"] == int(rid)), None)
    if raw is None:
        st.error("记录不存在或已被删除")
        if st.button("关闭", key="dlg_missing_close"):
            goto(edit=None)
        return
    with st.form("dlg_edit_" + code):
        values = {}
        chunks = [flds[i:i + 12] for i in range(0, len(flds), 12)]
        tabs = st.tabs([f"字段 {i * 12 + 1}-{min((i + 1) * 12, len(flds))}"
                        for i in range(len(chunks))])
        for tab, chunk in zip(tabs, chunks):
            with tab:
                for f in chunk:
                    values[f["key"]] = render_input(
                        f, raw.get(f["key"]), "dlg_edit_" + code + "_" + f["key"],
                        code=code, row=raw, user_name=current_user_name, is_edit=True)
        cs = st.columns([1, 4, 1])
        with cs[0]:
            cancel = st.form_submit_button("取消")
        with cs[2]:
            save = st.form_submit_button("保存", type="primary")
    if cancel:
        goto(edit=None)
    if save:
        body = {k: v for k, v in values.items() if v not in (None, "")}
        res = services.update_record(code, int(rid), body, current_role)
        if res["ok"]:
            st.success("保存成功")
            st.rerun()
        else:
            st.error("校验未通过：\n- " + "\n- ".join(res["errors"]))


@st.dialog("删除确认")
def _delete_dialog(code, rid):
    st.write("确定删除记录 ID " + str(rid) + " 吗？此操作不可撤销。")
    c1, c2, _ = st.columns([1, 1, 3])
    with c1:
        if st.button("确认删除", type="primary", key="dlg_del_confirm"):
            res = services.delete_record(code, [int(rid)])
            if res["ok"]:
                st.success("已删除 " + str(res["deleted"]) + " 行")
                goto(delete=None)
            else:
                st.error("删除失败")
    with c2:
        if st.button("取消", key="dlg_del_cancel"):
            goto(delete=None)


# 数据页且带 edit/delete 参数时弹出对话框（列表作为背景保留）
if view == "data" and edit_id and "edit" in buttons:
    _edit_dialog(module_code, edit_id, fields, role)
elif view == "data" and delete_id and "delete" in buttons:
    _delete_dialog(module_code, delete_id)
