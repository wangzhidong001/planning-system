"""投产规划填报系统 · Streamlit 原生版。

与 React 版共用同一套 Python 后端：注册表元数据、公式引擎、发号器、权限与导入导出服务。
云端部署时把 SQLite 数据文件复制到临时可写目录，避免依赖仓库只读文件系统。

部署命令：streamlit run streamlit_app.py --server.port 8501
"""
import os
import sys
import shutil
import tempfile
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

st.set_page_config(page_title="投产规划填报系统", page_icon="📋", layout="wide",
                   initial_sidebar_state="expanded")


# ---------- 工具 ----------
def format_value(value, fmt):
    """与前端 format.js 保持一致：千分位 / 两位小数 / 百分比（统一两位小数）。"""
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
        if "两位小数" in fmt:
            return f"{num:,.2f}"
        return f"{num:,.6f}".rstrip("0").rstrip(".")
    if "两位小数" in fmt:
        return f"{num:.2f}"
    return str(value)


def reset_page(code):
    st.session_state[f"page_{code}"] = 1


def visible_fields(code, role):
    return [f for f in registry.fields_of(code) if auth.field_visible(code, f["key"], role)]


def render_input(f, current, widget_key, disabled=False):
    """按字段类型渲染输入控件；公式/自动字段强制只读。"""
    name, key, typ = f["name"], f["key"], f["type"]
    if typ in ("formula", "auto_id"):
        tip = "系统实时计算，只读" if typ == "formula" else "系统发号，只读"
        st.text_input(name, value=tip, disabled=True, key=widget_key)
        return None
    if typ == "integer":
        val = int(current) if current not in (None, "") else None
        return st.number_input(name, value=val, step=1, key=widget_key, disabled=disabled)
    if typ == "number":
        val = float(current) if current not in (None, "") else None
        return st.number_input(name, value=val, format="%.6f", key=widget_key, disabled=disabled)
    if f["enum"]:
        opts = [""] + [str(e) for e in f["enum"]]
        idx = opts.index(str(current)) if current is not None and str(current) in opts else 0
        return st.selectbox(name, opts, index=idx, key=widget_key, disabled=disabled)
    return st.text_input(name, value="" if current is None else str(current),
                         key=widget_key, disabled=disabled,
                         placeholder=(f.get("desc") or "")[:80])


# ---------- 登录 ----------
if st.session_state.get("user") is None:
    st.title("投产规划填报系统")
    st.caption("元数据驱动 · 24 模块 / 518 字段 · 公式只读 · 稳定发号")
    col1, col2 = st.columns([1, 1])
    with col1:
        with st.form("login_form"):
            username = st.text_input("用户名")
            password = st.text_input("密码", type="password")
            if st.form_submit_button("登录", type="primary"):
                u = auth.USER_BY_NAME.get(username.strip())
                if u and u["password"] == password:
                    st.session_state.user = u
                    st.rerun()
                else:
                    st.error("用户名或密码错误")
    with col2:
        st.markdown("#### 体验账号（密码均为 `123456`）")
        st.dataframe(pd.DataFrame([{
            "用户名": u["username"], "角色": auth.ROLES[u["role"]]["name"],
            "姓名": u["name"], "数据范围": u["bu"],
        } for u in auth.USERS]), hide_index=True, width="stretch")
    st.stop()

user = st.session_state.user
role = user["role"]
buttons = auth.BUTTONS.get(role, [])


# ---------- 侧边栏 ----------
st.sidebar.markdown(f"### {user['name']} · {auth.ROLES[role]['name']}")
st.sidebar.caption(f"BU：{user['bu']} ｜ BG：{user['bg']}")

options = []
for top in registry.TOP_MODULES:
    for child in top["children"]:
        code = child["code"]
        if code in registry.MODULE_BY_CODE:
            if not auth.can_access_module(role, code):
                continue
            label = f"{top['name']} · {child['name']}（{len(registry.MODULE_BY_CODE[code]['fields'])} 字段）"
        else:
            label = f"{top['name']} · {child['name']}（预留）"
        options.append((label, code))
labels = [o[0] for o in options]
code_map = dict(options)
if st.session_state.get("module_label") not in labels:
    st.session_state.module_label = labels[0]
selected_label = st.sidebar.selectbox("功能模块", labels,
                                      index=labels.index(st.session_state.module_label))
code = code_map[selected_label]
st.session_state.module_label = selected_label

if st.sidebar.button("退出登录"):
    st.session_state.user = None
    st.rerun()
if role in ("admin", "super_admin") and st.sidebar.checkbox("重置模拟数据（危险操作）"):
    if st.sidebar.button("立即重置为内置 440 行"):
        seed.reset_and_seed()
        st.sidebar.success("已重置")
        st.rerun()

module = registry.MODULE_BY_CODE.get(code)
if module is None:
    st.title(selected_label)
    st.info("该模块字段定义尚未在飞书《各模块字段定义》中提供，保留菜单入口（预留空位）。")
    st.stop()

fields = visible_fields(code, role)
st.title(module["name"])
st.caption(f"所属：{module['top']} ｜ 字段：{len(fields)} ｜ 权限：{' / '.join(buttons)}")

tab_names = ["数据浏览"]
if "create" in buttons:
    tab_names.append("新增记录")
if "import" in buttons or "export" in buttons:
    tab_names.append("导入 / 导出")
tab_names.append("字段定义")
tabs = st.tabs(tab_names)


# ---------- 数据浏览 ----------
with tabs[0]:
    query_fields = [f for f in fields if f["query"]]
    queries = {}
    if query_fields:
        st.markdown("#### 查询条件（多个字段 AND 模糊匹配）")
        cols = st.columns(3)
        for i, f in enumerate(query_fields):
            with cols[i % 3]:
                qkey = f"q_{code}_{f['key']}"
                val = st.text_input(f["name"], value=st.session_state.get(qkey, ""),
                                    key=qkey, on_change=reset_page, args=(code,))
                if val and val.strip():
                    queries[f["key"]] = val.strip()
        if st.button("清空查询"):
            for f in query_fields:
                st.session_state[f"q_{code}_{f['key']}"] = ""
            st.session_state[f"page_{code}"] = 1
            st.rerun()

    page_key = f"page_{code}"
    page = st.session_state.get(page_key, 1)
    page_size = st.selectbox("每页行数", [20, 50, 100], index=0,
                             key=f"size_{code}", on_change=reset_page, args=(code,))
    with st.spinner("计算中…"):
        result = services.list_records(code, role, user, page=page, page_size=page_size,
                                       field_queries=queries or None)
    rows = result["rows"]

    table_data = []
    for r in rows:
        item = {"ID": r["_id"]}
        for f in fields:
            col = f["name"] + (" ƒ" if f["type"] == "formula" else "")
            if f["type"] == "formula":
                status = r.get("_fs_" + f["key"])
                v = r.get(f["key"])
                item[col] = (format_value(v, f["format"])
                             if status == "OK" and v not in (None, "")
                             else ("不适用" if status and status != "OK" else ""))
            else:
                item[col] = format_value(r.get(f["key"]), f["format"])
        table_data.append(item)
    st.dataframe(pd.DataFrame(table_data), hide_index=True, width="stretch",
                 height=min(620, 48 + len(table_data) * 34))

    c1, c2, c3 = st.columns([1, 1, 4])
    with c1:
        if st.button("上一页", disabled=page <= 1):
            st.session_state[page_key] = page - 1
            st.rerun()
    with c2:
        if st.button("下一页", disabled=page * page_size >= result["total"]):
            st.session_state[page_key] = page + 1
            st.rerun()
    with c3:
        st.caption(f"共 {result['total']:,} 行 ｜ 第 {page} / {max(1, (result['total'] + page_size - 1) // page_size)} 页 ｜ ƒ = 公式字段")

    if "delete" in buttons and rows:
        st.markdown("#### 删除记录")
        def _label(r):
            first = next((format_value(r.get(f["key"]), f["format"])
                          for f in fields if f["type"] != "formula" and r.get(f["key"])), "")
            return f"ID {r['_id']} · {first}"
        del_map = {_label(r): r["_id"] for r in rows}
        chosen = st.selectbox("选择记录", list(del_map.keys()), key=f"del_{code}")
        if st.checkbox("我确认删除所选记录", key=f"delc_{code}"):
            if st.button("删除", type="primary"):
                res = services.delete_record(code, [del_map[chosen]])
                st.success(f"已删除 {res['deleted']} 行") if res["ok"] else st.error("删除失败")
                st.rerun()


# ---------- 新增记录 ----------
idx = 1
if "create" in buttons:
    with tabs[idx]:
        st.markdown("#### 新增记录（公式字段与自动生成字段不可填写）")
        with st.form(f"create_form_{code}"):
            values = {}
            chunks = [fields[i:i + 12] for i in range(0, len(fields), 12)]
            tab_labels = [f"字段 {i * 12 + 1}-{min((i + 1) * 12, len(fields))}"
                          for i in range(len(chunks))]
            sub_tabs = st.tabs(tab_labels)
            for sub, chunk in zip(sub_tabs, chunks):
                with sub:
                    for f in chunk:
                        values[f["key"]] = render_input(f, None, f"new_{code}_{f['key']}")
            submitted = st.form_submit_button("提交新增", type="primary")
        if submitted:
            body = {k: v for k, v in values.items() if v not in (None, "")}
            res = services.create_record(code, body, role)
            if res["ok"]:
                st.success(f"新增成功，记录 ID：{res['id']}")
            else:
                st.error("校验未通过：\n- " + "\n- ".join(res["errors"]))
    idx += 1


# ---------- 导入 / 导出 ----------
if "import" in buttons or "export" in buttons:
    with tabs[idx]:
        if "export" in buttons:
            st.markdown("#### 导出")
            buf = services.export_excel(code)
            st.download_button("导出当前模块 Excel", data=buf.getvalue(),
                               file_name=f"{module['name']}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if "import" in buttons:
            st.markdown("#### 导入")
            tpl = services.download_template(code)
            st.download_button("下载导入模板（仅可导入字段）", data=tpl.getvalue(),
                               file_name=f"{module['name']}-导入模板.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            uploaded = st.file_uploader("上传 Excel（.xlsx）", type=["xlsx"], key=f"up_{code}")
            if uploaded is not None:
                try:
                    wb = load_workbook(BytesIO(uploaded.getvalue()), data_only=True)
                    sheet_rows = list(wb.active.iter_rows(values_only=True))
                    if not sheet_rows:
                        st.warning("文件为空")
                    else:
                        headers = [str(h).replace(" *", "").strip() if h is not None else ""
                                   for h in sheet_rows[0]]
                        records = [{headers[i]: r[i] for i in range(min(len(headers), len(r)))}
                                   for r in sheet_rows[1:] if any(c is not None for c in r)]
                        preview = services.import_preview(code, records)
                        c1, c2, c3 = st.columns(3)
                        c1.metric("总行数", preview["total"])
                        c2.metric("校验通过", preview["valid"])
                        c3.metric("校验失败", preview["invalid"])
                        st.dataframe(pd.DataFrame([{
                            "行号": x["row"],
                            "校验": "通过" if x["valid"] else "失败",
                            "错误": "；".join(x["errors"]),
                        } for x in preview["rows"]]), hide_index=True, width="stretch")
                        if st.button("确认导入通过的行", type="primary"):
                            ok_rows = [x["values"] for x in preview["rows"] if x["valid"]]
                            res = services.import_commit(code, ok_rows)
                            st.success(f"已导入 {res['imported']} 行") if res["ok"] else st.error("导入失败")
                            st.rerun()
                except Exception as exc:
                    st.error(f"解析失败：{exc}")
    idx += 1


# ---------- 字段定义 ----------
with tabs[idx]:
    st.markdown("#### 字段定义（来源：飞书《各模块字段定义》）")
    st.dataframe(pd.DataFrame([{
        "字段名称": f["name"], "类型": f["type"], "必填": "是" if f["required"] else "",
        "查询字段": "是" if f["query"] else "", "可导入": "是" if f["importable"] else "",
        "格式要求": f["format"] or "", "校验": f["validate"] or "",
        "说明": f.get("desc") or "",
    } for f in fields]), hide_index=True, width="stretch", height=620)
