"""通用业务服务：校验引擎、自动取值、CRUD 编排、导入解析、Excel 导出。"""
import io
import re
from datetime import datetime
from .registry import get_module, fields_of, MODULE_BY_CODE
from . import database as db


def to_display(row: dict, code: str, role: str = "admin"):
    """DB 行 -> 前端展示：fN key 转字段名，隐藏无权限字段，过滤系统列。
    公式字段展示引擎计算值（_formula_values），只读标记随字段元数据给出。"""
    from .auth import field_visible
    fields = fields_of(code)
    fv = row.get("_formula_values") or {}
    fs = row.get("_formula_status") or {}
    out = {"_id": row.get("_id"), "_status": row.get("_status", "草稿"),
           "_version": row.get("_version", 1), "_updated_at": row.get("_updated_at")}
    for f in fields:
        if not field_visible(code, f["key"], role):
            continue
        if f["type"] == "formula":
            val = fv.get(f["key"])
        else:
            val = row.get(f["key"])
        out[f["key"]] = val
        out[f["name"]] = val  # 同时提供中文名，便于阅读
        if f["type"] == "formula":
            out["_fs_" + f["key"]] = fs.get(f["key"], "NA")  # 公式状态
    return out


def validate_row(code: str, values: dict, all_rows: list, exclude_id=None):
    """返回错误列表。implement 枚举/重复/存在性/必填 校验。"""
    errors = []
    module = get_module(code)
    if not module:
        return errors
    fields = module["fields"]
    by_name = {f["name"]: f for f in fields}

    for f in fields:
        # 公式字段/自动生成字段为系统维护，跳过用户录入校验
        if f["type"] in ("formula", "auto_id"):
            continue
        v = values.get(f["key"])
        # 必填
        if f["required"] and (v is None or str(v).strip() == ""):
            errors.append(f"【{f['name']}】为必填项")
            continue
        if v in (None, ""):
            continue
        # 枚举值校验
        if f["validate"] == "枚举值校验" and f["enum"]:
            if str(v) not in [str(e) for e in f["enum"]]:
                errors.append(f"【{f['name']}】枚举值非法，允许：{'/'.join(f['enum'])}")
        # 数字/整数
        if f["type"] in ("number", "integer"):
            try:
                num = float(v)
                if f["type"] == "integer" and not float(num).is_integer():
                    errors.append(f"【{f['name']}】须为整数")
            except (ValueError, TypeError):
                errors.append(f"【{f['name']}】须为数字")

    # 重复性校验：规划项目名称
    dup_field = next((f for f in fields if f["validate"] == "重复性校验"), None)
    if dup_field:
        val = values.get(dup_field["key"])
        if val:
            for r in all_rows:
                if exclude_id is not None and r.get("_id") == exclude_id:
                    continue
                if str(r.get(dup_field["key"])) == str(val):
                    errors.append(f"【{dup_field['name']}】重复：{val}")
                    break

    # 存在性校验（通用）：从 validate 文本解析「《目标表》表中 目标字段 字段中存在」
    from .formula_engine import resolve_code, field_key
    for ef in fields:
        if not (ef["validate"] and "存在性校验" in ef["validate"]):
            continue
        val = values.get(ef["key"])
        if not val:
            continue
        vm = re.search(r"是否在《(.+?)》表中\s*(.+?)\s*字段中存在", ef["validate"])
        if not vm:
            continue
        target_code = resolve_code(vm.group(1))
        target_field_name = vm.group(2).strip()
        if not target_code:
            continue
        tfk = field_key(target_code, target_field_name)
        if not tfk:
            continue
        parents = db.all_rows(target_code)
        pool = {str(r.get(tfk)) for r in parents if r.get(tfk) not in (None, "")}
        if str(val) not in pool:
            errors.append(f"【{ef['name']}】不存在于{vm.group(1)}：{val}")
    return errors


def auto_fill(code: str, values: dict):
    """跨表自动取值（对齐 Excel 取值公式）。就地修改 values。"""
    fields = {f["name"]: f for f in fields_of(code)}
    # 1) 规划项目清单.规划项目等级 <- 项目分类标准.项目等级（按 IPD/非IPD + 项目类型 匹配）
    if code == "project_list":
        grade_f = fields.get("规划项目等级")
        ipd = values.get("f12")  # IPD/非IPD项目
        ptype = values.get("f13")  # 规划项目类型
        if grade_f and ipd and ptype:
            for r in db.all_rows("project_class"):
                if str(r.get("f1")) == str(ipd) and str(r.get("f2")) == str(ptype):
                    values["f14"] = r.get("f3")  # 项目等级
                    break
    # 2) 项目生命周期预测.IPD/非IPD项目、规划项目类型 <- 规划项目清单（按 规划项目名称 匹配）
    if code == "lifecycle_forecast":
        pname = values.get("f2")
        if pname:
            for r in db.all_rows("project_list"):
                if str(r.get("f2")) == str(pname):
                    values["f3"] = r.get("f12")  # IPD/非IPD项目
                    values["f4"] = r.get("f13")  # 规划项目类型
                    break


def extract_values(code: str, body: dict):
    """从请求体提取 fN 字段值。排除公式字段与自动生成字段（系统维护，不可手工写入）。"""
    fields = fields_of(code)
    keys = {f["key"] for f in fields if f["type"] not in ("formula", "auto_id")}
    return {k: body[k] for k in keys if k in body}


def list_records(code: str, role="admin", user=None, query="", page=1, page_size=50, field_queries=None):
    from .auth import row_scope
    rows = db.all_rows(code)
    # 行级过滤（演示：按 f7 BU名称 / f8 服务组织 字段匹配）
    scope = row_scope(user) if user else None
    if scope and user:
        fields = fields_of(code)
        bu_field = next((f["key"] for f in fields if f["name"] in ("BU名称", "BU部门")), None)
        if bu_field:
            rows = [r for r in rows if str(r.get(bu_field)) == str(user.get("bu"))]
    # 多字段独立模糊查询：每个查询框对应一个 query=True 字段，多字段之间 AND 关系
    if field_queries and isinstance(field_queries, dict):
        for fkey, val in field_queries.items():
            val = str(val).strip() if val is not None else ""
            if val:
                rows = [r for r in rows if val.lower() in str(r.get(fkey, "") or "").lower()]
    elif query:  # 兼容旧的单框全局模糊
        qfields = [f["key"] for f in fields_of(code) if f["query"]]
        rows = [r for r in rows if any(query.lower() in str(r.get(k, "")).lower() for k in qfields)]
    total = len(rows)
    start = (page - 1) * page_size
    page_rows = rows[start:start + page_size]
    # 公式字段：只读计算（不落库），写入 _formula_values 供展示
    from . import formula_engine
    formula_engine.enrich_formulas(code, page_rows)
    return {
        "total": total, "page": page, "page_size": page_size,
        "rows": [to_display(r, code, role) for r in page_rows],
    }


def apply_auto_ids(code: str, values: dict):
    """为「自动生成」字段发号（若用户未传）。生成后不随名称变化而改变。"""
    from . import idgen
    for f in fields_of(code):
        if f["type"] == "auto_id" and not values.get(f["key"]):
            values[f["key"]] = idgen.next_id(f["name"])


def create_record(code: str, body: dict, role="admin"):
    values = extract_values(code, body)
    errs = validate_row(code, values, db.all_rows(code))
    if errs:
        return {"ok": False, "errors": errs}
    apply_auto_ids(code, values)
    # 跨表取值与公式字段统一由 formula_engine 在展示层计算，不落库为手工事实（R2）
    values.update({"_created_at": datetime.now().isoformat(timespec="seconds"),
                   "_updated_at": datetime.now().isoformat(timespec="seconds"),
                   "_version": 1, "_status": "草稿", "_created_by": role})
    rid = db.save_row(code, values)
    return {"ok": True, "id": rid}


def update_record(code: str, row_id: int, body: dict, role="admin"):
    values = extract_values(code, body)
    rows = db.all_rows(code)
    errs = validate_row(code, values, rows, exclude_id=row_id)
    if errs:
        return {"ok": False, "errors": errs}
    # 公式字段由引擎实时计算（R2），此处不写入
    values.update({"_updated_at": datetime.now().isoformat(timespec="seconds"),
                   "_version": (next((r.get("_version") or 1 for r in rows if r["_id"] == row_id), 1) + 1)})
    n = db.update_row(code, row_id, values)
    return {"ok": n > 0, "id": row_id}


def delete_record(code: str, ids: list):
    n = db.delete_rows(code, ids)
    return {"ok": n > 0, "deleted": n}


def import_preview(code: str, records: list):
    """导入预览：逐行校验，返回每行结果与整体统计。"""
    existing = db.all_rows(code)
    results = []
    valid = 0
    for i, rec in enumerate(records, 1):
        values = {}
        fields = fields_of(code)
        # 支持中文名或 key
        for f in fields:
            if f["type"] in ("formula", "auto_id"):
                continue  # 只读字段不接受导入
            if f["key"] in rec:
                values[f["key"]] = rec[f["key"]]
            elif f["name"] in rec:
                values[f["key"]] = rec[f["name"]]
        errs = validate_row(code, values, existing)
        results.append({"row": i, "values": values, "errors": errs, "valid": not errs})
        if not errs:
            valid += 1
    return {"total": len(records), "valid": valid, "invalid": len(records) - valid, "rows": results}


def import_commit(code: str, records: list):
    now = datetime.now().isoformat(timespec="seconds")
    rows = []
    for rec in records:
        r = dict(rec)
        apply_auto_ids(code, r)  # 自动生成字段发号
        r.update({"_created_at": now, "_updated_at": now, "_version": 1, "_status": "草稿", "_created_by": "import"})
        rows.append(r)
    n = db.upsert_many(code, rows)
    return {"ok": True, "imported": n}


def export_excel(code: str):
    """导出为 Excel（xlsx）字节流。"""
    from openpyxl import Workbook
    module = get_module(code)
    wb = Workbook()
    ws = wb.active
    ws.title = module["name"][:31]
    fields = module["fields"]
    ws.append([f["name"] for f in fields])
    for r in db.all_rows(code):
        ws.append([r.get(f["key"]) for f in fields])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def download_template(code: str):
    from openpyxl import Workbook
    module = get_module(code)
    wb = Workbook()
    ws = wb.active
    ws.title = module["name"][:31]
    # 仅包含「是否导入字段」= 是 且非只读（公式/自动生成）的字段
    fields = [f for f in module["fields"] if f["importable"] and f["type"] not in ("formula", "auto_id")]
    headers = [f["name"] + (" *" if f["required"] else "") for f in fields]
    ws.append(headers)
    # 示例行：枚举字段写入首个枚举值
    sample = []
    for f in fields:
        if f["auto"]:
            sample.append("(自动取值)")
        elif f["enum"]:
            sample.append(f["enum"][0])
        else:
            sample.append("")
    ws.append(sample)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
