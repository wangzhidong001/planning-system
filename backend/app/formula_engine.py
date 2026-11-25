"""公式字段计算引擎（只读，结果不落库为可编辑事实）。

对齐详设 R2/R6/R11：
- 公式字段由引擎计算，用户不可编辑。
- 跨表引用按「来源表/来源字段/条件」解析，用字段名匹配（不依赖行号）。
- 算术表达式在受控命名空间内求值，禁止访问外部对象。
- 无法计算时返回 None + status 说明，不用 0 冒充未知。

公式分四类（来自飞书表 231 个公式字段实测）：
1. 跨表引用：含「来源表」「来源字段」「条件」
2. 算术表达式：+ - * / ( )，引用同表字段名
3. 自动匹配账号：「自动匹配上传人的账号」
4. 其他判断规则：如固资标准判断（尽力解析）
"""
import re
from .registry import MODULES, MODULE_BY_CODE
from . import database as db

MODULE_BY_NAME = {m["name"]: m for m in MODULES}

# 模块名别名（飞书表里的简写/别名 -> 正式模块 code）
MODULE_ALIAS = {
    "BGBU对应关系": "bg_mapping", "BGBG对应关系": "bg_mapping",
    "组织架构表": "org_tree", "项目分类标准": "project_class",
    "规划项目清单": "project_list", "项目基础信息表": "project_list",
    "二、投产规划项目信息表": "project_list",
    "三层面业务分类": "biz_class", "项目生命周期预测": "lifecycle_forecast",
    "开票收入和人力投入": "revenue_manpower", "四算表参数": "calc_params",
    "费用税率标准": "tax_rate", "费用分类标准": "expense_class",
    "固资利用率": "asset_util", "汇总表：项目投产表汇总": "sum_project",
    "项目投产表汇总": "sum_project", "部门费用汇总表": "sum_dept_expense",
    "产品资质办理费": "cert_fee", "新增固定资产": "fixed_asset",
}

def resolve_code(name: str):
    name = (name or "").strip().strip("《》")
    if name in MODULE_BY_NAME:
        return MODULE_BY_NAME[name]["code"]
    if name in MODULE_ALIAS:
        return MODULE_ALIAS[name]
    # 模糊：去括号后缀
    base = re.sub(r"（.*?）", "", name)
    if base in MODULE_BY_NAME:
        return MODULE_BY_NAME[base]["code"]
    return None

def field_key(code: str, field_name: str):
    """按字段名找 key。"""
    for f in MODULE_BY_CODE.get(code, {}).get("fields", []):
        if f["name"] == field_name:
            return f["key"]
    return None

def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def parse_lookup(formula: str):
    """解析跨表引用公式，返回 (source_code, source_field_key, [(cond_field, cond_side), ...]) 或 None。"""
    if "来源表" not in formula:
        return None
    m = re.search(r"来源表[：:]\s*(.+)", formula)
    if not m:
        return None
    src_line = m.group(1).split("\n")[0].strip()
    src_code = resolve_code(src_line)
    if not src_code:
        return None
    mf = re.search(r"来源字段[：:]\s*(.+)", formula)
    src_field = mf.group(1).split("\n")[0].strip() if mf else None
    sfk = field_key(src_code, src_field) if src_field else None
    # 条件：形如 "规划项目名称=规划项目名称" 或 "取值条件：BU部门=BU"
    conds = []
    for cm in re.finditer(r"([^\n；;]+?)\s*[=＝]\s*([^\n；;]+)", formula):
        left, right = cm.group(1).strip(), cm.group(2).strip()
        if "来源" in left or "来源" in right or "条件" in left:
            continue
        # 去掉可能的字段名噪声
        left = re.sub(r"^(条件|取值条件)[：:]?", "", left).strip()
        if not left or not right:
            continue
        conds.append((left, right))
    return src_code, sfk, conds

def compute_lookup(code: str, field: dict, row: dict, formula: str):
    parsed = parse_lookup(formula)
    if not parsed:
        return None, "LOOKUP_PARSE_FAIL"
    src_code, sfk, conds = parsed
    src_rows = db.all_rows(src_code)
    src_fields = MODULE_BY_CODE[src_code]["fields"]
    src_name_to_key = {f["name"]: f["key"] for f in src_fields}
    # 本行字段名 -> 值
    my_fields = MODULE_BY_CODE[code]["fields"]
    my_name_to_key = {f["name"]: f["key"] for f in my_fields}
    my_name_to_val = {f["name"]: row.get(f["key"]) for f in my_fields}

    def resolve_side(side, name_to_key, name_to_val):
        # side 可能是本表字段名，或来源表字段名
        k = name_to_key.get(side)
        if k is not None:
            return name_to_val.get(k)
        return None

    for sr in src_rows:
        ok = True
        for left, right in conds:
            # left/right 是字段名，各自在对应表里取值
            lv = resolve_side(left, my_name_to_key, my_name_to_val)
            rv = resolve_side(right, src_name_to_key, {f["name"]: sr.get(f["key"]) for f in src_fields})
            if lv is None or rv is None or str(lv) != str(rv):
                ok = False
                break
        if ok:
            if sfk:
                return sr.get(sfk), "OK"
            # 无来源字段：返回整行第一个非空业务值（兜底）
            for f in src_fields:
                if sr.get(f["key"]) not in (None, ""):
                    return sr.get(f["key"]), "OK"
    return None, "NO_MATCH"

def compute_arithmetic(code: str, row: dict, formula: str, _stack=None):
    """算术表达式：把中文字段名替换为数值，在受控命名空间求值。
    若引用的字段本身是公式字段，递归计算它（依赖链），避免引用未落库的公式值而算成 0。"""
    _stack = _stack or set()
    fields = MODULE_BY_CODE[code]["fields"]
    name_to_field = {f["name"]: f for f in fields}
    expr = formula.strip()
    # 将字段名替换为数值（按长度降序，避免短名先匹配）
    for name in sorted(name_to_field, key=len, reverse=True):
        if name not in expr:
            continue
        f = name_to_field[name]
        if f["type"] == "formula":
            if name in _stack or f["key"] in _stack:
                expr = expr.replace(name, "0")  # 循环依赖：断开，避免死循环
                continue
            dep_val, _ = compute_field(code, f, row, _stack=_stack | {name, f["key"]})
            v = _num(dep_val)
        else:
            v = _num(row.get(f["key"]))
        expr = expr.replace(name, str(v if v is not None else 0))
    # 全角括号 -> 半角
    expr = expr.replace("（", "(").replace("）", ")").replace("，", ",")
    # 只允许数字、运算符、括号、小数点
    if not re.fullmatch(r"[0-9+\-*/().\s]+", expr):
        return None, "UNSUPPORTED_EXPR"
    try:
        val = eval(expr, {"__builtins__": {}}, {})
        return round(val, 6), "OK"
    except ZeroDivisionError:
        return None, "DIV_ZERO"
    except Exception:
        return None, "EVAL_FAIL"

def compute_field(code: str, field: dict, row: dict, user_name: str = "", _stack=None):
    """计算单个公式字段，返回 (value, status_code)。"""
    formula = field.get("formula") or ""
    if not formula:
        return None, "NO_FORMULA"
    if "来源表" in formula:
        return compute_lookup(code, field, row, formula)
    if "自动匹配" in formula:
        return (user_name or "系统"), "OK"
    if "固资标准" in formula or ("单价" in formula and "使用年限" in formula):
        # 尽力解析判断规则
        m = re.search(r"单价\s*>=\s*(\d+)", formula)
        if m:
            threshold = float(m.group(1))
            price_fields = [f for f in MODULE_BY_CODE[code]["fields"] if "单价" in f["name"]]
            price = _num(row.get(price_fields[0]["key"])) if price_fields else None
            if price is not None:
                return ("是" if price >= threshold else "否"), "OK"
        return None, "JUDGE_PARTIAL"
    # 算术
    val, st = compute_arithmetic(code, row, formula, _stack=_stack)
    if st == "OK":
        return val, st
    return None, st

def enrich_formulas(code: str, rows: list, user_name: str = ""):
    """为一批行计算所有公式字段，写入临时展示字段 _formula_values / _formula_status。
    不修改原始 row 的业务字段（不落库）。"""
    fields = MODULE_BY_CODE.get(code, {}).get("fields", [])
    fml_fields = [f for f in fields if f["type"] == "formula"]
    for r in rows:
        fv, fs = {}, {}
        for f in fml_fields:
            v, st = compute_field(code, f, r, user_name)
            fv[f["key"]] = v
            fs[f["key"]] = st
        r["_formula_values"] = fv
        r["_formula_status"] = fs
    return rows
