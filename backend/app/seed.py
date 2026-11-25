"""模拟数据生成器（v2）：按新注册表 24 模块/518 字段生成符合口径的演示数据。
公式字段不落值（由 formula_engine 实时计算），自动生成字段由发号器生成。
"""
import random
from datetime import datetime
from .registry import MODULES, MODULE_BY_CODE, fields_of
from . import database as db

random.seed(20261125)

BU_LIST = ["云桌面事业部", "网络安全事业部", "服务产品事业部", "战略合作事业部", "智能产业事业部"]
BG_LIST = ["EBG", "CBG", "TBU", "SBG", "AIABU", "SID"]
SERVICE_ORG = ["运营商服务部", "政企服务部", "金融服务部", "教育服务部"]
COUNTRIES = [("CN", "中国"), ("US", "美国"), ("DE", "德国"), ("JP", "日本"), ("SG", "新加坡"), ("UK", "英国")]
IPD = ["IPD项目", "非IPD项目"]
PTYPE = ["自研产品", "联合开发", "定制化", "平台型", "解决方案"]
GRADE = ["A", "B", "C"]
LIFECYCLE = ["引入期", "成长期", "成熟期", "衰退期"]
LAYER = ["H1", "H2", "H3"]
BUDGET_ATTR = ["经营预算", "战略预算"]
PROJ_STATUS = ["规划中", "已立项", "已上市", "项目暂停"]
PROJ_BUDGET = ["专用预算", "公共预算", "共享预算"]
YESNO = ["是", "否"]
EXPENSE = [("FSCI", "研发费用", "服务采购", "技术服务"), ("FSCM", "市场费用", "市场采购", "市场推广"),
           ("FSCO", "运营费用", "运营采购", "运营服务"), ("FSCA", "管理费用", "行政采购", "行政办公")]

PROJECT_NAMES = [
    "下一代核心交换机研发", "云桌面虚拟化平台", "AI算力网关", "企业级防火墙升级",
    "数据中心光互联", "SD-WAN解决方案", "5G小基站", "智能运维平台",
    "园区网交换机", "无线AP新品", "网络安全态势感知", "混合云管理平台",
    "边缘计算节点", "存储网关", "视频会议终端", "工业互联网网关",
]

# 每个模块的行数（演示规模）；汇总表多行
ROW_COUNT = {
    "country": 6, "bg_mapping": 15, "org_tree": 20, "calc_params": 12,
    "project_class": 15, "expense_class": 12, "tax_rate": 8, "asset_util": 10,
    "biz_class": 18, "project_list": 24, "lifecycle_forecast": 30,
    "revenue_manpower": 40, "outsourcing": 25, "rd_expense": 25, "mold": 20,
    "cloud_rental": 18, "business_outsource": 20, "fixed_asset": 22,
    "cert_fee": 16, "cert_fee_auto": 16, "alloc_by_project": 12,
    "alloc_by_budget_id": 12, "sum_project": 24, "sum_dept_expense": 20,
}


def gen_value(f, row_idx):
    """按字段类型与名称语义生成值。公式字段返回 None。"""
    name, typ = f["name"], f["type"]
    if typ == "formula":
        return None  # 由计算引擎实时算
    if typ == "auto_id":
        return None  # 由发号器生成
    if f["enum"]:
        return random.choice(f["enum"])
    # 名称语义
    if "编码" in name or "ID" in name and typ == "string":
        prefix = {"一级": "L1", "二级": "L2", "三级": "L3", "四级": "L4"}.get(name[:2], "CODE")
        return f"{prefix}{random.randint(1000, 9999)}"
    if "英文简称" in name:
        return COUNTRIES[row_idx % len(COUNTRIES)][0]
    if "中文简称" in name:
        return COUNTRIES[row_idx % len(COUNTRIES)][1]
    if name in ("BG",):
        return random.choice(BG_LIST)
    if "BG总" in name:
        return f"{random.choice(['张','李','王','赵'])}总"
    if "BU总" in name:
        return f"{random.choice(['刘','陈','杨','黄'])}总"
    if name in ("BU名称", "BU部门", "归属BU"):
        return random.choice(BU_LIST)
    if "服务组织" in name:
        return random.choice(SERVICE_ORG)
    if any(k in name for k in ("负责人", "CPO", "填写人", "预算管理", "财经BP", "认证组", "运营管理", "预算负责人")):
        return f"{random.choice(['张','李','王','赵','刘','陈'])}{random.choice(['伟','芳','娜','强','敏'])}"
    if "人年" in name or "人力投入" in name:
        return round(random.uniform(1, 20), 2)
    if "适用年份" in name:
        return random.choice(["2025", "2026", "2027"])
    if "生效" in name:
        return "202501"
    if "失效" in name:
        return "209912"
    if "时间" in name or "日期" in name:
        return f"2026-{random.randint(1,12):02d}-{random.randint(1,28):02d}"
    if "税率" in name or "毛利率" in name or "费率" in name or "比例" in name or "利用率" in name or "分成" in name or "水线" in name:
        if "金额" not in name:
            return round(random.uniform(0.05, 0.45), 4)
    if "单价" in name:
        return round(random.uniform(100, 8000), 2)
    if "数量" in name or "编制数" in name:
        return random.randint(1, 20)
    if "金额" in name or "收入" in name or "费用" in name or "毛利额" in name:
        return round(random.uniform(100, 5000), 2)
    if "原因" in name or "目的" in name:
        return random.choice(["拓展运营商市场，提升份额", "提升产品核心竞争力", "技术预研与能力储备", "降本增效，优化结构", "重点客户定制需求"])
    if "解决方案" in name:
        return random.choice(["混合云解决方案", "数据中心解决方案", "园区网解决方案"])
    if "状态" in name:
        return random.choice(PROJ_STATUS)
    if "所处生命周期" in name:
        return random.choice(LIFECYCLE)
    if "三层面" in name:
        return random.choice(LAYER)
    if "预算属性" in name:
        return random.choice(BUDGET_ATTR)
    if "等级" in name:
        return random.choice(GRADE)
    if "颗粒度" in name:
        return random.choice(["项目", "项目大包", "其他-非IPD项目"])
    if "型号" in name:
        return f"RG-{random.choice(['S','E','W'])}-{random.randint(1000,9999)}"
    if "IPD" in name:
        return random.choice(IPD)
    if "项目类型" in name:
        return random.choice(PTYPE)
    if "科目" in name:
        e = EXPENSE[row_idx % len(EXPENSE)]
        return e[1] if "名称" in name else e[0]
    if "采购类型" in name:
        return random.choice(["服务采购", "市场采购", "运营采购", "行政采购"])
    if "采购分类" in name:
        return random.choice(["技术服务", "市场推广", "运营服务", "行政办公"])
    if "场景" in name:
        return random.choice(["通用计算", "AI训练", "存储", "网络"])
    if "岗位" in name:
        return random.choice(["软件开发工程师", "硬件测试工程师", "算法工程师"])
    if "城市" in name:
        return random.choice(["北京", "上海", "深圳", "杭州", "成都"])
    if "是否" in name:
        return random.choice(YESNO)
    if typ == "integer":
        return random.randint(1, 99)
    if typ == "number":
        return round(random.uniform(1, 1000), 2)
    if typ == "url":
        return "https://example.com/estimate-model"
    return f"{name}值{random.randint(1,99)}"


def build_rows(code, n, now):
    m = MODULE_BY_CODE[code]
    rows = []
    for i in range(n):
        values = {f["key"]: gen_value(f, i) for f in m["fields"]}
        # 关键字段语义覆盖
        if code == "project_list":
            values["f2"] = PROJECT_NAMES[i % len(PROJECT_NAMES)] + (f"-{i//len(PROJECT_NAMES)+1}" if i >= len(PROJECT_NAMES) else "")
            values["f1"] = f"PL-2026-{i+1:03d}"  # 规划项目ID（文本，非自动生成）
            values["f12"] = IPD[i % 2]
            values["f13"] = PTYPE[i % len(PTYPE)]
        if code in ("lifecycle_forecast", "revenue_manpower"):
            idx = i % 24
            values["f2"] = PROJECT_NAMES[idx % len(PROJECT_NAMES)] + (f"-{idx//len(PROJECT_NAMES)+1}" if idx >= len(PROJECT_NAMES) else "")
        # 费用类模块：引用字段绑定真实主数据值，保证存在性校验通过
        if code in ("outsourcing", "rd_expense", "mold", "cloud_rental", "business_outsource", "fixed_asset", "cert_fee"):
            biz_rows = db.all_rows('biz_class')
            org_rows = db.all_rows('org_tree')
            if biz_rows:
                values["f7"] = biz_rows[i % len(biz_rows)].get('f2')  # 一级业务名称
            # 规划项目名称字段（各模块 key 需按 name 找）
            prj = next((f for f in m["fields"] if f["name"] == "规划项目名称"), None)
            if prj and PROJECT_NAMES:
                idx = i % 24
                values[prj["key"]] = PROJECT_NAMES[idx % 16] + (f"-{idx//16+1}" if idx >= 16 else "")
            # 费用归属部门（二/三/四级）：按字段名找 key，取组织表真实名
            dep_map = {"费用归属部门（二级部门）": 'f4', "费用归属部门（三级部门）": 'f6',
                       "费用归属部门（四级部门）": 'f8', "费用归属部门（BU二级部门）": 'f4',
                       "费用归属部门（BU三级部门）": 'f6', "费用归属部门（BU四级部门）": 'f8'}
            for f in m["fields"]:
                if f["name"] in dep_map and org_rows:
                    values[f["key"]] = org_rows[i % len(org_rows)].get(dep_map[f["name"]])
        values.update({"_created_at": now, "_updated_at": now, "_version": 1,
                       "_status": "草稿", "_created_by": "seed"})
        rows.append(values)
    return rows


def reset_and_seed():
    """清空所有模块表后重新生成（含发号器）。"""
    from . import idgen
    with db.engine.begin() as conn:
        for m in MODULES:
            conn.execute(db.table(m["code"]).delete())
    counts = {}
    now = datetime.now().isoformat(timespec="seconds")
    for m in MODULES:
        code = m["code"]
        n = ROW_COUNT.get(code, 10)
        rows = build_rows(code, n, now)
        # 自动生成字段发号
        for r in rows:
            for f in m["fields"]:
                if f["type"] == "auto_id":
                    r[f["key"]] = idgen.next_id(f["name"])
        db.upsert_many(code, rows)
        counts[code] = n
    return counts
