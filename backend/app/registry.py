"""共享元数据注册表：从 registry.json 加载模块与字段定义。"""
import json
from pathlib import Path

_REG_PATH = Path(__file__).with_name("registry.json")
_DATA = json.loads(_REG_PATH.read_text(encoding="utf-8"))

MODULES = _DATA["modules"]
MODULE_BY_CODE = {m["code"]: m for m in MODULES}
FIELD_COUNT = _DATA["field_count"]

# 一级模块 -> 二级模块清单（含无字段的预留模块）
TOP_MODULES = [
    {
        "name": "主数据管理",
        "code": "md",
        "children": [{"code": m["code"], "name": m["name"]} for m in MODULES if m["top"] == "主数据管理"],
    },
    {
        "name": "投产表填报",
        "code": "planning",
        "children": [{"code": m["code"], "name": m["name"]} for m in MODULES if m["top"] == "投产表填报"],
    },
    {
        "name": "四算表填报",
        "code": "calc",
        "children": [
            {"code": "calc_project_param", "name": "四算项目参数"},
            {"code": "calc_product_config", "name": "产品配置"},
            {"code": "calc_cost", "name": "成本"},
            {"code": "calc_price", "name": "价格"},
            {"code": "calc_ramp", "name": "爬坡"},
            {"code": "calc_custom_income", "name": "定制化收入"},
            {"code": "calc_rd", "name": "研发投入"},
            {"code": "calc_period", "name": "期间核算"},
        ],
    },
    {
        "name": "权限管理",
        "code": "perm",
        "children": [
            {"code": "perm_user", "name": "用户管理"},
            {"code": "perm_role", "name": "角色管理"},
            {"code": "perm_scope", "name": "数据范围"},
        ],
    },
]

# 无字段模块（预留空位）的 code 集合
RESERVED_CODES = {c["code"] for top in TOP_MODULES for c in top["children"] if c["code"] not in MODULE_BY_CODE}


def get_module(code: str):
    return MODULE_BY_CODE.get(code)


def fields_of(code: str):
    m = get_module(code)
    return m["fields"] if m else []
