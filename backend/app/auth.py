"""权限模型（对齐详设 §9）：7 类角色 + RBAC(功能/按钮) + ABAC(行/字段)。"""
ROLES = {
    "super_admin": {"name": "超级管理员", "desc": "系统配置与授权，不直接改业务快照"},
    "admin": {"name": "管理员", "desc": "全流程管理与配置"},
    "budget_admin": {"name": "预算管理员", "desc": "预算属性、同步与对账"},
    "finance_bp": {"name": "财经BP", "desc": "填报审核、四算评审"},
    "bu_head": {"name": "BU总", "desc": "本 BU 数据查看与审批"},
    "bg_head": {"name": "BG总", "desc": "本 BG 数据查看与审批"},
    "certifier": {"name": "认证人", "desc": "项目认证与确认"},
}

# 功能权限：角色可访问的模块（* 表示全部）
FUNCTION_SCOPE = {
    "super_admin": "*", "admin": "*", "budget_admin": "*", "finance_bp": "*",
    "bu_head": ["project_list", "lifecycle_forecast", "biz_class", "calc_project_param"],
    "bg_head": ["project_list", "lifecycle_forecast", "biz_class"],
    "certifier": ["project_list", "lifecycle_forecast"],
}

# 按钮权限
BUTTONS = {
    "super_admin": ["view", "export"],
    "admin": ["view", "create", "edit", "delete", "import", "export", "submit", "review"],
    "budget_admin": ["view", "edit", "import", "export", "submit", "sync"],
    "finance_bp": ["view", "edit", "import", "export", "submit", "review"],
    "bu_head": ["view", "export", "approve"],
    "bg_head": ["view", "export", "approve"],
    "certifier": ["view", "certify"],
}

# 字段级权限：对敏感字段脱敏（字段 key -> 允许角色）。None 表示全员可见
FIELD_MASK = {
    "tax_rate.f2": ["super_admin", "admin", "budget_admin", "finance_bp"],  # 税率
}

# 行级数据范围：按 BU 过滤（bu_head/bg_head/certifier 仅看授权 BU）
ROW_SCOPE_BY_ROLE = {
    "bu_head": "own_bu", "bg_head": "own_bg", "certifier": "own_bu",
}

# 演示用户（登录用）
USERS = [
    {"username": "admin", "password": "123456", "role": "admin", "name": "系统管理员", "bu": "全部", "bg": "全部"},
    {"username": "super", "password": "123456", "role": "super_admin", "name": "超级管理员", "bu": "全部", "bg": "全部"},
    {"username": "budget", "password": "123456", "role": "budget_admin", "name": "张预算", "bu": "全部", "bg": "全部"},
    {"username": "bp", "password": "123456", "role": "finance_bp", "name": "李财经", "bu": "全部", "bg": "全部"},
    {"username": "bu", "password": "123456", "role": "bu_head", "name": "王BU总", "bu": "智能产业事业部", "bg": "EBG"},
    {"username": "bg", "password": "123456", "role": "bg_head", "name": "赵BG总", "bu": "全部", "bg": "EBG"},
    {"username": "cert", "password": "123456", "role": "certifier", "name": "陈认证", "bu": "智能产业事业部", "bg": "EBG"},
]
USER_BY_NAME = {u["username"]: u for u in USERS}


def can_access_module(role: str, code: str) -> bool:
    scope = FUNCTION_SCOPE.get(role, [])
    return scope == "*" or code in scope


def can_button(role: str, button: str) -> bool:
    return button in BUTTONS.get(role, [])


def field_visible(code: str, field_key: str, role: str) -> bool:
    allowed = FIELD_MASK.get(f"{code}.{field_key}")
    return allowed is None or role in allowed


def row_scope(user: dict) -> str | None:
    return ROW_SCOPE_BY_ROLE.get(user["role"])
