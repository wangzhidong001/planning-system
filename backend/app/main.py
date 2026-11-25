"""FastAPI 主应用：投产规划填报系统 API。"""
from fastapi import FastAPI, Header, HTTPException, Query, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from io import BytesIO
from contextlib import asynccontextmanager
import json
from urllib.parse import quote

from .registry import MODULES, MODULE_BY_CODE, TOP_MODULES, RESERVED_CODES, FIELD_COUNT
from . import auth, services, seed


@asynccontextmanager
async def lifespan(app: FastAPI):
    """启动时若库为空则初始化模拟数据（兼容新版 Starlette 已移除的 on_startup）。"""
    from .database import table, engine
    from sqlalchemy import func, select
    with engine.connect() as conn:
        cnt = conn.execute(select(func.count()).select_from(table("project_list"))).scalar()
    if cnt == 0:
        seed.reset_and_seed()
    yield


app = FastAPI(title="投产规划填报系统 API", version="1.0", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])


def current_user(x_user: str | None = Header(None)):
    """从请求头 X-User 取用户名，解析用户。简化鉴权（演示）。"""
    if not x_user or x_user not in auth.USER_BY_NAME:
        raise HTTPException(401, "未登录或登录已失效")
    return auth.USER_BY_NAME[x_user]


def require_module(code: str, user: dict):
    if code not in MODULE_BY_CODE and code not in RESERVED_CODES:
        raise HTTPException(404, "模块不存在")
    if code in MODULE_BY_CODE and not auth.can_access_module(user["role"], code):
        raise HTTPException(403, "无该模块访问权限")
    if code in RESERVED_CODES:
        raise HTTPException(200, {"reserved": True, "message": "该模块字段定义待补充（预留空位）"})


# ---------- 元数据 ----------
@app.get("/api/meta/modules")
def meta_modules(user=Header(None, alias="X-User")):
    u = current_user(user)
    tops = []
    for top in TOP_MODULES:
        children = []
        for c in top["children"]:
            if c["code"] in MODULE_BY_CODE or c["code"] in RESERVED_CODES:
                if c["code"] in MODULE_BY_CODE and not auth.can_access_module(u["role"], c["code"]):
                    continue
                children.append({**c, "has_fields": c["code"] in MODULE_BY_CODE,
                                 "field_count": len(MODULE_BY_CODE[c["code"]]["fields"]) if c["code"] in MODULE_BY_CODE else 0})
        tops.append({"name": top["name"], "code": top["code"], "children": children})
    return {"top_modules": tops, "field_count": FIELD_COUNT, "user": {"name": u["name"], "role": u["role"],
            "role_name": auth.ROLES[u["role"]]["name"], "bu": u["bu"], "bg": u["bg"],
            "buttons": auth.BUTTONS[u["role"]]}}


@app.get("/api/meta/module/{code}")
def meta_module(code: str, user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    m = MODULE_BY_CODE[code]
    fields = []
    for f in m["fields"]:
        if not auth.field_visible(code, f["key"], u["role"]):
            continue
        fields.append(f)
    return {"code": code, "name": m["name"], "top": m["top"], "fields": fields}


# ---------- 鉴权 ----------
@app.post("/api/login")
def login(body: dict):
    username = body.get("username", "").strip()
    password = body.get("password", "")
    u = auth.USER_BY_NAME.get(username)
    if not u or u["password"] != password:
        raise HTTPException(401, "用户名或密码错误")
    return {"ok": True, "username": u["username"], "name": u["name"], "role": u["role"],
            "role_name": auth.ROLES[u["role"]]["name"], "bu": u["bu"], "bg": u["bg"]}


@app.get("/api/roles")
def roles(user=Header(None, alias="X-User")):
    current_user(user)
    return auth.ROLES


# ---------- CRUD ----------
@app.get("/api/data/{code}")
def list_data(code: str, user=Header(None, alias="X-User"), q: str = "", fields_q: str = "",
              page: int = 1, page_size: int = 50):
    u = current_user(user)
    require_module(code, u)
    import json
    field_queries = None
    if fields_q:
        try:
            field_queries = json.loads(fields_q)
        except Exception:
            field_queries = None
    return services.list_records(code, u["role"], u, q, page, page_size, field_queries)


@app.post("/api/data/{code}")
def create_data(code: str, body: dict, user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    if not auth.can_button(u["role"], "create"):
        raise HTTPException(403, "无新增权限")
    return services.create_record(code, body, u["role"])


@app.put("/api/data/{code}/{row_id}")
def update_data(code: str, row_id: int, body: dict, user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    if not auth.can_button(u["role"], "edit"):
        raise HTTPException(403, "无编辑权限")
    return services.update_record(code, row_id, body, u["role"])


@app.delete("/api/data/{code}")
def delete_data(code: str, ids: list = Query(...), user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    if not auth.can_button(u["role"], "delete"):
        raise HTTPException(403, "无删除权限")
    return services.delete_record(code, ids)


# ---------- 导入 ----------
@app.post("/api/import/{code}/preview")
async def import_preview(code: str, file: UploadFile = File(...), user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    from openpyxl import load_workbook
    content = await file.read()
    wb = load_workbook(BytesIO(content), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {"total": 0, "valid": 0, "invalid": 0, "rows": []}
    headers = [str(h).strip().replace(" *", "") if h is not None else "" for h in rows[0]]
    records = []
    for r in rows[1:]:
        if all(c is None for c in r):
            continue
        records.append({headers[i]: r[i] for i in range(min(len(headers), len(r)))})
    return services.import_preview(code, records)


@app.post("/api/import/{code}/commit")
def import_commit(code: str, body: dict, user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    if not auth.can_button(u["role"], "import"):
        raise HTTPException(403, "无导入权限")
    records = body.get("rows", [])
    return services.import_commit(code, records)


# ---------- 导出 / 模板 ----------
@app.get("/api/export/{code}")
def export_data(code: str, user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    if not auth.can_button(u["role"], "export"):
        raise HTTPException(403, "无导出权限")
    buf = services.export_excel(code)
    name = quote(MODULE_BY_CODE[code]["name"]) + ".xlsx"
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename*=UTF-8''{name}"})


@app.get("/api/template/{code}")
def download_template(code: str, user=Header(None, alias="X-User")):
    u = current_user(user)
    require_module(code, u)
    buf = services.download_template(code)
    name = quote(MODULE_BY_CODE[code]["name"] + "-导入模板") + ".xlsx"
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename*=UTF-8''{name}"})


# ---------- 模拟数据 ----------
@app.post("/api/seed/reset")
def reset_seed(user=Header(None, alias="X-User")):
    u = current_user(user)
    if u["role"] not in ("admin", "super_admin"):
        raise HTTPException(403, "仅管理员可重置模拟数据")
    return {"ok": True, "counts": seed.reset_and_seed()}


@app.get("/api/health")
def health():
    return {"status": "ok", "modules": len(MODULE_BY_CODE), "fields": FIELD_COUNT}
