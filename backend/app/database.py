"""动态数据模型：按注册表为每个二级模块生成一张 SQLite 表。
系统列以 _ 开头，避免与业务字段 f1..fn 冲突。"""
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String, Float, Text
from sqlalchemy.dialects.sqlite import Insert
from pathlib import Path
import os
from .registry import MODULES, fields_of

# 本地开发默认用源码目录 data.db；云端（Streamlit Cloud）用环境变量指向可写临时目录。
DB_PATH = Path(os.environ.get("PLANNING_DB_PATH") or Path(__file__).with_name("data.db"))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
metadata = MetaData()

TYPE_MAP = {
    "string": String, "enum": String, "date": String, "time": String, "url": String,
    "integer": Integer, "number": Float,
    # 公式/自动生成字段在库中保留列（存计算值或发号值），但前端只读、引擎维护
    "formula": String, "auto_id": String,
}
SYS_COLS = ["_id", "_created_at", "_updated_at", "_version", "_status", "_created_by"]

_tables = {}

def build_tables():
    for m in MODULES:
        cols = [Column("_id", Integer, primary_key=True, autoincrement=True)]
        for f in m["fields"]:
            cols.append(Column(f["key"], TYPE_MAP.get(f["type"], String), nullable=True))
        for name in SYS_COLS[1:]:
            ctype = Integer if name == "_version" else String
            cols.append(Column(name, ctype, nullable=True))
        _tables[m["code"]] = Table(f"mod_{m['code']}", metadata, *cols)
    metadata.create_all(engine)

def table(code: str):
    return _tables.get(code)

def save_row(code: str, payload: dict):
    t = table(code)
    if t is None:
        return None
    with engine.begin() as conn:
        result = conn.execute(t.insert().values(**payload))
        return result.inserted_primary_key[0]

def update_row(code: str, row_id: int, payload: dict):
    t = table(code)
    if t is None:
        return 0
    with engine.begin() as conn:
        result = conn.execute(t.update().where(t.c._id == row_id).values(**payload))
        return result.rowcount

def delete_rows(code: str, ids: list):
    t = table(code)
    if t is None:
        return 0
    with engine.begin() as conn:
        return conn.execute(t.delete().where(t.c._id.in_(ids))).rowcount

def all_rows(code: str):
    t = table(code)
    if t is None:
        return []
    with engine.connect() as conn:
        rows = conn.execute(t.select()).mappings().all()
        return [dict(r) for r in rows]

def upsert_many(code: str, rows: list):
    """导入用：整批写入。"""
    t = table(code)
    if t is None:
        return 0
    with engine.begin() as conn:
        for r in rows:
            conn.execute(t.insert().values(**r))
    return len(rows)

build_tables()
