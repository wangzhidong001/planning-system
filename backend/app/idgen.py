"""稳定 ID 发号器（对齐详设 R1 红线）。

规则：
- 内部主键 _id 仍由 SQLite 自增（仅内部使用，不对外作业务编码）。
- 「自动生成」业务字段（预算ID/规划项目ID等）由本发号器生成：
  * 全局唯一、永不复用、不含名称含义
  * 格式：<前缀>-<年份>-<6位序号>，序号持久化到 seq 表，删行不改变
  * 已生成的 ID 若该行已写入，则保持不变（生成后不再改变）
"""
import re
from datetime import datetime
from . import database as db
from sqlalchemy import text

# 前缀映射：字段名 -> 编码前缀
PREFIX_RULES = [
    ("预算ID", "BUD"),
    ("规划项目ID", "PL"),
    ("项目流水ID", "PL"),
    ("项目编码", "PRJ"),
]

def prefix_for(field_name: str) -> str:
    for kw, p in PREFIX_RULES:
        if kw in field_name:
            return p
    return "ID"

def ensure_seq_table():
    with db.engine.begin() as conn:
        conn.execute(text(
            "CREATE TABLE IF NOT EXISTS id_seq ("
            "seq_key TEXT PRIMARY KEY, next_val INTEGER NOT NULL)"
        ))

def next_id(field_name: str) -> str:
    """生成一个稳定业务 ID。"""
    ensure_seq_table()
    prefix = prefix_for(field_name)
    year = datetime.now().year
    seq_key = f"{prefix}-{year}"
    with db.engine.begin() as conn:
        row = conn.execute(text("SELECT next_val FROM id_seq WHERE seq_key=:k"), {"k": seq_key}).fetchone()
        if row is None:
            val = 1
            conn.execute(text("INSERT INTO id_seq(seq_key,next_val) VALUES(:k,:v)"), {"k": seq_key, "v": val + 1})
        else:
            val = row[0]
            conn.execute(text("UPDATE id_seq SET next_val=:v WHERE seq_key=:k"), {"k": seq_key, "v": val + 1})
    return f"{prefix}-{year}-{val:06d}"

def is_id_field(field_name: str) -> bool:
    return any(kw in field_name for kw, _ in PREFIX_RULES)
