# 产研项目投产规划填报系统 V2

依据**飞书《各模块字段定义》**（[链接](https://ruijie.feishu.cn/wiki/WqrKwGUKUiUQ3PkML3Zc5wCZnDf?sheet=tshMVC)）两个工作表构建的**完整前后端系统**。核心理念：**元数据驱动 + 稳定 ID + 公式引擎**——把飞书表的字段定义固化为注册表，前后端通用组件据此自动生成页面、校验与计算，不逐个硬编码。

> V2 相对 V1（本地 Excel 11 模块/108 字段）是**全量重建**：以飞书表为唯一事实来源，扩到 **24 模块 / 518 字段**，其中 **231 个公式字段、7 个自动生成字段**。旧注册表已备份为 `registry_v1_108fields.json.bak`。

## 一、技术栈

| 层 | 技术 |
|---|---|
| 前端 | React 18 + Vite 6 + Ant Design 5 |
| 后端 | Python FastAPI + SQLAlchemy 2 + SQLite |
| 计算 | 公式引擎（跨表 lookup / 算术依赖链 / 判断规则） |
| 发号 | 稳定 ID 发号器（`前缀-年份-6位序号`，持久化不复用） |
| 数据交换 | Excel 导入导出（openpyxl / xlsx） |

## 二、已实现范围（按你的三项决策）

1. **全量重建**：24 模块 / 518 字段（主数据 8 模块 54 字段 + 投产填报 16 模块 464 字段）。
2. **公式字段只读 + 真实计算**：231 个公式字段不可手工填报，由后端引擎按口径实时计算，展示时带 ƒ 标记；无法计算时显示「不适用」并标注状态码，不用 0 冒充。
3. **自动生成字段用发号器**：预算ID / 规划项目ID 等按稳定规则发号，生成后不随名称变化而改变。
4. **数据**：内置 440 行模拟数据（引用字段绑定真实主数据，保证存在性校验通过），支持 Excel 导入（预览校验）与导出。

## 三、字段元数据模型（飞书 15 列 → 系统能力）

| 飞书列 | 系统实现 |
|---|---|
| 二级功能模块 / 字段名称 / 字段类型 | 动态表单（文本/整数/数字/日期/枚举/公式/时间/超链接/自动生成） |
| 字段说明 | 列标题与表单 label 的 Tooltip |
| 是否必填 | 表单 + 导入必填校验 |
| 是否筛选字段（支持模糊查询） | 列表页每个查询字段独立查询框，多字段 AND |
| 是否导入 / 导出 | 导入模板（仅可导入且非只读字段）、导出 Excel |
| 枚举 / 取值公式（跨表引用） | 下拉选项 + 枚举校验；跨表 lookup 引擎 |
| 导入时校验类型 | 枚举值 / 重复性 / 存在性（通用解析目标表）三类 |
| 格式要求 | 千分位 / 两位小数 / 百分比（展示层，存储保持原值） |
| 是否自动生成 / 自动生成规则 / 其他规则 | 发号器 + 判断规则解析 |

## 四、公式引擎（四类，实测覆盖）

| 类型 | 数量 | 处理 |
|---|---|---|
| 跨表引用 | 123 | 解析「来源表/来源字段/条件」，按字段名匹配（不依赖行号） |
| 算术表达式 | 69 | 受控命名空间求值，支持公式字段依赖链递归（如不含税=含税/(1+税率)） |
| 自动匹配账号 | 15 | 填入当前用户 |
| 判断规则 | 24 | 如固资标准：单价≥4000 → 「是」，尽力解析 |

- 计算结果写入展示层 `_formula_values`，**不落库为可编辑事实**（对齐详设 R2）。
- 状态码区分 `OK / NO_MATCH / DIV_ZERO / UNSUPPORTED_EXPR / JUDGE_PARTIAL`，非 OK 显示「不适用」。

## 五、权限（对齐详设 §9，RBAC + ABAC）

7 类角色：超级管理员 / 管理员 / 预算管理员 / 财经BP / BU总 / BG总 / 认证人。
- **功能权限**：角色可见模块不同（BU总仅见项目相关模块）。
- **按钮权限**：新增/编辑/删除/导入/导出/评审按角色授予（超级管理员只读）。
- **行权限**：BU总/BG总/认证人按 BU 过滤。
- **字段权限**：税率等敏感字段对部分角色脱敏。

体验账号（密码均 `123456`）：`admin` / `super` / `budget` / `bp` / `bu` / `bg` / `cert`。

## 六、启动方式

### 方式 A：两个终端分别启动（推荐）

```bash
cd planning-system/backend
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```
```bash
cd planning-system/frontend
npm run dev            # http://localhost:5173
```

### 方式 B：双击批处理

`start-backend.bat` + `start-frontend.bat`。

> 首次启动或换注册表后，若表结构变化需删除 `backend/app/data.db`，启动会自动 seed。

## 七、查询 / 模板 / 展示规则

1. **多字段独立查询**：每个「是否筛选字段」一个输入框，多字段 AND。
2. **模板过滤**：下载模板只含「是否导入=是」且非公式/自动生成的字段。
3. **全字段展示**：表格平铺全部字段；公式字段紫色 ƒ 标记，自动发号蓝色 Tag。
4. **格式严格按「格式要求」列**：`千分位；两位小数`→`3,986.37`；`百分比；`与`百分比；两位小数`→统一两位小数如 `40.00%`（逻辑在 `frontend/src/format.js`，仅展示层）。

## 八、API 一览

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | /api/health | 健康检查（模块数/字段数） |
| POST | /api/login | 登录 |
| GET | /api/meta/modules | 模块树（按权限过滤） |
| GET | /api/meta/module/{code} | 模块字段定义 |
| GET | /api/data/{code}?fields_q=&page= | 分页 + 多字段查询（含公式计算） |
| POST | /api/data/{code} | 新增（发号器 + 校验） |
| PUT | /api/data/{code}/{row_id} | 更新 |
| DELETE | /api/data/{code}?ids= | 删除 |
| POST | /api/import/{code}/preview | 导入预览校验 |
| POST | /api/import/{code}/commit | 导入提交 |
| GET | /api/export/{code} | 导出 Excel |
| GET | /api/template/{code} | 下载导入模板 |
| POST | /api/seed/reset | 重置模拟数据（管理员） |

## 九、待确认 / 预留

- **四算表填报、权限管理**：飞书表「功能模块」列出但未提供字段定义，保留菜单入口（预留空位），补字段后自动生成 CRUD。
- 公式引擎对「其他规则」24 条为尽力解析（如固资标准判断），复杂 `IF/FILTER` 嵌套公式按状态码提示，未逐条搬运旧公式（对齐详设「不逐条搬公式」红线）。

## 十、Streamlit 云端版（本次发布）

React + FastAPI 无法直接在 Streamlit Cloud 运行，因此新增 `streamlit_app.py`：**复用同一套 Python 后端**（注册表、公式引擎、发号器、权限、导入导出服务），用 Streamlit 原生组件重写填报界面。功能与 React 版等价：7 类账号登录、24 模块动态页面、公式只读/状态提示、多字段 AND 查询、全字段展示、格式要求严格格式化、Excel 导入导出与模板。

云端启动时把仓库自带的 `backend/app/data.db`（440 行模拟数据）复制到临时可写目录，通过环境变量 `PLANNING_DB_PATH` 指向它；容器重启后新增数据会丢失，模拟数据与体验账号始终可用。

### 本地运行

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py --server.port 8501
```

### Streamlit Cloud 部署

1. GitHub 仓库：`https://github.com/wangzhidong001/planning-system`（公开）。
2. 打开 https://share.streamlit.io → New app → 选择该仓库，Main file path 填 `streamlit_app.py`。
3. Python 版本选 3.11+，无需填写 Secrets（无外部密钥依赖）。
4. Deploy 后得到公开访问地址。
