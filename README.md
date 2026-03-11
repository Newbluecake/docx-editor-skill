# docx-editor-skill

一个基于 Model Context Protocol (MCP) 的服务器，为 Claude 提供细粒度的 Microsoft Word (.docx) 文档操作能力。

## 特性

- **会话管理**：维护有状态的文档编辑会话，支持并发操作
- **原子化操作**：精确控制段落、文本块、标题和表格的每个元素
- **混合上下文**：支持基于 ID 的精确操作和基于上下文的便捷操作
- **高级内容操作**：
  - **表格处理**：深拷贝表格、批量填充数据
  - **文本替换**：支持跨 Run 的智能文本替换（Text Stitching）
  - **模板填充**：完善的模板占位符处理能力
- **精确布局控制**：支持通过 `position` 参数（如 `after:para_123`）将元素插入到文档的任意位置
- **可视化上下文**：工具返回直观的 ASCII 树状图，展示操作前后的文档结构
- **格式化**：设置字体（粗体、斜体、大小、颜色）和对齐方式
- **布局控制**：调整页边距和插入分页符
- **Windows GUI**：提供独立的 Windows 启动器，无需配置环境即可使用

## 响应格式

**v2.0 重大更新**：所有 MCP 工具现在返回 **Markdown 格式**的响应（不再是 JSON），包含：

- **结构化元数据**：操作状态、元素 ID、操作类型等
- **ASCII 可视化**：使用 Unicode 框线字符展示文档结构
- **上下文感知**：自动显示操作位置周围的文档元素
- **Git diff 风格**：编辑操作显示修改前后的对比

### 响应示例

创建段落的响应：

```markdown
# 操作结果: Insert Paragraph

**Status**: ✅ Success
**Element ID**: para_abc123
**Operation**: Insert Paragraph
**Position**: end:document_body

---

## 📄 Document Context

📄 Document Context (showing 3 elements around para_abc123)

  ┌─────────────────────────────────────┐
  │ Paragraph (para_xyz789)             │
  ├─────────────────────────────────────┤
  │ Previous paragraph text             │
  └─────────────────────────────────────┘

>>> [CURSOR] <<<

  ┌─────────────────────────────────────┐
  │ Paragraph (para_abc123) ⭐ CURRENT   │
  ├─────────────────────────────────────┤
  │ New paragraph text                  │
  └─────────────────────────────────────┘

  ┌─────────────────────────────────────┐
  │ Paragraph (para_def456)             │
  ├─────────────────────────────────────┤
  │ Next paragraph text                 │
  └─────────────────────────────────────┘
```

### 解析响应

对于需要提取数据的场景，可以使用正则表达式：

```python
import re

# 提取元素 ID
match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', response)
element_id = match.group(1) if match else None

# 检查操作状态
is_success = '**Status**: ✅ Success' in response
is_error = '**Status**: ❌ Error' in response

# 提取元数据字段
def extract_field(response, field_name):
    pattern = rf'\*\*{field_name}\*\*:\s*(.+?)(?:\n|$)'
    match = re.search(pattern, response)
    return match.group(1).strip() if match else None
```

**注意**：测试辅助函数可在 `tests/helpers/markdown_extractors.py` 中找到。

## ⚠️ Breaking Changes (v3.0)

**文件管理架构重大更新**：

- ❌ **已移除**: `docx_create(file_path=...)` 参数
- ❌ **已移除**: `docx_list_files()` 工具
- ✅ **新增**: 全局单文件模式 + HTTP API 文件管理
- ✅ **新增**: Combined 传输模式（FastAPI + MCP）

**迁移指南**: 详见 [docs/migration-v2-to-v3.md](docs/migration-v2-to-v3.md)

---

## 快速开始

### 安装

使用安装脚本（推荐）：

```bash
./scripts/install.sh
```

或手动安装：

```bash
pip install .
```

### Windows 用户（GUI 启动器）

直接下载最新发布的 `DocxServerLauncher.exe`，双击运行即可。无需安装 Python 或任何依赖。

**v3.0 更新**: 启动器现在支持文件选择和集中文件管理。

### 运行服务器

服务器支持四种传输模式：

#### 1. STDIO 模式（默认，用于 Claude Desktop）

```bash
mcp-server-docx
# 或显式指定
mcp-server-docx --transport stdio
```

#### 2. SSE 模式（HTTP Server-Sent Events）

适用于需要通过 HTTP 访问的场景：

```bash
# 使用默认配置（127.0.0.1:8000）
mcp-server-docx --transport sse

# 指定自定义 host 和 port
mcp-server-docx --transport sse --host 0.0.0.0 --port 3000

# 使用环境变量
DOCX_MCP_TRANSPORT=sse DOCX_MCP_HOST=127.0.0.1 DOCX_MCP_PORT=3000 mcp-server-docx
```

启动后可通过 `http://127.0.0.1:8000` 访问（或你指定的 host:port）。

#### 3. Streamable HTTP 模式

```bash
# 使用默认配置
mcp-server-docx --transport streamable-http

# 指定 host、port 和挂载路径
mcp-server-docx --transport streamable-http --host 0.0.0.0 --port 8080 --mount-path /mcp
```

启动后可通过 `http://0.0.0.0:8080/mcp` 访问（如果指定了 mount-path）。

#### 4. Combined 模式（v3.0 新增，推荐用于 GUI 启动器）

Combined 模式提供 FastAPI REST API + MCP 的组合功能，适用于 GUI 启动器集成：

```bash
# 使用默认配置（127.0.0.1:8080）
mcp-server-docx --transport combined

# 指定初始文件
mcp-server-docx --transport combined --file /path/to/document.docx

# 指定 host 和 port
mcp-server-docx --transport combined --host 0.0.0.0 --port 8080
```

**Combined 模式特性**：
- **REST API**: `POST /api/file/switch`、`GET /api/status`、`POST /api/session/close`
- **MCP Server**: 挂载在 `/mcp` 路径
- **Health Check**: `GET /health`
- **文件管理**: 通过 HTTP API 切换当前活动文件

**使用场景**：
- GUI 启动器需要文件选择功能
- 需要通过 HTTP API 集中管理文件
- 多应用共享同一个 MCP 服务器实例

#### 查看所有选项

```bash
mcp-server-docx --help
mcp-server-docx --version
```

#### Windows GUI 启动器

Windows GUI 启动器会自动使用 SSE 模式启动服务器，你可以在界面中配置：
- **Host**: 通过"Allow LAN Access"复选框选择 127.0.0.1（本地）或 0.0.0.0（局域网）
- **Port**: 在端口输入框中指定端口号
- **Working Directory**: 服务器的工作目录

### Claude Desktop 集成

**v0.3.0 更新**: 启动器现在显示完整的 Claude CLI 启动命令，而不是自动启动。这提高了可靠性和灵活性。

1. 配置服务器设置（Host, Port）。
2. 如果需要，在 "Additional CLI Parameters" 中添加参数（如 `--dangerously-skip-permission`）。
3. 点击 **"Copy Command"** 将完整命令复制到剪贴板。
4. 在你的终端（PowerShell 或 CMD）中粘贴并运行该命令。

**命令示例**:

*Windows*:
```cmd
cmd.exe /c claude --mcp-config {"mcpServers":{"docx-server":{"url":"http://127.0.0.1:8000/sse","transport":"sse"}}} --dangerously-skip-permission
```

### 构建可执行文件

如果您想自己从源码构建 Windows 可执行文件：

```powershell
# 1. 确保已安装 Python 3.10+
# 2. 运行构建脚本
.\scripts\build_exe.ps1
```

构建产物将位于 `dist\DocxServerLauncher.exe`。

### 配置 Claude Desktop

在 Claude Desktop 的 MCP 配置中添加：

```json
{
  "mcpServers": {
    "docx": {
      "command": "mcp-server-docx"
    }
  }
}
```

## MCP 工具列表

### 生命周期管理

- `docx_create(auto_save=False)` - 创建新文档会话（⚠️ v3.0: 移除了 file_path 参数）
- `docx_save(session_id, file_path)` - 保存文档到文件
- `docx_get_context(session_id)` - 获取当前会话上下文信息

**v3.0 文件管理变更**：
- 文件选择现在由 Launcher GUI 或 `--file` CLI 参数管理
- `docx_create()` 使用全局活动文件（通过 `/api/file/switch` 设置）
- 移除了 `docx_list_files()` 工具（文件浏览由 Launcher 提供）

**v4.0 会话管理变更**：
- 移除 `docx_close()` 接口，会话自动管理
- 会话在 1 小时后自动过期，无需手动关闭

### 内容检索与浏览

- `docx_read_content(session_id)` - 读取文档全文
- `docx_find_paragraphs(session_id, query)` - 查找包含特定文本的段落
- `docx_find_table(session_id, text)` - 查找包含特定文本的表格
- `docx_get_table(session_id, index)` - 按索引获取表格

### 内容编辑

- `docx_insert_paragraph(session_id, text, position, style=None)` - 添加段落（position 必选）
- `docx_insert_heading(session_id, text, position, level=1)` - 添加标题（position 必选）
- `docx_insert_run(session_id, text, position)` - 向段落添加文本块（position 必选）
- `docx_insert_page_break(session_id, position)` - 插入分页符（position 必选）
- `docx_insert_image(session_id, image_path, width=None, height=None, position)` - 插入图片（position 必选）

### 高级编辑

- `docx_copy_paragraph(session_id, paragraph_id, position)` - 复制段落（保留格式）
- `docx_copy_table(session_id, table_id, position)` - 深拷贝表格（保留结构与格式）
- `docx_copy_elements_range(session_id, start_id, end_id, position)` - 复制元素区间（如整个章节）
- `docx_replace_text(session_id, old_text, new_text, scope_id=None)` - 智能文本替换（支持模板填充）
- `docx_batch_replace_text(session_id, replacements_json, scope_id=None)` - 批量文本替换（格式保留）
- `docx_update_paragraph_text(session_id, paragraph_id, new_text)` - 更新段落文本
- `docx_update_run_text(session_id, run_id, new_text)` - 更新 Run 文本
- `docx_extract_template_structure(session_id)` - 提取文档模板结构（智能识别标题、表格、段落）
- `docx_extract_format_template(session_id, element_id)` - 提取格式模板
- `docx_apply_format_template(session_id, element_id, template_json)` - 应用格式模板
- `docx_get_element_source(session_id, element_id)` - 获取元素来源元数据

### 表格操作

- `docx_insert_table(session_id, rows, cols, position)` - 创建表格（position 必选）
- `docx_insert_table_row(session_id, position)` - 添加行到表格末尾（position 必选）
- `docx_insert_table_col(session_id, position)` - 添加列到表格末尾（position 必选）
- `docx_insert_row_at(session_id, table_id, position, row_index=None, copy_format=False)` - 在指定位置插入行（支持 after:N, before:N, start, end）
- `docx_insert_col_at(session_id, table_id, position, col_index=None, copy_format=False)` - 在指定位置插入列（支持 after:N, before:N, start, end）
- `docx_delete_row(session_id, table_id, row_index)` - 删除指定行（自动清理 element_id）
- `docx_delete_col(session_id, table_id, col_index)` - 删除指定列（自动清理 element_id）
- `docx_fill_table(session_id, data, table_id=None, start_row=0)` - 批量填充表格数据
- `docx_get_cell(session_id, table_id, row, col)` - 获取单元格
- `docx_insert_paragraph_to_cell(session_id, text, position)` - 向单元格添加段落（position 必选）

### 格式化

- `docx_set_properties(session_id, properties, element_id=None)` - 通用属性设置（JSON 格式）
- `docx_set_font(...)` - 设置字体属性（快捷方式）
- `docx_set_alignment(...)` - 设置对齐方式（快捷方式）
- `docx_set_margins(...)` - 设置页边距

### Cursor 定位系统

- `docx_cursor_get(session_id)` - 获取当前光标位置
- `docx_cursor_move(session_id, element_id, position)` - 移动光标到指定位置

### 特殊位置 ID（Special Position IDs）⭐ v2.3 新增

为了简化连续操作，系统提供了三个特殊的位置 ID，无需手动提取和传递元素 ID：

| 特殊 ID | 说明 | 使用场景 |
|---------|------|----------|
| `last_insert` | 最后一次插入操作创建的元素 ID | 连续插入内容时，无需提取上一个元素的 ID |
| `last_update` | 最后一次更新操作涉及的元素 ID | 格式复制、批量操作时引用刚修改的元素 |
| `cursor` | 当前光标位置的元素 ID | 与光标系统配合，实现基于光标的插入 |

**使用示例**：

```python
# 传统方式（需要提取 ID）
p1_resp = docx_insert_paragraph(session_id, "First", position="end:document_body")
p1_id = extract_element_id(p1_resp)  # 提取 ID
p2_resp = docx_insert_paragraph(session_id, "Second", position=f"after:{p1_id}")
p2_id = extract_element_id(p2_resp)  # 再次提取
p3_resp = docx_insert_paragraph(session_id, "Third", position=f"after:{p2_id}")

# 使用 last_insert（简化）
docx_insert_paragraph(session_id, "First", position="end:document_body")
docx_insert_paragraph(session_id, "Second", position="after:last_insert")
docx_insert_paragraph(session_id, "Third", position="after:last_insert")
# 无需提取 ID，代码更简洁
```

**格式复制示例**：

```python
# 创建源段落并格式化
p1_resp = docx_insert_paragraph(session_id, "", position="end:document_body")
p1_id = extract_element_id(p1_resp)
run1_resp = docx_insert_run(session_id, "Bold Text", position=f"inside:{p1_id}")
run1_id = extract_element_id(run1_resp)
docx_set_font(session_id, run1_id, bold=True, size=16)  # last_update = run1_id

# 创建目标段落
p2_resp = docx_insert_paragraph(session_id, "", position=f"after:{p1_id}")
p2_id = extract_element_id(p2_resp)
docx_insert_run(session_id, "Normal Text", position=f"inside:{p2_id}")  # last_insert = run2_id

# 使用特殊 ID 复制格式（无需手动传递 run1_id 和 run2_id）
docx_format_copy(session_id, source_id=run1_id, target_id="last_insert")
```

**注意事项**：

1. **初始化检查**：在使用特殊 ID 前，必须先执行相应的操作（如插入或更新），否则会返回错误
2. **作用域**：特殊 ID 在会话级别维护，跨工具调用有效
3. **更新时机**：
   - `last_insert`：在 `docx_insert_*` 系列工具成功后更新
   - `last_update`：在 `docx_update_*`、`docx_set_*` 系列工具成功后更新
   - `cursor`：通过 `docx_cursor_move` 显式更新
4. **错误类型**：使用未初始化的特殊 ID 会返回 `SpecialIdNotInitialized` 错误

## 使用示例

### 示例 1：创建和编辑文档

```python
import re

# 辅助函数：从 Markdown 响应中提取元素 ID
def extract_element_id(response):
    match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', response)
    return match.group(1) if match else None

# 辅助函数：从 Markdown 响应中提取 session ID
def extract_session_id(response):
    match = re.search(r'\*\*Session Id\*\*:\s*(\S+)', response)
    return match.group(1) if match else None

# 创建新文档
session_response = docx_create()
session_id = extract_session_id(session_response)

# 添加标题
heading_response = docx_insert_heading(
    session_id,
    "文档标题",
    position="end:document_body",
    level=1
)
heading_id = extract_element_id(heading_response)

# 添加段落
para_response = docx_insert_paragraph(
    session_id,
    "这是第一段内容",
    position=f"after:{heading_id}"
)
para_id = extract_element_id(para_response)

# 保存文档
docx_save(session_id, "/path/to/output.docx")
```

### 示例 2：加载并编辑文档（v3.0 新方式）

```python
# 方式 1: 使用 Launcher GUI 选择文件
# Launcher 会调用 POST /api/file/switch 设置活动文件

# 方式 2: 使用 CLI 参数启动服务器
# mcp-server-docx --transport combined --file /path/to/template.docx

# 创建会话（使用当前活动文件）
session_response = docx_create()
session_id = extract_session_id(session_response)

# 提取文档结构（智能识别标题、表格、段落）
structure_json = docx_extract_template_structure(session_id)
structure = json.loads(structure_json)

# 查看提取的元素
for element in structure["document_structure"]:
    if element["type"] == "table":
        print(f"表格: {element['headers']}")  # 自动检测的表头
    elif element["type"] == "heading":
        print(f"标题 {element['level']}: {element['text']}")
```

**v2.x 旧方式（已废弃）**:
```python
# ❌ 不再支持
session_response = docx_create(file_path="/path/to/template.docx")
```

输出格式：
```json
{
  "metadata": {
    "extracted_at": "2026-01-21T...",
    "docx_version": "0.1.3"
  },
  "document_structure": [
    {
      "type": "heading",
      "level": 1,
      "text": "章节标题",
      "style": {"font": "Arial", "size": 16, "bold": true}
    },
    {
      "type": "table",
      "rows": 5,
      "cols": 3,
      "header_row": 0,
      "headers": ["姓名", "年龄", "部门"],
      "style": {...}
    },
    {
      "type": "paragraph",
      "text": "段落内容",
      "style": {"font": "宋体", "size": 12, "alignment": "left"}
    }
  ]
}
```

### 示例 3：高级编辑功能

#### 3.1 模板填充（智能替换）

```python
# 确保已通过 Launcher 或 --file 参数设置活动文件
session_response = docx_create()
session_id = extract_session_id(session_response)

# 智能替换 {{name}} 占位符，即使它跨越了多个 Run
docx_replace_text(session_id, "{{name}}", "张三")
docx_replace_text(session_id, "{{date}}", "2026-01-20")

docx_save(session_id, "/path/to/result.docx")
```

#### 3.2 表格克隆与填充

```python
session_response = docx_create()
session_id = extract_session_id(session_response)

# 获取模板中的第一个表格
table_response = docx_get_table(session_id, 0)
table_id = extract_element_id(table_response)

# 克隆表格用于填充新数据
new_table_response = docx_copy_table(session_id, table_id, position="end:document_body")
new_table_id = extract_element_id(new_table_response)

# 批量填充数据
data = json.dumps([
    ["李四", "28", "工程师"],
    ["王五", "32", "设计师"]
])
docx_fill_table(session_id, data, table_id=new_table_id, start_row=1)

docx_save(session_id, "/path/to/report.docx")
```

## 开发指南

### 安装开发环境

```bash
./scripts/install.sh
source venv/bin/activate
```

### 运行测试

```bash
./scripts/test.sh
```

或手动运行：

```bash
# 单元测试
python -m pytest tests/unit/ -v

# E2E 测试
python -m pytest tests/e2e/ -v
```

### 项目结构

```
docx-editor-skill/
├── src/docx_editor_skill/
│   ├── server.py          # MCP 主入口
│   ├── tools/             # 工具模块（按领域拆分）
│   │   ├── __init__.py
│   │   ├── session_tools.py      # 会话生命周期
│   │   ├── content_tools.py      # 内容检索与浏览
│   │   ├── paragraph_tools.py    # 段落操作
│   │   ├── run_tools.py          # 文本块操作
│   │   ├── table_tools.py        # 表格操作
│   │   ├── format_tools.py       # 格式化与样式
│   │   ├── advanced_tools.py     # 高级编辑（替换、图片）
│   │   ├── cursor_tools.py       # 光标定位系统
│   │   ├── copy_tools.py         # 复制与元数据
│   │   └── system_tools.py       # 系统状态
│   ├── core/              # 核心逻辑
│   │   ├── session.py     # 会话管理
│   │   ├── cursor.py      # 光标系统
│   │   ├── copier.py      # 对象克隆引擎
│   │   ├── replacer.py    # 文本替换引擎
│   │   └── properties.py  # 属性设置引擎
│   ├── preview/           # 实时预览
│   └── utils/             # 工具函数
├── src/docx_server_launcher/ # Windows GUI 启动器
├── tests/
│   ├── unit/              # 单元测试
│   ├── e2e/               # 端到端测试
│   └── integration/       # 集成测试
├── docs/                  # 文档
├── config/                # 配置文件
├── scripts/               # 脚本工具
└── CLAUDE.md              # Claude 开发指南
```

## 许可证

MIT License

## 相关资源

- [MCP 协议规范](https://modelcontextprotocol.io)
- [python-docx 文档](https://python-docx.readthedocs.io)
