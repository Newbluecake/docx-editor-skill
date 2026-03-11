# docx-editor-skill - Claude 开发指南

## 项目概述

docx-editor-skill 是一个基于 Model Context Protocol (MCP) 的服务器，为 Claude 提供细粒度的 Microsoft Word 文档操作能力。通过原子化的 API 设计，Claude 可以精确控制文档的每个元素。

### 核心目标

- **状态管理**：维护多个文档编辑会话，支持并发操作
- **原子化操作**：每个操作针对单一元素（段落、文本块、表格）
- **ID 映射系统**：将 python-docx 的内存对象映射为稳定的字符串 ID
- **MCP 协议兼容**：完全符合 MCP 规范，易于集成
- **模块化架构**：工具按领域拆分，便于维护和扩展

## 核心架构

### 1. Session 管理机制

```
Client (Claude)
    ↓ docx_create()
SessionManager
    ↓ 创建 UUID
Session {
    session_id: "abc-123"
    document: Document()
    object_registry: {}
    cursor: Cursor()      # 新增：光标位置管理
    last_accessed: timestamp
}
```

**关键特性**：
- 每个会话独立，互不干扰
- 自动过期机制（默认 1 小时）

**代码位置**：`src/docx_editor_skill/core/session.py`

### 2. 对象 ID 映射系统

这是本项目最关键的设计。python-docx 的对象（Paragraph、Run、Table）是临时的 Python 对象，没有稳定 ID。我们通过 `object_registry` 建立映射：

```python
# 创建段落时
paragraph = document.add_paragraph("Hello")
element_id = session.register_object(paragraph, "para")  # 返回 "para_a1b2c3d4"

# 后续操作时
paragraph = session.get_object("para_a1b2c3d4")
```

**ID 前缀约定**：
- `para_*` - 段落（Paragraph）
- `run_*` - 文本块（Run）
- `table_*` - 表格（Table）
- `cell_*` - 单元格（Cell）

#### 特殊位置 ID ⭐ v2.3 新增

为了简化连续操作，系统维护三个特殊的位置 ID：

| 特殊 ID | 维护位置 | 更新时机 | 用途 |
|---------|---------|---------|------|
| `last_insert` | `session.last_insert_id` | `docx_insert_*` 成功后 | 引用最后插入的元素 |
| `last_update` | `session.last_update_id` | `docx_update_*`、`docx_set_*` 成功后 | 引用最后更新的元素 |
| `cursor` | `session.cursor.element_id` | `docx_cursor_move` 调用后 | 引用光标位置 |

**实现要点**：

1. **解析特殊 ID**：在 `session.py` 的 `resolve_position()` 方法中处理
2. **更新时机**：在工具成功执行后，调用 `session.update_last_insert()` 或 `session.update_last_update()`
3. **错误处理**：使用前检查是否已初始化，未初始化返回 `SpecialIdNotInitialized` 错误
4. **测试覆盖**：确保单元测试和 E2E 测试覆盖所有特殊 ID 场景

**代码示例**：

```python
# 在工具中更新特殊 ID
def docx_insert_paragraph(session_id: str, text: str, position: str, style: str = None) -> str:
    session = session_manager.get_session(session_id)
    # ... 插入逻辑 ...
    para_id = session.register_object(paragraph, "para")

    # 更新 last_insert
    session.update_last_insert(para_id)

    return create_context_aware_response(session, "Paragraph inserted", element_id=para_id)

# 在 session.py 中解析特殊 ID
def resolve_position(self, position: str) -> Tuple[str, str]:
    if ":" not in position:
        raise ValueError(f"Invalid position format: {position}")

    relation, target = position.split(":", 1)

    # 解析特殊 ID
    if target == "last_insert":
        if not self.last_insert_id:
            raise ValueError("Special ID 'last_insert' not initialized")
        target = self.last_insert_id
    elif target == "last_update":
        if not self.last_update_id:
            raise ValueError("Special ID 'last_update' not initialized")
        target = self.last_update_id
    elif target == "cursor":
        if not self.cursor.element_id:
            raise ValueError("Special ID 'cursor' not initialized")
        target = self.cursor.element_id

    return relation, target
```

### 3. 原子化操作设计

每个工具只做一件事，避免复杂的组合参数：

```python
# 不好的设计（过于复杂）
docx_insert_formatted_paragraph(session_id, text, position="end:document_body", bold=True, size=14, alignment="center")

# 好的设计（原子化）
para_id = docx_insert_paragraph(session_id, "Hello", position="end:document_body")
run_id = docx_insert_run(session_id, text="World", position=f"inside:{para_id}")
docx_set_font(session_id, run_id, bold=True, size=14)
docx_set_alignment(session_id, para_id, "center")
```

### 4. 标准化 Markdown 响应格式 ⭐ v2.0 重大更新

**v2.0 重大更新**：所有工具现在返回 **Markdown 格式**的响应（不再是 JSON），包含结构化元数据和 ASCII 可视化。这使得响应更易读、更直观，同时保持了可解析性。

#### 响应结构

所有工具返回以下 Markdown 格式：

```markdown
# 操作结果: [Operation Name]

**Status**: ✅ Success  // 或 ❌ Error
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
```

#### 成功响应示例

```python
import re

# 创建段落
result = docx_insert_paragraph(session_id, "Hello World", position="end:document_body")

# 响应是 Markdown 格式：
# # 操作结果: Insert Paragraph
#
# **Status**: ✅ Success
# **Element ID**: para_abc123
# **Operation**: Insert Paragraph
# ...

# 提取元素 ID
match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', result)
para_id = match.group(1) if match else None
```

#### 错误响应示例

```python
# 尝试获取不存在的元素
result = docx_update_paragraph_text(session_id, "para_nonexistent", "New text")

# 响应是 Markdown 格式：
# # 操作结果: Error
#
# **Status**: ❌ Error
# **Error Type**: ElementNotFound
# **Message**: Paragraph para_nonexistent not found

# 检查错误
is_error = '**Status**: ❌ Error' in result
if is_error:
    # 提取错误类型
    match = re.search(r'\*\*Error Type\*\*:\s*(\w+)', result)
    error_type = match.group(1) if match else None
    # 根据错误类型处理
```

#### 错误类型分类

| 错误类型 | 说明 | 示例 |
|---------|------|------|
| `SessionNotFound` | 会话不存在或已过期 | 无效的 session_id |
| `ElementNotFound` | 元素 ID 不存在 | 引用已删除的段落 |
| `InvalidElementType` | 元素类型不匹配 | 对表格调用段落操作 |
| `ValidationError` | 参数验证失败 | 无效的对齐方式、颜色格式 |
| `FileNotFound` | 文件不存在 | 图片路径错误 |
| `CreationError` | 创建元素失败 | 内部错误 |
| `UpdateError` | 更新元素失败 | 内部错误 |
| `SpecialIdNotInitialized` | 特殊 ID 未初始化 | 在任何插入前使用 last_insert |

#### Agent 使用模式

```python
import re

# 辅助函数：提取元素 ID
def extract_element_id(response):
    match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', response)
    return match.group(1) if match else None

# 辅助函数：检查状态
def is_success(response):
    return '**Status**: ✅ Success' in response

def is_error(response):
    return '**Status**: ❌ Error' in response

# 1. 执行操作
result = docx_insert_paragraph(session_id, "Text", position="end:document_body")

# 2. 检查状态
if is_success(result):
    element_id = extract_element_id(result)
    # 继续操作
else:
    # 提取错误信息
    match = re.search(r'\*\*Message\*\*:\s*(.+?)(?:\n|$)', result)
    error_msg = match.group(1) if match else "Unknown error"
    # 错误处理逻辑

# 3. 获取上下文（如适用）
# 上下文信息已包含在 Markdown 响应的 Document Context 部分
```

#### 迁移指南

**旧代码（v1.x - JSON 格式）**：
```python
try:
    result = docx_insert_paragraph(session_id, "Text", position="end:document_body")
    data = json.loads(result)  # 解析 JSON
    para_id = data["data"]["element_id"]
except ValueError as e:
    print(f"Error: {e}")
```

**新代码（v2.0+ - Markdown 格式）**：
```python
import re

result = docx_insert_paragraph(session_id, "Text", position="end:document_body")  # 返回 Markdown

if is_success(result):
    para_id = extract_element_id(result)  # 使用正则提取
    # 可以直接查看 Markdown 响应中的可视化上下文
else:
    # 错误信息在 Markdown 中清晰展示
    match = re.search(r'\*\*Message\*\*:\s*(.+?)(?:\n|$)', result)
    error_msg = match.group(1) if match else "Unknown error"
```

**关键变化**：
1. **Markdown 格式**：响应现在是人类可读的 Markdown，包含 ASCII 可视化
2. **正则提取**：使用正则表达式提取元数据（element_id、status 等）
3. **错误分类**：error_type 字段便于自动化错误处理
4. **上下文可视化**：Document Context 部分提供直观的文档结构展示
5. **测试辅助**：`tests/helpers/markdown_extractors.py` 提供了提取函数

## 开发指南

### 环境配置与运行

本项目**必须**使用 [uv](https://github.com/astral-sh/uv) 进行依赖管理和任务执行。

```bash
# 1. 安装 uv
pip install uv

# 2. 安装项目依赖 (创建虚拟环境)
uv venv
uv pip install -e .[gui]
uv pip install pytest pytest-cov

# 3. 运行服务器 (多种模式)

# STDIO 模式 (默认)
uv run mcp-server-docx

# SSE 模式 (HTTP Server-Sent Events)
uv run mcp-server-docx --transport sse --host 127.0.0.1 --port 3000

# Streamable HTTP 模式
uv run mcp-server-docx --transport streamable-http --port 8080 --mount-path /mcp

# 查看所有选项
uv run mcp-server-docx --help

# 4. 运行 GUI 启动器
uv run docx-server-launcher
```

### 添加新工具

1. **在 `src/docx_editor_skill/tools/` 下的相应模块中定义工具**

   （如 `paragraph_tools.py`，或新建模块）

```python
from docx_editor_skill.core.response import (
    create_context_aware_response,
    create_error_response,
    create_success_response
)

def docx_new_feature(session_id: str, param: str) -> str:
    """
    工具描述（Claude 会读取这个）

    Args:
        session_id: 会话 ID
        param: 参数说明

    Returns:
        str: JSON 响应字符串
    """
    from docx_editor_skill.server import session_manager

    session = session_manager.get_session(session_id)
    if not session:
        return create_error_response(
            f"Session {session_id} not found",
            error_type="SessionNotFound"
        )

    try:
        # 实现逻辑
        result_obj = do_something()
        element_id = session.register_object(result_obj, "prefix")

        # 更新上下文
        session.update_context(element_id, action="create")

        # 返回标准化响应
        return create_context_aware_response(
            session,
            message="Operation completed successfully",
            element_id=element_id,
            # 可选：添加其他数据字段
            custom_field="value"
        )
    except Exception as e:
        logger.exception(f"Operation failed: {e}")
        return create_error_response(
            f"Operation failed: {str(e)}",
            error_type="CreationError"
        )
```

**响应格式化工具**：

| 函数 | 用途 | 返回 |
|------|------|------|
| `create_success_response(message, element_id=None, cursor=None, **extra)` | 创建成功响应 | JSON 字符串 |
| `create_error_response(message, error_type=None)` | 创建错误响应 | JSON 字符串 |
| `create_context_aware_response(session, message, element_id=None, **extra)` | 创建带光标上下文的响应 | JSON 字符串 |

**最佳实践**：
- 优先使用 `create_context_aware_response`（自动包含光标信息）
- 所有错误通过 `create_error_response` 返回，不要抛出异常
- 使用明确的 `error_type` 便于自动化处理
- 在 `extra_data` 中添加操作特定的字段（如 `changed_fields`、`replacements` 等）
```

2. **注册工具**

   - 如果是现有模块，无需额外操作（已自动扫描）
   - 如果是新模块，需要在 `src/docx_editor_skill/tools/__init__.py` 中注册：

```python
from . import new_module
new_module.register_tools(mcp)
```

3. **编写单元测试**

在 `tests/unit/` 创建测试文件：

```python
def test_new_feature():
    session_id = docx_create()
    result = docx_new_feature(session_id, "test")
    assert "expected" in result
```

4. **更新文档**

- 在 `README.md` 的工具列表中添加
- 如果是重要功能，在本文件添加说明

### 测试策略

**单元测试**（`tests/unit/`）：
- 测试每个工具的基本功能
- 测试错误处理（无效 session_id、element_id）
- 测试边界条件

**E2E 测试**（`tests/e2e/`）：
- 模拟真实使用场景
- 测试工具组合使用
- 验证生成的 .docx 文件

**运行测试**：
```bash
# 1. 安装测试依赖 (含 GUI 和测试工具)
uv pip install -e ".[gui,dev]"

# 2. 运行测试
# 注意：在无头 Linux 环境需指定 QT_QPA_PLATFORM=offscreen
QT_QPA_PLATFORM=offscreen uv run pytest

# 或运行脚本
./scripts/test.sh
```

### 调试技巧

1. **查看会话状态**

```python
# 临时添加调试工具
@mcp.tool()
def docx_debug_session(session_id: str) -> str:
    session = session_manager.get_session(session_id)
    return f"Objects: {list(session.object_registry.keys())}"
```

2. **日志配置**

修改 `config/dev.yaml` 中的日志级别为 `DEBUG`。错误日志会自动包含堆栈跟踪（Stack Trace）。

## MCP 协议注意事项

### 1. 工具命名规范

- 使用 `docx_` 前缀，避免与其他 MCP 服务器冲突
- 使用动词开头：`add`、`set`、`get`
- 保持简洁：`docx_insert_paragraph` 而非 `docx_insert_paragraph_to_document`

### 2. 错误处理

始终使用明确的错误消息，利用 `logger.exception` 记录完整堆栈：

```python
try:
    # ...
except Exception as e:
    logger.exception(f"Operation failed: {e}")
    raise ValueError(f"User friendly error message: {e}")
```

### 3. 会话生命周期

```
创建 → 操作 → 保存
  ↓      ↓      ↓
create  add_*  save
        set_*
```

**注意**：在 v4.0 架构中，会话会自动管理，无需显式关闭。会话会在 1 小时后自动过期。

## 快速参考

### ⚡ 优化后的工作流（推荐）

**v2.0 新增了复合工具，大幅简化常见操作。优先使用这些工具！**

**v3.0 Breaking Change**: 文件管理已迁移到全局单文件模式。详见下方说明。

### v3.0 文件管理新架构

**关键变更**:
- ❌ `docx_create(file_path=...)` 参数已移除
- ❌ `docx_list_files()` 工具已移除
- ✅ 使用 Launcher GUI 或 `--file` CLI 参数设置活动文件
- ✅ HTTP API (`/api/file/switch`) 用于文件切换

**新工作流**:
```python
# 方式 1: 通过 Launcher GUI 选择文件
# Launcher 会调用 POST /api/file/switch 设置活动文件

# 方式 2: 启动服务器时指定文件
# mcp-server-docx --transport combined --file /path/to/document.docx

# 创建会话（使用全局活动文件）
session_id = docx_create()
```

**创建格式化文档（新方式）**：
```python
session_id = docx_create()
# 一步创建格式化段落
para_id = docx_insert_formatted_paragraph(
    session_id, "重要文本", position="end:document_body",
    bold=True, size=16, color_hex="FF0000", alignment="center"
)
docx_save(session_id, "/path/to/output.docx")
```

**快速编辑文档（新方式）**：
```python
# 确保已通过 Launcher 或 --file 设置活动文件
session_id = docx_create()
# 查找并编辑，一步完成
result = docx_quick_edit(
    session_id, "old text",
    new_text="new text", bold=True
)
docx_save(session_id, "/path/to/doc.docx")
```

**轻量级结构提取（新方式）**：
```python
# 确保已通过 Launcher 或 --file 设置活动文件
session_id = docx_create()
# 只返回标题和表格，不返回段落内容
summary = docx_get_structure_summary(
    session_id,
    max_headings=10,
    max_tables=5,
    max_paragraphs=0
)
# Token 使用减少 90%
```

**智能表格填充（新方式）**：
```python
# 确保已通过 Launcher 或 --file 设置活动文件
session_id = docx_create()
data = json.dumps([
    ["Name", "Age", "City"],
    ["Alice", "30", "NYC"],
    ["Bob", "25", "LA"]
])
# 自动扩展行，无需手动计算
result = docx_smart_fill_table(
    session_id, "Employee",  # 通过文本查找表格
    data, has_header=True, auto_resize=True
)
docx_save(session_id, "/path/to/output.docx")
```

### 常用工具组合（原子操作）

**提取模板结构**：
```python
# v3.0: 确保已通过 Launcher 或 --file 设置活动文件
session_id = docx_create()
structure_json = docx_extract_template_structure(session_id)
structure = json.loads(structure_json)
```

**创建格式化文档（原子方式）**：
```python
session_id = docx_create()
para_id = docx_insert_paragraph(session_id, "", position="end:document_body")
run_id = docx_insert_run(session_id, "重要文本", position=f"inside:{para_id}")
docx_set_font(session_id, run_id, bold=True, size=16, color_hex="FF0000")
docx_save(session_id, "/path/to/output.docx")
```

**Cursor 定位系统（高级插入）**：
```python
# 1. 移动光标到指定元素之后
docx_cursor_move(session_id, element_id="para_123", position="after")

# 2. 使用 position 进行插入（不再依赖 cursor）
docx_insert_paragraph(session_id, "这是插入在中间的段落", position="after:para_123")
docx_insert_table(session_id, rows=3, cols=2, position="after:para_123")
```

**格式刷（Format Painter）**：
```python
# 将源对象（如 Run, Paragraph, Table）的格式复制到目标对象
docx_format_copy(session_id, source_id="run_src", target_id="run_target")
```

**创建表格**：
```python
table_id = docx_insert_table(session_id, rows=3, cols=2, position="end:document_body")
cell_id = docx_get_cell(session_id, table_id, row=0, col=0)
docx_insert_paragraph_to_cell(session_id, "单元格内容", position=f"inside:{cell_id}")
```

## 完整工具参考

本服务器提供 50 个 MCP 工具（v4.0 移除 docx_close），按功能领域分为 11 个模块：

### 0. Composite Tools（复合工具，5 个）⭐ 新增

**这些是最常用的高层工具，优先使用！**

| 工具 | 说明 | 效果 |
|------|------|------|
| `docx_insert_formatted_paragraph(session_id, text, position, bold, italic, size, color_hex, alignment, style)` | 一步创建格式化段落 | 4 次调用 → 1 次 |
| `docx_quick_edit(session_id, search_text, new_text, bold, italic, size, color_hex)` | 查找并编辑段落 | N+1 次 → 1 次 |
| `docx_get_structure_summary(session_id, max_headings, max_tables, max_paragraphs, include_content)` | 轻量级结构提取 | Token 减少 90% |
| `docx_smart_fill_table(session_id, table_identifier, data, has_header, auto_resize)` | 智能表格填充 | 自动扩展行 |
| `docx_format_range(session_id, start_text, end_text, bold, italic, size, color_hex)` | 批量格式化范围 | 批量操作 |

### 1. Session Tools（会话管理，3 个）

| 工具 | 说明 |
|------|------|
| `docx_create(auto_save=False)` | 创建新会话（⚠️ v3.0: 移除 file_path 参数）|
| `docx_save(session_id, file_path)` | 保存文档到指定路径 |
| `docx_get_context(session_id)` | 获取会话上下文信息 |

**v3.0 变更**: 文件选择由 Launcher GUI 或 `--file` CLI 参数管理，通过全局 `active_file` 状态共享。
**v4.0 变更**: 移除 `docx_close()` 接口，会话自动管理，1 小时后自动过期。

### 2. Content Tools（内容检索，3 个）⭐ 已优化

| 工具 | 说明 | 新增参数 |
|------|------|----------|
| `docx_read_content(session_id, max_paragraphs, start_from, include_tables)` | 读取文档全文 | 支持分页 |
| `docx_find_paragraphs(session_id, query, max_results, return_context)` | 查找包含指定文本的段落 | 限制结果数 |
| `docx_extract_template_structure(session_id, max_depth, include_content, max_items_per_type)` | 提取文档结构 | 可控详细程度 |

**v3.0 移除**: `docx_list_files()` 已移除，文件浏览由 Launcher GUI 提供。

### 3. Paragraph Tools（段落操作，6 个）

| 工具 | 说明 |
|------|------|
| `docx_insert_paragraph(session_id, text, position, style=None)` | 添加段落（position 必选） |
| `docx_insert_heading(session_id, text, position, level=1)` | 添加标题（position 必选） |
| `docx_update_paragraph_text(session_id, paragraph_id, new_text)` | 更新段落文本 |
| `docx_copy_paragraph(session_id, paragraph_id, position)` | 深拷贝段落（保留格式） |
| `docx_delete(session_id, element_id=None)` | 删除元素 |
| `docx_insert_page_break(session_id, position)` | 插入分页符（position 必选） |

### 4. Run Tools（文本块操作，3 个）

| 工具 | 说明 |
|------|------|
| `docx_insert_run(session_id, text, position)` | 向段落添加文本块（position 必选） |
| `docx_update_run_text(session_id, run_id, new_text)` | 更新 Run 文本 |
| `docx_set_font(session_id, run_id, size=None, bold=None, italic=None, color_hex=None)` | 设置字体属性 |

### 5. Table Tools（表格操作，13 个）⭐ v2.2 新增 4 个

| 工具 | 说明 |
|------|------|
| `docx_insert_table(session_id, rows, cols, position)` | 创建表格（position 必选） |
| `docx_get_table(session_id, index)` | 按索引获取表格 |
| `docx_find_table(session_id, text)` | 查找包含指定文本的表格 |
| `docx_get_cell(session_id, table_id, row, col)` | 获取单元格 |
| `docx_insert_paragraph_to_cell(session_id, text, position)` | 向单元格添加段落（position 必选） |
| `docx_insert_table_row(session_id, position)` | 添加行到表格末尾（position 必选） |
| `docx_insert_table_col(session_id, position)` | 添加列到表格末尾（position 必选） |
| `docx_insert_row_at(session_id, table_id, position, row_index, copy_format)` | ⭐ 在指定位置插入行（支持 after:N, before:N, start, end） |
| `docx_insert_col_at(session_id, table_id, position, col_index, copy_format)` | ⭐ 在指定位置插入列（支持 after:N, before:N, start, end） |
| `docx_delete_row(session_id, table_id, row_index)` | ⭐ 删除指定行（自动清理 element_id） |
| `docx_delete_col(session_id, table_id, col_index)` | ⭐ 删除指定列（自动清理 element_id） |
| `docx_fill_table(session_id, data, table_id=None, start_row=0)` | 批量填充表格数据 |
| `docx_copy_table(session_id, table_id, position)` | 深拷贝表格 |

### 6. Format Tools（格式化，6 个）

| 工具 | 说明 |
|------|------|
| `docx_set_alignment(session_id, paragraph_id, alignment)` | 设置段落对齐方式 |
| `docx_set_properties(session_id, properties, element_id=None)` | 通用属性设置（JSON） |
| `docx_format_copy(session_id, source_id, target_id)` | 复制格式（格式刷） |
| `docx_set_margins(session_id, top=None, bottom=None, left=None, right=None)` | 设置页边距 |
| `docx_extract_format_template(session_id, element_id)` | 提取格式模板 |
| `docx_apply_format_template(session_id, element_id, template_json)` | 应用格式模板 |

### 7. Advanced Tools（高级编辑，3 个）

| 工具 | 说明 |
|------|------|
| `docx_replace_text(session_id, old_text, new_text, scope_id=None)` | 智能文本替换（跨 Run） |
| `docx_batch_replace_text(session_id, replacements_json, scope_id=None)` | 批量文本替换 |
| `docx_insert_image(session_id, image_path, width=None, height=None, position)` | 插入图片（position 必选） |

### 8. Cursor Tools（光标定位，4 个）

| 工具 | 说明 |
|------|------|
| `docx_cursor_get(session_id)` | 获取当前光标位置 |
| `docx_cursor_move(session_id, element_id, position)` | 移动光标到指定位置 |

### 9. Copy Tools（复制与元数据，2 个）

| 工具 | 说明 |
|------|------|
| `docx_get_element_source(session_id, element_id)` | 获取元素来源元数据 |
| `docx_copy_elements_range(session_id, start_id, end_id, position)` | 复制元素区间 |

### 10. System Tools（系统状态，1 个）

| 工具 | 说明 |
|------|------|
| `docx_server_status()` | 获取服务器状态和环境信息 |

### 工具设计原则

1. **原子化操作**：每个工具只做一件事
2. **ID 映射系统**：所有对象通过稳定 ID 引用
3. **混合上下文**：支持显式 ID 和隐式上下文
4. **格式保留**：高级操作保留原始格式
5. **标准化响应**：所有工具返回 JSON 格式，包含状态、消息和数据 ⭐ 新增

详细参数和示例请参考 [README.md](../README.md) 的工具列表部分。

---

**最后更新**：2026-01-22

**v2.1 更新日志**：
- 所有工具现在返回标准化 JSON 响应（`{status, message, data}`）
- 错误通过 JSON 返回，不再抛出异常
- 自动包含光标上下文信息（`cursor.context`）
- 新增响应格式化工具（`create_context_aware_response` 等）
- 错误类型分类（`SessionNotFound`、`ElementNotFound` 等）
