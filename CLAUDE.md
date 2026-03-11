# docx-editor-skill - Claude 开发指南

## 项目概述

docx-editor-skill 是一个 Claude Code Skill 插件，通过 `docx` CLI 工具为 Claude 提供细粒度的 Microsoft Word 文档操作能力。安装到 `~/.claude/skills/` 后，Claude Code 可通过 Bash 工具调用 `docx` 命令处理 Word 文档。

### 核心目标

- **CLI 架构**：通过 `docx <subcommand> <file>` 命令行接口操作文档
- **原子化操作**：每个子命令针对单一元素（段落、文本块、表格）
- **确定性 ID**：元素 ID（`para_001`, `table_001`）在文档结构不变时保持稳定
- **自动保存**：修改操作执行后自动保存文件
- **智能命令**：提供 4 个高层命令（edit/insert/format/copy）简化常见操作

### 架构特点

**纯 CLI 工具**：
- 不依赖 MCP 协议或服务器
- 每次调用独立执行：打开文档 → 操作 → 保存 → 输出结果
- Claude 通过 Bash 工具调用 `docx` 命令
- 支持管道和脚本组合

**两层命令设计**：
1. **智能命令**（4 个）：`edit`、`insert`、`format`、`copy` - 人类友好，自动推断意图
2. **原子命令**（40+ 个）：`insert-paragraph`、`set-font`、`fill-table` - 精确控制，适合自动化

## 核心架构

### 1. CLI Session 模型

```
每次 CLI 调用：
    ↓ docx <command> <file>
打开文档
    ↓ 构建 element_id 映射
执行操作
    ↓ 修改文档（如适用）
保存文档（如 --no-save 未设置）
    ↓ 输出 Markdown 结果
```

**关键特性**：
- 无状态：每次调用独立，不保留会话
- 自动 ID 映射：基于文档结构生成稳定的 element_id
- 自动保存：修改操作默认保存（可用 `--no-save` 禁用）

**代码位置**：`src/docx_editor_skill/cli.py`

### 2. 元素 ID 系统

python-docx 的对象（Paragraph、Run、Table）没有稳定 ID。CLI 工具在每次执行时自动生成确定性 ID：

```bash
# 查看文档结构和 ID
docx read document.docx

# 输出示例：
# para_001: "第一段"
# para_002: "第二段"
# table_001: 3x2 表格
```

**ID 生成规则**：
- 基于文档顺序：第一个段落是 `para_001`，第二个是 `para_002`
- 前缀约定：`para_*`（段落）、`run_*`（文本块）、`table_*`（表格）、`cell_*`（单元格）
- 确定性：只要文档结构不变，ID 保持稳定

**使用示例**：

```bash
# 1. 查看文档，获取 ID
docx read doc.docx

# 2. 使用 ID 操作元素
docx set-font doc.docx para_001 --bold --size 16
docx update-paragraph doc.docx para_002 "新文本"
docx delete doc.docx table_001
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

### 3. 命令输出格式

所有 CLI 命令输出 Markdown 格式，包含操作结果和文档上下文：

```bash
$ docx insert-paragraph doc.docx "Hello World"

# 操作结果: Insert Paragraph

**Status**: ✅ Success
**Element ID**: para_003
**File**: doc.docx

---

## 📄 Document Context

  ┌─────────────────────────────────────┐
  │ Paragraph (para_002)                │
  ├─────────────────────────────────────┤
  │ Previous paragraph                  │
  └─────────────────────────────────────┘

  ┌─────────────────────────────────────┐
  │ Paragraph (para_003) ⭐ NEW          │
  ├─────────────────────────────────────┤
  │ Hello World                         │
  └─────────────────────────────────────┘
```

**输出特点**：
- 人类可读的 Markdown 格式
- 包含 element_id 用于后续操作
- 显示文档上下文（周围元素）
- 错误信息清晰明确

### 4. 智能命令 vs 原子命令

**智能命令**（推荐人类使用）：

```bash
# 自动查找并编辑
docx edit doc.docx "old text" --new "new text" --bold

# 自动识别内容类型插入
docx insert doc.docx "# Heading"        # 识别为标题
docx insert doc.docx "image.png"        # 识别为图片
docx insert doc.docx "Normal text"      # 识别为段落

# 智能格式化
docx format doc.docx "关键词" --bold --color red

# 智能复制
docx copy doc.docx para_001 --after para_005
```

**原子命令**（精确控制）：

```bash
# 精确操作特定元素
docx insert-paragraph doc.docx "Text" --after para_001
docx set-font doc.docx run_001 --bold --size 14
docx fill-table doc.docx table_001 data.json

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

### 环境配置与安装

本项目使用 [uv](https://github.com/astral-sh/uv) 进行依赖管理。

```bash
# 1. 安装 uv
pip install uv

# 2. 安装项目依赖
uv pip install -e .

# 3. 全局安装 docx 命令（推荐）
uv tool install --editable .

# 或使用安装脚本
./install.sh

# 4. 验证安装
docx --help
```

### 使用 CLI

```bash
# 智能命令（推荐）
docx edit doc.docx "old" --new "new" --bold
docx insert doc.docx "# Heading"
docx format doc.docx "keyword" --color red
docx copy doc.docx para_001 --after para_005

# 查看文档结构
docx read doc.docx
docx find doc.docx "search text"
docx summary doc.docx

# 原子命令（精确控制）
docx insert-paragraph doc.docx "Text" --after para_001
docx set-font doc.docx run_001 --bold --size 14
docx fill-table doc.docx table_001 data.json
```

### 添加新 CLI 命令

1. **在 `src/docx_editor_skill/cli.py` 中添加命令处理函数**

```python
def _cmd_new_feature(session, args):
    """处理新功能命令"""
    from docx_editor_skill.tools.some_tools import docx_new_feature
    return docx_new_feature(
        param=args.param,
        # ... 其他参数
    )
```

2. **注册子命令**

在 `cli_main()` 函数中添加：

```python
# 在 subparsers 部分添加
p = sub.add_parser("new-feature", help="Description of new feature")
p.add_argument("param", help="Parameter description")
p.set_defaults(func=_cmd_new_feature, mutating=True)  # mutating=True 表示会修改文档
```

3. **编写单元测试**

在 `tests/unit/` 创建测试文件：

```python
def test_new_feature():
    # 测试逻辑
    pass
```

4. **更新文档**

- 在 `README.md` 的命令列表中添加
- 如果是重要功能，在本文件添加说明

### 测试策略

**单元测试**（`tests/unit/`）：
- 测试每个命令的基本功能
- 测试错误处理
- 测试边界条件

**E2E 测试**（`tests/e2e/`）：
- 模拟真实使用场景
- 测试命令组合使用
- 验证生成的 .docx 文件

**运行测试**：
```bash
# 安装测试依赖
uv pip install -e ".[dev]"

# 运行测试
uv run pytest

# 或运行脚本
./scripts/test.sh
```
### 调试技巧

1. **查看文档结构**

```bash
# 查看所有元素和 ID
docx read doc.docx

# 查看表格结构
docx table-structure doc.docx table_001
```

2. **日志配置**

设置环境变量启用调试日志：

```bash
export LOG_LEVEL=DEBUG
docx read doc.docx
```

## 快速参考

### 智能命令（推荐）

```bash
# 编辑：查找并替换文本
docx edit doc.docx "old text" --new "new text" --bold

# 插入：自动识别内容类型
docx insert doc.docx "# Heading"        # 识别为标题
docx insert doc.docx "image.png"        # 识别为图片
docx insert doc.docx "Normal text"      # 识别为段落

# 格式化：按文本或 ID 应用样式
docx format doc.docx "keyword" --bold --color red

# 复制：单个元素或范围
docx copy doc.docx para_001 --after para_005
```

### 常用查询命令

```bash
# 读取文档内容
docx read doc.docx

# 查找文本
docx find doc.docx "search keyword"

# 查看文档结构
docx summary doc.docx
```

---

**最后更新**：2026-03-11
**版本**：v3.0（纯 CLI 架构）
