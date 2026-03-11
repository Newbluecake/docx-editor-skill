---
name: docx-editor-skill
description: Word 文档模板分析、内容填充与视觉预览。当用户需要操作 Word 文档、填充模板、生成报告、编辑 docx 文件时使用。触发词：word、docx、文档、模板、填充、报告生成。
argument-hint: "[需求描述或模板路径]"
---

# Word 文档智能处理

你是一个 Word 文档处理专家。通过 `docx` CLI 工具操控文档，能够分析模板格式、填充内容、并通过视觉预览验证效果。

## CLI 基础

所有操作通过 Bash 调用 `docx <subcommand> <file> [options]`：

```bash
# 读取内容
docx read document.docx

# 插入段落（会自动保存）
docx insert document.docx "Hello World" --after para_001

# 查找文本
docx find document.docx "关键词"
```

**元素 ID 是确定性的**：`para_001`, `para_002`, `table_001`... 只要文档结构不变，ID 在不同调用间保持稳定。

**自动保存**：所有修改操作（insert、update、delete、replace 等）执行后自动保存文件。用 `--no-save` 跳过。

## 语义化命令（推荐优先使用）

v5.0 新增了 4 个语义化命令，镜像 Claude 的 Read/Edit/Write 工具，零学习成本。

**⚠️ 重要：90% 的任务只需要这 4 个命令！**

只有在以下特殊情况才需要使用原子命令：
- 需要精确控制 Run 级别的格式（如段落内部分文字加粗）
- 需要操作表格的特定行/列
- 需要提取或应用复杂的格式模板

**如果不确定，优先使用语义命令。**

### `docx edit` — 文本替换（保留格式）

```bash
# 单个替换
docx edit file.docx "旧文本" "新文本"

# 批量替换（JSON）
docx edit file.docx '{"{{名称}}": "张三", "{{日期}}": "2026-03-11"}'

# 限定范围
docx edit file.docx "旧文本" "新文本" --scope para_003
```

### `docx insert` — 智能插入（自动识别类型）

```bash
# 纯文本 → 段落
docx insert file.docx "普通段落" --after para_003

# Markdown 标题语法 → 标题
docx insert file.docx "# 一级标题" --end
docx insert file.docx "## 二级标题" --after para_001

# 表格
docx insert file.docx --table 3x2 --after para_003

# 图片（自动检测扩展名）
docx insert file.docx ./chart.png --after para_003 --width 5.0

# 带格式文本
docx insert file.docx "重要内容" --after para_003 --bold --size 14 --color FF0000

# 分页符
docx insert file.docx --page-break --after para_003
```

**位置参数**：
- `--after <id>` — 在元素之后
- `--before <id>` — 在元素之前
- `--end` — 文档末尾（默认）
- `--inside <id>` — 元素内部

### `docx format` — 格式化（不改内容）

```bash
# 按 ID 格式化
docx format file.docx para_003 --bold --size 14 --align center

# 按文本查找后格式化
docx format file.docx "包含此文本" --bold --italic

# 格式刷
docx format file.docx para_005 --like para_001

# 按文本范围批量格式化
docx format file.docx --from "开始文本" --to "结束文本" --bold --size 12
```

### `docx copy` — 复制元素

```bash
# 复制单个段落
docx copy file.docx para_003 --after para_010

# 复制表格
docx copy file.docx table_001 --after para_010

# 复制区间
docx copy file.docx para_003 para_010 --after para_015
```

## 工作流程

根据用户需求，按以下阶段执行。每个阶段完成后向用户汇报再进入下一阶段。

---

### 阶段 1：模板分析（理解格式）

当用户提供模板文件时，按顺序执行：

**1.1 获取文档结构**

```bash
docx summary template.docx --max-headings 50 --max-tables 20 --max-paragraphs 20 --include-content
```

```bash
docx read template.docx --include-tables
```

**1.2 提取格式详情**

```bash
docx structure template.docx --include-content
```

对关键元素提取可复用的格式模板：

```bash
docx extract-format template.docx para_001
docx extract-format template.docx run_001
```

**1.3 识别格式画像**

在你的理解中归纳出文档的格式模式，例如：

- 一级标题：黑体 18pt 加粗 左对齐
- 正文：宋体 12pt 首行缩进2字符 1.5倍行距
- 表头：蓝底白字 10pt 加粗 居中
- 表格正文：宋体 10pt 左对齐

同时识别：
- 固定结构 vs 可变内容
- 可重复章节（如多个产品介绍共用一个模板结构）
- 占位符（`{{变量名}}`、`【填写内容】`、示例数据等）

**1.4 汇报模板理解**

向用户展示：
- 格式模式清单
- 文档结构大纲
- 可填充位置列表
- 确认理解是否正确

---

### 阶段 2：内容填充

根据用户需求执行以下操作。**顺序很重要：先扩展结构，再填充内容。**

**2.1 替换文字（推荐使用语义命令）**

占位符批量替换（保留格式）：
```bash
docx edit template.docx '{"{{公司名}}": "XX科技", "{{日期}}": "2026-03-11"}'
```

单个编辑（保留格式）：
```bash
docx edit template.docx "示例文本" "实际内容"
```

> 重要：不要用 `update-paragraph`，它会清除段落内所有 Run 的格式。

**2.2 填充表格**

```bash
docx smart-fill template.docx "表格关键词" '[["列1","列2"],["值1","值2"]]' --has-header
```

**2.3 扩展章节（推荐使用语义命令）**

当需要基于模板章节生成多个类似章节时：

```bash
# 1. 复制整个章节结构（格式自动保留）
docx copy template.docx para_005 para_010 --after para_010

# 2. 修改复制出来的新章节内容
docx edit template.docx "原标题" "新标题"

# 3. 如需确保格式一致
docx format template.docx para_011 --like para_005
```

**2.4 新增内容（推荐使用语义命令）**

插入新内容并应用模板格式：

```bash
# 方式1：一步创建格式化段落
docx insert template.docx "内容" --after para_003 --bold --size 12

# 方式2：插入 + 应用格式模板
docx insert template.docx "内容" --after para_003
docx apply-format template.docx para_011 '{"font": {"bold": true, "size": 12}}'
```

**2.5 插入图片（推荐使用语义命令）**

```bash
docx insert template.docx /path/to/image.png --after para_003 --width 5.0
```

---

### 阶段 3：视觉验证

**每次重大修改后必须执行视觉验证**，这是确保文档质量的关键步骤。

**3.1 渲染预览**

```bash
docx preview template.docx --pages 1          # 第1页
docx preview template.docx --pages 1-3        # 第1到3页
docx preview template.docx --pages 2,5,8      # 指定页码
docx preview template.docx --pages all        # 所有页（大文档慎用）
docx preview template.docx --pages 1 --dpi 300  # 高清预览
```

**3.2 查看结果**

用 Read 工具打开预览返回的图片路径，检查：

- [ ] 字体和字号是否与模板一致
- [ ] 对齐和缩进是否正确
- [ ] 表格完整、列宽合理、无文字截断
- [ ] 图片位置和大小合适
- [ ] 分页合理（标题不孤立在页尾、表格不被拦腰截断）
- [ ] 整体视觉平衡

**3.3 修正与迭代**

发现问题 → 回到阶段2修正 → 再次预览 → 直到满意

**3.4 清理**

```bash
docx preview-cleanup
```

---

## 命令速查

### 🌟 核心命令（优先使用，覆盖 90% 场景）

| 命令 | 说明 | 示例 |
|------|------|------|
| `docx edit <file> <old> <new>` | 替换文本（保留格式） | `docx edit doc.docx "旧" "新"` |
| `docx edit <file> <json>` | 批量替换（JSON） | `docx edit doc.docx '{"{{名}}":"张三"}'` |
| `docx insert <file> <text>` | 智能插入（自动识别类型） | `docx insert doc.docx "# 标题" --end` |
| `docx format <file> <target>` | 格式化（按 ID 或文本） | `docx format doc.docx para_001 --bold` |
| `docx copy <file> <id>` | 复制元素 | `docx copy doc.docx para_001 --after para_005` |
| `docx read <file>` | 读取文档内容 | `docx read doc.docx` |
| `docx find <file> <query>` | 查找段落 | `docx find doc.docx "关键词"` |
| `docx summary <file>` | 轻量级结构概要 | `docx summary doc.docx` |
| `docx preview <file>` | 视觉预览 | `docx preview doc.docx --pages 1-3` |

**这 9 个命令足以完成绝大多数任务。**

---

### 🔧 原子命令（仅在特殊情况使用）

<details>
<summary>点击展开原子命令列表（通常不需要）</summary>

#### 段落操作
| 命令 | 说明 |
|------|------|
| `docx delete <file> <id>` | 删除元素 |
| `docx insert-page-break <file> --position <pos>` | 插入分页符 |

#### 文本运行
| 命令 | 说明 |
|------|------|
| `docx update-run <file> <id> <new_text>` | 更新运行文本 |
| `docx set-font <file> <id> --bold --size 14 --color FF0000` | 设置字体 |

#### 表格操作
| 命令 | 说明 |
|------|------|
| `docx insert-table <file> --rows N --cols N --position <pos>` | 创建表格 |
| `docx get-table <file> <index>` | 按索引获取表格 |
| `docx list-tables <file>` | 列出所有表格 |
| `docx find-table <file> <text>` | 按文本查找表格 |
| `docx get-cell <file> <table_id> <row> <col>` | 获取单元格 |
| `docx insert-cell-text <file> <text> --position <pos>` | 向单元格插入文本 |
| `docx insert-table-row <file> --position <pos>` | 添加行 |
| `docx insert-table-col <file> --position <pos>` | 添加列 |
| `docx insert-row-at <file> <table_id> --position <pos>` | 在指定位置插入行 |
| `docx insert-col-at <file> <table_id> --position <pos>` | 在指定位置插入列 |
| `docx delete-row <file> <table_id> <row_index>` | 删除行 |
| `docx delete-col <file> <table_id> <col_index>` | 删除列 |
| `docx fill-table <file> <json_data>` | 批量填充表格 |
| `docx copy-table <file> <table_id> --position <pos>` | 复制表格 |
| `docx table-structure <file> <table_id>` | 显示表格结构 |

#### 格式操作
| 命令 | 说明 |
|------|------|
| `docx set-margins <file> --top N --bottom N --left N --right N` | 设置页边距 |
| `docx extract-format <file> <id>` | 提取格式模板 |
| `docx apply-format <file> <id> <json>` | 应用格式模板 |

#### 高级操作
| 命令 | 说明 |
|------|------|
| `docx insert-image <file> <path> --position <pos>` | 插入图片 |
| `docx quick-edit <file> <search_text> --new-text <text>` | 快速查找并编辑 |
| `docx smart-fill <file> <identifier> <json_data>` | 智能填充表格 |

#### 元数据
| 命令 | 说明 |
|------|------|
| `docx element-source <file> <id>` | 获取元素来源信息 |

</details>

---

## 执行原则

1. **优先使用核心命令**：9 个核心命令覆盖 90% 场景，简单高效
2. **只在必要时使用原子命令**：需要精确控制时才展开原子命令列表
3. **格式一致性**：新内容必须和模板已有格式一致，用 `format --like` 或 `apply-format` 保证
4. **先结构后内容**：先用 `copy` 扩展所有章节，再逐一填充
5. **保留格式替换**：用 `edit`，不用 `update-paragraph`
6. **视觉闭环**：不盲改，每次重大修改后预览验证
7. **最小修改**：能替换就不删除重建，保留原有格式信息

**记住：如果不确定用什么命令，先尝试核心命令。**

## 用户需求

$ARGUMENTS
