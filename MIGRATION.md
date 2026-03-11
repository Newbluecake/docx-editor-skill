# 命令迁移指南

## 概述

v6.0 版本对 CLI 命令进行了精简，从 51 个命令减少到 38 个（精简 25.5%）。所有被删除的命令都有智能命令完全替代，功能完全保留。

## 为什么精简？

1. **降低学习曲线**：更少的命令意味着更容易学习
2. **更清晰的语义**：智能命令更直观（`insert "# Title"` vs `insert-heading`）
3. **减少选择困难**：不再纠结用哪个命令
4. **保持功能完整**：所有功能通过智能命令保留

## 迁移对照表

### 段落操作（4 个命令）

| 旧命令 | 新命令 | 说明 |
|--------|--------|------|
| `docx insert-paragraph doc.docx "text" --after para_001` | `docx insert doc.docx "text" --after para_001` | 智能命令自动识别普通文本 |
| `docx insert-heading doc.docx "Title" --level 1 --after para_001` | `docx insert doc.docx "# Title" --after para_001` | 使用 Markdown 语法 |
| `docx update-paragraph doc.docx para_001 "new text"` | `docx edit doc.docx "old text" "new text"` | 保留格式的文本替换 |
| `docx copy-paragraph doc.docx para_001 --after para_005` | `docx copy doc.docx para_001 --after para_005` | 智能命令自动识别元素类型 |

### 文本运行（1 个命令）

| 旧命令 | 新命令 | 说明 |
|--------|--------|------|
| `docx insert-run doc.docx "text" --inside para_001` | `docx insert doc.docx "text" --inside para_001` | 智能命令支持 --inside 参数 |

### 格式化（3 个命令）

| 旧命令 | 新命令 | 说明 |
|--------|--------|------|
| `docx set-alignment doc.docx para_001 center` | `docx format doc.docx para_001 --align center` | 统一的格式化命令 |
| `docx set-properties doc.docx para_001 '{"bold": true}'` | `docx format doc.docx para_001 --bold` | 更简洁的参数 |
| `docx format-copy doc.docx para_001 para_002` | `docx format doc.docx para_002 --like para_001` | 格式刷功能 |

### 高级操作（4 个命令）

| 旧命令 | 新命令 | 说明 |
|--------|--------|------|
| `docx replace-text doc.docx "old" "new"` | `docx edit doc.docx "old" "new"` | 统一的编辑命令 |
| `docx batch-replace doc.docx '{"old":"new"}'` | `docx edit doc.docx '{"old":"new"}'` | 智能命令支持 JSON |
| `docx insert-formatted doc.docx "text" --bold --size 14 --after para_001` | `docx insert doc.docx "text" --bold --size 14 --after para_001` | 智能命令支持格式参数 |
| `docx format-range doc.docx --from "start" --to "end" --bold` | `docx format doc.docx --from "start" --to "end" --bold` | 智能命令支持范围 |

### 复制操作（1 个命令）

| 旧命令 | 新命令 | 说明 |
|--------|--------|------|
| `docx copy-range doc.docx para_001 para_005 --after para_010` | `docx copy doc.docx para_001 para_005 --after para_010` | 智能命令支持范围复制 |

## 迁移示例

### 示例 1：创建格式化文档

**旧方式**（使用原子命令）：
```bash
docx insert-paragraph doc.docx "" --end
docx insert-run doc.docx "重要文本" --inside para_003
docx set-font doc.docx run_001 --bold --size 16 --color FF0000
docx set-alignment doc.docx para_003 center
```

**新方式**（使用智能命令）：
```bash
docx insert doc.docx "重要文本" --bold --size 16 --color FF0000 --end
docx format doc.docx para_003 --align center
```

### 示例 2：批量替换文本

**旧方式**：
```bash
docx batch-replace doc.docx '{
  "{{公司名}}": "XX科技",
  "{{日期}}": "2026-03-12"
}'
```

**新方式**（完全相同）：
```bash
docx edit doc.docx '{
  "{{公司名}}": "XX科技",
  "{{日期}}": "2026-03-12"
}'
```

### 示例 3：复制并修改章节

**旧方式**：
```bash
docx copy-paragraph doc.docx para_005 --after para_010
docx update-paragraph doc.docx para_011 "新标题"
docx format-copy doc.docx para_005 para_011
```

**新方式**：
```bash
docx copy doc.docx para_005 --after para_010
docx edit doc.docx "旧标题" "新标题"
docx format doc.docx para_011 --like para_005
```

## 常见问题

### Q: 旧命令还能用吗？

A: 不能。v6.0 版本已经移除了这些命令。请使用智能命令替代。

### Q: 智能命令会影响性能吗？

A: 不会。智能命令的性能与原子命令相同，只是提供了更友好的接口。

### Q: 如何快速学习新命令？

A: 只需记住 4 个智能命令：
- `edit` - 编辑文本
- `insert` - 插入内容
- `format` - 格式化
- `copy` - 复制元素

这 4 个命令覆盖 90% 的使用场景。

### Q: 原子命令还有吗？

A: 有。表格操作、页边距设置等原子命令仍然保留，用于需要精确控制的场景。

### Q: 如何查看所有可用命令？

A: 使用 `docx --help-all` 查看完整命令列表。

## 获取帮助

- 查看简化帮助：`docx --help`
- 查看完整帮助：`docx --help-all`
- 查看特定命令帮助：`docx <command> --help`
- 查看文档：`SKILL.md`, `CLAUDE.md`, `README.md`

## 反馈

如果您在迁移过程中遇到问题，请在 GitHub 提交 Issue。
