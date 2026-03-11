# docx-editor-skill

A Claude Code Skill plugin for fine-grained Microsoft Word (.docx) document manipulation. Analyze templates, fill content, and visually verify results — all through natural language.

## Installation

### As a Claude Code Skill (Recommended)

Clone this repo into your Claude Code skills directory:

```bash
git clone <repo-url> ~/.claude/skills/docx-editor-skill
```

Then install the CLI dependency:

```bash
cd ~/.claude/skills/docx-editor-skill
uv pip install -e .
```

> **Prerequisite**: [uv](https://github.com/astral-sh/uv) must be installed (`pip install uv`).

### Verify Installation

```bash
docx --help
```

## What It Does

Once installed, Claude Code gains the `/docx` skill which enables:

- **Template Analysis** — Read document structure, identify headings, tables, placeholders, and formatting patterns
- **Content Filling** — Replace placeholders, fill tables, insert paragraphs/headings/images with precise formatting
- **Visual Verification** — Render document pages as images to verify layout before delivery
- **Format Preservation** — All edits preserve original formatting through smart text replacement

## Usage

Invoke the skill in Claude Code:

```
/docx analyze template.docx
/docx fill the template with the data from data.csv
/docx replace {{company}} with "Acme Corp" in report.docx
```

Or just describe what you need — Claude will use the skill automatically when it detects Word document tasks.

## CLI Commands

The `docx` CLI tool supports these subcommands:

| Category | Commands |
|----------|----------|
| **Read** | `read`, `find`, `structure`, `summary` |
| **Paragraphs** | `insert-paragraph`, `insert-heading`, `update-paragraph`, `copy-paragraph`, `delete` |
| **Runs** | `insert-run`, `update-run`, `set-font` |
| **Tables** | `insert-table`, `get-table`, `find-table`, `fill-table`, `smart-fill` |
| **Format** | `set-alignment`, `format-copy`, `extract-format`, `apply-format` |
| **Advanced** | `replace-text`, `batch-replace`, `insert-image`, `copy-range` |
| **Composite** | `insert-formatted`, `quick-edit`, `format-range` |
| **Preview** | `preview`, `preview-cleanup` |

Run `docx <command> --help` for detailed usage.

## Development

```bash
# Install dev dependencies
uv pip install -e ".[dev]"

# Run tests
uv run pytest tests/unit tests/e2e

# Run with QT offscreen (headless Linux)
QT_QPA_PLATFORM=offscreen uv run pytest
```

See [CLAUDE.md](CLAUDE.md) for the full developer guide.

## License

MIT License
