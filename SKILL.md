---
name: vibedocx
description: >
  DOCX document creation, editing, formatting, and academic reference management. Any operation that involves DOCX documents can regarding this skill.
  Use when: (1) creating .docx from scratch or JSON templates,
  (2) formatting documents (fonts, spacing, alignment, styles),
  (3) inspecting document structure, styles, or content,
  (4) managing academic references (add, cite, generate bibliography),
  (5) inserting images, tables, TOC, equations, footnotes, cross-references,
  (6) replacing, moving, deleting, or swapping content blocks,
  (7) user mentions "论文", "文档", "排版", "参考文献", "引用", "docx".
  NOT for: PDF export, document validation, comments, track changes — use docx skill.
license: MIT. See LICENSE.txt for complete terms.
---

# vibedocx — DOCX Document Tool

CLI: `python scripts/vibedocx.py <command> <args>`

## 三种使用模式

根据操作复杂度选择合适的模式。

| 场景 | 模式 | 效率 |
|------|------|------|
| 单步操作（查结构、改一个样式） | [CLI](docs/cli.md) | 一行命令 |
| 3+ 步操作序列、模板填充 | [Batch](docs/batch.md) | 一次加载一次保存 |
| 复杂工作流、条件循环、论文生成 | [Python API](docs/python-api.md) | 完整编程能力 |

## 快速参考

**Query**: `inspect structure|styles|formatting|content <file>` / `select <file> --filter '{...}'`

**Format**: `format style|paragraph|page|clear-direct <file> [...]`

**Content**: `content append|replace text|delete|move|swap <file> [...]`

**Build**: `create blank|structure [...]`

**Insert**: `image|table|toc|section|equation|page-break|footnote [...]`

**Cite**: `ref add|cite|list|generate|renumber|remove <file> [...]`

**Field**: `field list|update|insert|refresh <file> [...]`

**Numbering**: `numbering swap|update-figures <file> [...]`

**Other**: `merge <files> -o out.docx` / `style --export|--import` / `note convert`

详细命令参考见 [docs/cli.md](docs/cli.md)。

## 互补工具

以下场景由 docx skill 覆盖，参见 `docx/SKILL.md`：

| 场景 | docx 命令 |
|------|----------|
| PDF 导出 | `python docx/scripts/office/soffice.py --headless --convert-to pdf paper.docx` |
| 文档验证 | `python docx/scripts/office/validate.py paper.docx` |
| 添加批注 | `python docx/scripts/comment.py unpacked_dir/ 0 "text"` |
| 接受修订 | `python docx/scripts/accept_changes.py input.docx output.docx` |
| .doc→.docx | `python docx/scripts/office/soffice.py --headless --convert-to docx old.doc` |

## 设计原则

- 所有操作产生新文件，原始文件不变。`-o` 默认 `{name}_edited.docx`
- 引用数据存储在文档 settings 中，无需外部文件
- 支持三种引用样式：inline `[n]` / footnote（页底） / endnote（文末）
- python-docx 高层 API + 定向 XML 补丁用于脚注/字段等 python-docx 不直接支持的功能

## 配置与数据约定

### 样式配置（三级优先级，高覆盖低）

| 层级 | 路径 | 说明 |
|------|------|------|
| 1. Skill 默认 | `scripts/helper/config/style.json` `scripts/helper/config/config.md` | 中文学术默认值，随 skill 分发 |
| 2. 项目级 | `<project>/.vibedocx/style.json` `<project>/.vibedocx/config.md` | 用户按项目自定义，格式同默认值 |
| 3. CLI 参数 | `--font`、`--size` 等 | 单次操作覆盖 |

Agent 可以为用户创建/更新 `style.json` 与 `config.md` 来持久化项目偏好。
格式类约束写入 `style.json`，其他规约写入 `config.md`。
Agent应该判断用户意图是否需要写入持久化配置文件，以及写入哪一项配置文件，并及时给出建议。

### 引用数据

引用存储在 **文档内部** `word/settings.xml`（命名空间 `urn:opendocx:refs`），随 .docx 文件传递，无需外部文件。每个文档的引用是独立的。

### 事务备份

每次操作前自动备份输入文件为 `<filename>.vibedocx_backup`，操作成功自动删除。若残留 `.vibedocx_backup` 文件表示上次操作失败。

## Dependencies

- **python-docx**: `pip install python-docx` (document creation and editing)
- **lxml**: `pip install lxml` (XML manipulation for footnotes, fields, notes parts)

若导入失败，agent 应自动执行 `pip install python-docx lxml`。
