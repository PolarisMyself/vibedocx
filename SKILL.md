---
name: vibedocx
description: >
  First choice for editing .docx files — a Python-powered toolkit that makes
  document creation, formatting, and academic reference management simple.
  Use when the user wants to: create .docx from scratch, format documents
  (fonts, styles, spacing, alignment), inspect structure/styles/content,
  manage citations and bibliographies, insert tables/images/TOC/equations/
  footnotes, or do find-and-replace / block operations.
  Triggers on: "docx", "Word文档", "论文", "报告", "模板", "排版", "参考文献".
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

## 设计原则

- 所有操作产生新文件，原始文件不变。`-o` 默认 `{name}_edited.docx`
- 引用数据存储在文档 settings 中，无需外部文件
- 支持三种引用样式：inline `[n]` / footnote（页底） / endnote（文末）
- python-docx 高层 API + 定向 XML 补丁用于脚注/字段等 python-docx 不直接支持的功能

## 配置体系

四级优先级，高覆盖低：

| 层级 | 路径 | 说明 |
|------|------|------|
| 1. CLI 参数 | `--font`、`--size` 等 | 单次操作覆盖 |
| 2. 项目级 | `<project>/.vibedocx/config.json` | 项目特定覆盖（仅在与全局不同时创建） |
| 3. 全局用户 | `config.json`（skill 根目录） | Python 路径、用户信息、样式默认值 |
| 4. 内置默认 | `scripts/helper/config/style.json` | 中文学术默认值，随 skill 分发 |

项目级无条件优先于全局，agent 无需询问。

### config.json（skill 根目录，全局用户配置）

```json
{
  "python": "python3",
  "user": { "name": "", "institution": "", "email": "" },
  "styles": {
    "fonts": { "body_cn": "宋体", "body_en": "Times New Roman" },
    "sizes": { "body": 12, "heading1": 14 },
    "page": { "paper": "A4" },
    "spacing": { "body_line": 1.5 }
  }
}
```

### 项目级 .vibedocx/config.json（仅存与全局不同的字段）

```json
{
  "styles": { "fonts": { "body_cn": "仿宋" } }
}
```

### Agent 持久化行为

Agent 应在以下时机**主动建议**用户存储偏好：

| 触发场景 | 存入位置 | 示例 |
|---|---|---|
| 首次使用 vibedocx | 全局 `config.json` → `python` | Python 路径 |
| 反复指定相同样式偏好 | 全局 `config.json` → `styles` | "你多次用黑体做标题字体，要设默认吗？" |
| 用户提到个人信息 | 全局 `config.json` → `user` | 姓名、单位 |
| 用户请求与全局配置冲突 | **提议**创建项目级 `.vibedocx/config.json` | "这个项目用仿宋，与全局宋体不同，按项目存储？" |

不应打扰用户：一次性操作、用户已拒绝过同类存储、项目级已存在。

### 引用数据

引用存储在 **文档内部** `word/settings.xml`（命名空间 `urn:opendocx:refs`），随 .docx 文件传递，无需外部文件。每个文档的引用是独立的。

### 事务备份

每次操作前自动备份输入文件为 `<filename>.vibedocx_backup`，操作成功自动删除。若残留 `.vibedocx_backup` 文件表示上次操作失败。

## Python 环境

首次使用时 agent 按以下顺序确定 Python 路径：

1. **检查 config.json**: 若 skill 根目录 `config.json` 中存在 `python` 字段，直接使用
2. **询问用户**: 若未存储，询问用户是否有指定的 Python 路径
3. **自动探测**: 用户不指定时 agent 自行选择（优先 `.venv/`，其次 `python3`/`python`）
4. **持久化**: 确定后写入 `config.json`

每次执行 CLI 使用该路径：`{python} scripts/vibedocx.py <command>`

## Dependencies

- **python-docx**: `pip install python-docx`
- **lxml**: `pip install lxml`

若导入失败，agent 使用上述确定的 Python 环境自动安装缺失的依赖。
