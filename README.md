# vibedocx

一个Agent skill，用于处理Word文档。

Python 库与 CLI 工具，用于程序化创建、编辑和排版 Word 文档（`.docx`），面向中文学术写作场景，提供参考文献管理、图表编号、交叉引用、脚注等原生支持。

绝大多数操作基于python-docx实现，避免直接操作xml域，提高token效率。



- **无需 GUI** — 全部通过命令行或 Python 脚本操作
- **非破坏性** — 所有操作生成新文件，原始文件不会被修改
- **引用自包含** — 引用数据存储在文档内部（`word/settings.xml`），`.docx` 文件可独立传递
- **三层设计** — 一行 CLI 命令 / JSON 批量文件 / Python API，按需选用

## 安装

```bash
pip install python-docx lxml
```

或：

```bash
pip install -r requirements.txt
```

## 三种使用模式

| 场景 | 模式 | 入口 |
|------|------|------|
| 单步操作（查结构、改样式） | CLI | `python scripts/vibedocx.py <command> <args>` |
| 3+ 步操作序列 | Batch | `python scripts/vibe_batch.py batch.json` |
| 复杂工作流、论文生成 | Python API | `from helper import ...` |

### CLI 示例

```bash
# 创建空白 A4 文档
python scripts/vibedocx.py create blank -o paper.docx --paper A4

# 修改 Normal 样式
python scripts/vibedocx.py format style paper.docx --name "Normal" --font 宋体 --size 12 --line-spacing 1.5 -o paper.docx

# 追加标题和正文
python scripts/vibedocx.py content append paper.docx --type heading --text "第一章 引言" --level 1 -o paper.docx
python scripts/vibedocx.py content append paper.docx --type body --text "这是正文内容。" -o paper.docx

# 查看文档结构
python scripts/vibedocx.py inspect structure paper.docx
```

### Batch 示例

```json
{
  "source": "input.docx",
  "output": "output.docx",
  "operations": [
    {"command": "format style", "args": {"name": "Normal", "font": "宋体", "size": 12}},
    {"command": "content append", "args": {"type": "heading", "text": "第一章", "level": 1}},
    {"command": "ref add", "args": {"key": "ref1", "text": "作者. 标题[J]. 期刊, 2020."}},
    {"command": "ref generate", "args": {"heading": "参考文献"}}
  ]
}
```

### Python API 示例

```python
from helper.build import create_blank
from helper.content import append_content, table_create, page_break
from helper.style import format_style
from helper.cite import ref_add, ref_cite, ref_generate

OUT = "paper.docx"

create_blank(OUT, paper="A4")
format_style(OUT, style_name="Normal", font="宋体", size=12, line_spacing=1.5, output=OUT)

append_content(OUT, "heading", "第一章  绪论", level=1, output=OUT)
append_content(OUT, "body", "这是正文内容。", output=OUT)

table_create(OUT, headers=["项目", "数值"], rows=[["A", "100"]],
             caption="表1  测试数据", ref_id="tab:data", output=OUT)

ref_add(OUT, key="ref1", text="作者. 标题[J]. 期刊, 2020.", output=OUT)
ref_cite(OUT, "ref1", paragraph_index=2, style="footnote", after_text="正文内容。", output=OUT)
ref_generate(OUT, heading="参考文献", output=OUT)

page_break(OUT, output=OUT)
```

运行方式：

```bash
cd scripts && python your_script.py
```

## 功能概览

| 模块 | 功能 |
|------|------|
| **创建** | 空白文档（含页面设置）、从 JSON 模板创建结构化文档 |
| **内容** | 追加标题/正文、文本替换（普通/正则/书签）、删除/移动/交换内容块、合并文档 |
| **排版** | 样式修改、段落直接格式化、页面设置（纸张/边距/页眉页脚）、清除直接格式 |
| **插入** | 图片（含题注与引用 ID）、表格（创建/增行/增列/合并单元格）、目录、公式 (OMML)、脚注/尾注、分页/分节符 |
| **引用** | 添加参考文献、三种引用方式（行内 `[n]` / 脚注 / 尾注）、自动生成参考文献列表、重新编号 |
| **域** | 查看/更新/插入/刷新 Word 域（PAGE、NUMPAGES、DATE、TOC） |
| **编号** | 章节编号交换（同步更新所有子编号与图表引用）、按章节自动编号图表 |
| **查询** | 查看文档结构、样式、格式、段落内容；使用组合过滤器筛选元素 |
| **交叉引用** | 插入与更新图表、公式的交叉引用 |

## 项目结构

```
scripts/
├── vibedocx.py            # CLI 入口
├── vibe_batch.py          # Batch 模式入口
├── helper/                # 核心模块
│   ├── build.py           # 文档创建
│   ├── content.py         # 内容操作（追加、替换、删除、移动、插入等）
│   ├── style.py           # 样式与页面设置
│   ├── cite.py            # 参考文献管理
│   ├── field.py           # Word 域操作
│   ├── numbering.py       # 章节编号交换、图表自动编号
│   ├── filter.py          # 组合过滤器引擎 (AND/OR)
│   ├── query.py           # 只读查询
│   ├── config.py          # 三级配置系统
│   ├── units.py           # 字号/单位解析
│   └── logutil.py         # 日志
├── templates/             # XML 模板（脚注、尾注）
└── tests/                 # pytest 测试
docs/
├── cli.md                 # CLI 命令参考
├── python-api.md          # Python API 参考
└── batch.md               # Batch 模式参考
```

## 运行测试

```bash
python -m pytest scripts/tests/ -v
```

## 文档

| 文档 | 内容 |
|------|------|
| [SKILL.md](SKILL.md) | 主技能指南：三种模式、快速参考、配置系统 |
| [docs/cli.md](docs/cli.md) | CLI 完整命令参考与示例 |
| [docs/python-api.md](docs/python-api.md) | Python API 函数参考与完整示例 |
| [docs/batch.md](docs/batch.md) | Batch JSON 格式与支持的命令 |

## 设计原则

- 所有操作生成新文件，`-o` 默认输出 `{name}_edited.docx`
- 每次修改前自动备份（`.vibedocx_backup`），失败时自动恢复
- 引用数据存储在文档自身中，无需外部文件
- python-docx 高层 API + lxml XML 补丁，覆盖原生库不支持的脚注、域等功能
- 三级配置：Skill 默认 → 项目 `.vibedocx/style.json` → CLI 参数

## License

[MIT](LICENSE.txt)
