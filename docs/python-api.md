# Python API 脚本模式

适用于复杂工作流、条件判断、循环处理、论文生成。

## 环境设置

Agent 执行 Python API 时，需确保 `scripts/` 在 Python 路径中：

```bash
# 方式一：PYTHONPATH 追加
# Linux / macOS:
PYTHONPATH="scripts:$PYTHONPATH" python /tmp/gen_paper.py
# Windows PowerShell:
$env:PYTHONPATH = "scripts;$env:PYTHONPATH"; python /tmp/gen_paper.py

# 方式二：cd 到 scripts 目录（cwd 自动加入 sys.path，跨平台最简单）
cd scripts && python /tmp/gen_paper.py
```

脚本内直接 import，无需 `sys.path.insert`：

```python
from helper.build import create_blank, create_structure
from helper.content import (
    append_content, page_break, replace_text, replace_filtered, replace_bookmark,
    delete_block, move_block, swap_blocks,
    insert_image, insert_toc, insert_section, insert_footnote, insert_equation,
    table_create, table_add_row, table_add_column, table_merge_cells,
    xref_insert, xref_update, merge_documents, convert_notes,
)
from helper.style import format_style, format_page, format_paragraph, format_clear_direct
from helper.cite import ref_add, ref_cite, ref_list, ref_generate, ref_renumber, ref_remove
from helper.field import field_list, field_update, field_insert, field_refresh
from helper.numbering import update_chapter_numbers, update_figure_numbers
```

## 核心模式

每个函数都遵循 `func(file_path, ..., output=None)` 约定：
- `file_path`：输入文件
- `output`：输出文件（默认 `{name}_edited.docx`）
- 返回值：通常为 `int`（操作的计数）

```python
OUT = "paper.docx"
create_blank(OUT, paper="A4")

# 每次操作将 output 指向同一个文件，实现链式处理
format_style(OUT, style_name="Normal", font="宋体", size=12, output=OUT)
append_content(OUT, "heading", "第一章", level=1, output=OUT)
append_content(OUT, "body", "正文内容......", output=OUT)
ref_add(OUT, key="luxun", text="鲁迅. 中国小说史略[M]. 1973.", output=OUT)
ref_cite(OUT, "luxun", paragraph_index=2, style="footnote", output=OUT)
ref_generate(OUT, heading="参考文献", output=OUT)
```

---

## 完整函数参考

### build — 创建文档

```python
# 创建空白文档（返回 output 路径）
create_blank(output, paper='A4', margin_top=2.54, margin_bottom=2.54,
             margin_left=3.17, margin_right=3.17, orientation='portrait')
# → str

# 从 JSON 结构创建文档
create_structure(output, structure_json)  # dict or JSON string
# → str
```

### content — 内容操作

```python
# 追加段落（返回 1）
append_content(file_path, content_type, text, after_paragraph=None,
               level=1, font_cn=None, font_en=None, size=None,
               bold=None, italic=None, align=None, output=None)
# → int

# 查找替换（返回替换次数）
replace_text(file_path, mapping, output=None)          # mapping: {find: replace}
replace_filtered(file_path, filter_spec, mapping, output=None)
# → int

# 书签替换
replace_bookmark(file_path, bookmarks, output=None)    # bookmarks: {name: text}
# → int

# 块操作（返回操作元素数）
delete_block(file_path, filter_spec, output=None)
move_block(file_path, filter_spec, target_filter, position="after", output=None)
swap_blocks(file_path, filter1, filter2, output=None)
# → int / (int, int)

# 分页符（返回 1）
page_break(file_path, after_paragraph=None, output=None)
# → int
```

```python
# 图片插入（返回 1）
insert_image(file_path, image_path, after_paragraph=None,
             width=None, height=None, caption=None, ref_id=None, output=None)
# → int

# 目录（返回 1）
insert_toc(file_path, after_paragraph=None, levels=3, output=None)
# → int

# 分节符（返回 1）
insert_section(file_path, after_paragraph=None, break_type='nextPage',
               header_text=None, footer_text=None, output=None)
# → int

# 脚注（返回脚注 ID）
insert_footnote(file_path, paragraph_index, text, output=None)
# → int

# 公式（返回 1）
insert_equation(file_path, paragraph_index, omml_xml, output=None)
# → int
```

```python
# 表格操作
table_create(file_path, headers=None, rows=None, caption=None,
             after_paragraph=None, ref_id=None, output=None)
table_add_row(file_path, table_index, data=None, output=None)
table_add_column(file_path, table_index, header=None, output=None)
table_merge_cells(file_path, table_index, row_start, col_start, row_end, col_end, output=None)
# → int

# 交叉引用
xref_insert(file_path, ref_type, ref_id, text, paragraph_index, position='end', output=None)
xref_update(file_path, labels, output=None)  # labels: {ref_id: new_text}
# → int

# 合并文档（返回文档数量）
merge_documents(file_paths, output)
# → int

# 脚注/尾注互换（返回转换数量）
convert_notes(file_path, to_type, output=None)  # to_type: 'footnote' or 'endnote'
# → int
```

### style — 格式

```python
# 修改样式定义
format_style(file_path, style_name, font=None, font_west=None, size=None,
             bold=False, no_bold=False, italic=False, no_italic=False,
             align=None, line_spacing=None, space_before=None, space_after=None,
             indent_first=None, output=None)

# 段落直接格式
format_paragraph(file_path, paragraph_index, font=None, font_west=None,
                 size=None, bold=None, italic=None, align=None,
                 line_spacing=None, output=None)

# 页面设置
format_page(file_path, margin_top=None, margin_bottom=None, margin_left=None,
            margin_right=None, paper=None, orientation=None,
            header_text=None, footer_text=None, output=None)

# 清除直接格式
format_clear_direct(file_path, range_start=None, range_end=None,
                    style_filter=None, output=None)
```

### cite — 引用管理

```python
# 添加参考文献（返回引用总数）
ref_add(file_path, key, text, output=None)
# → int

# 插入引用（返回引用编号）
ref_cite(file_path, key, paragraph_index,
         after_text=None, style='inline', position='end', output=None)
# style: 'inline' | 'footnote' | 'endnote'
# → int

# 列出所有引用
ref_list(file_path)
# → dict: {references: {key: text}, citations: [...], count: int}

# 生成参考文献列表（返回引用数量）
ref_generate(file_path, heading='参考文献', style='numbered',
             after_heading=None, output=None)
# → int

# 重编号（返回修改数量）
ref_renumber(file_path, output=None)
# → int

# 删除引用（返回剩余数量）
ref_remove(file_path, key, output=None)
# → int
```

### field — 域操作

```python
# 列出所有域
field_list(file_path, field_type=None)
# → dict: {fields: [{id, type, instruction, location, ...}]}

# 更新域指令
field_update(file_path, field_id=None, field_type=None, nth=1, instruction=None, output=None)

# 插入域
field_insert(file_path, after_paragraph=None, field_type=None, fmt=None,
             location='body', output=None)

# 刷新域缓存
field_refresh(file_path, field_type=None, output=None)
# → int
```

### numbering — 编号

```python
# 交换章节编号（返回修改的 run 数）
update_chapter_numbers(file_path, ch_a, ch_b, chapter_names=None, patterns=None, output=None)
# → int

# 图表自动编号（返回更新的 caption 数）
update_figure_numbers(file_path, output=None)
# → int
```

### query — 只读查询

```python
from helper.query import inspect_structure, inspect_styles, inspect_formatting, inspect_content, select

inspect_structure(file_path)        # → dict
inspect_styles(file_path, style_name=None)  # → dict
inspect_formatting(file_path, range_start=None, range_end=None, paragraph_index=None)  # → dict
inspect_content(file_path, range_start=None, range_end=None, formatted=False)  # → dict
select(file_path, filter_spec, context_lines=0)  # → dict
```

---

## 自包含示例

以下代码写入临时文件后，`cd` 到 `scripts/` 目录执行：

```bash
cd scripts && python /tmp/gen_paper.py
```

```python
import os, tempfile

from helper.build import create_blank
from helper.content import append_content, page_break, table_create
from helper.style import format_style
from helper.cite import ref_add, ref_cite, ref_generate
from helper.field import field_insert
from helper.numbering import update_figure_numbers

OUT = os.path.join(tempfile.gettempdir(), "example.docx")

create_blank(OUT, paper="A4")
format_style(OUT, style_name="Normal", font="宋体", size=12, line_spacing=1.5, align="justify", output=OUT)
field_insert(OUT, field_type="PAGE", location="footer", output=OUT)

append_content(OUT, "heading", "第一章  绪论", level=1, output=OUT)
append_content(OUT, "body", "这是正文内容。", output=OUT)

table_create(OUT, headers=["项目", "数值"], rows=[["A", "100"]],
             caption="测试数据", ref_id="tab:data", output=OUT)

ref_add(OUT, key="ref1", text="作者. 标题[J]. 期刊, 2020.", output=OUT)
ref_cite(OUT, "ref1", paragraph_index=2, style="footnote",
         after_text="正文内容。", output=OUT)
ref_generate(OUT, heading="参考文献", output=OUT)

update_figure_numbers(OUT, output=OUT)
page_break(OUT, output=OUT)

print(f"Done: {OUT}")
```

Agent 将代码写入临时文件，然后 `cd scripts && python /tmp/gen_paper.py` 执行。
或者将scripts目录临时加入环境变量 `PYTHONPATH`。
