# CLI 单步命令参考

所有命令通过 `python scripts/vibedocx.py <command> <args>` 调用。

## Query（只读查询）

```bash
# 文档结构
python scripts/vibedocx.py inspect structure <file>

# 样式列表
python scripts/vibedocx.py inspect styles <file> [--style "Normal"]

# 段落格式
python scripts/vibedocx.py inspect formatting <file> [--range-start N] [--range-end N] [--paragraph N]

# 文本内容
python scripts/vibedocx.py inspect content <file> [--range-start N] [--range-end N] [--formatted]

# 过滤查询
python scripts/vibedocx.py select <file> --filter '{"style":"Heading 1"}' [--context N]
```

## Format（格式）

```bash
# 修改样式
python scripts/vibedocx.py format style <file> --name "Normal" --font 宋体 --size 12 --line-spacing 1.5 -o out

# 段落格式
python scripts/vibedocx.py format paragraph <file> --index 3 --bold --align center -o out

# 页面设置
python scripts/vibedocx.py format page <file> --paper A4 --margin-top 2.54 --header-text "页眉" -o out

# 清除直接格式
python scripts/vibedocx.py format clear-direct <file> [--style "Normal"] -o out
```

## Content（内容操作）

```bash
# 追加内容
python scripts/vibedocx.py content append <file> --type heading --text "第一章" --level 1 [--after-paragraph N] -o out
python scripts/vibedocx.py content append <file> --type body --text "正文..." [--font 宋体 --size 12] -o out

# 替换文本
python scripts/vibedocx.py replace text <file> --find "旧" --replace "新" -o out
python scripts/vibedocx.py replace text <file> --mapping '{"{{名}}":"张三"}' -o out

# 书签替换
python scripts/vibedocx.py replace bookmark <file> --bookmarks '{"key":"val"}' -o out

# 块操作
python scripts/vibedocx.py delete <file> --filter '{"heading_block":"章节"}' -o out
python scripts/vibedocx.py move <file> --filter '{...}' --to '{...}' [--position after|before] -o out
python scripts/vibedocx.py swap <file> --filter1 '{...}' --filter2 '{...}' -o out

# 编号
python scripts/vibedocx.py numbering swap <file> --ch-a 3 --ch-b 4 [--names '{"3":"标题三","4":"标题四"}'] -o out
python scripts/vibedocx.py numbering update-figures <file> -o out
```

## Build（创建）

```bash
python scripts/vibedocx.py create blank -o out.docx [--paper A4 --margin-top 2.54 ...]
python scripts/vibedocx.py create structure --file skeleton.json -o out.docx
```

## Insert（插入）

```bash
# 图片
python scripts/vibedocx.py image insert <file> --image fig.png [--after-paragraph N] [--width 10] [--height 8] [--caption "图1"] [--ref-id fig:xxx] -o out

# 表格
python scripts/vibedocx.py table create <file> --headers '["A","B"]' --rows '[["1","2"]]' [--caption "表1"] [--ref-id tab:xxx] -o out
python scripts/vibedocx.py table add-row <file> --table 0 --data '["a","b"]' -o out
python scripts/vibedocx.py table add-column <file> --table 0 [--header "C"] -o out
python scripts/vibedocx.py table merge <file> --table 0 --row-start 1 --col-start 0 --row-end 2 --col-end 1 -o out

# 目录
python scripts/vibedocx.py toc insert <file> [--after-paragraph N] [--levels 3] -o out

# 分节符
python scripts/vibedocx.py section insert <file> --after-paragraph N [--break-type nextPage|continuous] [--header-text "..."] [--footer-text "..."] -o out

# 公式
python scripts/vibedocx.py equation insert <file> --paragraph N --omml "<m:oMath>...</m:oMath>" -o out

# 分页符
python scripts/vibedocx.py page-break <file> [--after-paragraph N] -o out

# 脚注
python scripts/vibedocx.py footnote <file> --paragraph N --text "脚注内容" -o out
```

## Cite（引用管理）

```bash
# 添加参考文献
python scripts/vibedocx.py ref add <file> --key "luxun" --text "鲁迅. 中国小说史略[M]. 1973." -o out

# 插入引用（inline / footnote / endnote）
python scripts/vibedocx.py ref cite <file> --key "luxun" --paragraph 5 --style footnote [--after-text "鲁迅指出"] -o out

# 列出所有引用
python scripts/vibedocx.py ref list <file>

# 生成参考文献列表
python scripts/vibedocx.py ref generate <file> --heading "参考文献" [--after-heading "..."] -o out

# 重编号
python scripts/vibedocx.py ref renumber <file> -o out

# 删除引用
python scripts/vibedocx.py ref remove <file> --key "luxun" -o out
```

## Field（域操作）

```bash
python scripts/vibedocx.py field list <file> [--type PAGE|DATE|TOC]
python scripts/vibedocx.py field update <file> --type PAGE --nth 1 --instruction " PAGE " -o out
python scripts/vibedocx.py field insert <file> --type PAGE [--after-paragraph N] [--location header|footer|body] -o out
python scripts/vibedocx.py field refresh <file> [--type DATE] -o out
```

## 其他

```bash
# 合并文档
python scripts/vibedocx.py merge file1.docx file2.docx -o merged.docx

# 样式导入/导出
python scripts/vibedocx.py style <file> --export --to styles.json
python scripts/vibedocx.py style <file> --import --from styles.json -o out

# 脚注/尾注互换
python scripts/vibedocx.py note convert <file> --to footnote|endnote -o out

# 交叉引用
python scripts/vibedocx.py xref insert <file> --type figure --id "fig:arch" --text "图1" --paragraph 5 -o out
python scripts/vibedocx.py xref update <file> --labels '{"fig:arch":"图2"}' -o out
```

## 过滤器语法

用于 `select`、`delete`、`move`、`swap`、`replace text --filter`：

- AND: `{"style": "Heading 1", "text_contains": "引言"}`
- OR: `[{"style": "Heading 1"}, {"style": "Heading 2"}]`
- 段落级: `style`, `outline_level`, `text_contains`, `text_regex`, `bold`, `italic`, `font`, `paragraph_range`, `in_table`, `under_heading`
- Run 级: `run_text_contains`, `run_text_regex`, `run_bold`, `run_italic`, `run_font`, `run_size`
- 块级: `heading_block`, `heading_block_regex`, `heading_level`, `heading_outline_level`
