# Batch 批处理模式

适用于 3 步以上操作序列。一次加载文档，串行执行所有操作，一次保存。

```bash
python scripts/vibe_batch.py batch.json
```

## JSON 格式

```json
{
  "source": "input.docx",
  "output": "output.docx",
  "operations": [
    {"command": "format style", "args": {"name": "Normal", "font": "宋体", "size": 12}},
    {"command": "content append", "args": {"type": "heading", "text": "第一章 绪论", "level": 1}},
    {"command": "content append", "args": {"type": "body", "text": "正文内容..."}},
    {"command": "ref add", "args": {"key": "luxun", "text": "鲁迅..."}},
    {"command": "ref cite", "args": {"key": "luxun", "paragraph": 3, "style": "footnote"}},
    {"command": "ref generate", "args": {"heading": "参考文献"}},
    {"command": "page-break", "args": {}}
  ]
}
```

- `source`：输入文件路径
- `output`：输出文件路径
- `operations`：操作列表，按顺序执行

## 支持的命令

与 CLI 相同的命令名：

`format style`、`format page`、`format paragraph`、`format clear-direct`、
`replace text`、`replace bookmark`、`delete`、`move`、`swap`、
`content append`、`image insert`、`toc insert`、`section insert`、
`table create`、`table add-row`、`table add-column`、`table merge`、
`footnote`、`equation insert`、`page-break`、
`ref add`、`ref cite`、`ref generate`、`ref remove`、`ref renumber`、
`field insert`、`style export`、`style import`、
`numbering swap`、`numbering update-figures`、`merge`

每个操作的 `args` 与对应 CLI 命令的 `--参数` 一致，去掉前缀 `--`。

## 注意事项

- 操作按顺序执行，每个操作的输入是前一个操作的输出
- 中间文件写入临时目录，最后一步写入 `output`
- 任何操作失败会停止并报告错误位置
- 不支持在操作间做条件判断——需要条件逻辑请用 Python API
