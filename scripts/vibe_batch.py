"""Batch mode: execute a sequence of operations from a JSON definition.

Usage: python scripts/vibe_batch.py batch.json

The JSON file specifies source, output, and a list of operations.
The document is loaded once, all operations run in memory, then saved once.

Example batch.json:
{
  "source": "input.docx",
  "output": "output.docx",
  "operations": [
    {"command": "format style", "args": {"name": "Normal", "font": "宋体", "size": 12}},
    {"command": "content append", "args": {"type": "heading", "text": "第一章", "level": 1}},
    {"command": "content append", "args": {"type": "body", "text": "正文内容..."}},
    {"command": "ref add", "args": {"key": "luxun", "text": "鲁迅..."}},
    {"command": "ref cite", "args": {"key": "luxun", "paragraph": 3, "style": "footnote"}}
  ]
}
"""

import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def run_batch(batch_file):
    """Execute a batch operation file. Returns (count, message)."""
    with open(batch_file, 'r', encoding='utf-8') as f:
        batch = json.load(f)

    source = batch.get('source')
    output = batch.get('output', source)
    operations = batch.get('operations', [])

    if not source or not operations:
        return (0, "Error: 'source' and 'operations' are required")

    if not os.path.exists(source):
        return (0, f"Error: source file '{source}' not found")

    # Map command names to handler functions
    from helper.cli import (
        cmd_format_style, cmd_format_page, cmd_format_paragraph, cmd_format_clear_direct,
        cmd_replace_text, cmd_replace_bookmark, cmd_delete, cmd_move, cmd_swap,
        cmd_append_content, cmd_insert_image, cmd_insert_toc, cmd_insert_section,
        cmd_table_create, cmd_table_add_row, cmd_table_add_column, cmd_table_merge,
        cmd_insert_footnote, cmd_insert_equation, cmd_page_break,
        cmd_ref_add, cmd_ref_cite, cmd_ref_generate, cmd_ref_remove, cmd_ref_renumber,
        cmd_field_insert, cmd_style_export, cmd_style_import,
        cmd_numbering, cmd_numbering_update_figures, cmd_merge,
    )

    HANDLERS = {
        "format style": cmd_format_style,
        "format page": cmd_format_page,
        "format paragraph": cmd_format_paragraph,
        "format clear-direct": cmd_format_clear_direct,
        "replace text": cmd_replace_text,
        "replace bookmark": cmd_replace_bookmark,
        "delete": cmd_delete,
        "move": cmd_move,
        "swap": cmd_swap,
        "content append": cmd_append_content,
        "image insert": cmd_insert_image,
        "toc insert": cmd_insert_toc,
        "section insert": cmd_insert_section,
        "table create": cmd_table_create,
        "table add-row": cmd_table_add_row,
        "table add-column": cmd_table_add_column,
        "table merge": cmd_table_merge,
        "footnote": cmd_insert_footnote,
        "equation insert": cmd_insert_equation,
        "page-break": cmd_page_break,
        "ref add": cmd_ref_add,
        "ref cite": cmd_ref_cite,
        "ref generate": cmd_ref_generate,
        "ref remove": cmd_ref_remove,
        "ref renumber": cmd_ref_renumber,
        "field insert": cmd_field_insert,
        "style export": cmd_style_export,
        "style import": cmd_style_import,
        "numbering swap": cmd_numbering,
        "numbering update-figures": cmd_numbering_update_figures,
        "merge": cmd_merge,
    }

    # Use a temp file to chain operations (each reads output of previous)
    current = source
    import argparse
    for i, op in enumerate(operations):
        cmd = op.get("command")
        args_dict = op.get("args", {})

        if cmd not in HANDLERS:
            return (i, f"Error: unknown command '{cmd}' at operation {i}")

        is_last = (i == len(operations) - 1)
        out = output if is_last else os.path.join(tempfile.gettempdir(), f"_vibe_batch_{i}.docx")

        # Build fake argparse namespace
        ns = argparse.Namespace(file=current, output=out, **args_dict)
        try:
            HANDLERS[cmd](ns)
        except Exception as e:
            return (i, f"Error at operation {i} ({cmd}): {e}")

        if not is_last:
            current = out

    # Clean up intermediate files
    for i in range(len(operations) - 1):
        tmp = os.path.join(tempfile.gettempdir(), f"_vibe_batch_{i}.docx")
        if os.path.exists(tmp):
            os.unlink(tmp)

    return (len(operations), f"Executed {len(operations)} operations, saved to {output}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/vibe_batch.py <batch.json>")
        sys.exit(1)
    count, message = run_batch(sys.argv[1])
    print(message)
    sys.exit(0 if count > 0 else 1)
