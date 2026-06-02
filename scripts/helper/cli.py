"""CLI entry point for vibedocx — pure router.

Commands route to the appropriate module layer:
  query  → inspect / select / content reading
  style  → format style / page / clear-direct
  content → replace / move / delete / swap / numbering
  build  → create blank / structure
  cite   → ref add/cite/list/generate/renumber/remove
  field  → field list/update/insert/refresh
"""

import argparse
import json
import os
import shutil
import sys
from functools import wraps
from pathlib import Path

from helper import __version__
from helper.logutil import setup_logging, logger


def transactional(handler):
    """Decorator: backup input file before operation, restore on failure."""
    @wraps(handler)
    def wrapper(args):
        if not hasattr(args, 'file') or not args.file:
            return handler(args)
        src = args.file
        if not os.path.exists(src):
            return handler(args)
        backup = src + '.vibedocx_backup'
        shutil.copy2(src, backup)
        logger.debug(f'Backup created: {backup}')
        try:
            result = handler(args)
            if os.path.exists(backup):
                os.remove(backup)
            return result
        except Exception:
            logger.error(f'Operation failed, restoring from backup')
            output = getattr(args, 'output', src) if hasattr(args, 'output') else src
            if output != src and os.path.exists(src):
                shutil.copy2(src, output)
            elif os.path.exists(backup):
                shutil.copy2(backup, src)
            if os.path.exists(backup):
                os.remove(backup)
            raise
    return wrapper


# ── Helpers ──

def _parse_json(s, name="JSON"):
    try:
        return json.loads(s)
    except json.JSONDecodeError as e:
        print(f"Error: invalid {name}: {e}", file=sys.stderr)
        sys.exit(2)


def _default_output_path(input_path):
    p = Path(input_path)
    candidate = p.parent / f"{p.stem}_edited{p.suffix}"
    counter = 1
    while candidate.exists():
        candidate = p.parent / f"{p.stem}_edited{counter}{p.suffix}"
        counter += 1
    return str(candidate)


def _auto_output(args):
    if hasattr(args, "output") and args.output is None:
        if hasattr(args, 'file') and args.file:
            args.output = _default_output_path(args.file)


def _write_output(result, output_file):
    """Write JSON result to stdout and optionally to a UTF-8 file."""
    text = json.dumps(result, ensure_ascii=False, indent=2)
    print(text)
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(text)


# ── Command handlers ──

def cmd_inspect_structure(args):
    from helper.query import inspect_structure
    _write_output(inspect_structure(args.file), args.output_file)

def cmd_inspect_styles(args):
    from helper.query import inspect_styles
    _write_output(inspect_styles(args.file, style_name=args.style), args.output_file)

def cmd_inspect_formatting(args):
    from helper.query import inspect_formatting
    _write_output(inspect_formatting(args.file, range_start=args.range_start,
        range_end=args.range_end, paragraph_index=args.paragraph), args.output_file)

def cmd_inspect_content(args):
    from helper.query import inspect_content
    _write_output(inspect_content(args.file, range_start=args.range_start,
        range_end=args.range_end, formatted=args.formatted), args.output_file)


def cmd_inspect_table(args):
    from helper.query import inspect_table
    _write_output(inspect_table(args.file, table_index=args.table), args.output_file)

def cmd_format_style(args):
    from helper.style import format_style
    format_style(args.file, style_name=args.name, font=args.font, font_west=args.font_west,
        size=args.size, bold=args.bold, no_bold=args.no_bold, italic=args.italic,
        no_italic=args.no_italic, color=args.color, no_color=args.no_color,
        align=args.align, line_spacing=args.line_spacing,
        space_before=args.space_before, space_after=args.space_after,
        indent_first=args.indent_first, output=args.output)
    print(f"Saved to {args.output}", file=sys.stderr)

def cmd_format_clear_direct(args):
    from helper.style import format_clear_direct
    format_clear_direct(args.file, range_start=args.range_start, range_end=args.range_end,
        style_filter=args.style, output=args.output)
    print(f"Saved to {args.output}", file=sys.stderr)

def cmd_format_page(args):
    from helper.style import format_page
    format_page(args.file, margin_top=args.margin_top, margin_bottom=args.margin_bottom,
        margin_left=args.margin_left, margin_right=args.margin_right, paper=args.paper,
        orientation=args.orientation, header_text=args.header_text,
        footer_text=args.footer_text, output=args.output)
    print(f"Saved to {args.output}", file=sys.stderr)

def cmd_select(args):
    from helper.query import select
    _write_output(select(args.file, _parse_json(args.filter, "filter"),
        context_lines=args.context or 0), args.output_file)

def cmd_replace_text(args):
    from helper.content import replace_text, replace_filtered
    mapping = _parse_json(args.mapping, "mapping") if args.mapping else (
        {args.find: args.replace} if args.find and args.replace else None)
    if not mapping:
        print("Error: provide --find/--replace or --mapping", file=sys.stderr); sys.exit(2)
    if args.filter:
        count = replace_filtered(args.file, _parse_json(args.filter, "filter"),
                                 mapping=mapping, output=args.output)
    else:
        count = replace_text(args.file, mapping=mapping, output=args.output)
    print(f"Replaced {count} occurrence(s), saved to {args.output}", file=sys.stderr)

def cmd_replace_bookmark(args):
    from helper.content import replace_bookmark
    count = replace_bookmark(args.file, bookmarks=_parse_json(args.bookmarks, "bookmarks"),
                             output=args.output)
    print(f"Replaced {count} bookmark(s), saved to {args.output}", file=sys.stderr)

def cmd_field_list(args):
    from helper.field import field_list
    _write_output(field_list(args.file, field_type=args.type), args.output_file)

def cmd_field_update(args):
    from helper.field import field_update
    field_update(args.file, field_id=args.id, field_type=args.type, nth=args.nth,
                 instruction=args.instruction, output=args.output)
    print(f"Saved to {args.output}", file=sys.stderr)

def cmd_field_insert(args):
    from helper.field import field_insert
    field_insert(args.file, after_paragraph=args.after_paragraph, field_type=args.type,
                 fmt=args.format, location=args.location, output=args.output)
    print(f"Inserted field '{args.type}', saved to {args.output}", file=sys.stderr)

def cmd_field_refresh(args):
    from helper.field import field_refresh
    count = field_refresh(args.file, field_type=args.type, output=args.output)
    print(f"Refreshed {count} field(s), saved to {args.output}", file=sys.stderr)

def cmd_delete(args):
    from helper.content import delete_block
    count = delete_block(args.file, _parse_json(args.filter, "filter"), output=args.output)
    print(f"Deleted {count} element(s), saved to {args.output}", file=sys.stderr)

def cmd_move(args):
    from helper.content import move_block
    count = move_block(args.file, _parse_json(args.filter, "filter"),
                       _parse_json(args.to, "target"), position=args.position, output=args.output)
    print(f"Moved {count} element(s), saved to {args.output}", file=sys.stderr)

def cmd_swap(args):
    from helper.content import swap_blocks
    n1, n2 = swap_blocks(args.file, _parse_json(args.filter1, "filter1"),
                          _parse_json(args.filter2, "filter2"), output=args.output)
    print(f"Swapped blocks ({n1} + {n2} elements), saved to {args.output}", file=sys.stderr)

def cmd_numbering(args):
    from helper.numbering import update_chapter_numbers
    names = _parse_json(args.names, "names") if args.names else None
    count = update_chapter_numbers(args.file, ch_a=args.ch_a, ch_b=args.ch_b,
                                    chapter_names=names, output=args.output)
    print(f"Updated {count} run(s), saved to {args.output}", file=sys.stderr)

def cmd_create_blank(args):
    from helper.build import create_blank
    create_blank(args.output, paper=args.paper, margin_top=args.margin_top,
        margin_bottom=args.margin_bottom, margin_left=args.margin_left,
        margin_right=args.margin_right, orientation=args.orientation)
    print(f"Created blank document: {args.output}", file=sys.stderr)

def cmd_create_structure(args):
    from helper.build import create_structure
    if args.file:
        with open(args.file, 'r', encoding='utf-8') as f:
            structure = json.load(f)
    elif args.json:
        structure = _parse_json(args.json, "structure")
    else:
        print("Error: provide --file or --json", file=sys.stderr); sys.exit(2)
    create_structure(args.output, structure)
    print(f"Created document from structure: {args.output}", file=sys.stderr)


def cmd_create_from_template(args):
    from helper.build import create_from_template
    create_from_template(args.template, output=args.output)
    print(f"Created from template: {args.output}", file=sys.stderr)

def cmd_ref_add(args):
    from helper.cite import ref_add
    count = ref_add(args.file, key=args.key, text=args.text, output=args.output)
    print(f"Added reference '{args.key}' ({count} total), saved to {args.output}", file=sys.stderr)

def cmd_ref_cite(args):
    from helper.cite import ref_cite
    num = ref_cite(args.file, key=args.key, paragraph_index=args.paragraph,
                   position=args.position, after_text=args.after_text,
                   style=args.style, output=args.output)
    print(f"Inserted citation [{num}] for '{args.key}', saved to {args.output}", file=sys.stderr)

def cmd_ref_list(args):
    from helper.cite import ref_list
    _write_output(ref_list(args.file), args.output_file)

def cmd_ref_generate(args):
    from helper.cite import ref_generate
    count = ref_generate(args.file, heading=args.heading, style=args.style,
                          after_heading=args.after_heading, output=args.output)
    print(f"Generated bibliography with {count} references, saved to {args.output}", file=sys.stderr)

def cmd_ref_renumber(args):
    from helper.cite import ref_renumber
    count = ref_renumber(args.file, output=args.output)
    print(f"Renumbered {count} citation(s), saved to {args.output}", file=sys.stderr)

def cmd_ref_remove(args):
    from helper.cite import ref_remove
    remaining = ref_remove(args.file, key=args.key, output=args.output)
    print(f"Removed reference '{args.key}' ({remaining} remaining), saved to {args.output}", file=sys.stderr)

def cmd_insert_image(args):
    from helper.content import insert_image
    insert_image(args.file, image_path=args.image, after_paragraph=args.after_paragraph,
                 width=args.width, height=args.height, caption=args.caption,
                 ref_id=args.ref_id, output=args.output)
    print(f"Inserted image, saved to {args.output}", file=sys.stderr)

def cmd_insert_toc(args):
    from helper.content import insert_toc
    insert_toc(args.file, after_paragraph=args.after_paragraph,
               levels=args.levels, output=args.output)
    print(f"Inserted TOC, saved to {args.output}", file=sys.stderr)

def cmd_insert_section(args):
    from helper.content import insert_section
    insert_section(args.file, after_paragraph=args.after_paragraph,
                   break_type=args.break_type, header_text=args.header_text,
                   footer_text=args.footer_text, output=args.output)
    print(f"Inserted section break, saved to {args.output}", file=sys.stderr)

def cmd_table_add_row(args):
    from helper.content import table_add_row
    import json as _json
    data = _json.loads(args.data) if args.data else None
    table_add_row(args.file, table_index=args.table, data=data, output=args.output)
    print(f"Added row to table {args.table}, saved to {args.output}", file=sys.stderr)

def cmd_table_add_column(args):
    from helper.content import table_add_column
    table_add_column(args.file, table_index=args.table, header=args.header,
                     width=args.width, strategy=args.strategy, output=args.output)
    print(f"Added column to table {args.table}, saved to {args.output}", file=sys.stderr)

def cmd_table_merge(args):
    from helper.content import table_merge_cells
    table_merge_cells(args.file, table_index=args.table,
                      row_start=args.row_start, col_start=args.col_start,
                      row_end=args.row_end, col_end=args.col_end, output=args.output)
    print(f"Merged cells in table {args.table}, saved to {args.output}", file=sys.stderr)


def cmd_table_delete_row(args):
    from helper.content import table_delete_row
    table_delete_row(args.file, table_index=args.table,
                     row_index=args.row, output=args.output)
    print(f"Deleted row {args.row} from table {args.table}, saved to {args.output}", file=sys.stderr)


def cmd_table_delete_column(args):
    from helper.content import table_delete_column
    table_delete_column(args.file, table_index=args.table,
                        col_index=args.col, strategy=args.strategy,
                        output=args.output)
    print(f"Deleted column {args.col} from table {args.table}, saved to {args.output}", file=sys.stderr)


def cmd_table_format_cell(args):
    from helper.content import table_format_cell
    bold = True if args.bold else (False if args.no_bold else None)
    italic = True if args.italic else (False if args.no_italic else None)
    table_format_cell(args.file, table_index=args.table,
                      row=args.row, col=args.col,
                      font_cn=args.font_cn, font_en=args.font_en,
                      size=args.size, bold=bold, italic=italic,
                      align=args.align, vertical_align=args.vertical_align,
                      shading=args.shading, width=args.width,
                      output=args.output)
    print(f"Formatted cell ({args.row},{args.col}) in table {args.table}, saved to {args.output}",
          file=sys.stderr)

def cmd_xref_insert(args):
    from helper.content import xref_insert
    xref_insert(args.file, ref_type=args.type, ref_id=args.id, text=args.text,
                paragraph_index=args.paragraph, position=args.position, output=args.output)
    print(f"Inserted xref '{args.id}', saved to {args.output}", file=sys.stderr)

def cmd_xref_update(args):
    from helper.content import xref_update
    labels = _parse_json(args.labels, "labels")
    count = xref_update(args.file, labels=labels, output=args.output)
    print(f"Updated {count} xref(s), saved to {args.output}", file=sys.stderr)

def cmd_append_content(args):
    from helper.content import append_content
    bold = True if args.bold else (False if args.no_bold else None)
    italic = True if args.italic else (False if args.no_italic else None)
    append_content(args.file, content_type=args.type, text=args.text,
                   after_paragraph=args.after_paragraph, level=args.level,
                   font_cn=args.font, font_en=args.font_west, size=args.size,
                   bold=bold, italic=italic,
                   align=args.align, output=args.output)
    print(f"Appended {args.type}, saved to {args.output}", file=sys.stderr)

def cmd_insert_footnote(args):
    from helper.content import insert_footnote
    fn_id = insert_footnote(args.file, paragraph_index=args.paragraph,
                            text=args.text, output=args.output)
    print(f"Inserted footnote [{fn_id}], saved to {args.output}", file=sys.stderr)

def cmd_convert_notes(args):
    from helper.content import convert_notes
    count = convert_notes(args.file, to_type=args.to, output=args.output)
    print(f"Converted {count} note(s) to {args.to}, saved to {args.output}", file=sys.stderr)

def cmd_insert_equation(args):
    from helper.content import insert_equation
    insert_equation(args.file, paragraph_index=args.paragraph,
                    omml_xml=args.omml, output=args.output)
    print(f"Inserted equation, saved to {args.output}", file=sys.stderr)

def cmd_table_create(args):
    from helper.content import table_create
    data = _parse_json(args.rows, "rows") if args.rows else []
    headers = _parse_json(args.headers, "headers") if args.headers else []
    table_create(args.file, headers=headers, rows=data,
                 caption=args.caption, after_paragraph=args.after_paragraph,
                 ref_id=args.ref_id, output=args.output)
    print(f"Created table, saved to {args.output}", file=sys.stderr)

def cmd_numbering_update_figures(args):
    from helper.numbering import update_figure_numbers
    count = update_figure_numbers(args.file, output=args.output)
    print(f"Renumbered {count} figure(s)/table(s), saved to {args.output}", file=sys.stderr)

def cmd_page_break(args):
    from helper.content import page_break
    page_break(args.file, after_paragraph=args.after_paragraph, output=args.output)
    print(f"Inserted page break, saved to {args.output}", file=sys.stderr)

def cmd_format_paragraph(args):
    from helper.style import format_paragraph
    bold = True if args.bold else (False if args.no_bold else None)
    italic = True if args.italic else (False if args.no_italic else None)
    format_paragraph(args.file, paragraph_index=args.index,
                     font=args.font, font_west=args.font_west, size=args.size,
                     bold=bold, italic=italic, align=args.align,
                     line_spacing=args.line_spacing, output=args.output)
    print(f"Formatted paragraph {args.index}, saved to {args.output}", file=sys.stderr)

def cmd_merge(args):
    from helper.content import merge_documents
    files = args.files
    count = merge_documents(files, output=args.output)
    print(f"Merged {count} documents, saved to {args.output}", file=sys.stderr)

def cmd_style_export(args):
    from helper.style import style_export
    count = style_export(args.file, output_json=args.to)
    print(f"Exported {count} styles to {args.to}", file=sys.stderr)

def cmd_style_import(args):
    from helper.style import style_import
    count = style_import(args.file, json_path=args.from_json, output=args.output)
    print(f"Imported {count} styles, saved to {args.output}", file=sys.stderr)


# ──────────────────────────────────────────────
# CLI parser
# ──────────────────────────────────────────────

def main():
    # Ensure UTF-8 output on terminals that default to legacy code pages
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    parser = argparse.ArgumentParser(prog="vibedocx",
        description="DOCX formatting and field management CLI tool.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress non-error output")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # ── query (inspect / select) ──
    qp = subparsers.add_parser("inspect", help="Read document info")
    qs = qp.add_subparsers(dest="subcommand", required=True)

    p = qs.add_parser("structure", help="Document structure overview")
    p.add_argument("file"); p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_inspect_structure)

    p = qs.add_parser("styles", help="List style definitions")
    p.add_argument("file"); p.add_argument("--style", help="Filter by style name")
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_inspect_styles)

    p = qs.add_parser("formatting", help="Paragraph formatting details")
    p.add_argument("file"); p.add_argument("--range-start", type=int)
    p.add_argument("--range-end", type=int); p.add_argument("--paragraph", type=int)
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_inspect_formatting)

    p = qs.add_parser("content", help="Read document content")
    p.add_argument("file"); p.add_argument("--range-start", type=int)
    p.add_argument("--range-end", type=int); p.add_argument("--formatted", action="store_true")
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_inspect_content)

    p = qs.add_parser("table", help="Inspect table structure and cells")
    p.add_argument("file"); p.add_argument("--table", type=int, help="Table index (0-based)")
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_inspect_table)

    p = subparsers.add_parser("select", help="Query elements by filter")
    p.add_argument("file"); p.add_argument("--filter", required=True)
    p.add_argument("--context", type=int)
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_select)

    # ── style (format) ──
    sp = subparsers.add_parser("format", help="Modify formatting")
    ss = sp.add_subparsers(dest="subcommand", required=True)

    p = ss.add_parser("style", help="Modify a named style")
    p.add_argument("file"); p.add_argument("--name", required=True)
    p.add_argument("--font"); p.add_argument("--font-west")
    p.add_argument("--size", type=str); p.add_argument("--bold", action="store_true")
    p.add_argument("--no-bold", action="store_true"); p.add_argument("--italic", action="store_true")
    p.add_argument("--no-italic", action="store_true")
    p.add_argument("--color", help="Font color as RGB hex (e.g., '000000')")
    p.add_argument("--no-color", action="store_true", help="Remove font color")
    p.add_argument("--align", choices=["left", "center", "right", "justify"])
    p.add_argument("--line-spacing", type=float); p.add_argument("--space-before", type=float)
    p.add_argument("--space-after", type=float); p.add_argument("--indent-first", type=float)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_format_style)

    p = ss.add_parser("clear-direct", help="Clear direct formatting")
    p.add_argument("file"); p.add_argument("--range-start", type=int)
    p.add_argument("--range-end", type=int); p.add_argument("--style")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_format_clear_direct)

    p = ss.add_parser("page", help="Set page properties")
    p.add_argument("file"); p.add_argument("--margin-top", type=float)
    p.add_argument("--margin-bottom", type=float); p.add_argument("--margin-left", type=float)
    p.add_argument("--margin-right", type=float); p.add_argument("--paper")
    p.add_argument("--orientation", choices=["portrait", "landscape"])
    p.add_argument("--header-text"); p.add_argument("--footer-text")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_format_page)

    p = ss.add_parser("paragraph", help="Format a specific paragraph")
    p.add_argument("file"); p.add_argument("--index", type=int, required=True,
                   help="Paragraph index to format")
    p.add_argument("--font"); p.add_argument("--font-west"); p.add_argument("--size", type=str)
    p.add_argument("--bold", action="store_true"); p.add_argument("--no-bold", action="store_true")
    p.add_argument("--italic", action="store_true"); p.add_argument("--no-italic", action="store_true")
    p.add_argument("--align", choices=["left", "center", "right", "justify"])
    p.add_argument("--line-spacing", type=float)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_format_paragraph)

    # ── content (replace / block ops) ──
    rp = subparsers.add_parser("replace", help="Content replacement")
    rs = rp.add_subparsers(dest="subcommand", required=True)

    p = rs.add_parser("text", help="Find and replace text")
    p.add_argument("file"); p.add_argument("--find"); p.add_argument("--replace")
    p.add_argument("--mapping"); p.add_argument("--filter")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_replace_text)

    p = rs.add_parser("bookmark", help="Replace content at bookmarks")
    p.add_argument("file"); p.add_argument("--bookmarks", required=True)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_replace_bookmark)

    p = subparsers.add_parser("delete", help="Delete a block")
    p.add_argument("file"); p.add_argument("--filter", required=True)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_delete)

    p = subparsers.add_parser("move", help="Move a block")
    p.add_argument("file"); p.add_argument("--filter", required=True)
    p.add_argument("--to", required=True); p.add_argument("--position", choices=["before", "after"], default="after")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_move)

    p = subparsers.add_parser("swap", help="Swap two blocks atomically")
    p.add_argument("file"); p.add_argument("--filter1", required=True)
    p.add_argument("--filter2", required=True)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_swap)

    nup = subparsers.add_parser("numbering", help="Numbering operations")
    nus = nup.add_subparsers(dest="subcommand", required=True)

    p = nus.add_parser("swap", help="Update chapter numbers after swap")
    p.add_argument("file"); p.add_argument("--ch-a", type=int, required=True)
    p.add_argument("--ch-b", type=int, required=True); p.add_argument("--names")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_numbering)

    p = nus.add_parser("update-figures", help="Renumber figures and tables by chapter")
    p.add_argument("file")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_numbering_update_figures)

    # ── content (append) ──
    ctp = subparsers.add_parser("content", help="Content operations")
    cts = ctp.add_subparsers(dest="subcommand", required=True)

    p = cts.add_parser("append", help="Append heading or body paragraph")
    p.add_argument("file"); p.add_argument("--type", required=True, choices=["heading", "body"])
    p.add_argument("--text", required=True, help="Paragraph text")
    p.add_argument("--after-paragraph", type=int, help="Insert after paragraph index")
    p.add_argument("--level", type=int, default=1, help="Heading level (1-9)")
    p.add_argument("--font"); p.add_argument("--font-west"); p.add_argument("--size", type=str)
    p.add_argument("--bold", action="store_true"); p.add_argument("--no-bold", action="store_true")
    p.add_argument("--italic", action="store_true"); p.add_argument("--no-italic", action="store_true")
    p.add_argument("--align", choices=["left", "center", "right", "justify"])
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_append_content)

    # ── build (create) ──
    cp = subparsers.add_parser("create", help="Create new documents")
    cs = cp.add_subparsers(dest="subcommand", required=True)

    p = cs.add_parser("blank", help="Create a blank document")
    p.add_argument("-o", "--output", required=True); p.add_argument("--paper", default="A4")
    p.add_argument("--margin-top", type=float, default=2.54)
    p.add_argument("--margin-bottom", type=float, default=2.54)
    p.add_argument("--margin-left", type=float, default=3.17)
    p.add_argument("--margin-right", type=float, default=3.17)
    p.add_argument("--orientation", choices=["portrait", "landscape"], default="portrait")
    p.set_defaults(func=cmd_create_blank)

    p = cs.add_parser("structure", help="Create from JSON structure")
    p.add_argument("-o", "--output", required=True); p.add_argument("--file")
    p.add_argument("--json"); p.set_defaults(func=cmd_create_structure)

    p = cs.add_parser("from-template", help="Create blank doc from template")
    p.add_argument("template", help="Template .docx path")
    p.add_argument("-o", "--output", required=True)
    p.set_defaults(func=cmd_create_from_template)

    # ── cite (ref) ──
    fp = subparsers.add_parser("ref", help="Reference management")
    fs = fp.add_subparsers(dest="subcommand", required=True)

    p = fs.add_parser("add", help="Add a reference source")
    p.add_argument("file"); p.add_argument("--key", required=True)
    p.add_argument("--text", required=True); p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_ref_add)

    p = fs.add_parser("cite", help="Insert a citation marker")
    p.add_argument("file"); p.add_argument("--key", required=True)
    p.add_argument("--paragraph", type=int, required=True)
    p.add_argument("--position", choices=["end", "start"], default="end")
    p.add_argument("--after-text", help="Insert after specific text in paragraph")
    p.add_argument("--style", choices=["inline", "footnote", "endnote"], default="inline",
                   help="Citation style (inline=[n], footnote=page bottom, endnote=doc end)")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_ref_cite)

    p = fs.add_parser("list", help="List references and citations")
    p.add_argument("file")
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_ref_list)

    p = fs.add_parser("generate", help="Generate bibliography section")
    p.add_argument("file"); p.add_argument("--heading", default="参考文献")
    p.add_argument("--style", choices=["numbered", "author-year"], default="numbered")
    p.add_argument("--after-heading"); p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_ref_generate)

    p = fs.add_parser("renumber", help="Renumber all citations")
    p.add_argument("file"); p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_ref_renumber)

    p = fs.add_parser("remove", help="Remove a reference")
    p.add_argument("file"); p.add_argument("--key", required=True)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_ref_remove)

    # ── field ──
    fldp = subparsers.add_parser("field", help="Field operations")
    flds = fldp.add_subparsers(dest="subcommand", required=True)

    p = flds.add_parser("list", help="List all fields")
    p.add_argument("file"); p.add_argument("--type")
    p.add_argument("--output-file", help="Write output to UTF-8 file")
    p.set_defaults(func=cmd_field_list)

    p = flds.add_parser("update", help="Update a field's instruction")
    p.add_argument("file"); p.add_argument("--id", type=int); p.add_argument("--type")
    p.add_argument("--nth", type=int, default=1); p.add_argument("--instruction", required=True)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_field_update)

    p = flds.add_parser("insert", help="Insert a new field")
    p.add_argument("file"); p.add_argument("--after-paragraph", type=int)
    p.add_argument("--type", required=True); p.add_argument("--format")
    p.add_argument("--location", choices=["body", "header", "footer"], default="body")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_field_insert)

    p = flds.add_parser("refresh", help="Refresh field cached results")
    p.add_argument("file"); p.add_argument("--type")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_field_refresh)

    # ── content: image / toc / section ──
    ip = subparsers.add_parser("image", help="Image operations")
    ips = ip.add_subparsers(dest="subcommand", required=True)

    p = ips.add_parser("insert", help="Insert an image")
    p.add_argument("file"); p.add_argument("--image", required=True, help="Image file path")
    p.add_argument("--after-paragraph", type=int, help="Insert after paragraph index")
    p.add_argument("--width", type=float, help="Width in cm")
    p.add_argument("--height", type=float, help="Height in cm")
    p.add_argument("--caption", help="Caption text (base text without number)")
    p.add_argument("--ref-id", help="Reference ID for auto-numbering (e.g. fig:arch)")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_insert_image)

    # ── equation ──
    ep = subparsers.add_parser("equation", help="Equation operations")
    eps = ep.add_subparsers(dest="subcommand", required=True)

    p = eps.add_parser("insert", help="Insert an OMML equation")
    p.add_argument("file"); p.add_argument("--paragraph", type=int, required=True,
                   help="Target paragraph index")
    p.add_argument("--omml", required=True, help="OMML XML string")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_insert_equation)

    tp = subparsers.add_parser("toc", help="Table of contents")
    tps = tp.add_subparsers(dest="subcommand", required=True)

    p = tps.add_parser("insert", help="Insert TOC field")
    p.add_argument("file"); p.add_argument("--after-paragraph", type=int)
    p.add_argument("--levels", type=int, default=3, help="Heading levels (1-9)")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_insert_toc)

    sbp = subparsers.add_parser("section", help="Section break operations")
    sbs = sbp.add_subparsers(dest="subcommand", required=True)

    p = sbs.add_parser("insert", help="Insert section break")
    p.add_argument("file"); p.add_argument("--after-paragraph", type=int, required=True)
    p.add_argument("--break-type", choices=["nextPage", "continuous", "evenPage", "oddPage"],
                   default="nextPage")
    p.add_argument("--header-text", help="Header text for the new section")
    p.add_argument("--footer-text", help="Footer text for the new section")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_insert_section)

    # ── table ──
    tbp = subparsers.add_parser("table", help="Table operations")
    tbs = tbp.add_subparsers(dest="subcommand", required=True)

    p = tbs.add_parser("create", help="Create a new table")
    p.add_argument("file"); p.add_argument("--headers", help='JSON array of header texts')
    p.add_argument("--rows", help='JSON array of row arrays')
    p.add_argument("--caption", help="Table caption (base text without number)")
    p.add_argument("--after-paragraph", type=int, help="Insert after paragraph index")
    p.add_argument("--ref-id", help="Reference ID for auto-numbering (e.g. tab:compare)")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_create)

    p = tbs.add_parser("add-row", help="Add a row")
    p.add_argument("file"); p.add_argument("--table", type=int, required=True, help="Table index")
    p.add_argument("--data", help='JSON array of cell values')
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_add_row)

    p = tbs.add_parser("add-column", help="Add a column")
    p.add_argument("file"); p.add_argument("--table", type=int, required=True)
    p.add_argument("--header", help="Header text")
    p.add_argument("--width", type=float, help="Column width in cm")
    p.add_argument("--strategy", choices=["split", "expand", "refuse"], default="split",
                   help="split: break merged cells (default). expand: grow merges. refuse: error on merges")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_add_column)

    p = tbs.add_parser("merge", help="Merge cells")
    p.add_argument("file"); p.add_argument("--table", type=int, required=True)
    p.add_argument("--row-start", type=int, required=True)
    p.add_argument("--col-start", type=int, required=True)
    p.add_argument("--row-end", type=int, required=True)
    p.add_argument("--col-end", type=int, required=True)
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_merge)

    p = tbs.add_parser("delete-row", help="Delete a table row")
    p.add_argument("file"); p.add_argument("--table", type=int, required=True)
    p.add_argument("--row", type=int, required=True, help="Row index to delete (0-based)")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_delete_row)

    p = tbs.add_parser("delete-column", help="Delete a table column")
    p.add_argument("file"); p.add_argument("--table", type=int, required=True)
    p.add_argument("--col", type=int, required=True, help="Column index to delete (0-based)")
    p.add_argument("--strategy", choices=["shrink", "refuse"], default="shrink",
                   help="shrink: reduce gridSpan on merged cells (default). refuse: error on merged cells")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_delete_column)

    p = tbs.add_parser("format-cell", help="Format a table cell")
    p.add_argument("file"); p.add_argument("--table", type=int, required=True)
    p.add_argument("--row", type=int, required=True, help="Row index (0-based)")
    p.add_argument("--col", type=int, required=True, help="Column index (0-based)")
    p.add_argument("--font-cn", help="East Asian font")
    p.add_argument("--font-en", help="Western font")
    p.add_argument("--size", type=str, help="Font size (e.g. 12, '小四')")
    p.add_argument("--bold", action="store_true"); p.add_argument("--no-bold", action="store_true")
    p.add_argument("--italic", action="store_true"); p.add_argument("--no-italic", action="store_true")
    p.add_argument("--align", choices=["left", "center", "right", "justify"])
    p.add_argument("--vertical-align", choices=["top", "center", "bottom"])
    p.add_argument("--shading", help="Hex fill color (e.g. D9E2F3)")
    p.add_argument("--width", type=float, help="Cell width in cm")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_table_format_cell)

    # ── xref ──
    xrp = subparsers.add_parser("xref", help="Cross-reference operations")
    xrs = xrp.add_subparsers(dest="subcommand", required=True)

    p = xrs.add_parser("insert", help="Insert cross-reference")
    p.add_argument("file"); p.add_argument("--type", required=True,
                   choices=["figure", "table", "equation", "section"])
    p.add_argument("--id", required=True, help="Reference ID")
    p.add_argument("--text", required=True, help="Display text")
    p.add_argument("--paragraph", type=int, required=True)
    p.add_argument("--position", choices=["end", "start"], default="end")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_xref_insert)

    p = xrs.add_parser("update", help="Update cross-reference text")
    p.add_argument("file"); p.add_argument("--labels", required=True,
                   help='JSON dict {ref_id: new_text}')
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_xref_update)

    # ── footnote ──
    p = subparsers.add_parser("footnote", help="Insert footnote")
    p.add_argument("file"); p.add_argument("--paragraph", type=int, required=True)
    p.add_argument("--text", required=True, help="Footnote text")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_insert_footnote)

    # ── note (convert) ──
    np = subparsers.add_parser("note", help="Note operations")
    ns = np.add_subparsers(dest="subcommand", required=True)

    p = ns.add_parser("convert", help="Convert footnotes to endnotes or vice versa")
    p.add_argument("file")
    p.add_argument("--to", required=True, choices=["footnote", "endnote"],
                   help="Target note type")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_convert_notes)

    # ── page-break ──
    p = subparsers.add_parser("page-break", help="Insert a page break")
    p.add_argument("file"); p.add_argument("--after-paragraph", type=int,
                   help="Insert after paragraph index")
    p.add_argument("-o", "--output"); p.set_defaults(func=cmd_page_break)

    # ── merge ──
    p = subparsers.add_parser("merge", help="Merge multiple documents")
    p.add_argument("files", nargs="+", help="Input .docx files (in order)")
    p.add_argument("-o", "--output", required=True); p.set_defaults(func=cmd_merge)

    # ── style export/import ──
    p = subparsers.add_parser("style", help="Style export/import")
    p.add_argument("file")
    action_group = p.add_mutually_exclusive_group(required=True)
    action_group.add_argument("--export", dest="style_action", action="store_const", const="export",
                              help="Export styles to JSON")
    action_group.add_argument("--import", dest="style_action", action="store_const", const="import",
                              help="Import styles from JSON")
    p.add_argument("--to", help="Export to JSON file")
    p.add_argument("--from", dest="from_json", help="Import from JSON file")
    p.add_argument("-o", "--output"); p.set_defaults(func=lambda a: (
        cmd_style_export(a) if a.style_action == 'export' else cmd_style_import(a)
    ))

    # ── Execute ──
    args = parser.parse_args()
    setup_logging(verbose=args.verbose, quiet=args.quiet)
    logger.debug(f'Command: {args.command}')
    _auto_output(args)
    # Wrap in transaction for file-modifying operations
    handler = transactional(args.func) if hasattr(args, 'file') else args.func
    handler(args)


if __name__ == "__main__":
    main()
