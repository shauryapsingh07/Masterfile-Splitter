#!/usr/bin/env python3
"""Split a marketplace masterfile (.xlsx) into Complete and Incomplete SKUs.

A SKU's mandatory attributes are the cells the template itself highlights in
red. The template paints a cell red with a conditional-formatting rule that
looks up the SKU's category in the "Columns" sheet; this tool reads that same
rule, so it marks exactly the cells Excel shows in red, per SKU.

A SKU is Complete when every red cell in its row has something in it.
Cell values are never judged or changed: any non-blank entry counts as filled.

Used two ways:
    * command line:  python split_masterfile.py "Tesco UK All Catalog.xlsx" -o split
    * Streamlit app: app.py calls split_workbook()

Examples:
    python split_masterfile.py "Tesco UK All Catalog.xlsx"
    python split_masterfile.py "Tesco UK All Catalog.xlsx" "Macy's All Catalog.xlsx" -o split
    python split_masterfile.py -l files.txt
"""
import argparse
import csv
import html
import io
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
DATA_SHEET = "Data"
HEADER_ROWS = 2            # row 1 = attribute labels, row 2 = attribute codes
CATEGORY_COL = 1           # column A holds the SKU's category
MISSING_HEADER = "Missing Attributes"

ROW_RE = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
ATTR_RE = re.compile(r'([\w:]+)\s*=\s*"([^"]*)"')
T_RE = re.compile(r"<t\b[^>]*>(.*?)</t>", re.S)
V_RE = re.compile(r"<v>(.*?)</v>", re.S)
REF_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$")
CF_RE = re.compile(r"<conditionalFormatting\b[^>]*>(.*?)</conditionalFormatting>", re.S)
CFRULE_RE = re.compile(r"<cfRule\b([^>]*)>(.*?)</cfRule>", re.S)
FORMULA_RE = re.compile(r"<formula>(.*?)</formula>", re.S)
RULE_VALUE_RE = re.compile(r'=\s*"([^"]*)"\s*$')
LOOKUP_SHEET_RE = re.compile(
    r"OFFSET\(\s*'?([^'!]+)'?!\$A\$1\s*,\s*0\s*,\s*(\d+)\s*,\s*1\s*,\s*(\d+)\s*\)")


# ---------------------------------------------------------------- helpers

def col_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def attrs(text):
    return dict(ATTR_RE.findall(text))


def is_blank(value):
    return value is None or value.strip() == ""


def sheet_paths(z):
    """Map sheet name -> zip path of its XML part."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    paths = {}
    for sheet in wb.find("m:sheets", NS):
        target = targets.get(sheet.get("{%s}id" % NS["r"]), "")
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        paths[sheet.get("name")] = posixpath.normpath(path)
    return paths


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    t_tag = "{%s}t" % NS["m"]
    return ["".join(t.text or "" for t in si.iter(t_tag)) for si in root.findall("m:si", NS)]


def cell_text(cell_attrs, inner, sst):
    if not inner:
        return ""
    kind = cell_attrs.get("t", "n")
    if kind == "inlineStr":
        return html.unescape("".join(T_RE.findall(inner)))
    v = V_RE.search(inner)
    if not v:
        return ""
    value = html.unescape(v.group(1))
    if kind == "s":
        try:
            return sst[int(value)]
        except (ValueError, IndexError):
            return ""
    return value


def row_number(row_xml, previous):
    head = row_xml[:row_xml.index(">") + 1]
    r = attrs(head).get("r")
    return int(r) if r else previous + 1


def parse_cells(row_xml, sst):
    """Return {column number: text} for one <row>."""
    cells, col = {}, 0
    for m in CELL_RE.finditer(row_xml):
        a = attrs(m.group(1))
        ref = REF_RE.match(a.get("r", ""))
        col = col_to_num(ref.group(1)) if ref else col + 1
        cells[col] = cell_text(a, m.group(2), sst)
    return cells


def split_sheet_xml(xml):
    """Split worksheet XML into (prefix, [row xml...], suffix)."""
    m = re.search(r"<sheetData\b[^>]*?(/?)>", xml)
    if not m:
        raise ValueError("worksheet has no <sheetData>")
    if m.group(1) == "/":           # <sheetData/> : no rows at all
        return xml[:m.start()] + "<sheetData>", [], "</sheetData>" + xml[m.end():]
    end = xml.index("</sheetData>", m.end())
    return xml[:m.end()], ROW_RE.findall(xml, m.end(), end), xml[end:]


def read_rows(xml, sst):
    """Return list of (row number, row xml, {col: text})."""
    _, rows, _ = split_sheet_xml(xml)
    out, prev = [], 0
    for row_xml in rows:
        prev = row_number(row_xml, prev)
        out.append((prev, row_xml, parse_cells(row_xml, sst)))
    return out


# ---------------------------------------------------------- red rule

def dxf_colors(z):
    """Fill colour (RRGGBB or None) of each differential style used by conditional formatting."""
    root = ET.fromstring(z.read("xl/styles.xml"))
    dxfs = root.find("m:dxfs", NS)
    colors = []
    for dxf in (dxfs if dxfs is not None else []):
        rgb = None
        for tag in ("bgColor", "fgColor"):
            el = dxf.find(".//m:fill/m:patternFill/m:%s" % tag, NS)
            if el is not None and el.get("rgb"):
                rgb = el.get("rgb")[-6:]
                break
            if el is not None and el.get("indexed") in ("2", "10"):   # legacy palette red
                rgb = "FF0000"
                break
        colors.append(rgb)
    return colors


def is_red(rgb):
    if not rgb:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return r >= 0xC0 and r - max(g, b) >= 0x30


def find_red_rule(sheet_xml, colors):
    """Read the red conditional-formatting rule: which lookup value paints a cell red,
    which sheet it looks up, and which columns of that sheet hold the categories."""
    red_values, lookup = set(), None
    for block in CF_RE.findall(sheet_xml):
        for rule_attrs, body in CFRULE_RE.findall(block):
            dxf = attrs(rule_attrs).get("dxfId")
            if dxf is None or int(dxf) >= len(colors) or not is_red(colors[int(dxf)]):
                continue
            formula = html.unescape("".join(FORMULA_RE.findall(body)))
            value = RULE_VALUE_RE.search(formula)
            where = LOOKUP_SHEET_RE.search(formula)
            if not value or not where:
                raise ValueError("found a red highlight rule, but not in the expected "
                                 "category-lookup form: " + formula[:120])
            red_values.add(value.group(1).casefold())
            first = 1 + int(where.group(2))
            lookup = (where.group(1), first, first + int(where.group(3)) - 1)
    if not red_values:
        raise ValueError("no red highlight rule found in the Data sheet")
    return red_values, lookup


def red_columns_by_category(z, paths, sst, codes, red_values, lookup):
    """For every category: the Data columns that are red for a SKU in that category."""
    sheet, first_col, last_col = lookup
    if sheet not in paths:
        raise ValueError(f"lookup sheet '{sheet}' named in the red rule is missing")
    rows = read_rows(z.read(paths[sheet]).decode("utf-8"), sst)
    if not rows:
        raise ValueError(f"lookup sheet '{sheet}' is empty")
    header = rows[0][2]
    categories = {}
    for col in range(first_col, last_col + 1):
        name = header.get(col, "")
        if name and name.casefold() not in categories:
            categories[name.casefold()] = col
    attr_rows = {}
    for _, _, cells in rows[1:]:
        code = cells.get(1, "")
        if code and code.casefold() not in attr_rows:
            attr_rows[code.casefold()] = cells
    red = {}
    for cat, cat_col in categories.items():
        red[cat] = [col for col, code in sorted(codes.items())
                    if code and attr_rows.get(code.casefold(), {}).get(cat_col, "").casefold() in red_values]
    return red


# ------------------------------------------------------------ writing

def renumber(row_xml, new_r, drop_spans=False):
    head_end = row_xml.index(">") + 1
    head, body = row_xml[:head_end], row_xml[head_end:]
    if re.search(r'\br="\d+"', head):
        head = re.sub(r'\br="\d+"', f'r="{new_r}"', head, count=1)
    else:
        head = head.replace("<row", f'<row r="{new_r}"', 1)
    if drop_spans:
        head = re.sub(r'\sspans="[^"]*"', "", head)
    body = re.sub(r'(<c\b[^>]*?\br="\$?[A-Z]+)\$?\d+(")', rf"\g<1>{new_r}\2", body)
    return head + body


def append_cell(row_xml, cell):
    if not row_xml.rstrip().endswith("</row>"):          # self-closing <row .../>
        return row_xml.rstrip()[:-2] + ">" + cell + "</row>"
    i = row_xml.rindex("</row>")
    return row_xml[:i] + cell + row_xml[i:]


def text_cell(ref, text, style=None):
    s = f' s="{style}"' if style else ""
    esc = html.escape(text, quote=False)
    return f'<c r="{ref}" t="inlineStr"{s}><is><t xml:space="preserve">{esc}</t></is></c>'


def build_sheet(prefix, header_rows, data_rows, suffix, last_col, missing=None):
    """Reassemble the Data sheet with the given SKU rows renumbered from row 3.
    missing: list of 'Missing Attributes' texts (one per data row) for the Incomplete file."""
    extra_col = last_col + 1 if missing is not None else None
    parts = []
    for r, row_xml in header_rows:
        if extra_col:
            last = list(CELL_RE.finditer(row_xml))
            style = attrs(last[-1].group(1)).get("s") if last else None
            row_xml = append_cell(renumber(row_xml, r, drop_spans=True),
                                  text_cell(f"{num_to_col(extra_col)}{r}", MISSING_HEADER, style))
        parts.append(row_xml)
    for i, row_xml in enumerate(data_rows):
        new_r = HEADER_ROWS + 1 + i
        row_xml = renumber(row_xml, new_r, drop_spans=bool(extra_col))
        if extra_col:
            row_xml = append_cell(row_xml, text_cell(f"{num_to_col(extra_col)}{new_r}", missing[i]))
        parts.append(row_xml)

    last_row = HEADER_ROWS + len(data_rows)
    end_col = num_to_col(extra_col or last_col)

    def new_dim(m):
        start = m.group(2).split(":")[0]
        ref = REF_RE.match(start)
        if not ref or int(ref.group(2)) > last_row:
            start = "A1"
        return f"{m.group(1)}{start}:{end_col}{last_row}{m.group(3)}"

    prefix = re.sub(r'(<dimension\s+ref=")([^"]*)(")', new_dim, prefix, count=1)
    if extra_col and "</cols>" in prefix:
        maxes = [int(x) for x in re.findall(r'<col\b[^>]*?\bmax="(\d+)"', prefix)]
        if not maxes or max(maxes) < extra_col:
            prefix = prefix.replace(
                "</cols>", f'<col min="{extra_col}" max="{extra_col}" width="60" customWidth="1"/></cols>', 1)
    return prefix + "\n".join(parts) + suffix


def xlsx_bytes(src, replace_name, replace_text):
    """Copy every part of the source workbook unchanged, except the one replaced sheet."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in src.infolist():
            data = replace_text.encode("utf-8") if info.filename == replace_name else src.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    return buf.getvalue()


def summary_csv(result):
    total = result["total"]
    pct = lambda n: f"{(100.0 * n / total):.1f}%" if total else "0.0%"
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(["Masterfile", result["name"]])
    w.writerow(["Total SKUs", total])
    w.writerow(["Complete SKUs", result["complete"], pct(result["complete"])])
    w.writerow(["Incomplete SKUs", result["incomplete"], pct(result["incomplete"])])
    w.writerow(["Blank rows skipped", result["blank_rows"]])
    w.writerow([])
    w.writerow(["Rank", "Missing attribute", "Column", "Attribute code", "SKUs missing it", "% of all SKUs"])
    for rank, m in enumerate(result["missing_counts"], 1):
        w.writerow([rank, m["attribute"], m["column"], m["code"], m["skus"], pct(m["skus"])])
    return "﻿" + buf.getvalue()     # BOM so Excel opens it as UTF-8


# ------------------------------------------------------------ main logic

def split_workbook(source, name):
    """Split one masterfile.

    source: a path or a binary file-like object holding the .xlsx.
    name:   the original file name (used for output names).
    Returns a dict with the counts, the ranked missing attributes and the
    output files as bytes: complete_xlsx, incomplete_xlsx, summary_csv.
    """
    with zipfile.ZipFile(source) as z:
        paths = sheet_paths(z)
        if DATA_SHEET not in paths:
            raise ValueError(f"no '{DATA_SHEET}' sheet (sheets: {', '.join(paths)})")
        sst = shared_strings(z)
        data_name = paths[DATA_SHEET]
        data_xml = z.read(data_name).decode("utf-8")
        red_values, lookup = find_red_rule(data_xml, dxf_colors(z))

        prefix, row_xmls, suffix = split_sheet_xml(data_xml)
        rows, prev = [], 0
        for row_xml in row_xmls:
            prev = row_number(row_xml, prev)
            rows.append((prev, row_xml))
        header = {r: parse_cells(x, sst) for r, x in rows if r <= HEADER_ROWS}
        labels_row, codes = header.get(1, {}), header.get(HEADER_ROWS, {})
        labels = {c: labels_row.get(c) or codes.get(c, "") for c in set(labels_row) | set(codes)}
        if not codes:
            raise ValueError(f"row {HEADER_ROWS} of the Data sheet has no attribute codes")
        red = red_columns_by_category(z, paths, sst, codes, red_values, lookup)

        header_rows = [(r, x) for r, x in rows if r <= HEADER_ROWS]
        complete, incomplete, missing_texts = [], [], []
        counts, blank_rows, last_col = Counter(), 0, max(list(labels) + [1])
        cat_label = labels.get(CATEGORY_COL) or "Category Code"
        for r, row_xml in rows:
            if r <= HEADER_ROWS:
                continue
            cells = parse_cells(row_xml, sst)
            if all(is_blank(v) for v in cells.values()):
                blank_rows += 1
                continue
            last_col = max([last_col] + list(cells))
            category = (cells.get(CATEGORY_COL) or "").casefold()
            if is_blank(category):
                missing, note = [CATEGORY_COL], ""
            elif category not in red:
                missing, note = [CATEGORY_COL], " (category not found in the Columns sheet)"
            else:
                missing, note = [c for c in red[category] if is_blank(cells.get(c))], ""
            if missing:
                counts.update(missing)
                names = [labels.get(c) or num_to_col(c) for c in missing]
                if note:
                    names[0] = cat_label + note
                incomplete.append(row_xml)
                missing_texts.append(", ".join(names))
            else:
                complete.append(row_xml)

        complete_xlsx = xlsx_bytes(z, data_name, build_sheet(prefix, header_rows, complete, suffix, last_col))
        incomplete_xlsx = xlsx_bytes(z, data_name, build_sheet(prefix, header_rows, incomplete, suffix,
                                                               last_col, missing_texts))

    stem, ext = os.path.splitext(os.path.basename(name))
    result = {
        "name": os.path.basename(name),
        "stem": stem,
        "ext": ext or ".xlsx",
        "total": len(complete) + len(incomplete),
        "complete": len(complete),
        "incomplete": len(incomplete),
        "blank_rows": blank_rows,
        "missing_counts": [{"attribute": labels.get(c, ""), "column": num_to_col(c),
                            "code": codes.get(c, ""), "skus": n} for c, n in counts.most_common()],
        "complete_xlsx": complete_xlsx,
        "incomplete_xlsx": incomplete_xlsx,
    }
    result["summary_csv"] = summary_csv(result).encode("utf-8")
    return result


def output_names(result):
    stem, ext = result["stem"], result["ext"]
    return {"complete_xlsx": f"{stem} - Complete{ext}",
            "incomplete_xlsx": f"{stem} - Incomplete{ext}",
            "summary_csv": f"{stem} - Summary.csv"}


# ------------------------------------------------------------ command line

def split_file(path, args):
    result = split_workbook(path, path)
    os.makedirs(args.out_dir, exist_ok=True)
    saved = []
    for key, filename in output_names(result).items():
        out = os.path.join(args.out_dir, filename)
        if os.path.abspath(out) == os.path.abspath(path):
            raise ValueError("output would overwrite the input file; choose another --out-dir")
        with open(out, "wb") as f:
            f.write(result[key])
        saved.append(out)

    print(path)
    print(f"  SKUs: {result['total']}  Complete: {result['complete']}  Incomplete: {result['incomplete']}"
          + (f"  (blank rows skipped: {result['blank_rows']})" if result["blank_rows"] else ""))
    for m in result["missing_counts"][:args.top]:
        print(f"    missing {m['attribute']} [{m['column']}]: {m['skus']} SKU(s)")
    for out in saved:
        print(f"  Saved {out}")
    return result["total"], result["complete"], result["incomplete"]


def main():
    parser = argparse.ArgumentParser(description="Split masterfiles into Complete and Incomplete SKUs "
                                                 "using the template's red (mandatory) cells.")
    parser.add_argument("files", nargs="*", help="masterfile .xlsx paths")
    parser.add_argument("-l", "--list-file", help="text file with one masterfile path per line")
    parser.add_argument("-o", "--out-dir", default="split_output", help="output folder (default: split_output)")
    parser.add_argument("--top", type=int, default=10, help="how many top missing attributes to print (default 10)")
    args = parser.parse_args()

    targets = list(args.files)
    if args.list_file:
        with open(args.list_file, encoding="utf-8") as f:
            targets += [ln.strip().strip('"') for ln in f if ln.strip() and not ln.startswith("#")]
    if not targets:
        parser.error("provide at least one masterfile or --list-file")

    done = failed = 0
    grand = [0, 0, 0]
    for target in targets:
        try:
            grand = [a + b for a, b in zip(grand, split_file(target, args))]
            done += 1
        except PermissionError as e:
            print(f"Failed {target}: {e} (is an output file open in Excel?)", file=sys.stderr)
            failed += 1
        except Exception as e:
            print(f"Failed {target}: {e}", file=sys.stderr)
            failed += 1

    print(f"Done. Files split: {done}, failed: {failed}. "
          f"SKUs: {grand[0]}, Complete: {grand[1]}, Incomplete: {grand[2]}. Output: {args.out_dir}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
