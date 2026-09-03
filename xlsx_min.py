"""
Минимальный ридер/райтер .xlsx на стандартной библиотеке Python (без openpyxl и т.п.) —
в среде, где выполняется скрипт, нет доступа в интернет для установки пакетов.

Поддерживает то, что нужно для этой задачи:
  - чтение листов с точным позиционированием ячеек по атрибуту r="A1" (важно для
    "разреженных" строк, где часть ячеек пустая и просто отсутствует в XML —
    например строки pivot-таблиц с пустыми "продолженными" колонками);
  - разделяемые строки (sharedStrings) и inline-строки;
  - запись простой книги из нескольких листов с базовым форматированием
    (жирная шапка, заливка ячеек по статусу).
"""
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def col_letters_to_index(letters):
    """'A' -> 0, 'B' -> 1, 'AA' -> 26 ..."""
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def col_index_to_letters(idx):
    """0 -> 'A', 1 -> 'B', 26 -> 'AA' ..."""
    idx += 1
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _split_cell_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


class Workbook:
    """Открывает .xlsx для чтения и даёт доступ к листам по имени."""

    def __init__(self, path):
        self.path = path
        self._zip = zipfile.ZipFile(path)
        names = self._zip.namelist()
        self.shared_strings = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(self._zip.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", NS):
                texts = si.findall(".//m:t", NS)
                self.shared_strings.append("".join(t.text or "" for t in texts))
        wb_root = ET.fromstring(self._zip.read("xl/workbook.xml"))
        sheets_el = wb_root.find("m:sheets", NS)
        sheet_list = [(s.get("name"), s.get(f"{R_NS}id")) for s in sheets_el]
        rels_root = ET.fromstring(self._zip.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target") for r in rels_root}
        self.sheet_paths = {}
        for name, rid in sheet_list:
            target = rel_map[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target
            self.sheet_paths[name] = target
        self.sheet_names = [n for n, _ in sheet_list]

    def close(self):
        self._zip.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def _cell_value(self, c):
        t = c.get("t")
        if t == "inlineStr":
            is_el = c.find("m:is", NS)
            if is_el is None:
                return None
            texts = is_el.findall(".//m:t", NS)
            return "".join(x.text or "" for x in texts)
        v = c.find("m:v", NS)
        if v is None or v.text is None:
            return None
        val = v.text
        if t == "s":
            return self.shared_strings[int(val)]
        if t == "str":
            return val
        return val  # numeric (as string; caller converts if needed)

    def read_sheet(self, sheet_name):
        """
        Возвращает список строк. Каждая строка — dict {col_idx(int): value}.
        Индексация колонок 0-based, вычисляется из атрибута r="C5" ячейки,
        поэтому "дырки" в разреженных строках не сбивают смещение.
        Пустые строки листа пропускаются (не входят в результат) — используйте
        row_number из атрибута r строки, если нужен номер строки Excel (1-based).
        """
        sheet_path = self.sheet_paths[sheet_name]
        root = ET.fromstring(self._zip.read(sheet_path))
        sheet_data = root.find("m:sheetData", NS)
        rows = []
        for row_el in sheet_data.findall("m:row", NS):
            row_num = int(row_el.get("r"))
            row = {}
            for c in row_el.findall("m:c", NS):
                ref = c.get("r")
                letters, _ = _split_cell_ref(ref)
                col_idx = col_letters_to_index(letters)
                row[col_idx] = self._cell_value(c)
            rows.append((row_num, row))
        return rows


def row_get(row, idx, default=None):
    v = row.get(idx, default)
    return v if v is not None else default


def to_float(val, default=0.0):
    if val is None or val == "":
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------------------
# Запись простой книги
# ---------------------------------------------------------------------------

CONTENT_TYPES_TMPL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{sheet_overrides}
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1">
<numFmt numFmtId="164" formatCode="#,##0.00"/>
</numFmts>
<fonts count="3">
<font><sz val="10"/><name val="Arial"/></font>
<font><sz val="10"/><name val="Arial"/><b/></font>
<font><sz val="10"/><name val="Arial"/><b/><color rgb="FFFFFFFF"/></font>
</fonts>
<fills count="7">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6E0B4"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="5" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="6" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
</styleSheet>"""

# Стили (индексы cellXfs выше):
STYLE_DEFAULT = 0
STYLE_HEADER = 1
STYLE_NUMBER = 2
STYLE_NUMBER_OK = 3       # зелёный (совпадение)
STYLE_NUMBER_MISMATCH = 4  # красный/оранжевый (расхождение)
STYLE_WARN = 5             # жёлтый (не найдено / нет кода)
STYLE_GRAY = 6             # серый (исключено, справочно)
STYLE_BOLD = 7


def _cell_xml(row_idx, col_idx, value, style=None):
    ref = f"{col_index_to_letters(col_idx)}{row_idx}"
    s_attr = f' s="{style}"' if style else ""
    if value is None:
        return f'<c r="{ref}"{s_attr}/>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s_attr}><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


class SheetSpec:
    def __init__(self, name, headers, rows, col_styles=None, col_widths=None, header_comments=None):
        """
        headers: список заголовков колонок (строка)
        rows: список списков значений (той же длины, что headers)
        col_styles: список функций(row_dict) -> style_id по колонке, либо None
                     (по умолчанию STYLE_DEFAULT/STYLE_NUMBER определяется по типу)
        header_comments: список строк той же длины, что headers (или None на
                          позициях без комментария) — заметки Excel к ячейкам
                          заголовка (откуда берётся колонка и как считается).
        """
        self.name = name
        self.headers = headers
        self.rows = rows
        self.col_styles = col_styles or [None] * len(headers)
        self.col_widths = col_widths
        self.header_comments = header_comments


def _sheets_with_comments(sheets):
    return [
        i for i, s in enumerate(sheets)
        if s.header_comments and any(c for c in s.header_comments)
    ]


def _comments_xml(spec):
    items = []
    for ci, comment in enumerate(spec.header_comments):
        if not comment:
            continue
        ref = f"{col_index_to_letters(ci)}1"
        text = escape(str(comment))
        items.append(
            f'<comment ref="{ref}" authorId="0"><text><r><t xml:space="preserve">{text}</t></r></text></comment>'
        )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Сверка</author></authors>
<commentList>{"".join(items)}</commentList>
</comments>"""


def _vml_xml(spec):
    shapes = []
    shape_id = 1
    for ci, comment in enumerate(spec.header_comments):
        if not comment:
            continue
        shape_id += 1
        col0 = ci
        row0 = 0  # заголовок — строка 1 в Excel = индекс 0
        anchor = f"{col0+1}, 15, {row0}, 2, {col0+4}, 15, {row0+5}, 4"
        shapes.append(f"""<v:shape id="_x0000_s{shape_id}" type="#_x0000_t202"
 style='position:absolute;margin-left:80pt;margin-top:2pt;width:200pt;height:90pt;z-index:{shape_id};visibility:hidden'
 fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/>
<v:shadow on="t" color="black" obscured="t"/>
<v:path o:connecttype="none"/>
<v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>
<x:ClientData ObjectType="Note">
<x:MoveWithCells/>
<x:SizeWithCells/>
<x:Anchor>{anchor}</x:Anchor>
<x:AutoFill>False</x:AutoFill>
<x:Row>{row0}</x:Row>
<x:Column>{col0}</x:Column>
</x:ClientData>
</v:shape>""")
    return f"""<xml xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
<v:stroke joinstyle="miter"/>
<v:path gradientshapeok="t" o:connecttype="rect"/>
</v:shapetype>
{"".join(shapes)}
</xml>"""


def _sheet_rels_xml(vml_target, comments_target):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="{vml_target}"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="{comments_target}"/>
</Relationships>"""


def write_workbook(path, sheets):
    """sheets: список SheetSpec"""
    sheet_overrides = "\n".join(
        f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(len(sheets))
    )
    comment_sheet_idxs = _sheets_with_comments(sheets)
    comment_overrides = "\n".join(
        f'<Override PartName="/xl/comments{i+1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'
        for i in comment_sheet_idxs
    )
    vml_default = (
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
        if comment_sheet_idxs else ""
    )
    sheet_overrides = sheet_overrides + "\n" + comment_overrides
    content_types = CONTENT_TYPES_TMPL.format(sheet_overrides=sheet_overrides)
    content_types = content_types.replace(
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>\n' + vml_default,
    )

    wb_sheets_xml = "\n".join(
        f'<sheet name="{escape(s.name)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, s in enumerate(sheets)
    )
    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{wb_sheets_xml}</sheets>
</workbook>"""

    rels_items = "\n".join(
        f'<Relationship Id="rId{i+1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{i+1}.xml"/>'
        for i in range(len(sheets))
    )
    rels_items += (
        f'\n<Relationship Id="rId{len(sheets)+1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        f'Target="styles.xml"/>'
    )
    workbook_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels_items}
</Relationships>"""

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, s in enumerate(sheets):
            has_comments = i in comment_sheet_idxs
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", _sheet_xml(s, legacy_drawing_rid="rId1" if has_comments else None))
            if has_comments:
                z.writestr(f"xl/comments{i+1}.xml", _comments_xml(s))
                z.writestr(f"xl/drawings/vmlDrawing{i+1}.vml", _vml_xml(s))
                z.writestr(
                    f"xl/worksheets/_rels/sheet{i+1}.xml.rels",
                    _sheet_rels_xml(f"../drawings/vmlDrawing{i+1}.vml", f"../comments{i+1}.xml"),
                )


def _sheet_xml(spec, legacy_drawing_rid=None):
    parts = []
    # шапка
    header_cells = "".join(
        _cell_xml(1, ci, h, STYLE_HEADER) for ci, h in enumerate(spec.headers)
    )
    parts.append(f'<row r="1">{header_cells}</row>')
    for ri, row in enumerate(spec.rows, start=2):
        cells = []
        for ci, val in enumerate(row):
            style = None
            if spec.col_styles[ci] is not None:
                style = spec.col_styles[ci](row)
            elif isinstance(val, (int, float)):
                style = STYLE_NUMBER
            cells.append(_cell_xml(ri, ci, val, style))
        parts.append(f'<row r="{ri}">{"".join(cells)}</row>')
    rows_xml = "".join(parts)

    cols_xml = ""
    if spec.col_widths:
        col_defs = "".join(
            f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>'
            for i, w in enumerate(spec.col_widths)
        )
        cols_xml = f"<cols>{col_defs}</cols>"

    n_cols = max(len(spec.headers), 1)
    n_rows = len(spec.rows) + 1
    dim_ref = f"A1:{col_index_to_letters(n_cols-1)}{n_rows}"

    legacy_drawing_xml = (
        f'<legacyDrawing r:id="{legacy_drawing_rid}"/>' if legacy_drawing_rid else ""
    )

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="{dim_ref}"/>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
{cols_xml}
<sheetData>{rows_xml}</sheetData>
{legacy_drawing_xml}
</worksheet>"""
