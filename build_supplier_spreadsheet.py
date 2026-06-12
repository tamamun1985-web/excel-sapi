#!/usr/bin/env python3
"""
Build "Data_Supplier_Peternakan.xlsx" as a raw Office Open XML (.xlsx) package.

No third-party libraries are required (openpyxl / exceljs are unavailable in this
sandbox), so we assemble the ZIP/XML structure by hand.

Features implemented:
  * 2 sheets: "Supplier Pakan" and "Supplier Bakalan"
  * Frozen header row (freeze panes)
  * Excel Tables with auto-filter on every column (data auto-extends)
  * Currency (Rupiah) format on "Harga Terakhir"
  * Date format on "Tanggal Update Harga"
  * Text format on "No HP" (keeps leading zeros)
  * Dropdown data validation on category columns
  * Header styling + borders + bestFit column widths
"""

import zipfile
from xml.sax.saxutils import escape

N_ROWS = 50  # number of pre-formatted empty data rows

# ---------------------------------------------------------------------------
# Sheet definitions
# ---------------------------------------------------------------------------
# style keys: "text", "currency", "date", "general"
SHEETS = [
    {
        "name": "Supplier Pakan",
        "table": "TabelSupplierPakan",
        "headers": [
            ("Nama Supplier", "general", 22),
            ("No HP", "text", 16),
            ("Alamat", "general", 32),
            ("Jenis Pakan", "general", 16),
            ("Produk yang Dijual", "general", 24),
            ("Satuan", "general", 12),
            ("Harga Terakhir", "currency", 16),
            ("Tanggal Update Harga", "date", 20),
            ("Catatan", "general", 40),
        ],
        "validations": [
            ("D", ["Hijauan", "Konsentrat", "Dedak", "Ampas Tahu", "Ampas Kedelai",
                   "Bungkil Kopra", "Molases", "Singkong", "Mineral", "Vitamin",
                   "Obat", "Lainnya"]),
        ],
    },
    {
        "name": "Supplier Bakalan",
        "table": "TabelSupplierBakalan",
        "headers": [
            ("Nama Supplier", "general", 22),
            ("No HP", "text", 16),
            ("Alamat", "general", 32),
            ("Jenis Sapi", "general", 14),
            ("Bangsa Sapi", "general", 16),
            ("Kisaran Bobot (kg)", "general", 18),
            ("Harga Terakhir (Rp/kg BB)", "currency", 24),
            ("Tanggal Update Harga", "date", 20),
            ("Catatan", "general", 40),
        ],
        "validations": [
            ("D", ["Jantan", "Betina"]),
            ("E", ["Limousin", "Simmental", "PO", "Brahman", "Angus",
                   "Bali", "Madura", "Lainnya"]),
        ],
    },
]

# cellXfs style indices
S_DEFAULT = 0
S_HEADER = 1
S_GENERAL = 2
S_TEXT = 3
S_CURRENCY = 4
S_DATE = 5

BODY_STYLE = {
    "general": S_GENERAL,
    "text": S_TEXT,
    "currency": S_CURRENCY,
    "date": S_DATE,
}


def col_letter(idx):  # 1-based
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------------------
# styles.xml
# ---------------------------------------------------------------------------
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="&quot;Rp&quot;#,##0"/>
    <numFmt numFmtId="165" formatCode="dd/mm/yyyy"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2E5B3E"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFBFBFBF"/></left>
      <right style="thin"><color rgb="FFBFBFBF"/></right>
      <top style="thin"><color rgb="FFBFBFBF"/></top>
      <bottom style="thin"><color rgb="FFBFBFBF"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"""


def build_sheet_xml(sheet):
    headers = sheet["headers"]
    ncols = len(headers)
    last_col = col_letter(ncols)
    last_row = N_ROWS + 1
    ref = f"A1:{last_col}{last_row}"

    # columns widths (bestFit)
    cols = ['<cols>']
    for i, (_, _, width) in enumerate(headers, start=1):
        cols.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1" bestFit="1"/>')
    cols.append('</cols>')
    cols_xml = "".join(cols)

    # header row
    rows = []
    header_cells = []
    for i, (title, _, _) in enumerate(headers, start=1):
        ref_c = f"{col_letter(i)}1"
        header_cells.append(
            f'<c r="{ref_c}" s="{S_HEADER}" t="inlineStr"><is><t xml:space="preserve">{escape(title)}</t></is></c>'
        )
    rows.append(f'<row r="1" ht="30" customHeight="1">{"".join(header_cells)}</row>')

    # empty pre-formatted data rows
    for r in range(2, last_row + 1):
        cells = []
        for i, (_, kind, _) in enumerate(headers, start=1):
            s = BODY_STYLE[kind]
            cells.append(f'<c r="{col_letter(i)}{r}" s="{s}"/>')
        rows.append(f'<row r="{r}">{"".join(cells)}</row>')
    sheet_data = "".join(rows)

    # data validations
    dv_parts = []
    for col, options in sheet["validations"]:
        formula = "&quot;" + escape(",".join(options)) + "&quot;"
        sqref = f"{col}2:{col}{last_row}"
        dv_parts.append(
            f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
            f'errorTitle="Input tidak valid" error="Pilih nilai dari daftar dropdown." '
            f'promptTitle="Pilihan" prompt="Pilih salah satu opsi" sqref="{sqref}">'
            f'<formula1>{formula}</formula1></dataValidation>'
        )
    dv_xml = ""
    if dv_parts:
        dv_xml = f'<dataValidations count="{len(dv_parts)}">{"".join(dv_parts)}</dataValidations>'

    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="{ref}"/>
<sheetViews>
  <sheetView tabSelected="1" workbookViewId="0">
    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
  </sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
{cols_xml}
<sheetData>{sheet_data}</sheetData>
{dv_xml}
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
<tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>"""
    return xml


def build_table_xml(sheet, table_id):
    headers = sheet["headers"]
    ncols = len(headers)
    last_col = col_letter(ncols)
    ref = f"A1:{last_col}{N_ROWS + 1}"
    cols_xml = "".join(
        f'<tableColumn id="{i}" name="{escape(title)}"/>'
        for i, (title, _, _) in enumerate(headers, start=1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{table_id}" name="{sheet['table']}" displayName="{sheet['table']}" ref="{ref}" totalsRowShown="0">
  <autoFilter ref="{ref}"/>
  <tableColumns count="{ncols}">{cols_xml}</tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>"""


# ---------------------------------------------------------------------------
# Package-level XML
# ---------------------------------------------------------------------------
CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Supplier Pakan" sheetId="1" r:id="rId1"/>
    <sheet name="Supplier Bakalan" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

SHEET_RELS_TMPL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table{tid}.xml"/>
</Relationships>"""


def main():
    out = "Data_Supplier_Peternakan.xlsx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK_XML)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", STYLES_XML)
        for idx, sheet in enumerate(SHEETS, start=1):
            z.writestr(f"xl/worksheets/sheet{idx}.xml", build_sheet_xml(sheet))
            z.writestr(f"xl/worksheets/_rels/sheet{idx}.xml.rels",
                       SHEET_RELS_TMPL.format(tid=idx))
            z.writestr(f"xl/tables/table{idx}.xml", build_table_xml(sheet, idx))
    print("Saved", out)


if __name__ == "__main__":
    main()
