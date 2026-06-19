# -*- coding: utf-8 -*-
"""
Generator Sistem Excel Penggemukan Sapi Potong.
Membangun file .xlsx (OOXML) langsung tanpa library eksternal.
Keputusan desain:
- HPP akurat (biaya menempel ke sapi)
- Alokasi operasional & penyusutan berbasis hari pemeliharaan (cattle-days)
- Alokasi pakan dalam kelompok berbasis hari (Opsi A)
- Valuasi pakan: rata-rata tertimbang
- Kapasitas 10 tahun; tanpa hutang/piutang
"""
import zipfile, datetime, xml.dom.minidom as minidom

OUT = "/projects/sandbox/Sistem_Penggemukan_Sapi.xlsx"

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def col_letter(n):  # 1-indexed
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def serial(y, m, d):
    base = datetime.date(1899, 12, 30)
    return (datetime.date(y, m, d) - base).days

# ----------------------------------------------------------------------------
# Style manager
# ----------------------------------------------------------------------------
FONTS = [
    '<font><sz val="11"/><name val="Calibri"/></font>',                                              # 0 default
    '<font><b/><sz val="11"/><name val="Calibri"/></font>',                                          # 1 bold
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',                   # 2 bold white
    '<font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>',                   # 3 section
    '<font><b/><sz val="20"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',                   # 4 title white
    '<font><b/><sz val="11"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>',                   # 5 bold blue
    '<font><b/><sz val="26"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',                   # 6 KPI number
    '<font><i/><sz val="10"/><color rgb="FF7F6000"/><name val="Calibri"/></font>',                   # 7 note
    '<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',                        # 8 white normal
]
FILLS = [
    '<fill><patternFill patternType="none"/></fill>',                                                # 0
    '<fill><patternFill patternType="gray125"/></fill>',                                             # 1
    '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>',         # 2 header blue
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>',         # 3 input green
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>',         # 4 output gray
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>',         # 5 title navy
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>',         # 6 note yellow
    '<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>',         # 7 kpi blue
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>',         # 8 section band
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>',         # 9 total band
    '<fill><patternFill patternType="solid"><fgColor rgb="FF548235"/></patternFill></fill>',         # 10 green kpi
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC55A11"/></patternFill></fill>',         # 11 orange kpi
]
BORDERS = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',                                    # 0 none
    ('<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>'
     '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'),  # 1 thin
    ('<border><left/><right/><top style="thin"><color rgb="FF808080"/></top>'
     '<bottom style="double"><color rgb="FF808080"/></bottom><diagonal/></border>'),                 # 2 total
]
NUMFMTS = {
    164: '&quot;Rp&quot;#,##0',
    165: '#,##0',
    166: '0.0&quot; kg&quot;',
    167: 'dd/mm/yyyy',
    168: '0.0%',
    169: '#,##0&quot; hari&quot;',
    170: '0.000',
    171: '&quot;Rp&quot;#,##0.00',
}

class Styles:
    def __init__(self):
        self.xfs = []
        self.cache = {}
        self.add(0, 0, 0, 0)  # index 0 default

    def add(self, numFmt=0, font=0, fill=0, border=0, h=None, v=None, wrap=0, locked=1):
        key = (numFmt, font, fill, border, h, v, wrap, locked)
        if key in self.cache:
            return self.cache[key]
        idx = len(self.xfs)
        self.xfs.append(key)
        self.cache[key] = idx
        return idx

    def xml(self):
        nf = "".join(f'<numFmt numFmtId="{k}" formatCode="{v}"/>' for k, v in sorted(NUMFMTS.items()))
        fonts = "".join(FONTS)
        fills = "".join(FILLS)
        borders = "".join(BORDERS)
        xfs = []
        for (numFmt, font, fill, border, h, v, wrap, locked) in self.xfs:
            al = ""
            if h or v or wrap:
                a = []
                if h: a.append(f'horizontal="{h}"')
                if v: a.append(f'vertical="{v}"')
                if wrap: a.append('wrapText="1"')
                al = f'<alignment {" ".join(a)}/>'
            prot = f'<protection locked="{locked}"/>'
            xfs.append(
                f'<xf numFmtId="{numFmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" '
                f'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">'
                f'{al}{prot}</xf>')
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<numFmts count="{len(NUMFMTS)}">{nf}</numFmts>'
            f'<fonts count="{len(FONTS)}">{fonts}</fonts>'
            f'<fills count="{len(FILLS)}">{fills}</fills>'
            f'<borders count="{len(BORDERS)}">{borders}</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>')

S = Styles()
# Convenience style indices ---------------------------------------------------
ST = {
    'header'    : S.add(font=2, fill=2, border=1, h='center', v='center', wrap=1, locked=1),
    'sect'      : S.add(font=3, fill=8, border=1, h='left', v='center', locked=1),
    'title'     : S.add(font=4, fill=5, border=0, h='left', v='center', locked=1),
    'note'      : S.add(font=7, fill=6, border=1, h='left', v='center', wrap=1, locked=1),
    'label'     : S.add(font=1, fill=0, border=0, h='left', locked=1),
    'labelband' : S.add(font=1, fill=8, border=1, h='left', locked=1),
    # input cells (unlocked, green)
    'i_text'    : S.add(font=0, fill=3, border=1, h='left',  locked=0),
    'i_int'     : S.add(165, fill=3, border=1, h='right', locked=0),
    'i_rp'      : S.add(164, fill=3, border=1, h='right', locked=0),
    'i_kg'      : S.add(166, fill=3, border=1, h='right', locked=0),
    'i_date'    : S.add(167, fill=3, border=1, h='center', locked=0),
    'i_rp2'     : S.add(171, fill=3, border=1, h='right', locked=0),
    # output/formula cells (locked, gray)
    'o_text'    : S.add(font=0, fill=4, border=1, h='left',  locked=1),
    'o_int'     : S.add(165, fill=4, border=1, h='right', locked=1),
    'o_rp'      : S.add(164, fill=4, border=1, h='right', locked=1),
    'o_kg'      : S.add(166, fill=4, border=1, h='right', locked=1),
    'o_date'    : S.add(167, fill=4, border=1, h='center', locked=1),
    'o_days'    : S.add(169, fill=4, border=1, h='right', locked=1),
    'o_adg'     : S.add(170, fill=4, border=1, h='right', locked=1),
    'o_pct'     : S.add(168, fill=4, border=1, h='right', locked=1),
    'o_rp2'     : S.add(171, fill=4, border=1, h='right', locked=1),
    # report value cells (white bg)
    'v_rp'      : S.add(164, font=1, fill=0, border=1, h='right', locked=1),
    'v_rp_tot'  : S.add(164, font=1, fill=9, border=2, h='right', locked=1),
    'v_num'     : S.add(165, font=0, fill=0, border=1, h='right', locked=1),
    'v_adg'     : S.add(170, font=0, fill=0, border=1, h='right', locked=1),
    'v_rp2'     : S.add(171, font=0, fill=0, border=1, h='right', locked=1),
    'v_pct'     : S.add(168, font=0, fill=0, border=1, h='right', locked=1),
    'v_text'    : S.add(font=0, fill=0, border=1, h='left', locked=1),
    # kpi
    'kpi_t'     : S.add(font=2, fill=7, border=0, h='center', v='center', wrap=1, locked=1),
    'kpi_n'     : S.add(165, font=6, fill=7, border=0, h='center', v='center', locked=1),
    'kpi_rp'    : S.add(164, font=6, fill=7, border=0, h='center', v='center', locked=1),
    'kpi_t_g'   : S.add(font=2, fill=10, border=0, h='center', v='center', wrap=1, locked=1),
    'kpi_n_g'   : S.add(164, font=6, fill=10, border=0, h='center', v='center', locked=1),
    'kpi_t_o'   : S.add(font=2, fill=11, border=0, h='center', v='center', wrap=1, locked=1),
    'kpi_n_o'   : S.add(170, font=6, fill=11, border=0, h='center', v='center', locked=1),
    'bar'       : S.add(font=5, fill=0, border=0, h='left', locked=1),
}

# ----------------------------------------------------------------------------
# Cell + sheet builders
# ----------------------------------------------------------------------------
def cell(ref, value=None, kind='blank', st=0):
    if kind == 'blank' or value is None or value == '':
        return f'<c r="{ref}" s="{st}"/>'
    if kind == 'formula':
        f = value[1:] if str(value).startswith('=') else value
        return f'<c r="{ref}" s="{st}"><f>{esc(f)}</f></c>'
    if kind == 'str':
        return f'<c r="{ref}" s="{st}" t="inlineStr"><is><t xml:space="preserve">{esc(value)}</t></is></c>'
    # number / date serial
    return f'<c r="{ref}" s="{st}"><v>{value}</v></c>'

class Sheet:
    def __init__(self, name, cols=None, freeze_row=None, protect=False, tabcolor=None):
        self.name = name
        self.cols = cols or []          # list of (min,max,width)
        self.rows = {}                  # rownum -> list of cell-xml
        self.freeze_row = freeze_row
        self.protect = protect
        self.merges = []
        self.validations = []           # (sqref, formula1)
        self.tabcolor = tabcolor
        self.maxcol = 1
        self.maxrow = 1

    def put(self, r, c, value=None, kind='blank', st=0):
        ref = f'{col_letter(c)}{r}'
        self.rows.setdefault(r, []).append((c, cell(ref, value, kind, st)))
        self.maxcol = max(self.maxcol, c)
        self.maxrow = max(self.maxrow, r)

    def merge(self, r1, c1, r2, c2):
        self.merges.append(f'{col_letter(c1)}{r1}:{col_letter(c2)}{r2}')

    def dv(self, sqref, formula1):
        self.validations.append((sqref, formula1))

    def xml(self):
        # cols
        cols_xml = ""
        if self.cols:
            inner = "".join(f'<col min="{a}" max="{b}" width="{w}" customWidth="1"/>' for a, b, w in self.cols)
            cols_xml = f'<cols>{inner}</cols>'
        # sheetdata
        body = []
        for r in sorted(self.rows):
            cells = "".join(x for _, x in sorted(self.rows[r], key=lambda t: t[0]))
            body.append(f'<row r="{r}">{cells}</row>')
        sheetdata = f'<sheetData>{"".join(body)}</sheetData>'
        # views / freeze
        if self.freeze_row:
            pane = (f'<pane ySplit="{self.freeze_row}" topLeftCell="A{self.freeze_row+1}" '
                    f'activePane="bottomLeft" state="frozen"/>'
                    f'<selection pane="bottomLeft" activeCell="A{self.freeze_row+1}" sqref="A{self.freeze_row+1}"/>')
            views = f'<sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews>'
        else:
            views = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        # protection
        prot = ""
        if self.protect:
            prot = ('<sheetProtection sheet="1" objects="1" scenarios="1" '
                    'formatCells="0" formatColumns="0" formatRows="0" insertRows="0" '
                    'sort="0" autoFilter="0" selectLockedCells="0" selectUnlockedCells="0"/>')
        # merges
        mg = ""
        if self.merges:
            mg = f'<mergeCells count="{len(self.merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in self.merges) + '</mergeCells>'
        # data validations
        dvx = ""
        if self.validations:
            items = ""
            for sqref, f1 in self.validations:
                items += (f'<dataValidation type="list" allowBlank="1" showInputMessage="1" '
                          f'showErrorMessage="1" sqref="{sqref}"><formula1>{esc(f1)}</formula1></dataValidation>')
            dvx = f'<dataValidations count="{len(self.validations)}">{items}</dataValidations>'
        dim = f'A1:{col_letter(self.maxcol)}{self.maxrow}'
        sheetpr = ""
        if self.tabcolor:
            sheetpr = f'<sheetPr><tabColor rgb="{self.tabcolor}"/></sheetPr>'
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'{sheetpr}<dimension ref="{dim}"/>{views}'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'{cols_xml}{sheetdata}{prot}{mg}{dvx}'
            '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
            '</worksheet>')

# Helper to build a standard data sheet -------------------------------------
def data_sheet(name, columns, nrows, samples, tabcolor):
    """columns: list of dict(header,width,kind,style,formula,dv,inp_style)
       samples: list of dict(col_index(1-based)->(value,kind))"""
    widths = [(i + 1, i + 1, c['width']) for i, c in enumerate(columns)]
    sh = Sheet(name, cols=widths, freeze_row=1, protect=True, tabcolor=tabcolor)
    # header
    for i, c in enumerate(columns):
        sh.put(1, i + 1, c['header'], 'str', ST['header'])
    # data rows
    for r in range(2, nrows + 2):
        for i, c in enumerate(columns):
            ci = i + 1
            if c['kind'] == 'formula':
                sh.put(r, ci, c['formula'].format(r=r), 'formula', ST[c['style']])
            else:
                # input cell - sample value or blank
                sval = None
                if (r - 2) < len(samples) and ci in samples[r - 2]:
                    v, k = samples[r - 2][ci]
                    sh.put(r, ci, v, k, ST[c['style']])
                else:
                    sh.put(r, ci, None, 'blank', ST[c['style']])
            if c.get('dv'):
                sh.dv(f'{col_letter(ci)}2:{col_letter(ci)}{nrows + 1}', c['dv'])
    return sh

sheets = []  # ordered list


# ============================================================================
# CAPACITIES (10-year)
# ============================================================================
SAPI_N, HPP_N        = 500, 500
PAKANM_N             = 3000
PAKANK_N             = 10000
TIMB_N               = 6000
OPER_N               = 3000
JUAL_N               = 1000
PKM_N                = 50      # master pakan
ASET_N               = 50
BATCH_N              = 200     # master batch (kapasitas 200 batch)

# --- M_Sapi: A ID,B TglMasuk,C Bangsa,D BobotAwal,E HargaBeli,F Batch,
#             G StatusInput,H TglMati,I Terjual,J Status,K TglKeluar,L Hari ---
R_SAPI       = f'$A$2:$A${SAPI_N+1}'
R_SAPI_HARGA = f'M_Sapi!$E$2:$E${SAPI_N+1}'
R_SAPI_BATCH = f'M_Sapi!$F$2:$F${SAPI_N+1}'
R_SAPI_STATUS= f'M_Sapi!$J$2:$J${SAPI_N+1}'
R_SAPI_HARI  = f'M_Sapi!$L$2:$L${SAPI_N+1}'
# --- Penjualan: A Tgl,B ID,C Bobot,D Harga,E Total,F HPP,G Margin,H Margin% ---
R_JUALB = f'Penjualan!$B$2:$B${JUAL_N+1}'
R_JUALA = f'Penjualan!$A$2:$A${JUAL_N+1}'
R_JUALE = f'Penjualan!$E$2:$E${JUAL_N+1}'
# --- Pakan_Masuk: A Tgl,B Kode,C Qty,D HargaTotal,E Harga/kg ---
R_PKM_B = f'Pakan_Masuk!$B$2:$B${PAKANM_N+1}'
R_PKM_C = f'Pakan_Masuk!$C$2:$C${PAKANM_N+1}'
R_PKM_D = f'Pakan_Masuk!$D$2:$D${PAKANM_N+1}'
# --- Pakan_Keluar: A Tgl,B Batch,C Kode,D Qty,E Harga/kg,F Nilai ---
R_PKK_BATCH = f'Pakan_Keluar!$B$2:$B${PAKANK_N+1}'
R_PKK_KODE  = f'Pakan_Keluar!$C$2:$C${PAKANK_N+1}'
R_PKK_QTY   = f'Pakan_Keluar!$D$2:$D${PAKANK_N+1}'
R_PKK_NILAI = f'Pakan_Keluar!$F$2:$F${PAKANK_N+1}'
# --- Operasional: A Tgl,B Batch,C Jenis,D Jumlah,E Keterangan ---
R_OPER_BATCH  = f'Operasional!$B$2:$B${OPER_N+1}'
R_OPER_JUMLAH = f'Operasional!$D$2:$D${OPER_N+1}'
# --- HPP_Sapi: A ID,B Batch,C Status,D Hari,E HargaBeli,F Pakan,G Oper,H Susut,I HPP,J Bobot,K HPP/kg ---
R_HPP_I = f'HPP_Sapi!$I$2:$I${HPP_N+1}'
R_HPP_C = f'HPP_Sapi!$C$2:$C${HPP_N+1}'
R_HPP_A = f'HPP_Sapi!$A$2:$A${HPP_N+1}'
# --- M_AsetTetap: A Nama,B Tgl,C Harga,D Umur,E Residu,F Peny/Bln,G UmurTpk,H Akm,I NilaiBuku ---
R_ASET_F = f'M_AsetTetap!$F$2:$F${ASET_N+1}'
R_ASET_C = f'M_AsetTetap!$C$2:$C${ASET_N+1}'
# --- M_Batch: A Kode,B Ket,C Kandang,D Mulai,E Selesai,F Status,G Hari,H Bulan,
#              I JmlSapi,J CattleDays,K Pakan,L Oper,M Susut,N Total ---
R_BATCH_A = f'M_Batch!$A$2:$A${BATCH_N+1}'
R_BATCH_J = f'M_Batch!$J$2:$J${BATCH_N+1}'
R_BATCH_K = f'M_Batch!$K$2:$K${BATCH_N+1}'
R_BATCH_L = f'M_Batch!$L$2:$L${BATCH_N+1}'
R_BATCH_M = f'M_Batch!$M$2:$M${BATCH_N+1}'

PKM_LIST  = 'M_Pakan!$A$2:$A$' + str(PKM_N + 1)
SAPI_LIST = f'M_Sapi!$A$2:$A${SAPI_N+1}'
BATCH_LIST= f'M_Batch!$A$2:$A${BATCH_N+1}'

# ============================================================================
# 1. M_Pakan
# ============================================================================
cols = [
    {'header':'Kode Pakan','width':14,'kind':'input','style':'i_text'},
    {'header':'Nama Pakan','width':28,'kind':'input','style':'i_text'},
    {'header':'Jenis','width':16,'kind':'input','style':'i_text',
     'dv':'"Rumput,Konsentrat,Mineral,Hijauan Lain,Lain-lain"'},
    {'header':'Satuan','width':10,'kind':'input','style':'i_text'},
]
samples=[
    {1:('PK01','str'),2:('Rumput Gajah','str'),3:('Rumput','str'),4:('kg','str')},
    {1:('PK02','str'),2:('Konsentrat Sapi','str'),3:('Konsentrat','str'),4:('kg','str')},
    {1:('PK03','str'),2:('Mineral / Premix','str'),3:('Mineral','str'),4:('kg','str')},
]
sheets.append(data_sheet('M_Pakan', cols, PKM_N, samples, 'FF4472C4'))

# ============================================================================
# 3. M_AsetTetap
# ============================================================================
cols = [
    {'header':'Nama Aset','width':24,'kind':'input','style':'i_text'},
    {'header':'Tgl Beli','width':13,'kind':'input','style':'i_date'},
    {'header':'Harga Perolehan','width':16,'kind':'input','style':'i_rp'},
    {'header':'Umur Ekonomis (bln)','width':16,'kind':'input','style':'i_int'},
    {'header':'Nilai Residu','width':14,'kind':'input','style':'i_rp'},
    {'header':'Penyusutan/Bln','width':15,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",IFERROR(($C{r}-$E{r})/$D{r},0))'},
    {'header':'Umur Terpakai (bln)','width':16,'kind':'formula','style':'o_int',
     'formula':'IF($A{r}="","",IFERROR(DATEDIF($B{r},TODAY(),"m"),0))'},
    {'header':'Akm. Penyusutan','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",MIN($C{r}-$E{r},$F{r}*$G{r}))'},
    {'header':'Nilai Buku','width':15,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",$C{r}-$H{r})'},
]
samples=[
    {1:('Kandang Utama','str'),2:(serial(2024,1,10),'n'),3:(150000000,'n'),4:(120,'n'),5:(15000000,'n')},
    {1:('Mesin Chopper','str'),2:(serial(2024,1,15),'n'),3:(18000000,'n'),4:(60,'n'),5:(1000000,'n')},
    {1:('Timbangan Sapi','str'),2:(serial(2024,2,1),'n'),3:(12000000,'n'),4:(60,'n'),5:(500000,'n')},
    {1:('Karpet Karet Kandang','str'),2:(serial(2024,2,1),'n'),3:(6000000,'n'),4:(48,'n'),5:(0,'n')},
]
sheets.append(data_sheet('M_AsetTetap', cols, ASET_N, samples, 'FF4472C4'))

# ============================================================================
# 3b. M_Batch  (modul batch penggemukan -- mengunci biaya per periode)
# ============================================================================
cols = [
    {'header':'Kode Batch','width':12,'kind':'input','style':'i_text'},
    {'header':'Keterangan','width':26,'kind':'input','style':'i_text'},
    {'header':'Kandang','width':16,'kind':'input','style':'i_text'},
    {'header':'Tgl Mulai','width':12,'kind':'input','style':'i_date'},
    {'header':'Tgl Selesai','width':12,'kind':'input','style':'i_date'},
    {'header':'Status','width':10,'kind':'formula','style':'o_text',
     'formula':'IF($A{r}="","",IF($E{r}="","Aktif","Tutup"))'},
    {'header':'Lama (hari)','width':11,'kind':'formula','style':'o_days',
     'formula':'IF($A{r}="","",IF($D{r}="","",IF($E{r}="",TODAY()-$D{r},$E{r}-$D{r})))'},
    {'header':'Lama (bln)','width':10,'kind':'formula','style':'o_adg',
     'formula':'IF($A{r}="","",IFERROR($G{r}/30,0))'},
    {'header':'Jumlah Sapi','width':11,'kind':'formula','style':'o_int',
     'formula':f'IF($A{{r}}="","",COUNTIF({R_SAPI_BATCH},$A{{r}}))'},
    {'header':'Total Cattle-days','width':14,'kind':'formula','style':'o_days',
     'formula':f'IF($A{{r}}="","",SUMIFS({R_SAPI_HARI},{R_SAPI_BATCH},$A{{r}}))'},
    {'header':'Biaya Pakan','width':15,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",SUMIFS({R_PKK_NILAI},{R_PKK_BATCH},$A{{r}}))'},
    {'header':'Biaya Operasional','width':16,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",SUMIFS({R_OPER_JUMLAH},{R_OPER_BATCH},$A{{r}}))'},
    {'header':'Penyusutan','width':15,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",SUM({R_ASET_F})*$H{{r}})'},
    {'header':'Total Biaya Batch','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",$K{r}+$L{r}+$M{r})'},
]
samples=[
    {1:('B001','str'),2:('Batch Penggemukan #1','str'),3:('K1','str'),4:(serial(2025,7,1),'n')},
]
sheets.append(data_sheet('M_Batch', cols, BATCH_N, samples, 'FF4472C4'))

# ============================================================================
# 4. M_Sapi  (master + pembelian sapi)  -- tanpa Kelompok, Batch di col F
# ============================================================================
cols = [
    {'header':'ID Sapi','width':10,'kind':'input','style':'i_text'},
    {'header':'Tgl Masuk','width':12,'kind':'input','style':'i_date'},
    {'header':'Bangsa','width':14,'kind':'input','style':'i_text',
     'dv':'"Limousin,Simental,PO,Brahman,Bali,Madura,Lain-lain"'},
    {'header':'Bobot Awal','width':11,'kind':'input','style':'i_kg'},
    {'header':'Harga Beli','width':15,'kind':'input','style':'i_rp'},
    {'header':'Batch','width':10,'kind':'input','style':'i_text','dv':BATCH_LIST},
    {'header':'Status Input','width':11,'kind':'input','style':'i_text','dv':'"Aktif,Mati"'},
    {'header':'Tgl Mati','width':12,'kind':'input','style':'i_date'},
    {'header':'Terjual?','width':9,'kind':'formula','style':'o_int',
     'formula':f'IF($A{{r}}="","",COUNTIF({R_JUALB},$A{{r}}))'},
    {'header':'Status','width':11,'kind':'formula','style':'o_text',
     'formula':'IF($A{r}="","",IF($I{r}>0,"Terjual",IF($G{r}="Mati","Mati","Aktif")))'},
    {'header':'Tgl Keluar','width':12,'kind':'formula','style':'o_date',
     'formula':f'IF($A{{r}}="","",IF($I{{r}}>0,INDEX({R_JUALA},MATCH($A{{r}},{R_JUALB},0)),IF($G{{r}}="Mati",$H{{r}},"")))'},
    {'header':'Hari Pelihara','width':13,'kind':'formula','style':'o_days',
     'formula':'IF($A{r}="","",IF($K{r}="",TODAY()-$B{r},$K{r}-$B{r}))'},
]
samples=[
    {1:('S001','str'),2:(serial(2025,9,1),'n'),3:('Limousin','str'),4:(280,'n'),5:(28000000,'n'),6:('B001','str'),7:('Aktif','str')},
    {1:('S002','str'),2:(serial(2025,9,1),'n'),3:('Simental','str'),4:(290,'n'),5:(29500000,'n'),6:('B001','str'),7:('Aktif','str')},
    {1:('S003','str'),2:(serial(2025,8,15),'n'),3:('PO','str'),4:(330,'n'),5:(31000000,'n'),6:('B001','str'),7:('Aktif','str')},
    {1:('S004','str'),2:(serial(2025,8,15),'n'),3:('Brahman','str'),4:(345,'n'),5:(33000000,'n'),6:('B001','str'),7:('Aktif','str')},
    {1:('S005','str'),2:(serial(2025,7,1),'n'),3:('Limousin','str'),4:(410,'n'),5:(41000000,'n'),6:('B001','str'),7:('Aktif','str')},
    {1:('S006','str'),2:(serial(2025,7,1),'n'),3:('Simental','str'),4:(420,'n'),5:(42500000,'n'),6:('B001','str'),7:('Mati','str'),8:(serial(2026,1,15),'n')},
]
sheets.append(data_sheet('M_Sapi', cols, SAPI_N, samples, 'FF548235'))

# ============================================================================
# 5. Pakan_Masuk
# ============================================================================
cols = [
    {'header':'Tgl','width':12,'kind':'input','style':'i_date'},
    {'header':'Kode Pakan','width':13,'kind':'input','style':'i_text','dv':PKM_LIST},
    {'header':'Qty (kg)','width':12,'kind':'input','style':'i_kg'},
    {'header':'Harga Total','width':15,'kind':'input','style':'i_rp'},
    {'header':'Harga/kg','width':13,'kind':'formula','style':'o_rp2',
     'formula':'IF($C{r}="","",IFERROR($D{r}/$C{r},0))'},
]
samples=[
    {1:(serial(2025,9,1),'n'),2:('PK01','str'),3:(3000,'n'),4:(1500000,'n')},
    {1:(serial(2025,9,1),'n'),2:('PK02','str'),3:(1000,'n'),4:(3500000,'n')},
    {1:(serial(2025,9,5),'n'),2:('PK03','str'),3:(50,'n'),4:(750000,'n')},
    {1:(serial(2025,9,15),'n'),2:('PK01','str'),3:(3000,'n'),4:(1560000,'n')},
    {1:(serial(2025,9,20),'n'),2:('PK02','str'),3:(1200,'n'),4:(4320000,'n')},
]
sheets.append(data_sheet('Pakan_Masuk', cols, PAKANM_N, samples, 'FF548235'))

# ============================================================================
# 6. Pakan_Keluar (konsumsi harian per batch)  -- valuasi rata-rata
# ============================================================================
avg_num = f'SUMIFS({R_PKM_D},{R_PKM_B},$C{{r}})'
avg_den = f'SUMIFS({R_PKM_C},{R_PKM_B},$C{{r}})'
cols = [
    {'header':'Tgl','width':12,'kind':'input','style':'i_date'},
    {'header':'Batch','width':10,'kind':'input','style':'i_text','dv':BATCH_LIST},
    {'header':'Kode Pakan','width':13,'kind':'input','style':'i_text','dv':PKM_LIST},
    {'header':'Qty (kg)','width':12,'kind':'input','style':'i_kg'},
    {'header':'Harga Rata2/kg','width':14,'kind':'formula','style':'o_rp2',
     'formula':f'IF($C{{r}}="","",IFERROR({avg_num}/{avg_den},0))'},
    {'header':'Nilai Konsumsi','width':15,'kind':'formula','style':'o_rp',
     'formula':'IF($C{r}="","",$D{r}*$E{r})'},
]
samples=[
    {1:(serial(2025,9,2),'n'),2:('B001','str'),3:('PK01','str'),4:(40,'n')},
    {1:(serial(2025,9,2),'n'),2:('B001','str'),3:('PK02','str'),4:(12,'n')},
    {1:(serial(2025,9,2),'n'),2:('B001','str'),3:('PK01','str'),4:(45,'n')},
    {1:(serial(2025,9,2),'n'),2:('B001','str'),3:('PK02','str'),4:(14,'n')},
    {1:(serial(2025,9,2),'n'),2:('B001','str'),3:('PK01','str'),4:(55,'n')},
    {1:(serial(2025,9,2),'n'),2:('B001','str'),3:('PK02','str'),4:(18,'n')},
]
sheets.append(data_sheet('Pakan_Keluar', cols, PAKANK_N, samples, 'FF548235'))

# ============================================================================
# 7. Timbang
# ============================================================================
cols = [
    {'header':'ID Sapi','width':10,'kind':'input','style':'i_text','dv':SAPI_LIST},
    {'header':'Tgl Timbang','width':13,'kind':'input','style':'i_date'},
    {'header':'Bobot (kg)','width':12,'kind':'input','style':'i_kg'},
    {'header':'Bobot Awal','width':12,'kind':'formula','style':'o_kg',
     'formula':f'IF($A{{r}}="","",IFERROR(VLOOKUP($A{{r}},M_Sapi!$A$2:$D${SAPI_N+1},4,0),""))'},
    {'header':'Tgl Masuk','width':12,'kind':'formula','style':'o_date',
     'formula':f'IF($A{{r}}="","",IFERROR(VLOOKUP($A{{r}},M_Sapi!$A$2:$B${SAPI_N+1},2,0),""))'},
    {'header':'Pertambahan (kg)','width':15,'kind':'formula','style':'o_kg',
     'formula':'IF(OR($A{r}="",$C{r}=""),"",$C{r}-$D{r})'},
    {'header':'Hari','width':9,'kind':'formula','style':'o_days',
     'formula':'IF(OR($A{r}="",$B{r}=""),"",$B{r}-$E{r})'},
    {'header':'ADG Kumulatif','width':13,'kind':'formula','style':'o_adg',
     'formula':'IF(OR($G{r}="",$G{r}=0),"",$F{r}/$G{r})'},
]
samples=[
    {1:('S001','str'),2:(serial(2025,10,1),'n'),3:(312,'n')},
    {1:('S002','str'),2:(serial(2025,10,1),'n'),3:(325,'n')},
    {1:('S003','str'),2:(serial(2025,10,1),'n'),3:(372,'n')},
    {1:('S004','str'),2:(serial(2025,10,1),'n'),3:(388,'n')},
    {1:('S005','str'),2:(serial(2025,10,1),'n'),3:(458,'n')},
    {1:('S006','str'),2:(serial(2025,10,1),'n'),3:(470,'n')},
]
sheets.append(data_sheet('Timbang', cols, TIMB_N, samples, 'FF548235'))

# ============================================================================
# 8. Operasional  -- ditandai Batch agar biaya terkunci per periode
# ============================================================================
cols = [
    {'header':'Tgl','width':12,'kind':'input','style':'i_date'},
    {'header':'Batch','width':10,'kind':'input','style':'i_text','dv':BATCH_LIST},
    {'header':'Jenis Biaya','width':18,'kind':'input','style':'i_text',
     'dv':'"Gaji Karyawan,Listrik,Air,BBM,Obat & Vitamin,Perbaikan Kandang,Lain-lain"'},
    {'header':'Jumlah','width':15,'kind':'input','style':'i_rp'},
    {'header':'Keterangan','width':30,'kind':'input','style':'i_text'},
]
samples=[
    {1:(serial(2025,9,30),'n'),2:('B001','str'),3:('Gaji Karyawan','str'),4:(4500000,'n'),5:('Gaji 2 karyawan','str')},
    {1:(serial(2025,9,30),'n'),2:('B001','str'),3:('Listrik','str'),4:(650000,'n'),5:('Tagihan bulanan','str')},
    {1:(serial(2025,9,30),'n'),2:('B001','str'),3:('BBM','str'),4:(500000,'n'),5:('Solar chopper','str')},
    {1:(serial(2025,9,10),'n'),2:('B001','str'),3:('Obat & Vitamin','str'),4:(800000,'n'),5:('Vitamin B-complex','str')},
]
sheets.append(data_sheet('Operasional', cols, OPER_N, samples, 'FF548235'))

# ============================================================================
# 9. Penjualan
# ============================================================================
cols = [
    {'header':'Tgl Jual','width':12,'kind':'input','style':'i_date'},
    {'header':'ID Sapi','width':10,'kind':'input','style':'i_text','dv':SAPI_LIST},
    {'header':'Bobot Jual (kg)','width':14,'kind':'input','style':'i_kg'},
    {'header':'Harga/kg','width':13,'kind':'input','style':'i_rp'},
    {'header':'Total Jual','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($B{r}="","",$C{r}*$D{r})'},
    {'header':'HPP/ekor','width':16,'kind':'formula','style':'o_rp',
     'formula':f'IF($B{{r}}="","",IFERROR(INDEX({R_HPP_I},MATCH($B{{r}},{R_HPP_A},0)),0))'},
    {'header':'Margin','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($B{r}="","",$E{r}-$F{r})'},
    {'header':'Margin %','width':10,'kind':'formula','style':'o_pct',
     'formula':'IF(OR($B{r}="",$E{r}=0),"",$G{r}/$E{r})'},
]
samples=[]   # belum ada penjualan di contoh (semua sapi masih digemukkan)
sheets.append(data_sheet('Penjualan', cols, JUAL_N, samples, 'FF548235'))

print("Bagian transaksi selesai didefinisikan.")


# Extra ranges for output sheets
R_JUALC  = f'Penjualan!$C$2:$C${JUAL_N+1}'
R_TIMB_A = f'Timbang!$A$2:$A${TIMB_N+1}'
R_TIMB_C = f'Timbang!$C$2:$C${TIMB_N+1}'

# ============================================================================
# 10. Stok_Pakan  (kartu stok, valuasi rata-rata) -- mirror M_Pakan
# ============================================================================
cols = [
    {'header':'Kode','width':12,'kind':'formula','style':'o_text',
     'formula':'IF(M_Pakan!$A{r}="","",M_Pakan!$A{r})'},
    {'header':'Nama Pakan','width':24,'kind':'formula','style':'o_text',
     'formula':'IF($A{r}="","",M_Pakan!$B{r})'},
    {'header':'Total Masuk (kg)','width':15,'kind':'formula','style':'o_kg',
     'formula':f'IF($A{{r}}="","",SUMIFS({R_PKM_C},{R_PKM_B},$A{{r}}))'},
    {'header':'Nilai Masuk','width':16,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",SUMIFS({R_PKM_D},{R_PKM_B},$A{{r}}))'},
    {'header':'Harga Rata2/kg','width':14,'kind':'formula','style':'o_rp2',
     'formula':'IF($A{r}="","",IFERROR($D{r}/$C{r},0))'},
    {'header':'Total Keluar (kg)','width':15,'kind':'formula','style':'o_kg',
     'formula':f'IF($A{{r}}="","",SUMIFS({R_PKK_QTY},{R_PKK_KODE},$A{{r}}))'},
    {'header':'Stok Akhir (kg)','width':14,'kind':'formula','style':'o_kg',
     'formula':'IF($A{r}="","",$C{r}-$F{r})'},
    {'header':'Nilai Stok Akhir','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",$G{r}*$E{r})'},
    {'header':'Nilai Konsumsi','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",$F{r}*$E{r})'},
]
sheets.append(data_sheet('Stok_Pakan', cols, PKM_N, [], 'FFBF8F00'))

# ============================================================================
# 11. HPP_Sapi  (mesin HPP per ekor -- BERBASIS BATCH) -- mirror M_Sapi
# ============================================================================
bobot_terkini = (
    f'IF($A{{r}}="","",IF(COUNTIF({R_JUALB},$A{{r}})>0,'
    f'IFERROR(INDEX({R_JUALC},MATCH($A{{r}},{R_JUALB},0)),0),'
    f'IFERROR(LOOKUP(2,1/({R_TIMB_A}=$A{{r}}),{R_TIMB_C}),M_Sapi!$D{{r}})))'
)
# total & cattle-days batch diambil dari M_Batch (terkunci setelah batch ditutup)
b_pakan = f'INDEX({R_BATCH_K},MATCH($B{{r}},{R_BATCH_A},0))'
b_oper  = f'INDEX({R_BATCH_L},MATCH($B{{r}},{R_BATCH_A},0))'
b_susut = f'INDEX({R_BATCH_M},MATCH($B{{r}},{R_BATCH_A},0))'
b_days  = f'INDEX({R_BATCH_J},MATCH($B{{r}},{R_BATCH_A},0))'
cols = [
    {'header':'ID Sapi','width':10,'kind':'formula','style':'o_text',
     'formula':'IF(M_Sapi!$A{r}="","",M_Sapi!$A{r})'},
    {'header':'Batch','width':10,'kind':'formula','style':'o_text',
     'formula':'IF($A{r}="","",M_Sapi!$F{r})'},
    {'header':'Status','width':10,'kind':'formula','style':'o_text',
     'formula':'IF($A{r}="","",M_Sapi!$J{r})'},
    {'header':'Hari Pelihara','width':12,'kind':'formula','style':'o_days',
     'formula':'IF($A{r}="","",M_Sapi!$L{r})'},
    {'header':'Harga Beli','width':15,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",M_Sapi!$E{r})'},
    {'header':'Biaya Pakan (alokasi batch)','width':18,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",IFERROR({b_pakan}*$D{{r}}/{b_days},0))'},
    {'header':'Biaya Operasional (alokasi batch)','width':20,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",IFERROR({b_oper}*$D{{r}}/{b_days},0))'},
    {'header':'Penyusutan (alokasi batch)','width':18,'kind':'formula','style':'o_rp',
     'formula':f'IF($A{{r}}="","",IFERROR({b_susut}*$D{{r}}/{b_days},0))'},
    {'header':'HPP Total','width':16,'kind':'formula','style':'o_rp',
     'formula':'IF($A{r}="","",$E{r}+$F{r}+$G{r}+$H{r})'},
    {'header':'Bobot Terkini (kg)','width':15,'kind':'formula','style':'o_kg',
     'formula':bobot_terkini},
    {'header':'HPP per kg','width':13,'kind':'formula','style':'o_rp2',
     'formula':'IF($A{r}="","",IFERROR($I{r}/$J{r},0))'},
]
sheets.append(data_sheet('HPP_Sapi', cols, HPP_N, [], 'FFBF8F00'))

print("Output sheets (Stok_Pakan, HPP_Sapi) selesai.")


# ============================================================================
# 12. Lap_Keuangan  (Laba Rugi + Arus Kas + Neraca)
# ============================================================================
lk = Sheet('Lap_Keuangan', cols=[(1,1,3),(2,2,44),(3,3,3),(4,4,20)],
           protect=True, tabcolor='FFC00000')
lk.put(1,2,'LAPORAN KEUANGAN — Usaha Penggemukan Sapi','str',ST['title']); lk.merge(1,2,1,4)
def lk_row(r, label, value=None, lab_st='label', val_st='v_rp', sect=False):
    if sect:
        lk.put(r,2,label,'str',ST['sect']); lk.merge(r,2,r,4)
    else:
        lk.put(r,2,label,'str',ST[lab_st])
        if value is not None:
            lk.put(r,4,value,'formula',ST[val_st])
lk_row(3,'A. LABA RUGI',sect=True)
lk_row(4,'Pendapatan Penjualan Sapi', f'SUM({R_JUALE})')
lk_row(5,'(-) HPP Sapi Terjual', f'SUMIFS({R_HPP_I},{R_HPP_C},"Terjual")')
lk_row(6,'Laba Kotor','D4-D5',lab_st='label',val_st='v_rp')
lk_row(7,'(-) Kerugian Sapi Mati/Afkir', f'SUMIFS({R_HPP_I},{R_HPP_C},"Mati")')
lk_row(8,'LABA BERSIH','D6-D7',lab_st='label',val_st='v_rp_tot')
lk_row(10,'B. ARUS KAS (Metode Langsung)',sect=True)
lk_row(11,'Kas Masuk — Modal Disetor','MENU!$B$5')
lk_row(12,'Kas Masuk — Penjualan Sapi','D4')
lk_row(13,'Total Kas Masuk','D11+D12',val_st='v_rp')
lk_row(14,'Kas Keluar — Beli Sapi Bakalan', f'SUM(M_Sapi!$E$2:$E${SAPI_N+1})')
lk_row(15,'Kas Keluar — Beli Pakan', f'SUM({R_PKM_D})')
lk_row(16,'Kas Keluar — Biaya Operasional', f'SUM({R_OPER_JUMLAH})')
lk_row(17,'Kas Keluar — Beli Aset Tetap', f'SUM(M_AsetTetap!$C$2:$C${ASET_N+1})')
lk_row(18,'Total Kas Keluar','D14+D15+D16+D17',val_st='v_rp')
lk_row(19,'SALDO KAS AKHIR','D13-D18',val_st='v_rp_tot')
lk_row(21,'C. NERACA',sect=True)
lk_row(22,'ASET',lab_st='labelband'); lk.merge(22,2,22,4)
lk_row(23,'Kas','D19')
lk_row(24,'Persediaan Sapi (Aktif)', f'SUMIFS({R_HPP_I},{R_HPP_C},"Aktif")')
lk_row(25,'Persediaan Pakan', f'SUM(Stok_Pakan!$H$2:$H${PKM_N+1})')
lk_row(26,'Aset Tetap (Nilai Buku)', f'SUM({R_ASET_C})-SUM({R_BATCH_M})')
lk_row(27,'TOTAL ASET','D23+D24+D25+D26',val_st='v_rp_tot')
lk_row(29,'PASIVA & MODAL',lab_st='labelband'); lk.merge(29,2,29,4)
lk_row(30,'Modal Disetor','MENU!$B$5')
lk_row(31,'Laba Ditahan (akumulasi laba bersih)','D8')
lk_row(32,'TOTAL PASIVA & MODAL','D30+D31',val_st='v_rp_tot')
lk_row(34,'Cek Keseimbangan (harus = 0)','D27-D32',val_st='v_rp')
lk.put(36,2,'Catatan: Biaya pakan, operasional & penyusutan dikelompokkan per BATCH lalu menempel ke sapi '
             '(metode HPP akurat). Setelah batch ditutup (Tgl Selesai diisi), biaya & HPP batch tsb TERKUNCI '
             'sehingga margin penjualan historis tidak berubah walau ada batch/biaya periode baru. '
             'Penyusutan diakui sepanjang batch berjalan = (penyusutan/bln semua aset) x lama batch.',
       'str',ST['note']); lk.merge(36,2,38,4)
sheets.append(lk)

# ============================================================================
# 13. Analisis (per ekor + per BATCH FCG)
# ============================================================================
an = Sheet('Analisis', cols=[(1,1,9),(2,2,10),(3,3,10),(4,4,11),(5,5,11),
                              (6,6,11),(7,7,8),(8,8,11),(9,9,15),(10,10,15),(11,11,15),(12,12,10),
                              (13,13,3),(14,14,11),(15,15,10),(16,16,11),(17,17,14),(18,18,16),(19,19,15),(20,20,12)],
           freeze_row=1, protect=True, tabcolor='FF7030A0')
ehead = ['ID Sapi','Batch','Status','Bobot Awal','Bobot Akhir','Pertambahan',
         'Hari','ADG (kg/hr)','HPP Total','Harga Jual','Margin','Margin %']
for i,h in enumerate(ehead): an.put(1,i+1,h,'str',ST['header'])
bhead = ['Batch','Status','Jumlah Sapi','Total Gain (kg)','Total Biaya Pakan','FCG (Rp/kg gain)','ADG Rata2']
for i,h in enumerate(bhead): an.put(1,14+i,h,'str',ST['header'])
for r in range(2, SAPI_N+2):
    an.put(r,1, f'IF(M_Sapi!$A{r}="","",M_Sapi!$A{r})','formula',ST['o_text'])
    an.put(r,2, f'IF($A{r}="","",M_Sapi!$F{r})','formula',ST['o_text'])
    an.put(r,3, f'IF($A{r}="","",M_Sapi!$J{r})','formula',ST['o_text'])
    an.put(r,4, f'IF($A{r}="","",M_Sapi!$D{r})','formula',ST['o_kg'])
    an.put(r,5, f'IF($A{r}="","",HPP_Sapi!$J{r})','formula',ST['o_kg'])
    an.put(r,6, f'IF($A{r}="","",$E{r}-$D{r})','formula',ST['o_kg'])
    an.put(r,7, f'IF($A{r}="","",M_Sapi!$L{r})','formula',ST['o_days'])
    an.put(r,8, f'IF(OR($A{r}="",$G{r}=0),"",$F{r}/$G{r})','formula',ST['o_adg'])
    an.put(r,9, f'IF($A{r}="","",HPP_Sapi!$I{r})','formula',ST['o_rp'])
    an.put(r,10,f'IF($A{r}="","",IFERROR(INDEX({R_JUALE},MATCH($A{r},{R_JUALB},0)),0))','formula',ST['o_rp'])
    an.put(r,11,f'IF($A{r}="","",IF($J{r}=0,0,$J{r}-$I{r}))','formula',ST['o_rp'])
    an.put(r,12,f'IF(OR($A{r}="",$J{r}=0),"",$K{r}/$J{r})','formula',ST['o_pct'])
for r in range(2, BATCH_N+2):
    an.put(r,14,f'IF(M_Batch!$A{r}="","",M_Batch!$A{r})','formula',ST['o_text'])
    an.put(r,15,f'IF($N{r}="","",M_Batch!$F{r})','formula',ST['o_text'])
    an.put(r,16,f'IF($N{r}="","",M_Batch!$I{r})','formula',ST['o_int'])
    an.put(r,17,f'IF($N{r}="","",SUMIFS($F$2:$F${SAPI_N+1},$B$2:$B${SAPI_N+1},$N{r}))','formula',ST['o_kg'])
    an.put(r,18,f'IF($N{r}="","",M_Batch!$K{r})','formula',ST['o_rp'])
    an.put(r,19,f'IF($N{r}="","",IFERROR($R{r}/$Q{r},0))','formula',ST['o_rp2'])
    an.put(r,20,f'IF($N{r}="","",IFERROR(AVERAGEIFS($H$2:$H${SAPI_N+1},$B$2:$B${SAPI_N+1},$N{r}),0))','formula',ST['o_adg'])
sheets.append(an)

# ============================================================================
# 14. DASHBOARD
# ============================================================================
db = Sheet('DASHBOARD', cols=[(1,1,2),(2,2,16),(3,3,16),(4,4,16),(5,5,18),(6,6,30),(7,7,4)],
           protect=True, tabcolor='FF1F4E79')
db.put(1,2,'DASHBOARD USAHA PENGGEMUKAN SAPI','str',ST['title']); db.merge(1,2,1,6)
def card(r, c, title, formula, t_st, n_st):
    db.put(r,c,title,'str',ST[t_st]); db.merge(r,c,r,c+1)
    db.put(r+1,c,formula,'formula',ST[n_st]); db.merge(r+1,c,r+1,c+1)
# row 3-4 cards
card(3,2,'Sapi Aktif', f'COUNTIF(M_Sapi!$J$2:$J${SAPI_N+1},"Aktif")','kpi_t','kpi_n')
card(3,4,'Sapi Terjual', f'COUNTIF(M_Sapi!$J$2:$J${SAPI_N+1},"Terjual")','kpi_t','kpi_n')
card(3,6,'Saldo Kas','Lap_Keuangan!$D$19','kpi_t_g','kpi_n_g')
# row 6-7 cards
card(6,2,'Total Pendapatan','Lap_Keuangan!$D$4','kpi_t','kpi_rp')
card(6,4,'Laba Bersih','Lap_Keuangan!$D$8','kpi_t_g','kpi_n_g')
card(6,6,'Nilai Persediaan Sapi','Lap_Keuangan!$D$24','kpi_t','kpi_rp')
# row 9-10 cards
card(9,2,'Persediaan Pakan','Lap_Keuangan!$D$25','kpi_t','kpi_rp')
card(9,4,'ADG Rata-rata (kg/hr)', f'IFERROR(AVERAGE(Analisis!$H$2:$H${SAPI_N+1}),0)','kpi_t_o','kpi_n_o')
card(9,6,'Margin Rata-rata', f'IFERROR(AVERAGE(Analisis!$L$2:$L${SAPI_N+1}),0)','kpi_t','kpi_n')
# row 11-12 cards (sapi mati/afkir)
card(11,2,'Sapi Mati/Afkir', f'COUNTIF(M_Sapi!$J$2:$J${SAPI_N+1},"Mati")','kpi_t','kpi_n')
card(11,4,'Kerugian Sapi Mati','Lap_Keuangan!$D$7','kpi_t','kpi_rp')
# FCG table with bar (per BATCH)
db.put(13,2,'FEED COST PER GAIN (FCG) & PERFORMA PER BATCH','str',ST['sect']); db.merge(13,2,13,6)
for i,h in enumerate(['Batch','Jml Sapi','ADG','FCG (Rp/kg)','Visual FCG']):
    db.put(14,2+i,h,'str',ST['header'])
for dr in range(15, 27):       # tampilkan hingga 12 batch
    ar = dr-13                 # baris tabel batch di Analisis (mulai 2)
    db.put(dr,2,f'IF(Analisis!$N{ar}="","",Analisis!$N{ar})','formula',ST['o_text'])
    db.put(dr,3,f'IF($B{dr}="","",Analisis!$P{ar})','formula',ST['o_int'])
    db.put(dr,4,f'IF($B{dr}="","",Analisis!$T{ar})','formula',ST['o_adg'])
    db.put(dr,5,f'IF($B{dr}="","",Analisis!$S{ar})','formula',ST['o_rp2'])
    db.put(dr,6,f'IF($B{dr}="","",REPT("|",ROUND(IFERROR($E{dr}/MAX(Analisis!$S$2:$S${BATCH_N+1}),0)*25,0)))','formula',ST['bar'])
db.put(28,2,'Catatan: Dashboard otomatis. FCG = total biaya pakan batch / total pertambahan bobot batch. '
             'ADG = pertambahan bobot harian rata-rata. Margin = harga jual - HPP per ekor. '
             'Biaya & HPP batch yang sudah ditutup bersifat TERKUNCI.',
       'str',ST['note']); db.merge(28,2,29,6)
sheets.append(db)

print("Lap_Keuangan, Analisis, DASHBOARD selesai.")


# ============================================================================
# 0. MENU (navigasi + pengaturan)  -- diletakkan paling depan
# ============================================================================
mn = Sheet('MENU', cols=[(1,1,3),(2,2,26),(3,3,58),(4,4,18)], protect=True, tabcolor='FF1F4E79')
mn.put(1,2,'SISTEM PENGGEMUKAN SAPI POTONG — TERINTEGRASI','str',ST['title']); mn.merge(1,2,1,4)
mn.put(3,2,'PENGATURAN USAHA','str',ST['sect']); mn.merge(3,2,3,4)
mn.put(4,2,'PT/UD Contoh Ternak','str',ST['i_text'])
mn.put(4,3,'<-- Nama Usaha','str',ST['label'])
mn.put(5,2,500000000,'n',ST['i_rp'])
mn.put(5,3,'<-- Modal Disetor (Rp)  | dipakai di Neraca & Arus Kas','str',ST['label'])
mn.put(6,2,'Tahun 2025/2026','str',ST['i_text'])
mn.put(6,3,'<-- Periode Laporan','str',ST['label'])

mn.put(8,2,'DAFTAR SHEET & FUNGSINYA','str',ST['sect']); mn.merge(8,2,8,4)
mn.put(9,2,'Nama Sheet','str',ST['header'])
mn.put(9,3,'Fungsi','str',ST['header'])
mn.put(9,4,'Tipe','str',ST['header'])
directory = [
    ('M_Pakan','Master jenis pakan & satuan','INPUT'),
    ('M_AsetTetap','Master aset tetap + penyusutan otomatis','INPUT'),
    ('M_Batch','Master batch/periode penggemukan (pengunci biaya + kandang)','INPUT'),
    ('M_Sapi','Master & pembelian sapi (isi Batch)','INPUT'),
    ('Pakan_Masuk','Pembelian/stok pakan masuk','INPUT'),
    ('Pakan_Keluar','Konsumsi pakan harian (per batch)','INPUT'),
    ('Timbang','Penimbangan bobot berkala','INPUT'),
    ('Operasional','Biaya operasional (tandai Batch)','INPUT'),
    ('Penjualan','Penjualan sapi','INPUT'),
    ('Stok_Pakan','Kartu stok pakan (valuasi rata-rata)','OUTPUT'),
    ('HPP_Sapi','HPP per ekor berbasis batch (mesin hitung)','OUTPUT'),
    ('Lap_Keuangan','Laba Rugi, Arus Kas, Neraca','OUTPUT'),
    ('Analisis','ADG, FCG per batch, margin/ekor','OUTPUT'),
    ('DASHBOARD','Ringkasan KPI & visual usaha','OUTPUT'),
]
row = 10
for nm, fn, tp in directory:
    st_nm = ST['i_text'] if tp == 'INPUT' else ST['o_text']
    mn.put(row,2,nm,'str',st_nm)
    mn.put(row,3,fn,'str',ST['v_text'])
    mn.put(row,4,tp,'str',ST['o_text'])
    row += 1
row += 1
mn.put(row,2,'PETUNJUK & KETERANGAN','str',ST['sect']); mn.merge(row,2,row,4); row+=1
notes = [
    'Isi HANYA sel berwarna HIJAU (INPUT). Sel ABU-ABU dihitung otomatis (jangan diubah).',
    'Sheet OUTPUT terkunci (proteksi). Untuk membuka: menu Review > Unprotect Sheet (tanpa password).',
    'ALUR: buat Batch di M_Batch dulu -> tandai setiap sapi, pakan keluar, & operasional dengan Batch tsb.',
    'Metode: HPP akurat berbasis BATCH. Biaya pakan/operasional/penyusutan dikelompokkan per batch.',
    'Saat batch SELESAI: isi Tgl Selesai di M_Batch -> biaya & HPP batch itu TERKUNCI (margin historis tidak berubah).',
    'Alokasi operasional & penyusutan dalam batch: berdasarkan HARI pemeliharaan. Valuasi pakan: rata-rata tertimbang.',
    'Penyusutan batch = (penyusutan/bln semua aset) x lama batch (bulan). Kandang dicatat di M_Batch.',
    'Gunakan dropdown pada kolom Batch, Kode Pakan, ID Sapi, Jenis Biaya untuk hindari salah ketik.',
    'Kapasitas: 50 batch (3-4x/tahun x 10+ tahun). Cukup ketik di baris kosong berikutnya, formula sudah tersedia.',
]
for n in notes:
    mn.put(row,2,n,'str',ST['note']); mn.merge(row,2,row,4); row+=1

all_sheets = [mn] + sheets

# ============================================================================
# PACKAGING (OOXML zip)
# ============================================================================
def build():
    # workbook.xml
    sxml = "".join(
        f'<sheet name="{esc(s.name)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, s in enumerate(all_sheets))
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{sxml}</sheets>'
        '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    # workbook rels
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i, s in enumerate(all_sheets):
        rels.append(f'<Relationship Id="rId{i+1}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                    f'Target="worksheets/sheet{i+1}.xml"/>')
    sid = len(all_sheets) + 1
    rels.append(f'<Relationship Id="rId{sid}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                'Target="styles.xml"/>')
    rels.append('</Relationships>')
    workbook_rels = "".join(rels)
    # content types
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(len(all_sheets)))
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f'{overrides}'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>')
    now = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>Kiro</dc:creator><cp:lastModifiedBy>Kiro</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>Kiro Excel Generator</Application></Properties>')

    parts = {
        '[Content_Types].xml': content_types,
        '_rels/.rels': root_rels,
        'xl/workbook.xml': workbook,
        'xl/_rels/workbook.xml.rels': workbook_rels,
        'xl/styles.xml': S.xml(),
        'docProps/core.xml': core,
        'docProps/app.xml': app,
    }
    for i, s in enumerate(all_sheets):
        parts[f'xl/worksheets/sheet{i+1}.xml'] = s.xml()

    # validate well-formedness
    for name, xml in parts.items():
        if name.endswith('.xml') or name.endswith('.rels'):
            try:
                minidom.parseString(xml.encode('utf-8'))
            except Exception as e:
                raise SystemExit(f"XML tidak valid pada {name}: {e}")

    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, xml)
    return parts

parts = build()
import os
print(f"OK -> {OUT}")
print(f"Jumlah sheet : {len(all_sheets)}")
print(f"Ukuran file  : {os.path.getsize(OUT)/1024:.1f} KB")
print("Semua bagian XML well-formed.")
