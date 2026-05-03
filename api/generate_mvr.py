"""
generate_mvr.py  —  reproduces the Shaigan Method Validation Report
Usage:  python generate_mvr.py  <excel.xlsx>  [output.docx]
Requires shaigan_template.docx in the same folder.
"""
import sys, shutil, os
import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH as WDA
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ── Fixed constants ────────────────────────────────────────────────────────
SOP        = "3/QC/GEN/012"
DEPARTMENT = "Research & Development"
REVISION   = "00"

# ── Exact values from reference XML ───────────────────────────────────────
SZ      = 22          # 11 pt in half-points
SZ_SM   = 20          # 10 pt
FONT    = "Times New Roman"

F_GREY  = "BFBFBF"   # section / col headers
F_DARK  = "A5A5A5"   # IP / LOD darker headers
F_LABEL = "F2F2F2"   # cover label cells
F_WHITE = "FFFFFF"   # explicit white data cells
F_DATA  = "D9D9D9"   # repeatability mean/sd/rsd

CONC = {80:0.122, 90:0.135, 100:0.150, 110:0.165, 120:0.182}


# ═══════════════════════════════════════════════════
# LOW-LEVEL XML HELPERS
# ═══════════════════════════════════════════════════
def esc(s):
    s = str(s) if s is not None else ""
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def borders_xml(t="single", sz="4", c="auto"):
    a = f'w:val="{t}" w:sz="{sz}" w:space="0" w:color="{c}"'
    ns = nsdecls("w")
    return (f'<w:tblBorders {ns}>'
            f'<w:top {a}/><w:left {a}/><w:bottom {a}/>'
            f'<w:right {a}/><w:insideH {a}/><w:insideV {a}/>'
            f'</w:tblBorders>')

def set_borders(tbl, t="single", sz="4", c="auto"):
    pr = tbl._tbl.tblPr
    for old in pr.findall(qn("w:tblBorders")): pr.remove(old)
    pr.append(parse_xml(borders_xml(t, sz, c)))

def set_tblW(tbl, w, tp="dxa"):
    pr = tbl._tbl.tblPr
    for old in pr.findall(qn("w:tblW")): pr.remove(old)
    pr.insert(0, parse_xml(f'<w:tblW {nsdecls("w")} w:w="{w}" w:type="{tp}"/>'))

def set_grid(tbl, widths):
    old = tbl._tbl.find(qn("w:tblGrid"))
    if old is not None: tbl._tbl.remove(old)
    xml = (f'<w:tblGrid {nsdecls("w")}>'
           + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
           + "</w:tblGrid>")
    tr = tbl._tbl.find(qn("w:tr"))
    tbl._tbl.insert(list(tbl._tbl).index(tr), parse_xml(xml))

def set_cw(cell, w):
    pr = cell._tc.get_or_add_tcPr()
    for old in pr.findall(qn("w:tcW")): pr.remove(old)
    pr.insert(0, parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w}" w:type="dxa"/>'))

def set_shd(cell, fill):
    pr = cell._tc.get_or_add_tcPr()
    for old in pr.findall(qn("w:shd")): pr.remove(old)
    pr.append(parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))

def set_valign(cell, v="center"):
    pr = cell._tc.get_or_add_tcPr()
    for old in pr.findall(qn("w:vAlign")): pr.remove(old)
    pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{v}"/>'))

def set_rowh(row, val, rule="atLeast"):
    pr = row._tr.get_or_add_trPr()
    for old in pr.findall(qn("w:trHeight")): pr.remove(old)
    pr.append(parse_xml(
        f'<w:trHeight {nsdecls("w")} w:val="{val}" w:hRule="{rule}"/>'))

def set_gridspan(cell, n):
    pr = cell._tc.get_or_add_tcPr()
    for old in pr.findall(qn("w:gridSpan")): pr.remove(old)
    pr.append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="{n}"/>'))

def write_cell(cell, text, bold=False, italic=False,
               align=WDA.CENTER, sz=SZ,
               fill=None, valign="center"):
    if fill: set_shd(cell, fill)
    set_valign(cell, valign)
    tc = cell._tc
    for p in tc.findall(qn("w:p")): tc.remove(p)
    jc = {WDA.LEFT:"left",WDA.CENTER:"center",
          WDA.RIGHT:"right",WDA.JUSTIFY:"both"}.get(align,"left")
    b  = "<w:b/><w:bCs/>" if bold   else ""
    it = "<w:i/><w:iCs/>" if italic else ""
    tc.append(parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr><w:jc w:val="{jc}"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
        f'<w:rPr>{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'</w:pPr>'
        f'<w:r><w:rPr>{b}{it}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
        f'</w:p>'))

def write_cell_lines(cell, lines, align=WDA.LEFT,
                     sz=SZ, fill=None, valign="top"):
    """lines = [(text, bold, italic), ...]"""
    if fill: set_shd(cell, fill)
    set_valign(cell, valign)
    tc = cell._tc
    for p in tc.findall(qn("w:p")): tc.remove(p)
    jc = {WDA.LEFT:"left",WDA.CENTER:"center",
          WDA.RIGHT:"right",WDA.JUSTIFY:"both"}.get(align,"left")
    for text, bold, italic in lines:
        b  = "<w:b/><w:bCs/>" if bold   else ""
        it = "<w:i/><w:iCs/>" if italic else ""
        tc.append(parse_xml(
            f'<w:p {nsdecls("w")}>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'<w:rPr>{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'</w:pPr>'
            f'<w:r><w:rPr>{b}{it}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(str(text) if text else "")}</w:t></w:r>'
            f'</w:p>'))

def _insert_before_sectpr(doc, element):
    """Insert XML element before sectPr to maintain correct order."""
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    if sectPr is not None:
        body.insert(list(body).index(sectPr), element)
    else:
        body.append(element)

def body_p(doc, text="", bold=False, italic=False,
           align=WDA.LEFT, sz=SZ):
    jc = {WDA.LEFT:"left",WDA.CENTER:"center",
          WDA.RIGHT:"right",WDA.JUSTIFY:"both"}.get(align,"left")
    b  = "<w:b/><w:bCs/>" if bold   else ""
    it = "<w:i/><w:iCs/>" if italic else ""
    _insert_before_sectpr(doc, parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr><w:jc w:val="{jc}"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
        f'<w:rPr>{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'</w:pPr>'
        + (f'<w:r><w:rPr>{b}{it}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
           f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>' if text else "")
        + '</w:p>'))

def empty(doc, n=1):
    for _ in range(n): body_p(doc)

def add_tbl(doc, col_widths, rows_data,
            border="single", bsz="4", bcol="auto", tw=10627):
    """
    rows_data: list of rows, each row a list of cell specs.
    Cell spec: (text, bold, italic, fill, align, sz)
               or just a str for plain centered text
    """
    nr = len(rows_data); nc = len(col_widths)
    tbl = doc.add_table(rows=nr, cols=nc)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl, border, bsz, bcol)
    set_tblW(tbl, tw)
    set_grid(tbl, col_widths)
    for ri, row in enumerate(rows_data):
        for ci, spec in enumerate(row):
            cell = tbl.cell(ri, ci)
            set_cw(cell, col_widths[ci])
            if isinstance(spec, str):
                write_cell(cell, spec)
            else:
                t    = spec[0] if len(spec) > 0 else ""
                bold = spec[1] if len(spec) > 1 else False
                ita  = spec[2] if len(spec) > 2 else False
                fill = spec[3] if len(spec) > 3 else None
                aln  = spec[4] if len(spec) > 4 else WDA.CENTER
                s    = spec[5] if len(spec) > 5 else SZ
                write_cell(cell, t, bold=bold, italic=ita,
                           fill=fill, align=aln, sz=s)
    return tbl

# Cell spec helpers
def H(t):  return (t, True, False, F_GREY, WDA.CENTER, SZ)
def HD(t): return (t, True, False, F_DARK, WDA.CENTER, SZ)
def HL(t): return (t, True, False, F_LABEL, WDA.LEFT, SZ)
def HW(t): return (t, False,False, F_WHITE,WDA.CENTER, SZ)
def DT(t): return (t, True, False, F_DATA, WDA.CENTER, SZ)
def DW(t): return (str(t) if t is not None else "", False, False, F_WHITE, WDA.CENTER, SZ)
def D(t, bold=False, aln=WDA.CENTER):
    return (str(t) if t is not None else "", bold, False, None, aln, SZ)
def DL(t, bold=False): return D(t, bold, WDA.LEFT)

def fmt(v, dec=2):
    if v is None or v == "": return ""
    try: return f"{float(v):.{dec}f}"
    except: return str(v)


# ═══════════════════════════════════════════════════
# EXCEL READER
# ═══════════════════════════════════════════════════
def read_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    wi = wb["Product Info"]
    info = {}
    for row in wi.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1] is not None:
            info[str(row[0]).strip()] = row[1]

    PN  = str(info.get("Product Name", ""))
    AI  = str(info.get("Active Ingredient", ""))
    MVL = str(info.get("MVL Number", ""))
    TP  = str(info.get("Testing Procedure", ""))
    HID = str(info.get("HPLC ID", ""))
    WSL = str(info.get("Working Standard Lot", ""))
    SMB = str(info.get("SM Batch Number", ""))
    STR = str(info.get("Strength (mg)", ""))

    ws = wb["Validation Data"]
    cv = lambda c, r: ws.cell(row=r, column=c).value
    def fv(c, r, dec=4):
        v = cv(c, r)
        if v is None or v == "": return ""
        try: return round(float(v), dec)
        except: return str(v)

    # Locate sections
    sec = {}
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(row=r, column=1).value or "").strip()
        if   a.startswith("1. System"): sec["assay"] = r
        elif a.startswith("2. Lin"):    sec["lin"]   = r
        elif a.startswith("3. Acc"):    sec["acc"]   = r
        elif a.startswith("4. Int"):    sec["ip"]    = r
        elif a.startswith("5. For"):    sec["deg"]   = r
        elif a.startswith("6. Rob"):    sec["rob"]   = r
        elif a.startswith("7. LOD"):    sec["lod"]   = r

    # ── System Suitability Day 1 ──
    d1_hr = sec["assay"] + 1
    d1_ar = d1_hr + 1
    d1 = {
        "areas": [cv(c, d1_ar) for c in range(2, 8)],
        "av":  fv(2, d1_ar+1, 2), "sd":  fv(2, d1_ar+2, 2),
        "rsd": fv(2, d1_ar+3, 4), "purity": fv(2, d1_ar+4, 2),
    }
    # Day 2 starts 6 rows after Day 1 area row (area+av+sd+rsd+purity+blank)
    d2_hr = d1_ar + 6
    d2_ar = d2_hr + 1
    d2 = {
        "areas": [cv(c, d2_ar) for c in range(2, 8)],
        "av":  fv(2, d2_ar+1, 2), "sd":  fv(2, d2_ar+2, 2),
        "rsd": fv(2, d2_ar+3, 4), "purity": fv(2, d2_ar+4, 2),
    }

    # ── Linearity ──
    ls = sec["lin"] + 2      # first Spl row (skip section title + col header)
    lin_a = []
    for level, ac, rc in [(80,2,3),(90,4,5),(100,6,7)]:
        lin_a.append({
            "pct": level,
            "areas":     [cv(ac, ls+i) for i in range(3)],
            "results":   [fv(rc, ls+i, 4) for i in range(3)],
            "av_area":   fv(ac, ls+3, 2),
            "av_result": fv(rc, ls+3, 4),
            "rsd":       fv(ac, ls+4, 4),
        })
    ls2 = ls + 7             # skip spl1/spl2/spl3/av/rsd + blank + col header
    lin_b = []
    for level, ac, rc in [(110,2,3),(120,4,5)]:
        lin_b.append({
            "pct": level,
            "areas":     [cv(ac, ls2+i) for i in range(3)],
            "results":   [fv(rc, ls2+i, 4) for i in range(3)],
            "av_area":   fv(ac, ls2+3, 2),
            "av_result": fv(rc, ls2+3, 4),
            "rsd":       fv(ac, ls2+4, 4),
        })

    # ── Accuracy ──
    as_ = sec["acc"] + 2
    acc = []
    for level, ac, rc in [(80,2,3),(100,4,5),(120,6,7)]:
        acc.append({
            "pct": level,
            "areas":       [cv(ac, as_+i) for i in range(3)],
            "recoveries":  [fv(rc, as_+i, 2) for i in range(3)],
            "av_area":     fv(ac, as_+3, 2),
            "av_result":   fv(rc, as_+3, 2),
            "rsd":         fv(ac, as_+4, 4),
        })

    # ── Intermediate Precision ──
    ips = sec["ip"] + 2
    ip = []
    for name in ["Analyst 1 \u2013 Day 1",
                 "Analyst 2 \u2013 Day 1",
                 "Analyst 1 \u2013 Day 2"]:
        ip.append({
            "name":    name,
            "areas":   [cv(3, ips+i) for i in range(3)],
            "results": [fv(4, ips+i, 2) for i in range(3)],
            "avg_a":   fv(3, ips+3, 2), "avg_r": fv(4, ips+3, 2),
            "rsd_r":   fv(4, ips+5, 2),
        })
        ips += 8   # 3 samples + avg + sd + rsd + 2 blank

    # ── Forced Degradation ──
    ds = sec["deg"] + 2
    deg = []
    for name in ["Acid Degradation", "Base Degradation",
                 "H\u2082O\u2082 Degradation", "Heat Treated"]:
        deg.append({
            "name":    name,
            "areas":   [cv(3, ds+i) for i in range(3)],
            "results": [fv(4, ds+i, 2) for i in range(3)],
            "av_r":    fv(6, ds, 2),
        })
        ds += 3

    # ── Robustness ──
    rs = sec["rob"] + 2
    rob = []
    for i in range(0, 4, 2):
        rob.append({
            "flow":   str(cv(1, rs+i) or ""),
            "std_a":  cv(3, rs+i),
            "spl_a":  cv(3, rs+i+1),
            "result": fv(4, rs+i, 4),
        })

    # ── LOD / LOQ ──
    lr = sec["lod"] + 2
    lod = {"mg": fv(2,lr,6), "ug": fv(3,lr,4),
           "area": fv(4,lr,2), "sd": fv(5,lr,2), "slope": fv(6,lr,0)}
    loq = {"mg": fv(2,lr+1,6), "ug": fv(3,lr+1,4), "area": fv(4,lr+1,2)}

    lin_concs = ([{"pct":l["pct"],"av":l["av_area"]} for l in lin_a]
               + [{"pct":l["pct"],"av":l["av_area"]} for l in lin_b])

    return dict(PN=PN, AI=AI, MVL=MVL, TP=TP, HPLC=HID, WSL=WSL, SMB=SMB,
                STR=STR,
                d1=d1, d2=d2, lin_a=lin_a, lin_b=lin_b,
                acc=acc, ip=ip, deg=deg, rob=rob,
                lod=lod, loq=loq, lin_concs=lin_concs)


# ═══════════════════════════════════════════════════
# DOCUMENT BUILDER  —  exact replica of shaigan template
# ═══════════════════════════════════════════════════
def build(data, template, output):
    shutil.copy(template, output)
    doc = Document(output)
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    for ch in list(body):
        if ch is not sectPr: body.remove(ch)

    PN  = data["PN"]
    AI  = data["AI"];  MVL = data["MVL"]; TP = data["TP"]
    HID = data["HPLC"]; WSL = data["WSL"]; SMB = data["SMB"]
    STR = data["STR"]
    d1  = data["d1"];  d2  = data["d2"]

    # ── section header helper (grey left + description right) ──────────────
    def sec_hdr(num, title, desc, h=1097):
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_borders(tbl); set_tblW(tbl, 10627); set_grid(tbl, [2808, 7819])
        write_cell_lines(tbl.cell(0,0),
            [(f"{num}. {title}:", True, False),
             (f"       {AI}",     True, False)],
            align=WDA.LEFT, fill=F_GREY, valign="center")
        set_cw(tbl.cell(0,0), 2808)
        write_cell(tbl.cell(0,1), desc, align=WDA.JUSTIFY, valign="center")
        set_cw(tbl.cell(0,1), 7819)
        set_rowh(tbl.rows[0], h)

    # ══════════════════════════════════════════════
    # P[0]  MVL title
    # ══════════════════════════════════════════════
    body_p(doc, f"    Method Validation Report No: {MVL}", bold=True)

    # P[1] blank
    empty(doc, 1)

    # TBL[2]  info table — double border, F_LABEL labels
    tbl = doc.add_table(rows=3, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl, "double"); set_tblW(tbl, 10080); set_grid(tbl, [1525,3240,2160,3155])
    rows_info = [
        ("Product Name",    PN,          "Active Ingredient", AI),
        ("Testing Procedure", TP,        "SOP Followed",      SOP),
        ("Department",      DEPARTMENT,  "Revision No",       REVISION),
    ]
    for ri, (l1, v1, l2, v2) in enumerate(rows_info):
        write_cell(tbl.cell(ri,0), l1, bold=True, fill=F_LABEL, align=WDA.LEFT)
        set_cw(tbl.cell(ri,0), 1525)
        write_cell(tbl.cell(ri,1), v1, align=WDA.LEFT)
        set_cw(tbl.cell(ri,1), 3240)
        write_cell(tbl.cell(ri,2), l2, bold=True, fill=F_LABEL, align=WDA.LEFT)
        set_cw(tbl.cell(ri,2), 2160)
        write_cell(tbl.cell(ri,3), v2, align=WDA.LEFT)
        set_cw(tbl.cell(ri,3), 3155)
    set_rowh(tbl.rows[0], 735)
    set_rowh(tbl.rows[1], 400)
    set_rowh(tbl.rows[2], 400)

    # P[3] P[4] blank
    empty(doc, 2)

    # TBL[5]  signature table — double border
    sig = doc.add_table(rows=5, cols=2)
    sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(sig, "double"); set_tblW(sig, 10278); set_grid(sig, [5139,5139])
    sig_data = [
        ("Personnel\u2019s Involved in Validation Study", "Signature",             True,  530),
        ("Prepared By: R&D  Analyst",                     "",                      True,  1349),
        ("Reviewed By: SAS  Manager",                     "",                      True,  1349),
        ("Approved By: Sr.GM Quality Operations",         "",                      True,  1349),
        ("Date of Approval:",                             "",                      False, 1349),
    ]
    for ri, (t1, t2, bold, ht) in enumerate(sig_data):
        write_cell(sig.cell(ri,0), t1, bold=bold, align=WDA.LEFT)
        set_cw(sig.cell(ri,0), 5139)
        write_cell(sig.cell(ri,1), t2, bold=bold, align=WDA.CENTER)
        set_cw(sig.cell(ri,1), 5139)
        set_rowh(sig.rows[ri], ht)

    # P[6]–P[11]  6 blank lines (signature space)
    empty(doc, 6)

    # P[18]
    body_p(doc, "(Not Valid without Signature)", bold=True, align=WDA.CENTER)

    # P[19] P[20] blank
    empty(doc, 2)

    # TBL[21]  pre-verification — single col, no explicit border in ref (nil)
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl); set_tblW(tbl, 10296); set_grid(tbl, [10296])
    write_cell_lines(tbl.cell(0,0), [
        ("Pre-VERIFICATION Requirements", True, False),
        ("", False, False),
        ("Installation Qualification (IQ)", True, False),
        ("Documented Validation that all key aspects of an installed High performance Liquid "
         "chromatography system (Shimadzu) adhere to the approved design specification and that "
         "the recommendations of the Shimadzu have been suitably considered.", False, False),
        ("", False, False),
        ("Operational Qualification (OQ)", True, False),
        ("Operational Validation carried out after installation that shows High performance "
         "Liquid chromatography system Shimadzu Prominence-I LC-2030 Plus performs in accordance "
         "with Shimadzu specifications and process requirements and that the appropriate GMP "
         "systems (e.g. training, calibration, and maintenance, etc.) are in place.", False, False),
        ("", False, False),
        ("Calibration status of equipment", True, False),
        ("Equipment was calibrated and bears calibration sticker of the external calibrator.",
         False, False),
        ("", False, False),
        ("Facilities:", True, False),
        ("The Validation of the method of Determination was carried out in the Analytical Lab. "
         "of Shaigan Pharmaceuticals (PVT) LTD. in Rawalpindi, PAKISTAN.", False, False),
        ("", False, False),
        (f"Identification of Machine/ Equipment used: HPLC {HID}", True, False),
        ("", False, False),
        ("Traceability of Material or Product used for study:", True, False),
        (f"{AI} Working Standard # {WSL}", True, False),
        (f"{AI} Starting Material B # {SMB}", True, False),
        ("", False, False),
        ("Apparatus/Glassware used:", True, False),
        ("50ml volumetric flask type \u201CA\u201D", False, False),
        ("25ml volumetric flask type \u201CA\u201D", False, False),
        ("1ml pipette type \u201CA\u201D", False, False),
        ("1000 ml cylinder type \u201CA\u201D", False, False),
        ("1000ml beaker type \u201CA\u201D", False, False),
        ("", False, False),
        ("Precautions:", True, False),
        ("Use Pyrex type \u201CA\u201D glassware and broken glassware should not be used. "
         "During analysis temperature should be in between 20-25\u00B0C. Use dried glassware. "
         "Use HPLC grade solvents only and must be filtered with 0.45\u00B5 Filter. Always degas "
         "the mobile phase. Never run the column dry. Use a 0.2\u00B5 sized syringe filter for "
         "sample. Always filter your sample before injection.", False, False),
    ], align=WDA.LEFT, valign="top")

    # P[16]–P[17]  2 blank lines
    empty(doc, 2)

    # TBL[30]  Method of Analysis — outer 2-row table containing nested tables
    # Ref cols: ['10627', '6285', '1912', '1903', '1903']
    # Row1 = single spanning cell with nested grey title + gradient table
    # Row2 = blank
    outer = doc.add_table(rows=2, cols=1)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(outer); set_tblW(outer, 10627); set_grid(outer, [10627])
    set_cw(outer.cell(0,0), 10627); set_valign(outer.cell(0,0), "top")
    set_cw(outer.cell(1,0), 10627)
    set_rowh(outer.rows[1], 50)

    # Build content inside Row1 cell
    tc = outer.cell(0,0)._tc
    for p in tc.findall(qn("w:p")): tc.remove(p)

    # Nested grey title table (cols=['6285'])
    ns = nsdecls("w")
    ntbl1 = (
        f'<w:tbl {ns}>'
        f'<w:tblPr><w:tblW w:w="6285" w:type="dxa"/>'
        + borders_xml() +
        f'</w:tblPr><w:tblGrid><w:gridCol w:w="6285"/></w:tblGrid>'
        f'<w:tr><w:tc>'
        f'<w:tcPr><w:tcW w:w="6285" w:type="dxa"/>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{F_GREY}"/></w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
        f'<w:rPr><w:b/><w:bCs/><w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr>'
        f'<w:t>Method of Analysis to be Validate</w:t></w:r></w:p>'
        f'</w:tc></w:tr></w:tbl>'
    )
    tc.append(parse_xml(ntbl1))

    # Method text paragraphs
    for text, bold in [
        ("Column:             ", True),
        ("Elution:            ", True),
        ("Flow Rate:          ", True),
        ("Detection:          ", True),
        ("Injection Volume:   ", True),
        ("Column Temp:        ", True),
        ("", False),
        ("Solution A:         Weigh accurately about 1.3 g of potassium dihydrogen phosphate "
         "and 0.7 g of disodium hydrogen phosphate transfer it in to 1000 ml of volumetric flask. "
         "Add and dissolve it in to water. Filter it and sonicate for degas.", True),
        ("", False),
        ("Solution B:         Acetonitrile.", True),
        ("", False),
        ("Mobile Phase:                                                       Table 1", True),
    ]:
        b = "<w:b/><w:bCs/>" if bold else ""
        tc.append(parse_xml(
            f'<w:p {ns}>'
            f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'<w:rPr>{b}<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr>{b}<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'))

    # Nested gradient table (cols=['1912','1903','1903'])
    gcols = [1912, 1903, 1903]
    grad = [("Time (Minutes)","Solution A (%)","Solution B (%)"),
            ("0","80","20"),("2","80","20"),("22","25","75"),
            ("27","25","75"),("30","80","20"),("35","80","20")]
    gtbl = (f'<w:tbl {ns}><w:tblPr><w:tblW w:w="5718" w:type="dxa"/>'
            + borders_xml() + f'</w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="1912"/><w:gridCol w:w="1903"/>'
            f'<w:gridCol w:w="1903"/></w:tblGrid>')
    for ri, (t0,t1,t2) in enumerate(grad):
        fill = F_GREY if ri == 0 else ""
        bold = ri == 0
        def gc(t, w):
            shd = (f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
                   if fill else "")
            b = "<w:b/><w:bCs/>" if bold else ""
            return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}</w:tcPr>'
                    f'<w:p><w:pPr><w:jc w:val="center"/>'
                    f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
                    f'</w:pPr>'
                    f'<w:r><w:rPr>{b}<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr>'
                    f'<w:t>{esc(t)}</w:t></w:r></w:p></w:tc>')
        gtbl += f'<w:tr>{gc(t0,1912)}{gc(t1,1903)}{gc(t2,1903)}</w:tr>'
    gtbl += '</w:tbl>'
    tc.append(parse_xml(gtbl))

    # P[31] P[32]  blank
    empty(doc, 2)

    # TBL[33]  standard/sample prep — single col
    tbl33 = doc.add_table(rows=1, cols=1)
    tbl33.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl33); set_tblW(tbl33, 10627); set_grid(tbl33, [10627])
    write_cell_lines(tbl33.cell(0,0), [
        (f"Standard solution: Weigh accurately about {STR} mg of {AI} reference standard/working "
         "standard, transfer it in to 100 ml of volumetric flask. Add 20 % of total volume "
         "acetonitrile, mix well. Dilute with mobile phase A up to the mark. Mix well & filter "
         "through 0.22\u00B5m syringe filter.", True, False),
        ("", False, False),
        ("Sample Preparation:", True, False),
        (f"Weigh accurately about {STR} mg of sample, transfer it in to 100 ml of volumetric flask. "
         "Add 20 % of total volume acetonitrile, mix well. Dilute with mobile phase A up to the "
         "mark. Mix well & filter through 0.22\u00B5m syringe filter.", True, False),
        ("", False, False),
        ("Procedure for Injection:", True, False),
        (f"Separately inject equal volumes of system suitability solution, standard and sample "
         f"solution i.e.10\u00B5L and record chromatogram. Measure the peak response of {AI} "
         "at 242nm.", False, False),
        ("", False, False),
        ("Calculation:", True, False),
        ("                    Result = AT/AS   x Cs / CT x P", False, False),
        ("Where,", False, False),
        ("AT: Area of sample solution", False, False),
        ("AS: Area of standard solution", False, False),
        ("CS: Concentration of standard solution (mg/ml)", False, False),
        ("CT: Concentration of sample solution (mg/ml)", False, False),
        ("    P: Purity of working standard", False, False),
        ("", False, False),
        ("Method Specification: Manufacturer\u2019s Specification", True, False),
    ], align=WDA.LEFT, valign="top")

    # 19 blanks before params
    empty(doc, 19)

    # TBL  Parameters Under Study — header row + content row in one table
    tbl35 = doc.add_table(rows=2, cols=1)
    tbl35.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl35); set_tblW(tbl35, 10627); set_grid(tbl35, [10627])
    write_cell(tbl35.cell(0,0), "Parameters Under Study", bold=True, fill=F_GREY)
    set_rowh(tbl35.rows[1], 1077)
    write_cell_lines(tbl35.cell(1,0), [
        ("1.  Linearity and Range", True, False),
        ("2.  Accuracy and recovery", True, False),
        ("3.  Precision", True, False),
        ("    3.1. Repeatability", True, False),
        ("    3.2. Intermediate Precision (Ruggedness)", True, False),
        ("          3.2.1 Within Days Variation", True, False),
        ("          3.2.2 By different analyst", True, False),
        ("4.  Robustness", True, False),
        ("         4.1 Changing Mobile Phase Composition", False, False),
        ("5.  Force Degradation", True, False),
        ("6.  Detection Limit", True, False),
        ("7.  Quantitation Limit", True, False),
        ("8.  Specificity", True, False),
    ], align=WDA.LEFT, valign="top")

    # 19 blank lines (push to next page)
    empty(doc, 19)

    # ══════════════════════════════════════════════
    # TBL[41]  1. LINEARITY section header
    # ══════════════════════════════════════════════
    sec_hdr("1", "Linearity and Range",
        "For the establishment of linearity concentration ranges of a test substance "
        "(from 80% to 120 %) were prepared to determine assay of drug substance. "
        "The range of this analytical method lies between the 80% to 120% of the test "
        "concentration.", h=1097)

    # P[42]  blank
    empty(doc, 1)

    # TBL[43]  linearity data — cols 3947,1475,5205  rows 6
    lin_rows = [[H("Concentration %"), H("Conc.mg/ml"), H("Peak area; where n=2")]]
    rh_lin   = [428, 261, 273, 261, 261, 273]
    for lv in data["lin_a"] + data["lin_b"]:
        lin_rows.append([D(str(lv["pct"])),
                         D(fmt(CONC.get(lv["pct"],0), 3)),
                         D(fmt(lv["av_area"], 0))])
    t = add_tbl(doc, [3947,1475,5205], lin_rows)
    for ri, row in enumerate(t.rows):
        set_rowh(row, rh_lin[ri] if ri < len(rh_lin) else 261)

    # P[44]  blank
    empty(doc, 1)

    # P[45]  Graphical Presentation
    body_p(doc, "Graphical Presentation", bold=True, align=WDA.CENTER)

    # 16 blank lines (space for chart)
    empty(doc, 16)

    # P[49]  Correlation Coefficient
    body_p(doc, "Correlation Coefficient", bold=True, align=WDA.CENTER)

    # TBL[50]  acceptance — cols 6461,4166
    add_tbl(doc, [6461,4166], [
        [H("Acceptance Criteria"), H("Result")],
        [D("Correlation Coefficient\u2265 0.997", bold=True),
         D("0.999", bold=True)],
    ])

    # 1 blank before accuracy header
    empty(doc, 1)

    # ══════════════════════════════════════════════
    # TBL[56]  2. ACCURACY section header — cols 3082,7545
    # ══════════════════════════════════════════════
    tbl56 = doc.add_table(rows=1, cols=2)
    tbl56.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl56); set_tblW(tbl56, 10627); set_grid(tbl56, [3082,7545])
    write_cell_lines(tbl56.cell(0,0),
        [(f"2. Accuracy and Recovery", True, False),
         (f"       {AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl56.cell(0,0), 3082)
    write_cell(tbl56.cell(0,1),
        "In the study of accuracy / recovery, Known amount of sample corresponding to three "
        "concentration levels i.e. 80%, 100 % and 120 % were taken. Results were obtained using "
        "three concentration levels (80%, 100%, and 120 %) and three replicate of each "
        "concentration. Assessment of accuracy was established by evaluating the recovery of the "
        "analyte (Percentage recovery) across the range of assay. "
        "Acceptance criteria: (theoretical amount \u00B1 2%)",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl56.cell(0,1), 7545)
    set_rowh(tbl56.rows[0], 899)

    # 1 blank after accuracy header
    empty(doc, 1)

    # TBL[59]  "Accuracy and Recovery" label — cols 4410  double border
    add_tbl(doc, [4410], [[H("Accuracy and Recovery")]],
            border="double", tw=4410)

    # Sample I, II, III tables — cols 2206,2091,2089,1282,2959
    bounds = {80:("78","82"), 100:("98","102"), 120:("118","122")}
    labels = ["Sample I:", "Sample II:", "Sample III:"]
    for idx, lvl in enumerate(data["acc"]):
        # P[61]/[64]/[67]
        body_p(doc, labels[idx], bold=True)
        # TBL[62]/[65]/[68]
        arows = [[H("Level %"),H("Conc.mg/ml"),H("Peak Area"),H("%RSD"),H("%age Recovery")]]
        for i in range(3):
            arows.append([
                D(str(lvl["pct"]) if i==0 else ""),
                D(fmt(CONC.get(lvl["pct"],0), 3) if i==0 else ""),
                D(fmt(lvl["areas"][i], 0)),
                D(fmt(lvl["rsd"], 2) if i==0 else ""),
                D(fmt(lvl["recoveries"][i], 2)),
            ])
        add_tbl(doc, [2206,2091,2089,1282,2959], arows)
        # P[63]/[66] blank (no blank after last)
        if idx < 2: empty(doc, 1)

    # P[69]  blank
    empty(doc, 1)

    # P[70]  Grand average label
    body_p(doc, "Grand average %age recovery of three Samples:", bold=True)

    # TBL[71]  grand average — cols 6606,4021  rows 4
    avg_rows = [[H("Acceptance Criteria \u00B1 2%"), H("Average % Recovery")]]
    for lvl in data["acc"]:
        lo, hi = bounds.get(lvl["pct"], ("",""))
        avg_rows.append([
            DL(f"For {lvl['pct']} % Level    {lo} \u2013 {hi} %"),
            D(fmt(lvl["av_result"], 2) + "%", bold=True),
        ])
    t = add_tbl(doc, [6606,4021], avg_rows)
    for row in t.rows: set_rowh(row, 350)

    # 1 blank line
    empty(doc, 1)

    # ══════════════════════════════════════════════
    # TBL[79]  3. PRECISION section header — cols 2772,7855
    # ══════════════════════════════════════════════
    tbl_prec = doc.add_table(rows=1, cols=2)
    tbl_prec.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl_prec); set_tblW(tbl_prec, 10627); set_grid(tbl_prec, [2808, 7819])
    write_cell(tbl_prec.cell(0,0), "3. Precision (Repeatability):", bold=True, fill=F_GREY, align=WDA.LEFT)
    set_cw(tbl_prec.cell(0,0), 2808)
    write_cell(tbl_prec.cell(0,1),
        "Peak areas of Six replicates of homogenous sample were determined. Mean value, Standard "
        "deviation and relative standard were calculated. The relative standard deviation of six "
        "determinations at 100% of the test concentration should not be greater than 2% for drug "
        "product.", align=WDA.JUSTIFY, valign="center")
    set_cw(tbl_prec.cell(0,1), 7819)
    set_rowh(tbl_prec.rows[0], 674)

    # 1 blank after precision header
    empty(doc, 1)

    # TBL[83]  "3.1 Precision (Repeatability)" — white fill, cols 10627
    tbl83 = doc.add_table(rows=1, cols=1)
    tbl83.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl83); set_tblW(tbl83, 10627); set_grid(tbl83, [10627])
    set_rowh(tbl83.rows[0], 458)
    write_cell(tbl83.cell(0,0), "Precision (Repeatability)", fill=F_DATA, align=WDA.LEFT)

    # P[84]  blank
    empty(doc, 1)

    # TBL[85]  repeatability — cols 5578,4833  rows 10
    rep_rows = [[D("Replicate"), D("Peak Area")]]
    rh_rep = [277,277,277,277,277,277,277,548,683,530]
    for i, a in enumerate(d1["areas"]):
        rep_rows.append([D(f"Replicate {i+1}"), D(fmt(a, 0))])
    rep_rows += [
        [("Mean",           True,False,F_DATA,WDA.CENTER,SZ),
         (fmt(d1["av"],2),  True,False,F_DATA,WDA.CENTER,SZ)],
        [("Standard Deviation",                    True,False,F_DATA,WDA.LEFT,SZ),
         (fmt(d1["sd"],2),                         False,False,F_DATA,WDA.CENTER,SZ)],
        [("Relative Standard deviation (RSD%)",    True,False,F_DATA,WDA.LEFT,SZ),
         (fmt(d1["rsd"],4),                        True,False,F_DATA,WDA.CENTER,SZ)],
    ]
    t = add_tbl(doc, [5694,4933], rep_rows)
    for ri, row in enumerate(t.rows):
        set_rowh(row, rh_rep[ri] if ri < len(rh_rep) else 277)

    # P[86]–P[89]  4 blank lines
    empty(doc, 4)

    # TBL[90]  acceptance — cols 5643,4984
    add_tbl(doc, [5643,4984], [
        [H("Acceptance Criteria"), H("Result")],
        [D("RSD < 2.0 %", bold=True), D(fmt(d1["rsd"],2))],
    ])

    # 8 blank lines after repeatability acceptance
    empty(doc, 8)

    # ══════════════════════════════════════════════
    # TBL[101]  3.2 Intermediate Precision header — cols 2628,7452
    # ══════════════════════════════════════════════
    tbl101 = doc.add_table(rows=1, cols=2)
    tbl101.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl101); set_tblW(tbl101, 10627); set_grid(tbl101, [2771,7856])
    write_cell_lines(tbl101.cell(0,0),
        [("  3.2. Intermediate Precision (Ruggedness)", True, False),
         (f"       {AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl101.cell(0,0), 2771)
    write_cell(tbl101.cell(0,1),
        "Intermediate precision (Ruggedness) expresses within laboratory variation as on "
        "different days and with different analyst.",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl101.cell(0,1), 7856)
    set_rowh(tbl101.rows[0], 674)

    # P[102]  blank
    empty(doc, 1)

    # TBL[103]  "3.2.1 Within Days Variation" — double border, cols 4932
    add_tbl(doc, [4932],
            [[(f"3.2.1 Within Days Variation ({AI})", True,False,F_GREY,WDA.LEFT,SZ)]],
            border="double", tw=4932)
    doc.tables[-1].rows[0]  # ref h=600
    set_rowh(doc.tables[-1].rows[0], 600)

    # 1 blank after 4932 table
    empty(doc, 1)

    # TBL[107]  Within Days table — cols 1467,2152,2339,2341,2069  rows 8
    # Row1=header(A5A5A5), Row2=spacer, then Day1×3, Day2×3
    ip_days = [data["ip"][0], data["ip"][2]]
    day_lbls = ["Day 1", "Day 2"]
    wd_rows = [
        [HD("Day"),HD("Sample Area"),HD("Assay%"),HD("Average Assay %"),HD("%RSD")],
    ]
    for idx, ip in enumerate(ip_days):
        for i in range(3):
            wd_rows.append([
                (day_lbls[idx] if i==0 else "", True,False,F_DARK,WDA.CENTER,SZ),
                DW(fmt(ip["areas"][i], 0)),
                DW(fmt(ip["results"][i], 2)),
                DW(fmt(ip["avg_r"], 2)) if i==0 else DW(""),
                DW(fmt(ip["rsd_r"], 2)) if i==0 else DW(""),
            ])
    t = add_tbl(doc, [1504,2206,2397,2399,2121], wd_rows)
    set_rowh(t.rows[0], 810)
    for ri in range(1, len(t.rows)): set_rowh(t.rows[ri], 315)

    # P[108]  single space
    body_p(doc, " ")

    # TBL[109]  acceptance — cols 6606,4021  rows 3
    add_tbl(doc, [6606,4021],
        [[H("Acceptance Criteria"), H("Results%")]]
        + [[D("RSD < 2.0 %", bold=True), D(fmt(ip["rsd_r"],2))] for ip in ip_days])
    for row in doc.tables[-1].rows: set_rowh(row, 350)

    # 14 blank lines before analyst section
    empty(doc, 14)

    # TBL[125]  "3.2.2 By Different Analyst" — double border, cols 3690
    add_tbl(doc, [3690],
            [[(f"3.2.2. By Different Analyst ({AI})", True,False,F_GREY,WDA.LEFT,SZ)]],
            border="double", tw=3690)
    set_rowh(doc.tables[-1].rows[0], 303)

    # P[126]  spaces (ref has a long spaces string here)
    body_p(doc, "                                                                  ")

    # TBL[128]  analyst table — cols 2242,1733,2376,2050,2226  rows 7
    ip_an = data["ip"][:2]
    an_rows = [[HD("Analyst"),HD("Peak Area"),HD("Assay %"),HD("Average Assay %"),HD("%RSD")]]
    for idx, ip in enumerate(ip_an):
        for i in range(3):
            an_rows.append([
                (str(idx+1) if i==0 else "", True,False,F_DARK,WDA.CENTER,SZ),
                DW(fmt(ip["areas"][i], 0)),
                DW(fmt(ip["results"][i], 2)),
                DW(fmt(ip["avg_r"], 2)) if i==0 else DW(""),
                DW(fmt(ip["rsd_r"], 2)) if i==0 else DW(""),
            ])
    t = add_tbl(doc, [2242,1733,2376,2050,2226], an_rows)
    set_rowh(t.rows[0], 1106)
    for ri in range(1, len(t.rows)): set_rowh(t.rows[ri], 315)

    # P[129] P[130]  blank
    empty(doc, 2)

    # TBL[131]  acceptance — cols 6926,3701  rows 3
    add_tbl(doc, [6926,3701],
        [[H("Acceptance Criteria"), H("Result %")]]
        + [[D("RSD < 2.0 %", bold=True), D(fmt(ip["rsd_r"],2))] for ip in ip_an])

    # P[132]–P[146]  15 blank lines
    empty(doc, 15)

    # ══════════════════════════════════════════════
    # TBL[147]  4. ROBUSTNESS header — cols 2772,7855
    # ══════════════════════════════════════════════
    tbl147 = doc.add_table(rows=1, cols=2)
    tbl147.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl147); set_tblW(tbl147, 10627); set_grid(tbl147, [2772,7855])
    write_cell_lines(tbl147.cell(0,0),
        [("  4. ROBUSTNESS", True, False),
         (f"       {AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl147.cell(0,0), 2772)
    write_cell(tbl147.cell(0,1),
        "Robustness of an analytical procedure is a measure of its capacity to remain unaffected "
        "by small but deliberate variations in procedural parameters. (Acceptance criteria: Method "
        "should be robust and it should not show a change in assay from the actual method "
        "parameters by more than 2%)",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl147.cell(0,1), 7855)
    set_rowh(tbl147.rows[0], 980)

    # 1 blank after robustness header
    empty(doc, 1)

    # TBL[150]  Robustness data — cols 2739,821,2706,221,4111  rows 7  nil border
    # Row1=(AI), Row2=header, Row3=spacer, Row4–7=data (flow+std+spl pairs)
    rob_xml = (
        f'<w:tbl {nsdecls("w")}>'
        f'<w:tblPr><w:tblW w:w="10627" w:type="dxa"/>'
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="nil"/><w:left w:val="nil"/>'
        f'<w:bottom w:val="nil"/><w:right w:val="nil"/>'
        f'<w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        f'</w:tblBorders></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="2739"/><w:gridCol w:w="821"/>'
        f'<w:gridCol w:w="2706"/><w:gridCol w:w="221"/><w:gridCol w:w="4111"/></w:tblGrid>'
    )
    ns2 = nsdecls("w")
    def _rc(text, w, bold=False, fill="", span=1):
        shd = (f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else "")
        gs  = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
        b   = "<w:b/><w:bCs/>" if bold else ""
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{gs}{shd}'
                f'<w:vAlign w:val="center"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="center"/>'
                f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
                f'</w:pPr><w:r><w:rPr>{b}'
                f'<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr>'
                f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p></w:tc>')

    rob_xml += (f'<w:tr><w:trPr><w:trHeight w:val="308" w:hRule="atLeast"/></w:trPr>'
                + _rc(f"({AI})", 2706) + _rc("",4111) + '</w:tr>')
    rob_xml += (f'<w:tr><w:trPr><w:trHeight w:val="832" w:hRule="atLeast"/></w:trPr>'
                + _rc("Mobile Phase Flow Rate",2739,bold=True,fill=F_DARK)
                + _rc("Peak Area",3748,bold=True,fill=F_DARK,span=3)
                + _rc("Assay %",4111,bold=True,fill=F_DARK)
                + '</w:tr>')
    rob_xml += (f'<w:tr><w:trPr><w:trHeight w:val="283" w:hRule="atLeast"/></w:trPr>'
                + _rc("",2739) + _rc("",3748,span=3) + _rc("",4111) + '</w:tr>')
    rh_rob = [441,441,460,478]
    for ri, r in enumerate(data["rob"]):
        ht1 = rh_rob[ri*2]   if ri*2   < len(rh_rob) else 441
        ht2 = rh_rob[ri*2+1] if ri*2+1 < len(rh_rob) else 441
        rob_xml += (f'<w:tr><w:trPr><w:trHeight w:val="{ht1}" w:hRule="atLeast"/></w:trPr>'
                    + _rc(r["flow"],2739,bold=True,fill=F_DARK)
                    + _rc(fmt(r["std_a"],0),3748,fill=F_WHITE,span=3)
                    + _rc(fmt(r["result"],4),4111,fill=F_WHITE)
                    + '</w:tr>')
        rob_xml += (f'<w:tr><w:trPr><w:trHeight w:val="{ht2}" w:hRule="atLeast"/></w:trPr>'
                    + _rc("",2739)
                    + _rc(fmt(r["spl_a"],0),3748,fill=F_WHITE,span=3)
                    + _rc("",4111,fill=F_WHITE)
                    + '</w:tr>')
    rob_xml += '</w:tbl>'
    _insert_before_sectpr(doc, parse_xml(rob_xml))

    # 18 blank lines
    empty(doc, 18)

    # ══════════════════════════════════════════════
    # TBL[171]  5. FORCED DEGRADATION header — cols 2011,8616
    # ══════════════════════════════════════════════
    tbl171 = doc.add_table(rows=1, cols=2)
    tbl171.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl171); set_tblW(tbl171, 10627); set_grid(tbl171, [2358,8269])
    write_cell_lines(tbl171.cell(0,0),
        [("5. Forced Degradation    ", True, False),
         (f"       {AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl171.cell(0,0), 2358)
    write_cell(tbl171.cell(0,1),
        "Forced degradation is a measure to analyze the effect of different solvents on the "
        "active pharmaceutical ingredient. Sample Solutions were subjected to stress conditions "
        "in 0.1N HCl, 0.1N NaOH and Hydrogen Peroxide Solution 3%. The chromatographic analysis "
        "of the sample solutions were carried out under same conditions.",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl171.cell(0,1), 8269)
    set_rowh(tbl171.rows[0], 980)

    # P[172]  blank
    empty(doc, 1)

    # P[173]  AI name
    body_p(doc, AI, bold=False)

    # TBL[174]  forced degradation data — cols 2600,2600,2600,2601  rows 7
    deg_rows = [[H("Degradation Type"),H("Peak Area"),H("Assay %"),H("Average Assay %")]]
    for d in data["deg"]:
        for i, (a, r) in enumerate(zip(d["areas"], d["results"])):
            if a is None and r == "": continue
            deg_rows.append([
                (d["name"], True,False,F_GREY,WDA.LEFT,SZ),   # BFBFBF every row
                D(fmt(a, 0)),
                D(fmt(r, 2)),
                D(fmt(d["av_r"],2), bold=True) if i==0 else D(""),
            ])
    add_tbl(doc, [2657,2656,2656,2658], deg_rows)

    # 15 blank lines
    empty(doc, 15)

    # ══════════════════════════════════════════════
    # TBL[199]  6. LOD header — cols 2772,7855
    # ══════════════════════════════════════════════
    tbl199 = doc.add_table(rows=1, cols=2)
    tbl199.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl199); set_tblW(tbl199, 10627); set_grid(tbl199, [2772,7855])
    write_cell_lines(tbl199.cell(0,0),
        [(f"    6. Limit of Detection         {AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl199.cell(0,0), 2772)
    write_cell(tbl199.cell(0,1),
        "It is the lowest concentration of analyte in a sample that can be detected but not "
        "necessarily quantified. Limit of detection is determined by multiplying standard "
        "deviation with 3.3 and dividing by the slope of the curve. A specific calibration curve "
        "is studied using samples containing an analyte in the range of the limit of detection.",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl199.cell(0,1), 7855)
    set_rowh(tbl199.rows[0], 674)

    # P[200]  blank
    empty(doc, 1)

    # TBL[201]  "Limit of Detection" sub-header — double border, cols 3330
    add_tbl(doc, [3330], [[H("Limit of Detection ")]], border="double", tw=3330)

    # 2 blank before data
    empty(doc, 2)

    # TBL[206]  LOD data — cols 1090,1209,1390,1435,1520,1256,1311,1416  rows 7
    lod = data["lod"]; lc = data["lin_concs"]
    lod_cols = [1090,1209,1390,1435,1520,1256,1311,1416]
    lod_rows = [
        [HD("Sample"),HD("Conc.\n%"),HD("Conc.\nmg/ml"),HD("Peak Area"),
         HD("Standard\nDeviation"),HD("Slope"),
         HD("Conc.\n(\u00B5g/ml)\nat LOD"),HD("Area\nat LOD")],
    ]
    for i, c in enumerate(lc):
        lod_rows.append([
            (str(i+1),False,False,F_DARK,WDA.CENTER,SZ),
            DW(str(c["pct"])),
            DW(fmt(CONC.get(c["pct"],0), 3)),
            DW(fmt(c["av"], 0)),
            DW(fmt(lod["sd"],2))    if i==0 else DW(""),
            DW(fmt(lod["slope"],0)) if i==0 else DW(""),
            DW(fmt(lod["ug"],4))    if i==0 else DW(""),
            DW(fmt(lod["area"],0))  if i==0 else DW(""),
        ])
    t = add_tbl(doc, lod_cols, lod_rows)
    set_rowh(t.rows[0], 1185)
    for ri, ht in enumerate([412,250,270,270,270]):
        if ri+1 < len(t.rows): set_rowh(t.rows[ri+1], ht)

    # 1 blank before footnote
    empty(doc, 1)
    body_p(doc, "* Value of standard deviation is taken from system precision.", italic=True, sz=SZ_SM)

    # 12 blank lines after LOD footnote
    empty(doc, 12)

    # ══════════════════════════════════════════════
    # TBL[234]  7. LOQ header — cols 2772,7855
    # ══════════════════════════════════════════════
    tbl234 = doc.add_table(rows=1, cols=2)
    tbl234.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl234); set_tblW(tbl234, 10627); set_grid(tbl234, [2772,7855])
    write_cell_lines(tbl234.cell(0,0),
        [(f"7. Limit of Quantitation{AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl234.cell(0,0), 2772)
    write_cell(tbl234.cell(0,1),
        "It is the lowest concentration of analyte in a sample that can be quantified with "
        "accuracy and precision. Limit of Quantitation is determined by multiplying standard "
        "deviation with 10 and dividing by the slope of the curve.",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl234.cell(0,1), 7855)
    set_rowh(tbl234.rows[0], 674)

    # P[235]  blank
    empty(doc, 1)

    # TBL[236]  "Limit of Quantitation" sub-header — double border, cols 3330
    add_tbl(doc, [3330], [[H("Limit of Quantitation")]], border="double", tw=3330)
    set_rowh(doc.tables[-1].rows[0], 870)

    # 1 blank before LOQ data
    empty(doc, 1)

    # TBL[242]  LOQ data — same structure as LOD
    loq = data["loq"]
    loq_rows = [
        [HD("Sample"),HD("Conc.\n%"),HD("Conc.\nmg/ml"),HD("Peak Area"),
         HD("Standard\nDeviation"),HD("Slope"),
         HD("Conc.\n(\u00B5g/ml)\nat LOQ"),HD("Area\nat LOQ")],
    ]
    for i, c in enumerate(lc):
        loq_rows.append([
            (str(i+1),False,False,F_DARK,WDA.CENTER,SZ),
            DW(str(c["pct"])),
            DW(fmt(CONC.get(c["pct"],0), 3)),
            DW(fmt(c["av"], 0)),
            DW(fmt(lod["sd"],2))    if i==0 else DW(""),
            DW(fmt(lod["slope"],0)) if i==0 else DW(""),
            DW(fmt(loq["ug"],4))    if i==0 else DW(""),
            DW(fmt(loq["area"],0))  if i==0 else DW(""),
        ])
    t = add_tbl(doc, lod_cols, loq_rows)
    set_rowh(t.rows[0], 1185)
    for ri, ht in enumerate([412,250,270,270,270]):
        if ri+1 < len(t.rows): set_rowh(t.rows[ri+1], ht)

    # 1 blank before LOQ footnote
    empty(doc, 1)
    body_p(doc, "* Value of standard deviation is taken from system precision.", italic=True, sz=SZ_SM)

    # 16 blank lines after LOQ footnote
    empty(doc, 16)

    # ══════════════════════════════════════════════
    # TBL[267]  8. SPECIFICITY header — cols 2772,7855
    # ══════════════════════════════════════════════
    tbl267 = doc.add_table(rows=1, cols=2)
    tbl267.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl267); set_tblW(tbl267, 10627); set_grid(tbl267, [2772,7855])
    write_cell_lines(tbl267.cell(0,0),
        [("8.  SPECIFICITY", True, False),
         (f"       {AI}", True, False)],
        align=WDA.LEFT, fill=F_GREY, valign="center")
    set_cw(tbl267.cell(0,0), 2772)
    write_cell(tbl267.cell(0,1),
        "Specificity was determined by spiking the sample with appropriate levels of excipients. "
        "Results show that the procedure is unaffected by the presence of excipients.",
        align=WDA.JUSTIFY, valign="center")
    set_cw(tbl267.cell(0,1), 7855)
    set_rowh(tbl267.rows[0], 674)

    # 2 blanks after specificity header
    empty(doc, 2)

    # TBL  Mobile phase Interference label — D9D9D9 fill
    tbl269 = doc.add_table(rows=1, cols=1)
    tbl269.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(tbl269); set_tblW(tbl269, 10627); set_grid(tbl269, [10627])
    set_rowh(tbl269.rows[0], 458)
    write_cell(tbl269.cell(0,0), "Mobile phase Interference", fill=F_DATA, align=WDA.LEFT)
    empty(doc, 1)

    # Combined 4-row specificity table — cols 4291,1062,5078,196
    sp_tbl = add_tbl(doc, [4291,1062,5078,196], [
        [D("Mobile Phase"), D(""), D("Peak Area"), D("")],
        [D("Blank"),        D(""), D("--------"), D("")],
        [H("Acceptance Criteria"), H(""), H("Result"), H("")],
        [DL("Method should be specific and effect of mobile phase should be negligible."),
         D(""), D(""), D("")],
    ])
    set_rowh(sp_tbl.rows[2], 346); set_rowh(sp_tbl.rows[3], 1061)
    # Merge cols: 0+1 and 2+3 for all rows, remove extra cells
    for ri in range(4):
        set_gridspan(sp_tbl.cell(ri,0), 2); set_cw(sp_tbl.cell(ri,0), 4291+1062)
        set_gridspan(sp_tbl.cell(ri,2), 2); set_cw(sp_tbl.cell(ri,2), 5078+196)
        row_tcs = sp_tbl.rows[ri]._tr.findall(qn("w:tc"))
        for extra_tc in row_tcs[2:3]:
            sp_tbl.rows[ri]._tr.remove(extra_tc)
    # Row 3 col1 rich text result
    r3c1 = [c for c in sp_tbl.rows[3].cells][1]
    set_cw(r3c1, 5078+196); set_valign(r3c1, "center")
    tc = r3c1._tc
    for p in tc.findall(qn("w:p")): tc.remove(p)
    parts = [("The effect of ",False),("Mobile Phase/diluent",True),
             (" is negligible and did not show significant effect at the elution of ",False),
             (f"{AI}",True),(". It confirms that there is no interference of ",False),
             ("Mobile Phase/diluent",True),(f" at appropriate retention time of {AI}",False),
             (", thus the procedure is unaffected by the presence of solvents.",False)]
    p_xml = (f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="both"/>'
             f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/></w:pPr>')
    for txt, bold in parts:
        b = "<w:b/><w:bCs/>" if bold else ""
        p_xml += (f'<w:r><w:rPr>{b}<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr>'
                  f'<w:t xml:space="preserve">{esc(txt)}</w:t></w:r>')
    p_xml += '</w:p>'
    tc.append(parse_xml(p_xml))

    # blank + space + spaces + 13 blanks before conclusion
    empty(doc, 1); body_p(doc, " "); body_p(doc, "   "); empty(doc, 13)

    # ══════════════════════════════════════════════
    # CONCLUSION — TBL[277] Observation, TBL[279] Impact,
    #              TBL[281] Corrective, TBL[283] Conclusion+Statement
    # ══════════════════════════════════════════════
    def con_tbl(cols, heights, label, text):
        t = doc.add_table(rows=2, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_borders(t); set_tblW(t, 10627); set_grid(t, cols)
        set_rowh(t.rows[0], heights[0]); set_rowh(t.rows[1], heights[1])
        write_cell(t.cell(0,0), label, bold=True, fill=F_GREY, align=WDA.LEFT)
        set_cw(t.cell(0,0), cols[0])
        write_cell(t.cell(0,1), ""); set_cw(t.cell(0,1), cols[1])
        # Row2 spans both cols
        set_gridspan(t.cell(1,0), 2); set_cw(t.cell(1,0), 10627)
        write_cell(t.cell(1,0), text, align=WDA.JUSTIFY)
        extra = t.rows[1]._tr.findall(qn("w:tc"))
        if len(extra) > 1: t.rows[1]._tr.remove(extra[1])

    # TBL[277]
    con_tbl([3618,7009],[116,1466],"Observation/Deviation: ",
        f"No deviation was observed during the validation of method of analysis for {AI}.")
    # P[278]
    empty(doc, 1)

    # TBL[279]
    con_tbl([3528,7099],[116,1466],"Impact of Deviation: ","Not Applicable")
    # P[280]
    empty(doc, 1)

    # TBL[281]
    con_tbl([3528,7099],[116,539],"Corrective Action: ","Not Applicable")
    # P[282]
    empty(doc, 1)

    # TBL[283]  Conclusion + Statement — cols 4135,6334  rows 4
    ct = doc.add_table(rows=4, cols=2)
    ct.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(ct); set_tblW(ct, 10627); set_grid(ct, [4135,6492])
    set_rowh(ct.rows[0],116); set_rowh(ct.rows[1],1196)
    set_rowh(ct.rows[2],305); set_rowh(ct.rows[3],2672)
    # Row0: label | blank
    write_cell(ct.cell(0,0),"Conclusion: ",bold=True,fill=F_GREY,align=WDA.LEFT)
    set_cw(ct.cell(0,0),4135)
    write_cell(ct.cell(0,1),""); set_cw(ct.cell(0,1),6334)
    # Row1: spans 2
    set_gridspan(ct.cell(1,0),2); set_cw(ct.cell(1,0),10627)
    write_cell(ct.cell(1,0),
        f"A comprehensive study for the Validation of the test method is performed that is being "
        f"used for the determination of {AI} Starting Material.", align=WDA.JUSTIFY)
    ex = ct.rows[1]._tr.findall(qn("w:tc"))
    if len(ex)>1: ct.rows[1]._tr.remove(ex[1])
    # Row2: STATEMENT label | blank
    write_cell(ct.cell(2,0),"STATEMENT OF SUITABILITY",bold=True,fill=F_GREY,align=WDA.LEFT)
    set_cw(ct.cell(2,0),4135)
    write_cell(ct.cell(2,1),""); set_cw(ct.cell(2,1),6492)
    # Row3: spans 2
    set_gridspan(ct.cell(3,0),2); set_cw(ct.cell(3,0),10627)
    write_cell(ct.cell(3,0),
        f"All the results are analyzed by the application of statistical technique i.e. Mean, "
        f"Standard Deviation, Relative Standard Deviation and observed that all the results are "
        f"within limits, moreover the calculated values of Coefficient Correlation, Standard "
        f"Deviation & RSD are within limits. All the results of the tests carried out for the "
        f"validation of the method are in complete agreement with the required limits and criteria. "
        f"Method being used for the determination of {AI} produces consistent, reproducible and "
        f"reliable results therefore it is suitable for its intended purpose i.e. quantification "
        f"of {AI}.", align=WDA.JUSTIFY)
    ex = ct.rows[3]._tr.findall(qn("w:tc"))
    if len(ex)>1: ct.rows[3]._tr.remove(ex[1])

    # P[284]  blank
    empty(doc, 1)

    # sectPr must be last
    if sectPr is not None:
        body.remove(sectPr)
        body.append(sectPr)

    doc.save(output)
    print(f"Saved: {output}")


# ═══════════════════════════════════════════════════
if __name__ == "__main__":
    BASE     = os.path.dirname(os.path.abspath(__file__))
    excel    = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, "AMV_report.xlsx")
    output   = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, "Method_Validation_Report.docx")
    template = os.path.join(BASE, "shaigan_template.docx")
    data = read_excel(excel)
    print(f"Product: {data['AI']} | MVL: {data['MVL']}")
    build(data, template, output)
