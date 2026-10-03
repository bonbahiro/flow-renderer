# -*- coding: utf-8 -*-
"""
flow_renderer_excel.py  —  PowerPoint 版（flow_renderer.py）で描いた業務フロー図を、
                           編集可能な Excel 図形として .xlsx に載せ替える

使い方（flow_renderer.py の render と組み合わせる）:
    warnings = render(spec, "tmp.pptx")          # flow_renderer.py
    pptx_to_xlsx("tmp.pptx", "out.xlsx")         # このファイル

特徴:
- PowerPoint の図形・コネクタを、そのまま Excel の図形・コネクタとして書き出す。
  Excel 上でも図形を動かすと矢印が追従する。見た目は PowerPoint 版と同じ。
- シートは目盛線なし・A3横・1ページに収めて印刷する設定。
- 標準ライブラリ＋lxml（python-pptx の依存として導入済み）だけで動く。
"""
import zipfile
from copy import deepcopy
from xml.sax.saxutils import escape
from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
PAPER = {"A3": 8, "A4": 9}


def _to_xdr(el):
    """p: 名前空間の図形要素を xdr: に付け替える（p:nvPr は Excel に無いので削除）"""
    el = deepcopy(el)
    for nv in el.iter(f"{{{NS_P}}}nvPr"):
        nv.getparent().remove(nv)
    for e in el.iter():
        if isinstance(e.tag, str) and e.tag.startswith(f"{{{NS_P}}}"):
            e.tag = f"{{{NS_XDR}}}" + e.tag.split("}", 1)[1]
    return el


def _unrotate(cxn):
    """回転付きコネクタを、回転なしの同じ経路のコネクタに置き換える（Excel 系ビューアでの崩れ防止）
       ・カギ線1回（bentConnector2）：向きを逆にして横→縦の経路にし、矢印を始点側に付け替える
       ・カギ線2回（bentConnector3）：bentConnector4（adj1=0）で 縦→横→縦 の経路にする"""
    xfrm = cxn.find(f".//{{{NS_A}}}xfrm")
    rot = (int(xfrm.get("rot", "0")) // 60000) % 360
    if rot == 0:
        return
    off, ext = xfrm.find(f"{{{NS_A}}}off"), xfrm.find(f"{{{NS_A}}}ext")
    x, y, w, h = int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))
    fh, fv = xfrm.get("flipH") == "1", xfrm.get("flipV") == "1"
    cx, cy = x + w / 2, y + h / 2

    def turn(px, py):
        dx, dy = px - w / 2, py - h / 2
        if rot == 90: dx, dy = -dy, dx
        elif rot == 180: dx, dy = -dx, -dy
        elif rot == 270: dx, dy = dy, -dx
        return cx + dx, cy + dy

    sx, sy = (w if fh else 0), (h if fv else 0)
    p0, p1 = turn(sx, sy), turn(w - sx, h - sy)
    geom = cxn.find(f".//{{{NS_A}}}prstGeom")
    prst = geom.get("prst")
    av = geom.find(f"{{{NS_A}}}avLst")
    adj = 50000
    if av is not None:
        for gd in av:
            if gd.get("name") == "adj1": adj = int(gd.get("fmla").split()[1])
    else:
        av = etree.SubElement(geom, f"{{{NS_A}}}avLst")
    for gd in list(av): av.remove(gd)

    if prst == "bentConnector2":
        s, e = p1, p0                                   # 向きを反転
        ln = cxn.find(f".//{{{NS_A}}}ln")
        tail = ln.find(f"{{{NS_A}}}tailEnd") if ln is not None else None
        if tail is not None:
            tail.tag = f"{{{NS_A}}}headEnd"
        nv = cxn.find(f".//{{{NS_XDR}}}cNvCxnSpPr")
        st, en = nv.find(f"{{{NS_A}}}stCxn"), nv.find(f"{{{NS_A}}}endCxn")
        if st is not None and en is not None:
            (si, sx_), (ei, ex_) = (st.get("id"), st.get("idx")), (en.get("id"), en.get("idx"))
            st.set("id", ei); st.set("idx", ex_); en.set("id", si); en.set("idx", sx_)
    else:                                               # bentConnector3 → bentConnector4
        s, e = p0, p1
        geom.set("prst", "bentConnector4")
        g1 = etree.SubElement(av, f"{{{NS_A}}}gd"); g1.set("name", "adj1"); g1.set("fmla", "val 0")
        g2 = etree.SubElement(av, f"{{{NS_A}}}gd"); g2.set("name", "adj2"); g2.set("fmla", f"val {adj}")

    for k in ("rot", "flipH", "flipV"):
        if k in xfrm.attrib: del xfrm.attrib[k]
    if e[0] < s[0]: xfrm.set("flipH", "1")
    if e[1] < s[1]: xfrm.set("flipV", "1")
    off.set("x", str(int(min(s[0], e[0])))); off.set("y", str(int(min(s[1], e[1]))))
    ext.set("cx", str(int(abs(e[0] - s[0])))); ext.set("cy", str(int(abs(e[1] - s[1]))))


def _to_freeform(cxn):
    """（回転除去済みの）コネクタを、同じ経路の折れ線図形に置き換える。
       どのビューアでも同じ見た目になるが、図形を動かしても線は追従しない。"""
    xfrm = cxn.find(f".//{{{NS_A}}}xfrm")
    off, ext = xfrm.find(f"{{{NS_A}}}off"), xfrm.find(f"{{{NS_A}}}ext")
    x, y, w, h = int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))
    fh, fv = xfrm.get("flipH") == "1", xfrm.get("flipV") == "1"
    geom = cxn.find(f".//{{{NS_A}}}prstGeom")
    prst = geom.get("prst")
    adj = {gd.get("name"): int(gd.get("fmla").split()[1]) / 100000 for gd in geom.iter(f"{{{NS_A}}}gd")}
    if prst == "bentConnector2":
        pts = [(0, 0), (w, 0), (w, h)]
    elif prst == "bentConnector3":
        xm = w * adj.get("adj1", 0.5); pts = [(0, 0), (xm, 0), (xm, h), (w, h)]
    elif prst == "bentConnector4":
        x1, y2 = w * adj.get("adj1", 0.5), h * adj.get("adj2", 0.5)
        pts = [(0, 0), (x1, 0), (x1, y2), (w, y2), (w, h)]
    else:
        pts = [(0, 0), (w, h)]
    pts = [(x + (w - px if fh else px), y + (h - py if fv else py)) for px, py in pts]
    clean = []
    for p in pts:
        if not clean or (abs(p[0] - clean[-1][0]) > 1 or abs(p[1] - clean[-1][1]) > 1):
            clean.append(p)
    pts = clean if len(clean) > 1 else pts[:2]
    bx, by = min(p[0] for p in pts), min(p[1] for p in pts)
    bw, bh = max(1, int(max(p[0] for p in pts) - bx)), max(1, int(max(p[1] for p in pts) - by))

    sp = etree.Element(f"{{{NS_XDR}}}sp")
    nv = etree.SubElement(sp, f"{{{NS_XDR}}}nvSpPr")
    old = cxn.find(f".//{{{NS_XDR}}}cNvPr")
    nv.append(deepcopy(old))
    etree.SubElement(nv, f"{{{NS_XDR}}}cNvSpPr")
    spPr = etree.SubElement(sp, f"{{{NS_XDR}}}spPr")
    xf = etree.SubElement(spPr, f"{{{NS_A}}}xfrm")
    etree.SubElement(xf, f"{{{NS_A}}}off", x=str(int(bx)), y=str(int(by)))
    etree.SubElement(xf, f"{{{NS_A}}}ext", cx=str(bw), cy=str(bh))
    cg = etree.SubElement(spPr, f"{{{NS_A}}}custGeom")
    for t in ("avLst", "gdLst", "ahLst", "cxnLst"):
        etree.SubElement(cg, f"{{{NS_A}}}{t}")
    etree.SubElement(cg, f"{{{NS_A}}}rect", l="0", t="0", r="r", b="b")
    pl = etree.SubElement(cg, f"{{{NS_A}}}pathLst")
    path = etree.SubElement(pl, f"{{{NS_A}}}path", w=str(bw), h=str(bh), fill="none")
    for i, (px, py) in enumerate(pts):
        step = etree.SubElement(path, f"{{{NS_A}}}{'moveTo' if i == 0 else 'lnTo'}")
        etree.SubElement(step, f"{{{NS_A}}}pt", x=str(int(px - bx)), y=str(int(py - by)))
    etree.SubElement(spPr, f"{{{NS_A}}}noFill")
    ln = cxn.find(f".//{{{NS_A}}}ln")
    if ln is not None:
        spPr.append(deepcopy(ln))
    return sp


def pptx_to_xlsx(pptx_path, xlsx_path, sheet_name="業務フロー", paper="A3", connectors=True):
    """connectors=True ：矢印は Excel のコネクタ（図形を動かすと追従）
       connectors=False：矢印は折れ線図形（どのビューアでも同じ見た目。追従はしない）"""
    with zipfile.ZipFile(pptx_path) as z:
        slide = etree.fromstring(z.read("ppt/slides/slide1.xml"))
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find(f"{{{NS_P}}}sldSz")
    landscape = int(sz.get("cx")) >= int(sz.get("cy"))

    wsdr = etree.Element(f"{{{NS_XDR}}}wsDr", nsmap={"xdr": NS_XDR, "a": NS_A, "r": NS_R})
    tree = slide.find(f".//{{{NS_P}}}spTree")
    for child in tree:
        tag = etree.QName(child).localname
        if tag not in ("sp", "cxnSp"):
            continue
        shape = _to_xdr(child)
        if tag == "cxnSp":
            _unrotate(shape)
            if not connectors:
                shape = _to_freeform(shape)
        xfrm = shape.find(f".//{{{NS_A}}}xfrm")
        off, ext = xfrm.find(f"{{{NS_A}}}off"), xfrm.find(f"{{{NS_A}}}ext")
        anchor = etree.SubElement(wsdr, f"{{{NS_XDR}}}absoluteAnchor")
        etree.SubElement(anchor, f"{{{NS_XDR}}}pos", x=off.get("x"), y=off.get("y"))
        etree.SubElement(anchor, f"{{{NS_XDR}}}ext", cx=ext.get("cx"), cy=ext.get("cy"))
        anchor.append(shape)
        etree.SubElement(anchor, f"{{{NS_XDR}}}clientData")
    drawing = etree.tostring(wsdr, xml_declaration=True, encoding="UTF-8", standalone=True)

    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'xmlns:r="{NS_R}">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        '<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="70" zoomScaleNormal="70"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<sheetData/>'
        '<pageMargins left="0.3" right="0.3" top="0.3" bottom="0.3" header="0" footer="0"/>'
        f'<pageSetup paperSize="{PAPER.get(paper, 8)}" orientation="{"landscape" if landscape else "portrait"}" '
        'fitToWidth="1" fitToHeight="1"/>'
        '<drawing r:id="rId1"/>'
        '</worksheet>')
    files = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'xmlns:r="{NS_R}"><sheets><sheet name="{escape(sheet_name[:31])}" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>',
        "xl/styles.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Meiryo UI"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>',
        "xl/worksheets/sheet1.xml": sheet,
        "xl/worksheets/_rels/sheet1.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            '</Relationships>',
        "xl/drawings/drawing1.xml": drawing,
    }
    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return xlsx_path


if __name__ == "__main__":
    import sys
    print(pptx_to_xlsx(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "flow.xlsx",
                       connectors="--lines" not in sys.argv))
