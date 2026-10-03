# -*- coding: utf-8 -*-
"""
flow_renderer_drawio.py  —  業務フロー図（Layer-1〜3）を JSON 仕様から draw.io ファイルに描画する

PowerPoint 版（flow_renderer.py）と同じ「フロー仕様JSON」をそのまま受け取り、
同じ色・図形・文字サイズで .drawio ファイルを書き出す。

使い方:
    warnings = render(spec_dict_or_json_path, "out.drawio")
    for w in warnings: print(w)

特徴:
- システム領域は「コンテナ」。draw.io 上で領域を動かすと中のノードも一緒に動く。
- 矢印は図形に接続済み（source/target 指定）。図形を動かすと矢印が追従する。
- 矢印の交差は線をまたぐ形（jumpStyle=arc）で描く。
- 描画後に検査（ID不整合・重なり・領域はみ出し・孤立ノード・矢印がノードを横切る）を行い、
  警告リストを返す。

依存: Python 標準ライブラリのみ
"""
import json
from copy import deepcopy
from xml.sax.saxutils import escape, quoteattr

# ------------------------------------------------------------------ スタイル定義（PowerPoint 版と共通）
FONT = "Meiryo UI"
PX = 100                    # 1インチ = 100px（A3横 = 1654×1169）
PT = PX / 72                # 1pt → px

SYSTEM_COLORS = {
    "sales": "FEF9C3", "bi": "DCFCE7", "erp": "E0F2FE", "scheduler": "FFEDD5", "mes": "EDE9FE",
    "plm": "86EFAC", "portal": "FCE7F3", "factory": "BEF264", "manual": "F1F5F9",
}
SYSTEM_LABELS = {
    "sales": "販売・案件管理", "bi": "分析・BI", "erp": "基幹ERP", "scheduler": "生産計画",
    "mes": "製造実行", "plm": "仕様書・BOM", "portal": "取引先ポータル",
    "factory": "工場業務", "manual": "手作業・Excel",
}
INK = "334155"; SUB = "64748B"; RED = "DC2626"; GRAY = "94A3B8"; DARK = "1E293B"
EDGE_STYLE = {"flow": ("475569", 1.0, False), "if": ("2563EB", 1.0, True), "goods": ("A16207", 2.0, False)}
EDGE_NAMES = {"flow": "業務・情報の流れ", "if": "システム連携", "goods": "物の流れ"}
NODE_DEFAULT_SIZE = {
    "process": (4, 1.3), "data": (3.4, 2), "doc": (3.4, 1.8),
    "decision": (3.4, 2), "terminal": (3, 1.2), "note": (4, 1.5),
}
SHAPE_STYLE = {
    "process":  "rounded=1;arcSize=18;",
    "data":     "shape=cylinder3;boundedLbl=1;backgroundOutline=1;size=7;",
    "doc":      "shape=document;boundedLbl=1;size=0.2;",
    "decision": "shape=rhombus;perimeter=rhombusPerimeter;",
    "terminal": "rounded=1;arcSize=50;",
    "note":     "rounded=0;",
}
PAGES = {"A3": (16.54, 11.69), "A4": (11.69, 8.27), "16:9": (13.333, 7.5)}
SIDE_XY = {"T": (0.5, 0), "B": (0.5, 1), "L": (0, 0.5), "R": (1, 0.5)}


def _fs(pt): return round(pt * PT)                       # pt → draw.io fontSize
def _lw(pt): return round(pt * PT, 1)                    # pt → draw.io strokeWidth


def _font(size_pt, color, bold=False):
    return f"fontFamily={FONT};fontSize={_fs(size_pt)};fontColor=#{color};" + ("fontStyle=1;" if bold else "")


def _label(text):
    return "<br>".join(escape(l) for l in str(text).split("\n"))


class _Grid:
    def __init__(self, spec):
        pw, ph = PAGES.get(spec.get("page", "A3"), PAGES["A3"])
        g = spec.get("grid", {})
        self.cols, self.rows = g.get("cols", 48), g.get("rows", 30)
        self.left, self.top = 0.4 * PX, 1.0 * PX
        self.cw = (pw - 0.8) * PX / self.cols
        self.rh = (ph - 1.0 - 0.4) * PX / self.rows
        self.pw, self.ph = round(pw * PX), round(ph * PX)

    def box(self, col, row, w, h):
        return (round(self.left + col * self.cw, 1), round(self.top + row * self.rh, 1),
                round(w * self.cw, 1), round(h * self.rh, 1))


# ------------------------------------------------------------------ 経路計算（検査用・PowerPoint 版と同じ考え方）
def _port(b, s):
    x, y, w, h = b
    return {"T": (x + w / 2, y), "B": (x + w / 2, y + h), "L": (x, y + h / 2), "R": (x + w, y + h / 2)}[s]


def _auto_sides(a, b):
    ax, ay, aw, ah = a; bx, by, bw, bh = b
    if bx >= ax + aw or ax >= bx + bw:
        return ("R", "L") if bx + bw / 2 > ax + aw / 2 else ("L", "R")
    return ("B", "T") if by + bh / 2 > ay + ah / 2 else ("T", "B")


def _route(p0, s0, p1, s1, off):
    (x0, y0), (x1, y1) = p0, p1
    h0, h1 = s0 in "LR", s1 in "LR"
    if s0 == s1:
        if h0:
            mx = max(x0, x1) + off if s0 == "R" else min(x0, x1) - off
            return [p0, (mx, y0), (mx, y1), p1]
        my = max(y0, y1) + off if s0 == "B" else min(y0, y1) - off
        return [p0, (x0, my), (x1, my), p1]
    if abs(x0 - x1) < 0.5 or abs(y0 - y1) < 0.5:
        return [p0, p1]
    if h0 and h1:
        mx = (x0 + x1) / 2; return [p0, (mx, y0), (mx, y1), p1]
    if not h0 and not h1:
        my = (y0 + y1) / 2; return [p0, (x0, my), (x1, my), p1]
    if h0:
        return [p0, (x1, y0), p1]
    return [p0, (x0, y1), p1]


def _overlap(a, b):
    ax, ay, aw, ah = a; bx, by, bw, bh = b
    return ax < bx + bw and bx < ax + aw and ay < by + bh and by < ay + ah


def _contains(o, i):
    ox, oy, ow, oh = o; ix, iy, iw, ih = i
    return ox <= ix + 0.5 and oy <= iy + 0.5 and ix + iw <= ox + ow + 0.5 and iy + ih <= oy + oh + 0.5


def _seg_hits(p, q, b, shrink=3):
    x, y, w, h = b
    x0, y0, x1, y1 = x + shrink, y + shrink, x + w - shrink, y + h - shrink
    (px, py), (qx, qy) = p, q
    if abs(px - qx) < 0.5:
        return x0 < px < x1 and min(py, qy) < y1 and max(py, qy) > y0
    if abs(py - qy) < 0.5:
        return y0 < py < y1 and min(px, qx) < x1 and max(px, qx) > x0
    return False


# ------------------------------------------------------------------ 本体
def render(spec, out_path):
    if isinstance(spec, str):
        with open(spec, encoding="utf-8") as f:
            spec = json.load(f)
    spec = deepcopy(spec)
    warn, cells = [], []
    G = _Grid(spec)
    meta = spec.get("meta", {})

    def cell(cid, value, style, x, y, w, h, parent="1"):
        cells.append(f'<mxCell id={quoteattr(cid)} value={quoteattr(value)} style={quoteattr(style)} '
                     f'vertex="1" parent={quoteattr(parent)}><mxGeometry x="{x}" y="{y}" width="{w}" '
                     f'height="{h}" as="geometry"/></mxCell>')

    text_base = "text;html=1;strokeColor=none;fillColor=none;whiteSpace=wrap;"

    # タイトル・情報行
    cell("title", "■ " + escape(meta.get("title", "業務フロー")),
         text_base + "align=left;verticalAlign=middle;" + _font(20, DARK, True),
         0.4 * PX, 0.22 * PX, G.pw - 7.8 * PX, 0.5 * PX)
    info = "　｜　".join(x for x in [
        meta.get("layer", ""), meta.get("state", ""), meta.get("id", ""),
        ("v" + str(meta["version"])) if meta.get("version") else "",
        meta.get("author", ""), meta.get("date", "")] if x)
    cell("info", escape(info), text_base + "align=left;verticalAlign=middle;" + _font(9, SUB),
         0.4 * PX, 0.66 * PX, G.pw - 7.8 * PX, 0.28 * PX)

    # 組織枠
    for i, o in enumerate(spec.get("orgs", [])):
        x, y, w, h = G.box(o["col"], o.get("row", 0), o["w"], o.get("h", G.rows - o.get("row", 0)))
        frame = "rounded=0;html=1;whiteSpace=wrap;fillColor=none;strokeColor=#94A3B8;dashed=1;dashPattern=4 3;"
        if "row" in o and w > h * 2:      # 横長＝L3のスイムレーン：名前は左端に縦書き
            cell(f"org_{i}", "", frame, x, y, w, h)
            cell(f"orgl_{i}", escape(o["name"]), "text;html=1;whiteSpace=wrap;strokeColor=none;fillColor=#F1F5F9;"
                 "horizontal=0;align=center;verticalAlign=middle;" + _font(12, DARK, True), x, y, 0.32 * PX, h)
        else:
            cell(f"org_{i}", escape(o["name"]), frame + "verticalAlign=top;align=center;spacingTop=4;"
                 + _font(13, DARK), x, y, w, h)

    # システム領域（コンテナ）
    sys_rects = []          # (cid, sid, box)
    for i, sd in enumerate(spec.get("systems", [])):
        box = G.box(sd["col"], sd["row"], sd["w"], sd["h"])
        cat = sd.get("cat", "manual")
        if cat not in SYSTEM_COLORS:
            warn.append(f"[区分] システム {sd.get('id')} の区分 '{cat}' は未定義 → manual で描画"); cat = "manual"
        cid = f"sys_{i}"
        cell(cid, escape(sd.get("name", "")),
             f"rounded=1;absoluteArcSize=1;arcSize=14;html=1;whiteSpace=wrap;fillColor=#{SYSTEM_COLORS[cat]};"
             "strokeColor=none;container=1;collapsible=0;recursiveResize=0;verticalAlign=top;align=left;"
             "spacingLeft=10;spacingTop=4;" + _font(11, DARK, True), *box)
        sys_rects.append((cid, sd["id"], box))

    # ノード（座標は絶対で計算し、所属領域の子にする場合は相対へ変換）
    nodes, boxes, node_cells = {}, {}, []
    for n in spec.get("nodes", []):
        t = n.get("type", "process")
        if t not in SHAPE_STYLE:
            warn.append(f"[種別] ノード {n['id']} の type '{t}' は未定義 → process"); t = "process"
        if n["id"] in nodes:
            warn.append(f"[ID] ノードID {n['id']} が重複"); continue
        dw, dh = NODE_DEFAULT_SIZE[t]
        box = G.box(n["col"], n["row"], n.get("w", dw), n.get("h", dh))
        status = n.get("status", "")
        stroke = SUB if t in ("data", "doc") else INK
        fcolor, lw, dash, fill = INK, 0.75, False, "FFFFFF"
        if status in ("new", "changed"): stroke, fcolor, lw = RED, RED, 1.5
        if status == "removed": stroke, fcolor, dash = GRAY, GRAY, True
        if t == "terminal": fill, fcolor, stroke = DARK, "FFFFFF", DARK
        if t == "note": fill = "FFFBEB"
        bold = bool(n.get("emphasis")) or t == "terminal"
        size = n.get("size", 13 if n.get("emphasis") else 10)
        style = (SHAPE_STYLE[t] + f"html=1;whiteSpace=wrap;fillColor=#{fill};strokeColor=#{stroke};"
                 f"strokeWidth={_lw(lw)};" + ("dashed=1;" if dash else "") + _font(size, fcolor, bold))
        parent, (x, y, w, h) = "1", box
        sid = n.get("sys")
        if sid:
            cands = [r for r in sys_rects if r[1] == sid]
            if not cands:
                warn.append(f"[ID] ノード {n['id']} の sys '{sid}' が systems に無い")
            else:
                inside = [r for r in cands if _contains(r[2], box)]
                if not inside:
                    warn.append(f"[配置] ノード {n['id']} がシステム領域 {sid} からはみ出している")
                else:
                    parent = inside[0][0]
                    x, y = round(box[0] - inside[0][2][0], 1), round(box[1] - inside[0][2][1], 1)
        node_cells.append((n["id"], _label(n["label"]), style, x, y, w, h, parent))
        nodes[n["id"]] = n; boxes[n["id"]] = box

    ids = list(boxes)
    for i, a in enumerate(ids):
        for b in ids[i + 1:]:
            if _overlap(boxes[a], boxes[b]):
                warn.append(f"[重なり] ノード {a} と {b} が重なっている")

    # 矢印（ノードより先に置いて背面にする）
    used = set()
    for k, e in enumerate(spec.get("edges", [])):
        a, b = e.get("from"), e.get("to")
        if a not in boxes or b not in boxes:
            warn.append(f"[ID] 矢印 {a}→{b} の端点が存在しない"); continue
        used.update([a, b])
        s0, s1 = _auto_sides(boxes[a], boxes[b])
        s0, s1 = e.get("from_side", s0), e.get("to_side", s1)
        path = _route(_port(boxes[a], s0), s0, _port(boxes[b], s1), s1, 0.5 * G.cw)
        for nid, nb in boxes.items():
            if nid in (a, b): continue
            if any(_seg_hits(path[i], path[i + 1], nb) for i in range(len(path) - 1)):
                warn.append(f"[横切り] 矢印 {a}→{b} がノード {nid} の上を通過（座標か from_side/to_side を調整）")
        color, width, dash = EDGE_STYLE.get(e.get("kind", "flow"), EDGE_STYLE["flow"])
        if e.get("status") in ("new", "changed"): color = RED
        ex, ey = SIDE_XY[s0]; nx, ny = SIDE_XY[s1]
        style = ("edgeStyle=orthogonalEdgeStyle;rounded=1;arcSize=8;orthogonalLoop=1;jettySize=auto;html=1;"
                 f"exitX={ex};exitY={ey};exitDx=0;exitDy=0;exitPerimeter=0;"
                 f"entryX={nx};entryY={ny};entryDx=0;entryDy=0;entryPerimeter=0;"
                 f"strokeColor=#{color};strokeWidth={_lw(width)};endArrow=block;endFill=1;endSize=5;"
                 "jumpStyle=arc;jumpSize=8;labelBackgroundColor=#FFFFFF;" + ("dashed=1;" if dash else "")
                 + _font(8.5, "475569"))
        cells.append(f'<mxCell id={quoteattr(f"E{k:03d}_{a}_{b}")} value={quoteattr(_label(e.get("label", "")))} '
                     f'style={quoteattr(style)} edge="1" parent="1" source={quoteattr(a)} target={quoteattr(b)}>'
                     f'<mxGeometry relative="1" as="geometry"/></mxCell>')

    for nid, val, style, x, y, w, h, parent in node_cells:
        cell(nid, val, style, x, y, w, h, parent)

    for nid in ids:
        nd = nodes[nid]
        if nid not in used and nd.get("type") != "note" and not nd.get("standalone"):
            warn.append(f"[孤立] ノード {nid} に矢印が1本も無い")

    _legend(cell, G, spec, cells)

    xml = ('<mxfile host="flow_renderer_drawio">'
           f'<diagram id="flow" name={quoteattr(meta.get("id") or meta.get("title", "flow"))}>'
           f'<mxGraphModel dx="1600" dy="1100" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" '
           f'arrows="1" fold="1" page="1" pageScale="1" pageWidth="{G.pw}" pageHeight="{G.ph}" math="0" shadow="0">'
           '<root><mxCell id="0"/><mxCell id="1" parent="0"/>' + "".join(cells) +
           '</root></mxGraphModel></diagram></mxfile>')
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(xml)
    return warn


def _legend(cell, G, spec, raw):
    if spec.get("legend") is False: return
    cats = []
    for sd in spec.get("systems", []):
        c = sd.get("cat", "manual")
        if c not in cats and c in SYSTEM_COLORS: cats.append(c)
    kinds = sorted({e.get("kind", "flow") for e in spec.get("edges", [])}, key=list(EDGE_STYLE).index)
    has_new = any(n.get("status") in ("new", "changed") for n in spec.get("nodes", []))
    items = [("sys", c) for c in cats] + [("edge", k) for k in kinds] + ([("new", None)] if has_new else [])
    per_row, iw, ih = 5, 1.4 * PX, 0.24 * PX
    x0, y0 = G.pw - 0.4 * PX - per_row * iw, 0.2 * PX
    for i, (kind, v) in enumerate(items):
        x = round(x0 + (i % per_row) * iw, 1); y = round(y0 + (i // per_row) * ih, 1)
        if kind == "sys":
            cell(f"lg_{i}", "", f"rounded=1;html=1;fillColor=#{SYSTEM_COLORS[v]};strokeColor=none;",
                 x, y + 4, 28, 16); label = SYSTEM_LABELS[v]
        elif kind == "edge":
            col, w, d = EDGE_STYLE[v]
            st = (f"html=1;endArrow=block;endFill=1;endSize=5;strokeColor=#{col};strokeWidth={_lw(w)};"
                  + ("dashed=1;" if d else ""))
            raw.append(
                f'<mxCell id="lg_{i}" value="" style={quoteattr(st)} edge="1" parent="1">'
                f'<mxGeometry relative="1" as="geometry"><mxPoint x="{x}" y="{y + 12}" as="sourcePoint"/>'
                f'<mxPoint x="{x + 28}" y="{y + 12}" as="targetPoint"/></mxGeometry></mxCell>')
            label = EDGE_NAMES[v]
        else:
            cell(f"lg_{i}", "", f"rounded=1;html=1;fillColor=#FFFFFF;strokeColor=#{RED};strokeWidth={_lw(1.5)};",
                 x, y + 4, 28, 16); label = "新規・変更"
        cell(f"lgt_{i}", escape(label), "text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;"
             + _font(8, SUB), x + 32, y, iw - 34, ih)


if __name__ == "__main__":
    import sys
    ws = render(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "flow.drawio")
    print("\n".join(ws) if ws else "警告なし")
