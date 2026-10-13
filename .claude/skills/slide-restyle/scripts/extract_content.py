# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx>=1.0.2", "Pillow>=10", "PyYAML>=6"]
# ///
"""Extract the content of an input deck for layout classification (deterministic; no judgment).

    uv run extract_content.py inputs/x.pptx [--out work/x] [--render]

Writes into --out (default work/<stem>/):
  content.json  every slide's shapes (flattened, EMU boxes), full text with run styles, tables, pictures,
                charts, SmartArt text, notes, plus derived `features` (role candidates, grids, chains)
  media/        pictures, named <md5[:8]>.<ext>
  src_png/      (--render) the input rendered by render.sh — source images for `image_full` fallbacks

Schema: reference/content-schema.md. Nothing here decides a layout; `features` only reports
geometry and style facts that the classifier (Claude, writing plan.json) reads.
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
import statistics
import subprocess
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract_design import NS, Theme, fill_of, image_px, line_of, q, read_text  # noqa: E402

EMU_IN = 914400
SCHEMA = "slide-restyle/content@1"
ALIGN_TOL = int(0.08 * EMU_IN)    # edges within 0.08" count as aligned
TOUCH_TOL = int(0.15 * EMU_IN)    # connector end within 0.15" of a shape edge attaches to it
SIZE_TOL = 0.12                   # same-size test for grid cells (relative)
MONO = ("consolas", "courier", "menlo", "monaco", "sf mono", "d2coding", "jetbrains mono", "source code pro")
PAGE_NO = re.compile(r"^\s*\d{1,3}\s*(/\s*\d{1,3})?\s*$")
# "② 용어집 — 후보 만들기" → marker ②, section 용어집, sub 후보 만들기
SECTION = re.compile(r"^\s*(?P<marker>[①-⑳❶-❿]|\d{1,2}[.)]|[A-Z][.)]|Part\s+\d+|\d{1,2}\s)\s*"
                     r"(?P<section>.+?)(?:\s+[—–\-:|]\s+(?P<sub>.+))?\s*$")


def inch(v):
    return round(v / EMU_IN, 2)


# ---------------------------------------------------------------- shape walk

class Ctx:
    def __init__(self, prs, media_dir: Path):
        self.prs = prs
        self.media_dir = media_dir
        self.prs_default = prs.part._element.find("p:defaultTextStyle", NS)
        self.next_id = 0

    def sid(self, slide_n):
        self.next_id += 1
        return f"s{slide_n}.{self.next_id}"


def mark_code(text):
    """Tag runs set in a monospace face (scripts, paths) — the classifier treats them as code."""
    if not text:
        return text
    for p in text["paragraphs"]:
        for r in p["runs"]:
            if r.get("font") and r["font"].lower().startswith(MONO):
                r["code"] = True
            for k in ("size_raw", "caps", "spc"):
                if r.get(k) is None:
                    r.pop(k, None)
    text["plain"] = "\n".join("".join(r["text"] for r in p["runs"]) for p in text["paragraphs"])
    return text


def cell_text(cell, theme, tctx, ctx):
    class _Shim:  # read_text wants a shape; a table cell has the same text_frame API
        is_placeholder = False
        text_frame = cell.text_frame
        _element = cell._tc
    return mark_code(read_text(_Shim, theme, tctx, ctx.prs_default, 1.0))


def read_table(sh, theme, tctx, ctx):
    tbl = sh.table
    rows = []
    for ri, row in enumerate(tbl.rows):
        cells = []
        for ci, cell in enumerate(row.cells):
            tc = cell._tc
            tcpr = tc.find("a:tcPr", NS)
            sf = tcpr.find("a:solidFill", NS) if tcpr is not None else None
            c = {"text": cell_text(cell, theme, tctx, ctx)}
            if sf is not None:
                c["fill"] = theme.color(sf)
            for attr in ("gridSpan", "rowSpan"):
                if tc.get(attr):
                    c[attr] = int(tc.get(attr))
            if tc.get("hMerge") or tc.get("vMerge"):
                c["merged_into_prev"] = True
            cells.append(c)
        rows.append({"h": row.height, "cells": cells})
    tblpr = tbl._tbl.find("a:tblPr", NS)
    return {"n_rows": len(tbl.rows), "n_cols": len(tbl.columns), "col_w": [c.width for c in tbl.columns],
            "first_row_header": bool(tblpr is not None and tblpr.get("firstRow") in ("1", "true")),
            "rows": rows}


def read_chart(sh):
    out = {"chart_type": None, "title": None, "series": []}
    try:
        ch = sh.chart
        out["chart_type"] = str(ch.chart_type).split(" ")[0].split(".")[-1]
        if ch.has_title and ch.chart_title.has_text_frame:
            out["title"] = ch.chart_title.text_frame.text
        for plot in ch.plots:
            cats = [str(c) for c in plot.categories]
            for s in plot.series:
                out["series"].append({"name": s.name, "categories": cats, "values": list(s.values)})
    except Exception as e:  # charts with external data or unusual types
        out["error"] = str(e)
    return out


def read_smartart(sh, slide):
    """All text of a SmartArt diagram, from its data-model part (order = data model order)."""
    rel = sh._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds")
    if rel is None:
        return {"texts": []}
    rid = rel.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}dm")
    try:
        part = slide.part.related_part(rid)
        el = etree.fromstring(part.blob)
    except Exception:
        return {"texts": []}
    texts = []
    for pt in el.iter("{http://schemas.openxmlformats.org/drawingml/2006/diagram}pt"):
        t = "".join(x.text or "" for x in pt.iter(q("a:t")))
        if t.strip():
            texts.append(t)
    return {"texts": texts}


def arrow_ends(sh):
    ln = sh._element.find(".//a:ln", NS)
    if ln is None:
        return {}
    out = {}
    for end in ("headEnd", "tailEnd"):
        e = ln.find(f"a:{end}", NS)
        if e is not None and e.get("type") not in (None, "none"):
            out[end] = e.get("type")
    return out


def connector_geom(sh, box):
    """Begin/end points. A line's xfrm box runs begin→end unless flipped."""
    x, y, w, h = box
    xfrm = sh._element.find(".//a:xfrm", NS)
    fh = xfrm is not None and xfrm.get("flipH") in ("1", "true")
    fv = xfrm is not None and xfrm.get("flipV") in ("1", "true")
    bx, ex = (x + w, x) if fh else (x, x + w)
    by, ey = (y + h, y) if fv else (y, y + h)
    out = {"begin": [bx, by], "end": [ex, ey], **arrow_ends(sh)}
    for tag, key in (("a:stCxn", "begin_cxn"), ("a:endCxn", "end_cxn")):
        c = sh._element.find(f".//{tag}", NS)
        if c is not None:
            out[key] = int(c.get("id"))
    return out


def walk(shapes, slide, slide_n, ctx: Ctx, theme, tctx, xf=(1.0, 1.0, 0.0, 0.0), group=None, z=None):
    sx, sy, tx, ty = xf
    z = z if z is not None else [0]
    for sh in shapes:
        try:
            x, y, w, h = sh.left, sh.top, sh.width, sh.height
        except Exception:
            x = None
        if x is None:
            continue
        box = [round(x * sx + tx), round(y * sy + ty), round(w * sx), round(h * sy)]
        z[0] += 1
        rec = {"id": ctx.sid(slide_n), "xml_id": sh.shape_id, "name": sh.name, "z": z[0], "box": box}
        if group:
            rec["group"] = group
        rot = getattr(sh, "rotation", 0) or 0
        if rot:
            rec["rot"] = round(rot, 1)
        if sh.is_placeholder:
            rec["ph"] = str(sh.placeholder_format.type).split(" ")[0].split(".")[-1]
            rec["ph_idx"] = sh.placeholder_format.idx
        st = sh.shape_type
        gd = sh._element.find(".//a:graphicData", NS)
        uri = gd.get("uri", "") if gd is not None else ""

        if st == MSO_SHAPE_TYPE.GROUP:
            g = sh._element.find(q("p:grpSpPr")).find("a:xfrm", NS)
            off, ext = g.find("a:off", NS), g.find("a:ext", NS)
            choff, chext = g.find("a:chOff", NS), g.find("a:chExt", NS)
            csx = int(ext.get("cx")) / max(1, int(chext.get("cx")))
            csy = int(ext.get("cy")) / max(1, int(chext.get("cy")))
            gx = int(off.get("x")) - int(choff.get("x")) * csx
            gy = int(off.get("y")) - int(choff.get("y")) * csy
            child_xf = (sx * csx, sy * csy, tx + gx * sx, ty + gy * sy)
            rec["kind"] = "group"
            yield rec
            yield from walk(sh.shapes, slide, slide_n, ctx, theme, tctx, child_xf, rec["id"], z)
            continue

        if "diagram" in uri:
            rec["kind"] = "smartart"
            rec["smartart"] = read_smartart(sh, slide)
        elif "chart" in uri:
            rec["kind"] = "chart"
            rec["chart"] = read_chart(sh)
        elif getattr(sh, "has_table", False) and sh.has_table:
            rec["kind"] = "table"
            rec["table"] = read_table(sh, theme, tctx, ctx)
        elif uri and "ole" in uri.lower():
            rec["kind"] = "ole"
        elif st == MSO_SHAPE_TYPE.PICTURE or (hasattr(sh, "image") and st != MSO_SHAPE_TYPE.PLACEHOLDER):
            try:
                blob, ext = sh.image.blob, sh.image.ext
            except Exception:
                rec["kind"] = "picture"
                rec["picture"] = {"unreadable": True}
                yield rec
                continue
            md5 = hashlib.md5(blob).hexdigest()
            name = f"{md5[:8]}.{ext}"
            ctx.media_dir.mkdir(parents=True, exist_ok=True)
            (ctx.media_dir / name).write_bytes(blob)
            rec["kind"] = "picture"
            rec["picture"] = {"file": f"media/{name}", "md5": md5, "ext": ext, "px": image_px(blob),
                              "crop": [round(v, 4) for v in (sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom)]}
        elif st == MSO_SHAPE_TYPE.LINE or sh._element.tag == q("p:cxnSp"):
            rec["kind"] = "connector"
            rec["line"] = line_of(sh, theme)
            rec["connector"] = connector_geom(sh, box)
        else:
            geom = sh._element.find(".//a:prstGeom", NS)
            rec["geom"] = geom.get("prst") if geom is not None else ("custom" if st == MSO_SHAPE_TYPE.FREEFORM else None)
            rec["fill"], rec["line"] = fill_of(sh, theme), line_of(sh, theme)
            has_text = sh.has_text_frame and sh.text_frame.text.strip()
            if has_text:
                rec["text"] = mark_code(read_text(sh, theme, tctx, ctx.prs_default, 1.0))
            if has_text and not (rec["fill"] or rec["line"]):
                rec["kind"] = "text"           # bare text box
            elif rec["fill"] or rec["line"] or rec["geom"] == "custom":
                rec["kind"] = "shape"          # visible shape, maybe with text
            else:
                rec["kind"] = "empty"
        yield rec


# ---------------------------------------------------------------- features (facts, not decisions)

def text_size(rec):
    """Largest explicit run size in a text record (pt), or None."""
    sizes = [r["size"] for p in rec.get("text", {}).get("paragraphs", []) for r in p["runs"] if r.get("size")]
    return max(sizes) if sizes else None


def split_status(rec):
    """Title whose trailing runs are styled differently (smaller/other color) → main + status text."""
    paras = rec["text"]["paragraphs"]
    if len(paras) != 1 or len(paras[0]["runs"]) < 2:
        return None
    runs = paras[0]["runs"]
    sig = lambda r: (r.get("size"), (r.get("color") or {}).get("hex"))
    k = 1
    while k < len(runs) and sig(runs[k]) == sig(runs[0]):
        k += 1
    if k == len(runs):
        return None
    tail = runs[k:]
    if any(sig(r) == sig(runs[0]) for r in tail):
        return None
    main = "".join(r["text"] for r in runs[:k]).strip()
    status = "".join(r["text"] for r in tail).strip()
    if not main or not status:
        return None
    return {"main": main, "status": status, "status_color": (tail[0].get("color") or {}).get("hex")}


def parse_section(title):
    m = SECTION.match(title or "")
    if not m:
        return None
    out = {"marker": m.group("marker").strip(), "section": m.group("section").strip()}
    if m.group("sub"):
        out["sub"] = m.group("sub").strip()
    return out


def roles(shapes, W, H):
    """Role candidates by name, placeholder type, position and size — each with the reason it matched."""
    texty = [s for s in shapes if s.get("text") and s["kind"] in ("text", "shape")]
    out = {}
    pages = [s for s in texty if PAGE_NO.match(s["text"]["plain"]) and s["box"][1] > H * 0.85]
    if pages:
        out["page_number"] = {"id": pages[0]["id"], "by": "digits in bottom 15%", "text": pages[0]["text"]["plain"].strip()}
    title = next((s for s in texty if s.get("ph") in ("TITLE", "CENTER_TITLE")), None)
    by = "placeholder"
    if title is None:
        title = next((s for s in texty if s["name"].lower() == "title"), None)
        by = "name"
    if title is None:
        top = [s for s in texty if s["box"][1] < H * 0.45 and text_size(s)]
        if top:
            title = max(top, key=lambda s: (text_size(s), -s["box"][1]))
            by = "largest text in top 45%"
    if title:
        t = {"id": title["id"], "by": by, "text": title["text"]["plain"]}
        st = split_status(title)
        if st:
            t["split"] = st
        sec = parse_section(st["main"] if st else title["text"]["plain"])
        if sec:
            t["section"] = sec
        out["title"] = t
        tb = title["box"]
        below = [s for s in texty if s is not title and s["kind"] == "text"
                 and 0 <= s["box"][1] - (tb[1] + tb[3]) < 0.3 * EMU_IN and abs(s["box"][0] - tb[0]) < ALIGN_TOL]
        lead = next((s for s in texty if s.get("ph") == "SUBTITLE"), None) or \
            next((s for s in texty if s["name"].lower() in ("lead", "subtitle")), None) or \
            (below[0] if below else None)
        if lead:
            out["lead"] = {"id": lead["id"], "text": lead["text"]["plain"],
                           "by": "placeholder/name" if lead not in below else "text box right under title"}
    return out


def cluster(vals, tol):
    """Group sorted values into clusters whose members are within tol of the first member."""
    groups = []
    for v in sorted(vals):
        if groups and v - groups[-1][0] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [g[0] for g in groups]


def find_grids(shapes):
    """Rectangular arrangements of same-size boxes (≥2×2, or a single row/column of ≥3).

    Works per column-width class: boxes in one column share left edge and width; rows share top edge.
    Returns cells as shape ids in row-major order plus header text boxes sitting right above columns.
    """
    boxes = [s for s in shapes if s["kind"] in ("shape", "text", "picture") and s["box"][2] > 0.3 * EMU_IN
             and s["box"][3] > 0.2 * EMU_IN and not s.get("group_role")]
    used, grids = set(), []
    tops = cluster([s["box"][1] for s in boxes], ALIGN_TOL)
    lefts = cluster([s["box"][0] for s in boxes], ALIGN_TOL)
    near = lambda a, b: abs(a - b) <= ALIGN_TOL
    close = lambda a, b: abs(a - b) <= SIZE_TOL * max(a, b)
    # cells stacked in one column must look alike: same kind, same width, same height
    same_cell = lambda a, b: a["kind"] == b["kind"] and close(a["box"][2], b["box"][2]) and close(a["box"][3], b["box"][3])
    at = {}
    for s in boxes:
        r = next(i for i, t in enumerate(tops) if near(s["box"][1], t) or s["box"][1] >= t and (i + 1 == len(tops) or s["box"][1] < tops[i + 1]))
        c = next(i for i, l in enumerate(lefts) if near(s["box"][0], l) or s["box"][0] >= l and (i + 1 == len(lefts) or s["box"][0] < lefts[i + 1]))
        if near(s["box"][1], tops[r]) and near(s["box"][0], lefts[c]):
            at.setdefault((r, c), []).append(s)
    # a row band = cells sharing a top; keep boxes of similar height within the band
    for r in range(len(tops)):
        row = sorted([(c, v[0]) for (rr, c), v in at.items() if rr == r and v[0]["id"] not in used], key=lambda t: t[0])
        if len(row) < 2:
            continue
        cols = [c for c, _ in row]
        # extend downward while the same columns repeat with matching widths
        rows = [r]
        for r2 in range(r + 1, len(tops)):
            ok = all((r2, c) in at and at[(r2, c)][0]["id"] not in used and same_cell(at[(r2, c)][0], at[(r, c)][0])
                     for c in cols)
            if ok:
                rows.append(r2)
            elif len(rows) > 1:
                break
        cells = [[at[(rr, c)][0] for c in cols] for rr in rows]
        heights_ok = all(abs(cl["box"][3] - row_[0]["box"][3]) <= SIZE_TOL * row_[0]["box"][3] for row_ in cells for cl in row_)
        if len(rows) == 1:
            ws = [cl["box"][2] for cl in cells[0]]
            if len(cols) < 3 or max(ws) - min(ws) > SIZE_TOL * statistics.mean(ws) or not heights_ok:
                continue
        flat = [cl for row_ in cells for cl in row_]
        if not all(cl.get("text") for cl in flat):
            continue
        for cl in flat:
            used.add(cl["id"])
        first = cells[0]
        headers = []
        for cl in first:
            h = next((s for s in shapes if s["kind"] == "text" and s["id"] not in used and near(s["box"][0], cl["box"][0])
                      and 0 <= cl["box"][1] - (s["box"][1] + s["box"][3]) < 0.25 * EMU_IN), None)
            headers.append(h["id"] if h else None)
        grids.append({
            "rows": len(rows), "cols": len(cols),
            "col_x": [first[i]["box"][0] for i in range(len(cols))], "col_w": [first[i]["box"][2] for i in range(len(cols))],
            "row_y": [cells[i][0]["box"][1] for i in range(len(rows))], "row_h": [cells[i][0]["box"][3] for i in range(len(rows))],
            "cells": [[cl["id"] for cl in row_] for row_ in cells],
            "headers": headers if any(headers) else None,
        })
    return grids


def edge_hit(pt, box):
    x, y = pt
    bx, by, bw, bh = box
    inside_x = bx - TOUCH_TOL <= x <= bx + bw + TOUCH_TOL
    inside_y = by - TOUCH_TOL <= y <= by + bh + TOUCH_TOL
    if not (inside_x and inside_y):
        return None
    return min(abs(x - bx), abs(x - bx - bw), abs(y - by), abs(y - by - bh))


def link_connectors(shapes):
    """Attach each connector end to the nearest shape edge (or its stCxn/endCxn target)."""
    targets = [s for s in shapes if s["kind"] in ("shape", "text", "picture", "table", "chart", "smartart")]
    by_xml = {s["xml_id"]: s for s in targets}
    links = []
    for c in (s for s in shapes if s["kind"] == "connector"):
        g = c["connector"]
        ends = {}
        for end in ("begin", "end"):
            cx = g.get(f"{end}_cxn")
            if cx in by_xml:
                ends[end] = by_xml[cx]["id"]
                continue
            best = min(((edge_hit(g[end], t["box"]), t) for t in targets if edge_hit(g[end], t["box"]) is not None),
                       key=lambda t: t[0], default=None)
            ends[end] = best[1]["id"] if best else None
        g["from"], g["to"] = ends["begin"], ends["end"]
        if g.get("headEnd") and not g.get("tailEnd"):
            g["from"], g["to"] = g["to"], g["from"]   # arrowhead at begin → reverse direction
        links.append((g["from"], g["to"]))
    return links


def chains(shapes, links):
    """Maximal simple paths A→B→C… through shapes joined end-to-end by connectors (length ≥ 3)."""
    succ, pred = collections.defaultdict(list), collections.defaultdict(list)
    for a, b in links:
        if a and b and a != b:
            succ[a].append(b)
            pred[b].append(a)
    out = []
    for start in list(succ):
        if len(pred[start]) == 1 and len(succ[pred[start][0]]) == 1:
            continue  # not a chain head
        path, cur = [start], start
        while len(succ[cur]) == 1 and len(pred[succ[cur][0]]) == 1 and succ[cur][0] not in path:
            cur = succ[cur][0]
            path.append(cur)
        if len(path) >= 3:
            out.append(path)
    return out


def overlaps(shapes):
    vis = [s for s in shapes if s["kind"] in ("shape", "text", "picture")]
    n = 0
    for i, a in enumerate(vis):
        ax, ay, aw, ah = a["box"]
        for b in vis[i + 1:]:
            bx, by, bw, bh = b["box"]
            if ax < bx + bw and bx < ax + aw and ay < by + bh and by < ay + ah:
                n += 1
    return n


def style_classes(shapes):
    """Visible shapes grouped by (geom, fill, line color, dash) — e.g. color-coded step kinds."""
    cls = collections.defaultdict(list)
    for s in shapes:
        if s["kind"] != "shape":
            continue
        key = (s.get("geom"), (s.get("fill") or {}).get("hex"),
               (s.get("line") or {}).get("hex"), (s.get("line") or {}).get("dash"))
        cls[key].append(s["id"])
    return [{"geom": k[0], "fill": k[1], "line": k[2], "dash": k[3], "ids": v}
            for k, v in sorted(cls.items(), key=lambda kv: -len(kv[1]))]


def slide_features(shapes, W, H):
    kinds = collections.Counter(s["kind"] for s in shapes)
    links = link_connectors(shapes)
    rl = roles(shapes, W, H)
    grids = find_grids(shapes)
    in_grid = {i for g in grids for row in g["cells"] for i in row}
    role_ids = {v["id"] for v in rl.values()}
    in_grid |= {h for g in grids for h in (g["headers"] or []) if h}
    loose = [s["id"] for s in shapes if s["kind"] in ("shape", "text") and s["id"] not in in_grid and s["id"] not in role_ids]
    return {
        "counts": dict(kinds),
        "chars": sum(len(s["text"]["plain"]) for s in shapes if s.get("text")),
        "roles": rl,
        "grids": grids,
        "chains": chains(shapes, links),
        "connectors_linked": sum(1 for a, b in links if a and b),
        "connectors_dangling": sum(1 for a, b in links if not (a and b)),
        "overlaps": overlaps(shapes),
        "loose_shapes": loose,
        "style_classes": style_classes(shapes),
    }


def repeated_texts(slides):
    """Identical texts on ≥2 slides (legends, footers) — not page numbers."""
    seen = collections.defaultdict(list)
    for sl in slides:
        for s in sl["shapes"]:
            if s.get("text") and not PAGE_NO.match(s["text"]["plain"]):
                seen[s["text"]["plain"].strip()].append(f"{sl['n']}:{s['id']}")
    return [{"text": t, "at": at} for t, at in seen.items() if len({a.split(':')[0] for a in at}) >= 2]


# ---------------------------------------------------------------- main

def extract(path: Path, out: Path) -> dict:
    prs = Presentation(str(path))
    W, H = prs.slide_width, prs.slide_height
    ctx = Ctx(prs, out / "media")
    themes = {id(m): Theme(m) for m in prs.slide_masters}
    slides = []
    for n, s in enumerate(prs.slides, 1):
        ctx.next_id = 0
        m = s.slide_layout.slide_master
        shapes = list(walk(s.shapes, s, n, ctx, themes[id(m)], {"master_el": m._element}))
        notes = None
        if s.has_notes_slide:
            notes = s.notes_slide.notes_text_frame.text.strip() or None
        slides.append({"n": n, "layout_name": s.slide_layout.name, "hidden": s._element.get("show") == "0",
                       "shapes": shapes, "notes": notes, "features": slide_features(shapes, W, H)})
    core = prs.core_properties
    return {
        "schema": SCHEMA,
        "source": str(path),
        "slide_size": {"w": W, "h": H, "w_in": inch(W), "h_in": inch(H)},
        "units": {"box": "EMU [x, y, w, h]", "emu_per_in": EMU_IN, "size": "pt"},
        "meta": {"title": core.title or None, "author": core.author or None, "slides": len(slides)},
        "slides": slides,
        "deck_features": {"repeated_texts": repeated_texts(slides)},
    }


def digest(doc) -> str:
    """One line per slide, for the terminal."""
    lines = []
    for sl in doc["slides"]:
        f = sl["features"]
        c = f["counts"]
        t = f["roles"].get("title", {}).get("text", "").split("\n")[0][:40]
        g = " ".join(f"grid{x['rows']}x{x['cols']}" for x in f["grids"])
        ch = " ".join(f"chain{len(x)}" for x in f["chains"])
        kinds = " ".join(f"{k}={v}" for k, v in sorted(c.items()))
        lines.append(f"S{sl['n']:<2} {t:<40} | {kinds} | {g} {ch} | overlaps={f['overlaps']} loose={len(f['loose_shapes'])}")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck", type=Path)
    ap.add_argument("--out", type=Path, help="default: work/<stem>")
    ap.add_argument("--render", action="store_true", help="also render the input to <out>/src_png via render.sh")
    a = ap.parse_args()
    out = a.out or Path("work") / a.deck.stem
    out.mkdir(parents=True, exist_ok=True)
    doc = extract(a.deck, out)
    (out / "content.json").write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{out / 'content.json'}: {doc['meta']['slides']} slides")
    print(digest(doc))
    if a.render:
        subprocess.run([str(Path(__file__).with_name("render.sh")), str(a.deck), str(out / "src_png")], check=True)


if __name__ == "__main__":
    main()
