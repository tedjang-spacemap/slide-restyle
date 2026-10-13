# /// script
# requires-python = ">=3.10"
# dependencies = ["PyYAML>=6"]
# ///
"""Check a plan.json against its content.json and DESIGN.md (deterministic).

    uv run check_plan.py work/x/plan.json [--design DESIGN.md] [--md work/x/plan.md]

Checks:
  verbatim   every {"t", "src"} text in the plan is a substring of that source shape's text
             (whitespace-collapsed) — the plan may split and move text, never rewrite it
  coverage   every source shape with text is used by a text ref, shown by an image_full slide
             (`covers`), or listed in the slide's `dropped` with a reason
  layouts    layout / layout_variant exist in DESIGN.md; capacity limits (rows, cards, flow items)
  files      image_full `image` files exist (relative to the plan's folder)
Prints the classification table; --md also writes it as Markdown for user review.
Schema: reference/plan-schema.md. Exit 1 on errors.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_tokens import load as load_tokens, provenance  # noqa: E402

WS = re.compile(r"\s+")


def norm(s: str) -> str:
    return WS.sub(" ", s or "").strip()


def text_refs(node, where=""):
    """Yield (where, ref) for every {"t": ..., "src": ...} object in a plan node."""
    if isinstance(node, dict):
        if "t" in node and "src" in node:
            yield where, node
            return
        for k, v in node.items():
            if k in ("dropped", "covers"):
                continue
            yield from text_refs(v, f"{where}.{k}" if where else k)
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from text_refs(v, f"{where}[{i}]")


def count(slide, key):
    r = slide.get("regions", {})
    return len(r.get(key, []) or [])


def capacity_errors(sl, layout, tokens):
    errs = []
    lay = tokens["layouts"][layout]
    var = sl.get("layout_variant")
    cap = dict(lay.get("capacity") or {})
    if var:
        cap.update((lay.get("variants", {}).get(var) or {}).get("capacity") or {})
    checks = {"io_steps": [("rows", "rows")], "card_grid": [("cards", "cards")],
              "cover": [("flow", "flow_items")], "table": [("rows", "rows")], "toc": [("items", "items")]}
    for region, cap_key in checks.get(layout, []):
        n = count(sl, region)
        if cap_key in cap and n > cap[cap_key]:
            errs.append(f"{region}: {n} > capacity {cap[cap_key]}")
    return errs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("plan", type=Path)
    ap.add_argument("--design", type=Path, default=Path("DESIGN.md"))
    ap.add_argument("--md", type=Path, help="write the classification table as Markdown")
    a = ap.parse_args()

    plan = json.loads(a.plan.read_text(encoding="utf-8"))
    base = a.plan.resolve().parent
    content = json.loads((base / plan.get("content", "content.json")).read_text(encoding="utf-8"))
    tokens = load_tokens(a.design)
    variant = plan.get("variant") or tokens["meta"]["default_chrome"]
    shapes = {s["id"]: (sl["n"], s) for sl in content["slides"] for s in sl["shapes"]}
    errors, warnings = [], []

    used = {}       # src id -> list of normalized text pieces used
    covered = set()
    dropped = {}
    for sl in plan["slides"]:
        tag = f"out {sl['out']} (S{sl['src']})"
        layout = sl.get("layout")
        if layout not in tokens["layouts"]:
            errors.append(f"{tag}: unknown layout {layout!r}")
            continue
        var = sl.get("layout_variant")
        if var and var not in (tokens["layouts"][layout].get("variants") or {}):
            errors.append(f"{tag}: unknown variant {layout}.{var}")
        errors += [f"{tag}: {e}" for e in capacity_errors(sl, layout, tokens)]
        for where, ref in text_refs({k: v for k, v in sl.items() if k in ("slots", "regions")}):
            sid = ref["src"]
            if sid not in shapes:
                errors.append(f"{tag} {where}: unknown src {sid}")
                continue
            src_text = norm(shapes[sid][1].get("text", {}).get("plain", ""))
            t = norm(ref["t"])
            if t and t not in src_text:
                errors.append(f"{tag} {where}: text not verbatim in {sid}: {ref['t']!r}")
            used.setdefault(sid, []).append(t)
        if layout == "image_full":
            img = sl.get("regions", {}).get("image", {}).get("file")
            if not img or not (base / img).exists():
                errors.append(f"{tag}: image_full file missing: {img!r}")
            cov = sl.get("covers", [])
            if cov == "all":
                cov = [s["id"] for s in content["slides"][sl["src"] - 1]["shapes"]]
            covered.update(cov)
        for d in sl.get("dropped", []):
            if not d.get("why"):
                errors.append(f"{tag}: dropped {d.get('src')} without why")
            dropped[d["src"]] = d.get("why")

    for sid, (n, s) in shapes.items():
        if not s.get("text"):
            continue
        plain = norm(s["text"]["plain"])
        if sid in dropped or sid in covered:
            continue
        if sid not in used:
            errors.append(f"S{n} {sid} {s['name']!r}: text not used, covered or dropped: {plain[:50]!r}")
            continue
        got = sum(len(t) for t in used[sid])
        bare = len(plain.replace(" ", ""))
        if got and sum(len(t.replace(" ", "")) for t in used[sid]) < 0.9 * bare:
            warnings.append(f"S{n} {sid} {s['name']!r}: only part of the text is used ({got}/{len(plain)} chars)")

    rows = ["| 출력 | 원본 | 레이아웃 | collision/선택 변형 | 넘침 | 수작업 | 근거 |", "|---|---|---|---|---|---|---|"]
    for sl in plan["slides"]:
        layout = sl.get("layout", "?")
        lv = f"{layout}.{sl['layout_variant']}" if sl.get("layout_variant") else layout
        prov = provenance(tokens["layouts"].get(layout, {}), variant) if layout in tokens["layouts"] else "?"
        ov = sl.get("overflow", {})
        ov_s = ov.get("risk", "-") + (f" ({', '.join(ov.get('plan', []))})" if ov.get("plan") else "")
        manual = sl.get("manual") or ""
        rows.append(f"| {sl['out']} | S{sl['src']} | `{lv}` | {'실측' if prov == 'measured' else '추정'} | {ov_s} | {manual} | {sl.get('reason', '')} |")
    table = "\n".join(rows)
    print(f"{a.plan}: {len(plan['slides'])} output slides from {content['meta']['slides']} source slides, variant {variant}")
    print(table)
    if a.md:
        a.md.write_text(f"# 분류표 — {Path(plan['source']).name}\n\n변형: `{variant}`\n\n{table}\n", encoding="utf-8")
    for w in warnings:
        print(f"WARN  {w}")
    for e in errors:
        print(f"ERROR {e}")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
