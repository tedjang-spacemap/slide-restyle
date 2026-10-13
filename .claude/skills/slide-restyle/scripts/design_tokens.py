# /// script
# requires-python = ">=3.10"
# dependencies = ["PyYAML>=6"]
# ///
"""Load and validate the `yaml design-tokens` block of a DESIGN.md.

    uv run design_tokens.py DESIGN.md        # validate, print a short report, exit 1 on errors

As a module: `load(path) -> dict` (raises ValueError on a missing/duplicate block).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import yaml

FENCE = re.compile(r"^```yaml design-tokens\s*\n(.*?)^```\s*$", re.S | re.M)
HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
REQUIRED = ("meta", "colors", "fonts", "type", "chrome", "components", "layouts", "overflow")


def load(path: str | Path) -> dict:
    text = Path(path).read_text(encoding="utf-8")
    blocks = FENCE.findall(text)
    if len(blocks) != 1:
        raise ValueError(f"{path}: expected exactly one ```yaml design-tokens block, found {len(blocks)}")
    return yaml.safe_load(blocks[0])


def _boxes(node, where):
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "box":
                yield f"{where}.box", v
            else:
                yield from _boxes(v, f"{where}.{k}")
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from _boxes(v, f"{where}[{i}]")


def validate(tokens: dict, root: Path) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    for key in REQUIRED:
        if key not in tokens:
            errors.append(f"missing top-level key: {key}")
    if errors:
        return errors, warnings

    W, H = tokens["meta"]["canonical_slide"]["w"], tokens["meta"]["canonical_slide"]["h"]
    colors, fonts, types = tokens["colors"], tokens["fonts"], tokens["type"]

    for name, c in colors.items():
        for k in ("hex", "fill"):
            if k in c and not HEX.match(str(c[k])):
                errors.append(f"colors.{name}.{k}: not a 6-digit hex: {c[k]!r}")
        if "source" not in c:
            errors.append(f"colors.{name}: missing source")

    for name, t in types.items():
        if t.get("font") not in fonts:
            errors.append(f"type.{name}.font: unknown font role {t.get('font')!r}")
        col = t.get("color")
        if col and col not in colors and not str(col).startswith("status.") and not HEX.match(str(col)):
            errors.append(f"type.{name}.color: unknown color {col!r}")
        if "source" not in t:
            errors.append(f"type.{name}: missing source")
        if t.get("source") == "estimated" and "basis" not in t:
            errors.append(f"type.{name}: estimated without basis")

    for section in ("chrome", "components", "layouts"):
        for where, box in _boxes(tokens[section], section):
            if not (isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) for v in box)):
                errors.append(f"{where}: box must be [x, y, w, h], got {box!r}")
                continue
            x, y, w, h = box
            if x < -0.01 or y < -0.01 or x + w > W + 0.01 or y + h > H + 0.01:
                errors.append(f"{where}: {box} outside {W}×{H}")

    for name, lay in tokens["layouts"].items():
        src = lay.get("source")
        if src not in ("measured", "estimated"):
            errors.append(f"layouts.{name}: source must be measured|estimated")
        if src == "measured" and "evidence" not in lay:
            errors.append(f"layouts.{name}: measured without evidence")
        if src == "estimated" and "basis" not in lay:
            errors.append(f"layouts.{name}: estimated without basis")
        ch = lay.get("chrome")
        if ch not in (None, "none") and ch not in tokens["chrome"]:
            errors.append(f"layouts.{name}.chrome: unknown chrome {ch!r}")
        for rname, region in (lay.get("regions") or {}).items():
            st = region.get("style") if isinstance(region, dict) else None
            if isinstance(st, str) and st not in types:
                errors.append(f"layouts.{name}.regions.{rname}.style: unknown type style {st!r}")

    text = yaml.safe_dump(tokens, allow_unicode=True)
    for asset in sorted(set(re.findall(r"design/assets/[\w.]+", text))):
        if not (root / asset).exists():
            errors.append(f"asset not found: {asset}")
    base = tokens["meta"].get("base_template")
    if base and not (root / base).exists():
        errors.append(f"base_template not found: {base}")
    return errors, warnings


def main():
    path = Path(sys.argv[1] if len(sys.argv) > 1 else "DESIGN.md")
    try:
        tokens = load(path)
    except (ValueError, yaml.YAMLError) as e:
        sys.exit(f"ERROR {e}")
    errors, warnings = validate(tokens, path.resolve().parent)
    lays = tokens.get("layouts", {})
    measured = [k for k, v in lays.items() if v.get("source") == "measured"]
    estimated = [k for k, v in lays.items() if v.get("source") == "estimated"]
    print(f"{path}: {len(tokens.get('colors', {}))} colors, {len(tokens.get('type', {}))} type styles, "
          f"{len(lays)} layouts (measured {len(measured)}: {', '.join(measured)}; estimated {len(estimated)}: {', '.join(estimated)})")
    for w in warnings:
        print(f"WARN  {w}")
    for e in errors:
        print(f"ERROR {e}")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
