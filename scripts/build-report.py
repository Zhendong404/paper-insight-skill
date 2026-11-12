#!/usr/bin/env python3
"""Build one single-page paper insight slide from a report JSON.

DSH-native port of the reference ``scripts/build-report.mjs`` builder. It copies
the read-only reference deck (assets/reference.pptx) and edits the existing
shapes in place with python-pptx, so the original masters, layouts, theme,
group nesting, fills and line geometry survive.

Usage:
    python3 build-report.py --input report.json --output out/draft.pptx [--skill-dir DIR]

The output is the editable deck plus ``<output-stem>.layout.json`` (shape
inventory and content budgets). Render a preview separately, e.g. with the
office-render skill, and compare it with assets/preview.png before delivery.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

RED = "C00000"
BLACK = "1D1D1A"
DARK = "151515"
YAH = "Microsoft YaHei"
YAH_UI = "Microsoft YaHei UI"

# Source shape ids in the reference slide (see references/style-spec.json).
TITLE = 2
INSIGHT = 4          # inside group 9
PROBLEM = 13
METHOD = 15
RESULT_MAIN = 17
RESULT_SECONDARY = 22
KEY_TECH = 24
SOURCE = 26

FIGURE_SLOTS = {
    "context_figure": 14,
    "method_figure": 16,
    "evaluation_figure": 19,
    "result_figure": 21,
}

TEXT_BUDGETS = {
    "title": (16, 30),
    "insight": (45, 65),
    "implication": (45, 65),
    "problem_total": (100, 150),
    "method_total": (130, 180),
    "result_main": (35, 60),
    "result_secondary": (35, 60),
    "key_technologies_total": (90, 165),
}


# --------------------------------------------------------------------------- #
# XML helpers
# --------------------------------------------------------------------------- #
def _sub(parent, tag, **attrs):
    el = etree.SubElement(parent, qn(tag))
    for key, value in attrs.items():
        if value is not None:
            el.set(key, str(value))
    return el


def _run(p, text, *, size, font, color, bold=False, italic=False):
    r = _sub(p, "a:r")
    rPr = _sub(r, "a:rPr", lang="zh-CN", altLang="en-US",
               sz=int(round(size * 100)), b="1" if bold else "0",
               i="1" if italic else None, dirty="0")
    fill = _sub(rPr, "a:solidFill")
    _sub(fill, "a:srgbClr", val=color)
    _sub(rPr, "a:latin", typeface=font)
    _sub(rPr, "a:ea", typeface=font)
    _sub(r, "a:t").text = text
    return r


def _para(txBody, *, runs, bullet=None, marL=None, indent=None,
          alignment=None, line_spacing=150000):
    p = _sub(txBody, "a:p")
    pPr = _sub(p, "a:pPr", marL=marL, indent=indent, algn=alignment)
    if line_spacing:
        spc = _sub(pPr, "a:lnSpc")
        _sub(spc, "a:spcPct", val=line_spacing)
    if bullet:
        char, typeface = bullet
        _sub(pPr, "a:buFont", typeface=typeface)
        _sub(pPr, "a:buChar", char=char)
    for spec in runs:
        _run(p, **spec)
    return p


def replace_text(shape, paragraphs):
    """Replace every paragraph of a shape's text frame, keeping its bodyPr."""
    txBody = shape.text_frame._txBody
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)
    for spec in paragraphs:
        _para(txBody, **spec)


ARROW = ("\u00d8", "Wingdings")   # source template's arrow glyph
DOT = ("\u2022", "Arial")


def _heading(text, color=RED):
    return {"runs": [dict(text=text, size=9, font=YAH, color=color, bold=True)],
            "alignment": "l"}


def _bullet_item(text, *, color, bold, marker, large=False):
    return {"runs": [dict(text=text, size=9, font=YAH, color=color, bold=bold)],
            "bullet": marker,
            "marL": 285750 if large else 171450,
            "indent": -285750 if large else -171450}


# --------------------------------------------------------------------------- #
# Shape lookup across groups
# --------------------------------------------------------------------------- #
def _iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from _iter_shapes(shape.shapes)


def find_shape(slide, shape_id):
    for shape in _iter_shapes(slide.shapes):
        if shape.shape_id == shape_id:
            return shape
    raise SystemExit(f"reference slide is missing source shape {shape_id}")


def delete_shape(shape):
    el = shape._element
    el.getparent().remove(el)


# --------------------------------------------------------------------------- #
# Pictures
# --------------------------------------------------------------------------- #
def _image_px(path):
    from PIL import Image
    with Image.open(path) as im:
        return im.size


def replace_picture(slide, pic, path, alt_text, box_in=None):
    """Swap a picture's image bytes while keeping its source shape identity.

    ``box_in`` optionally overrides the frame (absolute inches ``[x, y, w, h]``)
    so a figure can grow into unused space inside its column.
    """
    image_part, rId = slide.part.get_or_add_image_part(str(path))
    blip = pic._element.find(".//" + qn("a:blip"))
    if blip is None:
        raise SystemExit(f"picture {pic.shape_id} has no blip")
    old_rId = blip.get(qn("r:embed"))
    blip.set(qn("r:embed"), rId)
    if old_rId and old_rId != rId:
        slide.part.drop_rel(old_rId)
    pic._element.nvPicPr.cNvPr.set("descr", alt_text)
    pic.crop_left = pic.crop_top = pic.crop_right = pic.crop_bottom = 0.0

    if box_in is not None:
        values = [float(v) for v in box_in]
        if len(values) != 4 or any(v < 0 for v in values) or values[2] <= 0 or values[3] <= 0:
            raise SystemExit("figures.*.box_in: expected [x, y, w, h] in inches")
        pic.left, pic.top, pic.width, pic.height = (int(round(v * 914400)) for v in values)

    # "contain": keep the frame centre but never distort the figure.
    img_w, img_h = _image_px(path)
    if img_w <= 0 or img_h <= 0:
        raise SystemExit(f"{path}: unreadable image size")
    frame_l, frame_t = pic.left, pic.top
    frame_w, frame_h = pic.width, pic.height
    if frame_w <= 0 or frame_h <= 0:
        return
    img_ratio = img_w / img_h
    frame_ratio = frame_w / frame_h
    if img_ratio > frame_ratio:      # image wider than frame
        new_w = frame_w
        new_h = int(round(frame_w / img_ratio))
    else:
        new_h = frame_h
        new_w = int(round(frame_h * img_ratio))
    pic.left = frame_l + (frame_w - new_w) // 2
    pic.top = frame_t + (frame_h - new_h) // 2
    pic.width = new_w
    pic.height = new_h



def _abs_box_in(shape):
    """Page-absolute box in inches, composing group child transforms."""
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    el = shape._element
    while el is not None:
        parent = el.getparent()
        if parent is None or parent.tag != qn("p:grpSp"):
            break
        xfrm = parent.find(qn("p:grpSpPr") + "/" + qn("a:xfrm"))
        if xfrm is None:
            break
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        ch_off, ch_ext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
        ox, oy = int(off.get("x")), int(off.get("y"))
        ex, ey = int(ext.get("cx")), int(ext.get("cy"))
        cx = int(ch_off.get("x")) if ch_off is not None else 0
        cy = int(ch_off.get("y")) if ch_off is not None else 0
        cw = int(ch_ext.get("cx")) if ch_ext is not None else ex
        chh = int(ch_ext.get("cy")) if ch_ext is not None else ey
        sx = ex / cw if cw else 1.0
        sy = ey / chh if chh else 1.0
        left = ox + (left - cx) * sx
        top = oy + (top - cy) * sy
        width = width * sx
        height = height * sy
        el = parent
    return [round(v / 914400, 4) for v in (left, top, width, height)]

# --------------------------------------------------------------------------- #
# Report validation
# --------------------------------------------------------------------------- #
def _line(value, key):
    if not isinstance(value, str) or not value.strip():
        raise SystemExit(f"{key}: expected a non-empty string")
    if "\r" in value or "\n" in value:
        raise SystemExit(f"{key}: expected a single line")
    return value.strip()


def _items(report, key, maximum):
    value = report.get(key)
    if not isinstance(value, list) or not 1 <= len(value) <= maximum:
        raise SystemExit(f"{key}: expected 1 to {maximum} items")
    return [_line(s, f"{key}[{i}]") for i, s in enumerate(value)]


# --------------------------------------------------------------------------- #
# Build
# --------------------------------------------------------------------------- #
def build(report_path: Path, out_path: Path, skill_dir: Path) -> dict:
    manifest_path = skill_dir / "artifact-template.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    reference = (skill_dir / manifest["reference"]).resolve()
    out_path = out_path.resolve()
    if out_path == reference:
        raise SystemExit("the reference deck is read-only; choose another output path")

    report = json.loads(report_path.read_text(encoding="utf-8"))
    base = report_path.resolve().parent

    prs = Presentation(str(reference))
    if len(prs.slides) != 1:
        raise SystemExit("expected the single-slide paper insight reference")
    slide = prs.slides[0]

    title = _line(report.get("title"), "title")
    insight = _line(report.get("insight"), "insight")
    implication = _line(report.get("implication"), "implication")
    insight = _strip_prefix(insight, "洞察")
    implication = _strip_prefix(implication, "启示|启发")
    method_title = _strip_prefix(_line(report.get("method_title"), "method_title"), "②")
    result_main = _line(report.get("result_main"), "result_main")
    problems = _items(report, "problem", 3)
    methods = _items(report, "method", 5)
    key_tech = _items(report, "key_technologies", 3)
    source = _line(report.get("source"), "source")
    notes = report.get("notes")
    if not isinstance(notes, str) or not notes.strip():
        raise SystemExit("notes: provide the paper evidence and caveats")

    # Title: 17 pt bold red, inherits the layout font.
    replace_text(find_shape(slide, TITLE), [{
        "runs": [dict(text=title, size=17, font=YAH, color=RED, bold=True)],
        "alignment": "l",
    }])

    # Insight / implication band (shape lives inside the reference group).
    replace_text(find_shape(slide, INSIGHT), [
        {"runs": [dict(text=f"洞察 | {insight}", size=12, font=YAH_UI,
                       color=BLACK, bold=True)],
         "bullet": DOT, "marL": 285750, "indent": -285750},
        {"runs": [dict(text=f"启示 | {implication}", size=12, font=YAH_UI,
                       color=RED, bold=True)],
         "bullet": DOT, "marL": 285750, "indent": -285750},
    ])

    # ① Problem
    replace_text(find_shape(slide, PROBLEM), [
        _heading("①问题描述"),
        *[_bullet_item(text, color=RED if i == 0 else BLACK, bold=True,
                       marker=ARROW) for i, text in enumerate(problems)],
    ])

    # ② Method / mechanism
    method_paras = [_heading("②" + method_title)]
    for i, text in enumerate(methods):
        method_paras.append(_bullet_item(
            text, color=RED if i == 1 else BLACK, bold=(i == 0),
            marker=ARROW if i == 0 else DOT))
    replace_text(find_shape(slide, METHOD), method_paras)

    # ③ Technical results
    replace_text(find_shape(slide, RESULT_MAIN), [
        _heading("③技术收益"),
        _bullet_item(result_main, color=BLACK, bold=True, marker=ARROW),
    ])
    secondary = report.get("result_secondary")
    if secondary is None:
        delete_shape(find_shape(slide, RESULT_SECONDARY))
    else:
        replace_text(find_shape(slide, RESULT_SECONDARY), [
            _bullet_item(_line(secondary, "result_secondary"), color=RED,
                         bold=True, marker=ARROW),
        ])

    # Key technologies: 12 pt bold, bullets, roomy indent.
    replace_text(find_shape(slide, KEY_TECH), [{
        "runs": [dict(text=text, size=12, font=YAH_UI, color=DARK, bold=True)],
        "bullet": DOT, "marL": 285750, "indent": -285750,
    } for text in key_tech])

    # Source line: 10 pt italic.
    replace_text(find_shape(slide, SOURCE), [{
        "runs": [dict(text=source, size=10, font=YAH_UI, color=DARK,
                      italic=True)],
    }])

    # Figures: every slot must be an explicit replacement or null.
    figures = report.get("figures")
    if not isinstance(figures, dict):
        raise SystemExit("figures: provide all four slots, null to delete")
    figure_notes = []
    for role, shape_id in FIGURE_SLOTS.items():
        if role not in figures:
            raise SystemExit(f"figures.{role}: provide a replacement or explicit null")
        value = figures[role]
        shape = find_shape(slide, shape_id)
        if value is None:
            delete_shape(shape)
            continue
        if not isinstance(value, dict):
            raise SystemExit(f"figures.{role}: expected an object or null")
        asset = (base / _line(value.get("path"), f"figures.{role}.path")).resolve()
        if asset.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            raise SystemExit(f"figures.{role}: use PNG or JPEG")
        if not asset.is_file():
            raise SystemExit(f"figures.{role}: missing asset {asset}")
        alt = _line(value.get("alt"), f"figures.{role}.alt")
        fig_source = _line(value.get("source"), f"figures.{role}.source")
        replace_picture(slide, shape, asset, alt, box_in=value.get("box_in"))
        figure_notes.append(f"{alt}：{fig_source}")

    # Evidence list in the slide notes.
    slide.notes_slide.notes_text_frame.text = "\n".join([notes.strip(), *figure_notes])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))

    # Re-open and verify the deliverable.
    reopened = Presentation(str(out_path))
    shapes = list(_iter_shapes(reopened.slides[0].shapes))
    texts = {s.shape_id: s.text_frame.text for s in shapes if s.has_text_frame}
    if len(reopened.slides) != 1:
        raise SystemExit("saved deck does not contain exactly one slide")
    if texts.get(TITLE) != title:
        raise SystemExit("saved deck lost the title text")
    if "AIPerf" in "".join(texts.values()) and "AIPerf" not in \
            json.dumps(report, ensure_ascii=False):
        raise SystemExit("saved deck still carries reference-only AIPerf text")

    layout = {
        "schemaVersion": 1,
        "output": str(out_path),
        "reference": str(reference),
        "slide_size_in": [round(reopened.slide_width / 914400, 6),
                          round(reopened.slide_height / 914400, 6)],
        "canvas_emu": [reopened.slide_width, reopened.slide_height],
        "slides": 1,
        "note_char_count": len(slide.notes_slide.notes_text_frame.text),
        "shapes": [{
            "id": s.shape_id,
            "name": s.name,
            "role": next((k for k, v in {
                "title": TITLE, "insight_and_implication": INSIGHT,
                "problem": PROBLEM, "method": METHOD, "result_main": RESULT_MAIN,
                "result_secondary": RESULT_SECONDARY, "key_technologies": KEY_TECH,
                "source": SOURCE, **FIGURE_SLOTS}.items() if v == s.shape_id), None),
            "box_in": _abs_box_in(s),
            "text": texts.get(s.shape_id, ""),
            "char_count": len(texts.get(s.shape_id, "")),
        } for s in shapes],
        "budgets": {key: {"min": lo, "max": hi,
                          "measured": _measure(key, report)}
                    for key, (lo, hi) in TEXT_BUDGETS.items()},
    }
    stem = out_path.with_suffix("")
    stem.with_suffix(".layout.json").write_text(
        json.dumps(layout, ensure_ascii=False, indent=1), encoding="utf-8")
    return layout


def _strip_prefix(text, prefixes):
    for prefix in prefixes.split("|"):
        for sep in (" | ", "：", ":", " "):
            head = prefix + sep
            if text.startswith(head):
                return text[len(head):].strip()
            bare = prefix
            if text.startswith(bare) and len(text) > len(bare):
                rest = text[len(bare):].lstrip(" |｜：:")
                if rest and rest != text:
                    return rest.strip()
    return text


def _measure(key, report):
    def total(name):
        value = report.get(name)
        return sum(len(s) for s in value) if isinstance(value, list) else 0
    measured = {
        "title": len(report.get("title") or ""),
        "insight": len(report.get("insight") or ""),
        "implication": len(report.get("implication") or ""),
        "problem_total": total("problem"),
        "method_total": total("method"),
        "result_main": len(report.get("result_main") or ""),
        "result_secondary": len(report.get("result_secondary") or ""),
        "key_technologies_total": total("key_technologies"),
    }[key]
    return measured


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True, help="report JSON path")
    ap.add_argument("--output", required=True, help="output .pptx path")
    ap.add_argument("--skill-dir", default=str(Path(__file__).resolve().parent.parent),
                    help="skill root holding artifact-template.json (default: parent of scripts/)")
    args = ap.parse_args(argv)
    report_path = Path(args.input)
    out_path = Path(args.output)
    if out_path.suffix.lower() != ".pptx":
        raise SystemExit("--output must end with .pptx")
    layout = build(report_path, out_path, Path(args.skill_dir))
    over = [f"{k}={v['measured']}>{v['max']}" for k, v in layout["budgets"].items()
            if v["measured"] > v["max"]]
    under = [f"{k}={v['measured']}<{v['min']}" for k, v in layout["budgets"].items()
             if v["measured"] < v["min"]]
    print(out_path)
    if over:
        print("warning: above the reference content budget (check fit on render): "
              + ", ".join(over), file=sys.stderr)
    if under:
        print("note: below the reference content budget (fine if evidence is thin): "
              + ", ".join(under), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
