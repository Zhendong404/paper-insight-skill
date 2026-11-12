#!/usr/bin/env python3
"""Fit the four figure slots to the free space left inside each column.

The reference deck reserves fixed picture frames; when a paper's figure has a
different aspect ratio, ``add``-style placement keeps the frame and letterboxes
the figure, leaving it small. This script keeps the picture's shape identity
(and thus its z-order, alt text and relationship) but resizes it to the largest
contain-box that fits between the column's text block and the panel border.

Slot geometry is taken from references/style-spec.json (page inches):
  problem column   x 0.4781-4.4636   text 13 bottom 3.7169
  method column    x 4.6207-9.7342   text 15 bottom 3.4897
  right column     x 9.8914-12.9830  text 17 bottom 2.8081, text 22 bottom 4.2531
Panels end at y = 5.5794.

Usage:
    python3 fit-figures.py --input draft.pptx [--output draft.pptx]
"""
from __future__ import annotations

import argparse
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

EMU = 914400
PANEL_BOTTOM = 5.55
GAP = 0.10
SLOTS = {
    # shape_id: (column_left, column_right, top_min, bottom_max)
    14: (0.4781, 4.4636, 3.7169 + GAP, PANEL_BOTTOM),   # context / problem
    16: (4.6207, 9.7342, 3.4897 + GAP, PANEL_BOTTOM),   # method / solution
    19: (9.8914, 12.9830, 2.8081 + 0.06, 4.2531 - 0.14),  # evaluation / results
    21: (9.8914, 12.9830, 4.2531 + 0.08, PANEL_BOTTOM),   # secondary result
}


def find_picture(slide, shape_id):
    for shape in slide.shapes:
        if shape.shape_id == shape_id and shape.shape_type == 13:
            return shape
    return None


def fit(slide):
    moved = []
    for shape_id, (left, right, top, bottom) in SLOTS.items():
        pic = find_picture(slide, shape_id)
        if pic is None:
            continue
        img = pic.image
        if not img.size:
            continue
        iw, ih = img.size
        ratio = iw / ih
        avail_w = right - left
        avail_h = bottom - top
        if avail_w <= 0 or avail_h <= 0:
            continue
        if ratio >= avail_w / avail_h:      # width-limited
            w = avail_w
            h = w / ratio
        else:                                # height-limited
            h = avail_h
            w = h * ratio
        pic.left = Emu(int(round((left + (avail_w - w) / 2) * EMU)))
        pic.top = Emu(int(round(top * EMU)))
        pic.width = Emu(int(round(w * EMU)))
        pic.height = Emu(int(round(h * EMU)))
        moved.append((shape_id, round(w, 2), round(h, 2)))
    return moved


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", default=None)
    args = ap.parse_args()
    out = Path(args.output or args.input)
    prs = Presentation(args.input)
    if len(prs.slides) != 1:
        raise SystemExit("expected a single-slide deck")
    moved = fit(prs.slides[0])
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(out)
    for shape_id, w, h in moved:
        print(f"  shape {shape_id}: {w} x {h} in")


if __name__ == "__main__":
    main()
