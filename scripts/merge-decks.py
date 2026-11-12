#!/usr/bin/env python3
"""Merge single-slide insight decks (same template) into one multi-slide deck.

All input decks must come from the same reference template: they share master,
layout, theme and the notes master. Each deck's first slide is copied in order,
including its pictures, alt text, animations-free shape tree and speaker notes.

Usage:
    python3 merge-decks.py --output combined.pptx 01.pptx 02.pptx ... 07.pptx

The output keeps the first deck's slide layout; every following slide is
re-created from that layout, its auto placeholders removed, and the source
shape tree copied with picture relationships re-created in the new package.
"""
from __future__ import annotations

import argparse
import copy
import io

from pptx import Presentation
from pptx.oxml.ns import qn

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SHAPE_TAGS = {qn("p:sp"), qn("p:pic"), qn("p:grpSp"), qn("p:cxnSp"),
              qn("p:graphicFrame"), qn("p:contentPart")}


def _iter_rids(element):
    for node in element.iter():
        for attr in list(node.attrib):
            if attr.startswith("{%s}" % R_NS):
                yield node, attr, node.attrib[attr]


def append_slide(prs, layout, source_slide):
    """Append a copy of ``source_slide`` built on ``layout``; returns new slide."""
    new_slide = prs.slides.add_slide(layout)
    spTree = new_slide.shapes._spTree
    # drop the auto-created placeholders; the source tree is authoritative
    for child in list(spTree):
        if child.tag in SHAPE_TAGS:
            spTree.remove(child)

    for child in list(source_slide.shapes._spTree):
        if child.tag not in SHAPE_TAGS:
            continue
        clone = copy.deepcopy(child)
        for node, attr, old_rid in list(_iter_rids(clone)):
            try:
                part = source_slide.part.related_part(old_rid)
            except KeyError:
                continue
            if part.content_type.startswith("image/"):
                _, new_rid = new_slide.part.get_or_add_image_part(io.BytesIO(part.blob))
                node.set(attr, new_rid)
        spTree.append(clone)

    if source_slide.has_notes_slide:
        text = source_slide.notes_slide.notes_text_frame.text
        if text:
            new_slide.notes_slide.notes_text_frame.text = text
    return new_slide


def merge(paths, out_path):
    prs = Presentation(paths[0])
    layout = prs.slides[0].slide_layout
    for path in paths[1:]:
        src = Presentation(path)
        if len(src.slides) != 1:
            raise SystemExit(f"{path}: expected a single-slide deck")
        if src.slide_width != prs.slide_width or src.slide_height != prs.slide_height:
            raise SystemExit(f"{path}: canvas differs from the first deck")
        append_slide(prs, layout, src.slides[0])
    prs.save(out_path)
    return len(prs.slides)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output", required=True)
    ap.add_argument("decks", nargs="+", help="single-slide .pptx files, in order")
    args = ap.parse_args()
    if len(args.decks) < 2:
        raise SystemExit("need at least two decks")
    count = merge(args.decks, args.output)
    print(args.output)
    print(f"  slides: {count}")


if __name__ == "__main__":
    main()
