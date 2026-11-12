#!/usr/bin/env python3
"""Crop a paper figure for the paper-insight slide.

Rasterises one page of a PDF with the DSH-bundled LibreOffice Kit, then crops a
fractional box and trims the white margin. Optionally crops from an existing
page image (``--page-image``) when the page was rendered elsewhere.

Examples:
    python3 crop-figure.py --paper paper.pdf --page 2 --box 0.51,0.07,0.96,0.20 --out fig1.png
    python3 crop-figure.py --page-image page-0005.png --box 0.49,0.06,1.0,0.44 \
        --out fig4.png --trim

Coordinates are fractions (0-1) of the full page: ``left,top,right,bottom``.
``--trim`` then cuts the residual white margin inside that box.
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image


# --------------------------------------------------------------------------- #
# LibreOffice Kit discovery (same layout the paper-insight-slide skill uses)
# --------------------------------------------------------------------------- #
def find_office_kit():
    node = os.environ.get("DSH_OFFICE_KIT_NODE")
    cli = os.environ.get("DSH_OFFICE_KIT_CLI")
    if node and cli and os.path.isfile(node) and os.path.isfile(cli):
        return {"node": node, "cli": cli}
    roots = ["/Applications/DeepSeek Harness.app/Contents/Resources"]
    roots += sorted(glob.glob("/Applications/DeepSeek*Harness*.app/Contents/Resources"))
    for root in roots:
        cli = os.path.join(root, "app.asar.unpacked", "dsh", "node_modules",
                           "@deepseek-ai", "libreoffice-kit", "lib", "cli.js")
        if not os.path.isfile(cli):
            continue
        patterns = [
            os.path.join(root, "runtime", "*", "dependencies", "node", "bin", "node"),
            os.path.join(root, "runtime", "primary-runtime", "dependencies", "node", "bin", "node"),
        ]
        for pattern in patterns:
            for candidate in sorted(glob.glob(pattern)):
                if os.path.isfile(candidate):
                    return {"node": candidate, "cli": cli}
        node = shutil.which("node")
        if node:
            return {"node": node, "cli": cli}
    return None


def render_pdf_page(pdf, page, dpi):
    kit = find_office_kit()
    if not kit:
        raise SystemExit(
            "找不到 DSH 自带的 LibreOffice Kit 渲染器；"
            "可设置 DSH_OFFICE_KIT_NODE / DSH_OFFICE_KIT_CLI，或先用 office-render 技能导出页面 PNG")
    tmp = tempfile.mkdtemp(prefix=".crop-figure-")
    outdir = os.path.join(tmp, "out")
    try:
        proc = subprocess.run(
            [kit["node"], kit["cli"], "render", "--input", os.path.abspath(pdf),
             "--output-dir", outdir, "--pages", str(int(page)), "--dpi", str(int(dpi))],
            capture_output=True, text=True, timeout=300)
        if proc.returncode != 0:
            raise SystemExit(f"渲染失败：{(proc.stderr or proc.stdout)[:800]}")
        manifest = json.loads(proc.stdout[proc.stdout.find("{"):])
        images = manifest.get("images") or []
        if not images:
            raise SystemExit("渲染没有返回图像")
        # copy out of the temporary tree before it is removed
        result = os.path.join(tmp, "page.png")
        shutil.copyfile(images[0]["path"], result)
        return result, tmp
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise


def parse_box(text):
    parts = [p.strip() for p in text.split(",")]
    if len(parts) != 4:
        raise SystemExit("--box 需要 4 个 0-1 的小数：left,top,right,bottom")
    values = [float(p) for p in parts]
    if not all(0.0 <= v <= 1.0 for v in values):
        raise SystemExit("--box 的四个值都应在 0-1 之间")
    left, top, right, bottom = values
    if right <= left or bottom <= top:
        raise SystemExit("--box 需要 left<right 且 top<bottom")
    return left, top, right, bottom


def trim_white(img, padding=6, threshold=246):
    """Trim near-white margins, keeping a small padding in pixels."""
    rgb = img.convert("RGB")
    w, h = rgb.size
    pixels = rgb.load()
    min_x, min_y, max_x, max_y = w, h, -1, -1
    for y in range(h):
        for x in range(w):
            r, g, b = pixels[x, y]
            if r < threshold or g < threshold or b < threshold:
                if x < min_x:
                    min_x = x
                if x > max_x:
                    max_x = x
                if y < min_y:
                    min_y = y
                if y > max_y:
                    max_y = y
    if max_x < 0:
        return rgb
    left = max(0, min_x - padding)
    top = max(0, min_y - padding)
    right = min(w, max_x + 1 + padding)
    bottom = min(h, max_y + 1 + padding)
    return rgb.crop((left, top, right, bottom))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--paper", help="source PDF")
    ap.add_argument("--page-image", help="existing rendered page PNG (alternative to --paper)")
    ap.add_argument("--page", type=int, default=1, help="one-based page number for --paper")
    ap.add_argument("--dpi", type=int, default=200, help="render DPI for --paper (default 200)")
    ap.add_argument("--box", required=True, help="crop box as fractions left,top,right,bottom")
    ap.add_argument("--out", required=True, help="output PNG path")
    ap.add_argument("--trim", action="store_true", help="trim white margins after cropping")
    ap.add_argument("--max-width", type=int, default=2400, help="downscale wider results")
    args = ap.parse_args(argv)

    if not args.paper and not args.page_image:
        raise SystemExit("需要 --paper 或 --page-image")
    left, top, right, bottom = parse_box(args.box)

    tmpdir = None
    try:
        if args.page_image:
            page_path = args.page_image
        else:
            page_path, tmpdir = render_pdf_page(args.paper, args.page, args.dpi)
        page_img = Image.open(page_path).convert("RGB")
        w, h = page_img.size
        box_px = (int(round(left * w)), int(round(top * h)),
                  int(round(right * w)), int(round(bottom * h)))
        if box_px[2] <= box_px[0] or box_px[3] <= box_px[1]:
            raise SystemExit(f"裁剪框太小：{box_px}")
        cropped = page_img.crop(box_px)
        if args.trim:
            cropped = trim_white(cropped)
        if args.max_width and cropped.width > args.max_width:
            ratio = args.max_width / cropped.width
            cropped = cropped.resize((args.max_width, int(round(cropped.height * ratio))),
                                     Image.LANCZOS)
        out = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        cropped.save(out, "PNG", optimize=True)
        print(json.dumps({"output": out, "size_px": list(cropped.size),
                          "aspect": round(cropped.width / cropped.height, 4),
                          "source_page_image": os.path.abspath(page_path),
                          "page": args.page if args.paper else None}, ensure_ascii=False))
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
