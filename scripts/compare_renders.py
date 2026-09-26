#!/usr/bin/env python3
"""Create full-page and regional visual comparison artifacts without judging fidelity."""

from __future__ import annotations

# Adapted from image-to-editable-ppt (MIT), commit b6aba71f.
# Copyright (c) 2026 Image to Editable PPT contributors.
# See THIRD_PARTY_NOTICES.md for the complete license.
import argparse
import hashlib
import html
import json
import math
import re
import shutil
import sys
from pathlib import Path

try:
    from PIL import Image, ImageChops, ImageEnhance, ImageOps, ImageStat
except ImportError as exc:
    raise SystemExit("Pillow is required for comparison QA. Install it explicitly; OCR is not required.") from exc


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_name(value: str, index: int) -> str:
    readable = re.sub(r"[^a-zA-Z0-9._-]+", "-", value).strip("-") or "region"
    suffix = hashlib.sha256(value.encode("utf-8")).hexdigest()[:8]
    return f"{index:02d}-{readable}-{suffix}"


def metrics(reference: Image.Image, rendered: Image.Image) -> dict[str, float]:
    difference = ImageChops.difference(reference, rendered)
    means = ImageStat.Stat(difference).mean
    rms = ImageStat.Stat(difference).rms
    return {
        "meanAbsoluteErrorNormalized": round(sum(means) / len(means) / 255, 6),
        "rmsNormalized": round(math.sqrt(sum(value * value for value in rms) / len(rms)) / 255, 6),
    }


def save_comparison_set(reference: Image.Image, rendered: Image.Image, output: Path, prefix: str) -> dict[str, str | float]:
    reference_path = output / f"{prefix}-reference.png"
    rendered_path = output / f"{prefix}-rendered.png"
    overlay_path = output / f"{prefix}-overlay.png"
    difference_path = output / f"{prefix}-difference.png"
    side_path = output / f"{prefix}-side-by-side.png"
    reference.save(reference_path)
    rendered.save(rendered_path)
    Image.blend(reference, rendered, 0.5).save(overlay_path)
    difference = ImageChops.difference(reference, rendered)
    ImageEnhance.Contrast(ImageOps.autocontrast(difference)).enhance(2).save(difference_path)
    gap = 24
    side = Image.new("RGB", (reference.width * 2 + gap, reference.height), "white")
    side.paste(reference, (0, 0))
    side.paste(rendered, (reference.width + gap, 0))
    side.save(side_path)
    return {
        "reference": reference_path.name,
        "rendered": rendered_path.name,
        "sideBySide": side_path.name,
        "overlay": overlay_path.name,
        "difference": difference_path.name,
        **metrics(reference, rendered),
    }


def build_html(report: dict, output: Path) -> None:
    rows = []
    full = report.get("fullPage")
    if full:
        rows.append(("Full page", full))
    for region in report.get("regions", []):
        rows.append((f"Region: {region['id']}", region))
    cards = []
    for title, item in rows:
        images = "".join(
            f'<figure><figcaption>{html.escape(label)}</figcaption><img src="{html.escape(str(item[key]))}" alt="{html.escape(title)} {html.escape(label)}"></figure>'
            for label, key in (("Reference", "reference"), ("Rendered", "rendered"), ("Side by side", "sideBySide"), ("50% overlay", "overlay"), ("Enhanced difference", "difference"))
            if key in item
        )
        score = ""
        if "meanAbsoluteErrorNormalized" in item:
            score = f"<p>Auxiliary pixel measures: MAE {item['meanAbsoluteErrorNormalized']:.6f}; RMS {item['rmsNormalized']:.6f}. These are not pass/fail thresholds.</p>"
        cards.append(f"<section><h2>{html.escape(title)}</h2>{score}<div class=grid>{images}</div></section>")
    warning = ""
    if report["status"] == "aspect-ratio-mismatch":
        warning = "<p class=warning>Aspect ratios differ. No resize, overlay, difference, or region comparison was produced because that could hide a canvas error.</p>"
    document = f"""<!doctype html>
<meta charset="utf-8">
<title>Visio source comparison</title>
<style>
body{{font:15px system-ui,sans-serif;margin:24px;background:#f4f5f7;color:#202124}}h1,h2{{margin:.2em 0 .5em}}section{{background:white;border:1px solid #d9dde3;border-radius:10px;margin:20px 0;padding:18px}}.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px}}figure{{margin:0}}figcaption{{font-weight:650;margin-bottom:6px}}img{{display:block;max-width:100%;height:auto;border:1px solid #c9cdd3;background:#fff}}.warning{{padding:12px;background:#fff2cc;border:1px solid #e3bd43}}
</style>
<h1>Reference / rendered comparison</h1>
{warning}
<p>Reference: {html.escape(report['reference']['path'])}<br>Rendered: {html.escape(report['rendered']['path'])}</p>
{''.join(cards)}
"""
    (output / "comparison.html").write_text(document, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare a source page image with a rendered page without assigning a pass threshold.")
    parser.add_argument("reference", type=Path)
    parser.add_argument("rendered", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--page", type=int, default=1)
    parser.add_argument("--aspect-tolerance", type=float, default=0.005, help="Maximum relative aspect-ratio difference; default 0.5%%.")
    options = parser.parse_args()
    if not math.isfinite(options.aspect_tolerance) or options.aspect_tolerance < 0:
        parser.error("--aspect-tolerance must be finite and non-negative")
    reference_path = options.reference.resolve()
    rendered_path = options.rendered.resolve()
    output = options.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    def opaque(path):
        with Image.open(path) as source:
            rgba = ImageOps.exif_transpose(source).convert("RGBA")
            background = Image.new("RGBA", rgba.size, "white")
            return Image.alpha_composite(background, rgba).convert("RGB")
    reference = opaque(reference_path)
    rendered = opaque(rendered_path)
    reference_copy = output / f"source-full{reference_path.suffix.lower() or '.png'}"
    rendered_copy = output / f"rendered-full{rendered_path.suffix.lower() or '.png'}"
    if reference_path != reference_copy:
        shutil.copy2(reference_path, reference_copy)
    if rendered_path != rendered_copy:
        shutil.copy2(rendered_path, rendered_copy)
    reference_ratio = reference.width / reference.height
    rendered_ratio = rendered.width / rendered.height
    relative_ratio_difference = abs(reference_ratio - rendered_ratio) / reference_ratio
    report = {
        "status": "measured",
        "note": "Pixel metrics are diagnostic only. Review the full page and every critical region visually.",
        "reference": {"path": str(reference_path), "sha256": sha256(reference_path), "width": reference.width, "height": reference.height, "aspectRatio": reference_ratio, "copy": reference_copy.name},
        "rendered": {"path": str(rendered_path), "sha256": sha256(rendered_path), "width": rendered.width, "height": rendered.height, "aspectRatio": rendered_ratio, "copy": rendered_copy.name},
        "relativeAspectRatioDifference": relative_ratio_difference,
        "aspectTolerance": options.aspect_tolerance,
        "regions": [],
    }
    if relative_ratio_difference > options.aspect_tolerance:
        report["status"] = "aspect-ratio-mismatch"
        (output / "comparison.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
        build_html(report, output)
        print(json.dumps(report, indent=2))
        return 2

    scale = min(reference.width / rendered.width, reference.height / rendered.height)
    scaled_size = (max(1, round(rendered.width * scale)), max(1, round(rendered.height * scale)))
    resized = rendered.resize(scaled_size, Image.Resampling.LANCZOS)
    aligned = Image.new("RGB", reference.size, "white")
    paste_at = ((reference.width - resized.width) // 2, (reference.height - resized.height) // 2)
    aligned.paste(resized, paste_at)
    report["alignment"] = {"method": "proportional-fit-no-stretch", "scale": scale, "renderedSize": list(scaled_size), "padding": [paste_at[0], paste_at[1]]}
    report["fullPage"] = save_comparison_set(reference, aligned, output, "full-page")

    if options.manifest:
        manifest_path = options.manifest.resolve()
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        page = next((item for item in manifest.get("pages", []) if item.get("page") == options.page), None)
        if page is None:
            raise SystemExit(f"Manifest has no page {options.page}: {manifest_path}")
        canvas = page.get("canvas", {})
        width_px, height_px = canvas.get("widthPx"), canvas.get("heightPx")
        if not isinstance(width_px, (int, float)) or not isinstance(height_px, (int, float)) or width_px <= 0 or height_px <= 0:
            raise SystemExit("Manifest page canvas is invalid.")
        manifest_ratio = width_px / height_px
        relative_manifest_ratio_difference = abs(manifest_ratio - reference_ratio) / reference_ratio
        if relative_manifest_ratio_difference > options.aspect_tolerance:
            raise SystemExit(
                f"Manifest canvas ratio {manifest_ratio:.6f} does not match reference ratio {reference_ratio:.6f}; region crops were not produced."
            )
        report["manifestCanvasRelativeAspectRatioDifference"] = relative_manifest_ratio_difference
        for region_index, region in enumerate(page.get("regions", []), start=1):
            bbox = region.get("bbox", {})
            values = [bbox.get(k) for k in ("x", "y", "w", "h")]
            if any(isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) for v in values):
                raise SystemExit("Region bbox requires four finite numbers")
            x, y, w, h = values
            if min(x, y) < 0 or min(w, h) <= 0 or x+w > width_px or y+h > height_px:
                raise SystemExit("Region bbox lies outside the source canvas")
            box = (
                max(0, round(bbox["x"] / width_px * reference.width)),
                max(0, round(bbox["y"] / height_px * reference.height)),
                min(reference.width, round((bbox["x"] + bbox["w"]) / width_px * reference.width)),
                min(reference.height, round((bbox["y"] + bbox["h"]) / height_px * reference.height)),
            )
            prefix = f"region-{safe_name(str(region.get('id', 'region')), region_index)}"
            item = save_comparison_set(reference.crop(box), aligned.crop(box), output, prefix)
            item.update({"id": region.get("id"), "critical": region.get("critical"), "sourcePixelBox": list(box)})
            report["regions"].append(item)
        report["manifest"] = {"path": str(manifest_path), "sha256": sha256(manifest_path), "page": options.page}

    (output / "comparison.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    build_html(report, output)
    print(json.dumps({"status": report["status"], "comparison": str(output / 'comparison.html'), "report": str(output / 'comparison.json'), "regions": len(report["regions"])}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
