from __future__ import annotations

import io
import re
from typing import Any

from PIL import Image

try:
    import cairosvg
except ImportError:  # pragma: no cover
    cairosvg = None

NBSP = "\u00A0"


def protect_turkish_claim_transition(text: str, *, min_tail_words: int = 5) -> str:
    """Prevent a short ``olup, özelliği;`` orphan at the end of a claim preamble.

    Binding only ``sistemi olup, özelliği;`` is not enough: Word can move that
    three-word tail to a new line.  The final ``min_tail_words`` words *before*
    ``olup`` are therefore bound with non-breaking spaces as one tail.  If the
    paragraph wraps, the last line contains a meaningful phrase rather than the
    short transition alone.
    """
    value = str(text or "")
    match = re.search(r"\s+olup,\s+özelliği;", value, flags=re.IGNORECASE)
    if not match:
        return value

    left = value[: match.start()].rstrip()
    right = value[match.end() :]
    word_matches = list(re.finditer(r"\S+", left))
    if not word_matches:
        return value
    take = min(max(1, int(min_tail_words)), len(word_matches))
    start = word_matches[-take].start()
    prefix = left[:start]
    tail = left[start:]
    tail = re.sub(r"\s+", NBSP, tail)
    return prefix + tail + NBSP + "olup," + NBSP + "özelliği;" + right


def protected_claim_tail_word_count(text: str) -> int:
    """Return the number of NBSP-bound words before ``olup`` in the protected tail."""
    value = str(text or "")
    marker = NBSP + "olup," + NBSP + "özelliği;"
    idx = value.casefold().find(marker.casefold())
    if idx < 0:
        return 0
    left = value[:idx]
    # The protected tail is the final regular-space-delimited token group.
    start = max(left.rfind(" "), left.rfind("\t"), left.rfind("\n"))
    tail = left[start + 1 :]
    return len([x for x in tail.split(NBSP) if x])


def method_step_numbers(method_steps: list[dict[str, Any]] | None) -> list[str]:
    numbers: list[str] = []
    for idx, step in enumerate(method_steps or [], start=1):
        number = str((step or {}).get("number", "") or "").strip() or str(1000 + idx)
        numbers.append(number)
    return numbers


def build_method_flow_svg(method_steps: list[dict[str, Any]] | None, language: str = "Türkçe") -> bytes:
    """Create a deterministic black/white vector method-flow figure.

    Each process-step reference is placed *inside* its own hollow box.  The figure
    deliberately contains no prose labels: the exact technical wording stays in
    REFERENCE NUMERALS / DETAILED DESCRIPTION / CLAIMS, while the arrows convey order.
    """
    refs = method_step_numbers(method_steps)
    if not refs:
        raise ValueError("Yöntem akış şekli için en az bir işlem adımı gereklidir.")

    width = 720
    box_w = 300
    box_h = 82
    top = 55
    gap = 48
    x = (width - box_w) // 2
    height = top * 2 + len(refs) * box_h + max(0, len(refs) - 1) * gap

    blocks: list[str] = []
    blocks.append(
        '<defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">'
        '<path d="M0,0 L0,6 L9,3 z" fill="#000000"/></marker></defs>'
    )
    for idx, ref in enumerate(refs):
        y = top + idx * (box_h + gap)
        blocks.append(
            f'<rect x="{x}" y="{y}" width="{box_w}" height="{box_h}" rx="6" ry="6" '
            'fill="#ffffff" stroke="#000000" stroke-width="3"/>'
        )
        blocks.append(
            f'<text x="{width/2:.1f}" y="{y + box_h/2 + 10:.1f}" text-anchor="middle" '
            'font-family="Arial" font-size="30" font-weight="normal" fill="#000000">'
            f'{ref}</text>'
        )
        if idx < len(refs) - 1:
            y1 = y + box_h
            y2 = y + box_h + gap - 8
            blocks.append(
                f'<line x1="{width/2:.1f}" y1="{y1}" x2="{width/2:.1f}" y2="{y2}" '
                'stroke="#000000" stroke-width="3" marker-end="url(#arrow)"/>'
            )

    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">'
        '<rect width="100%" height="100%" fill="#ffffff"/>'
        + ''.join(blocks)
        + '</svg>'
    )
    return svg.encode("utf-8")


def build_method_flow_png(method_steps: list[dict[str, Any]] | None, language: str = "Türkçe", *, output_width: int = 1400) -> bytes:
    if cairosvg is None:
        raise ValueError("Yöntem akış şekli için CairoSVG bağımlılığı bulunamadı.")
    return cairosvg.svg2png(bytestring=build_method_flow_svg(method_steps, language), output_width=output_width)


def _sample_rgb(data: bytes) -> Image.Image:
    with Image.open(io.BytesIO(data)) as im:
        rgb = im.convert("RGB")
        rgb.thumbnail((700, 700))
        return rgb.copy()


def material_color_ratio(data: bytes, *, channel_delta: int = 14) -> float:
    """Ratio of visibly chromatic pixels, ignoring near-white background."""
    rgb = _sample_rgb(data)
    total = 0
    colored = 0
    for r, g, b in rgb.getdata():
        if r > 248 and g > 248 and b > 248:
            continue
        total += 1
        if max(r, g, b) - min(r, g, b) > channel_delta:
            colored += 1
    return colored / max(total, 1)


def needs_line_art_normalization(data: bytes, *, ratio_threshold: float = 0.012) -> bool:
    return material_color_ratio(data) > ratio_threshold


def is_monochrome_enough(data: bytes, *, ratio_threshold: float = 0.012) -> bool:
    return material_color_ratio(data) <= ratio_threshold
