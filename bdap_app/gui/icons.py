"""Modulo per la generazione e gestione degli asset grafici e icone per la GUI BDAP."""

from __future__ import annotations

import math
from typing import Dict, Tuple
from PIL import Image, ImageDraw, ImageTk

# Palette colori di riferimento
PRIMARY_BLUE = "#1976d2"
PRIMARY_HOVER = "#1565c0"
PRIMARY_DISABLED = "#90caf9"
BG_BLOB = "#e0edfc"
TEXT_NAVY = "#0f2b5c"
TEXT_SLATE = "#5a6e85"
TEXT_MUTED = "#64748b"
BORDER_LIGHT = "#dbe7f5"
WHITE = "#ffffff"


def _hex_to_rgba(hex_color: str, alpha: int = 255) -> Tuple[int, int, int, int]:
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (r, g, b, alpha)


class IconCache:
    """Cache per le immagini PhotoImage create, per prevenire garbage collection."""
    _cache: Dict[str, ImageTk.PhotoImage] = {}

    @classmethod
    def get(cls, key: str) -> ImageTk.PhotoImage | None:
        return cls._cache.get(key)

    @classmethod
    def store(cls, key: str, photo: ImageTk.PhotoImage) -> ImageTk.PhotoImage:
        cls._cache[key] = photo
        return photo


def create_header_main_icon(size: int = 68) -> ImageTk.PhotoImage:
    key = f"header_main_{size}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    color = _hex_to_rgba(PRIMARY_BLUE)

    doc_x0, doc_y0 = int(s * 0.10), int(s * 0.08)
    doc_x1, doc_y1 = int(s * 0.74), int(s * 0.88)
    corner = int(s * 0.20)
    stroke = int(s * 0.065)
    r = int(s * 0.08)

    draw.rounded_rectangle([doc_x0, doc_y0, doc_x1, doc_y1], radius=r, outline=color, width=stroke)
    draw.line([(doc_x1 - corner, doc_y0), (doc_x1 - corner, doc_y0 + corner)], fill=color, width=stroke)
    draw.line([(doc_x1 - corner, doc_y0 + corner), (doc_x1, doc_y0 + corner)], fill=color, width=stroke)

    line_x0 = int(s * 0.24)
    line_x1 = int(s * 0.56)
    for y_rel in [0.30, 0.44, 0.58]:
        y = int(s * y_rel)
        draw.rounded_rectangle([line_x0, y - stroke // 3, line_x1, y + stroke // 3], radius=stroke // 3, fill=color)

    gx, gy = int(s * 0.73), int(s * 0.73)
    g_radius = int(s * 0.20)
    g_hole = int(s * 0.08)

    draw.ellipse([gx - g_radius - stroke, gy - g_radius - stroke, gx + g_radius + stroke, gy + g_radius + stroke], fill=(242, 246, 252, 255))

    num_teeth = 8
    teeth_len = int(s * 0.065)
    for i in range(num_teeth):
        angle = i * (2 * math.pi / num_teeth)
        dx = math.cos(angle)
        dy = math.sin(angle)
        tx0 = gx + dx * (g_radius - teeth_len // 2)
        ty0 = gy + dy * (g_radius - teeth_len // 2)
        tx1 = gx + dx * (g_radius + teeth_len)
        ty1 = gy + dy * (g_radius + teeth_len)
        draw.line([tx0, ty0, tx1, ty1], fill=color, width=int(stroke * 1.3))

    draw.ellipse([gx - g_radius, gy - g_radius, gx + g_radius, gy + g_radius], fill=color)
    draw.ellipse([gx - g_hole, gy - g_hole, gx + g_hole, gy + g_hole], fill=(242, 246, 252, 255))

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_header_illustration(width: int = 190, height: int = 105) -> ImageTk.PhotoImage:
    key = f"header_ill_{width}_{height}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 3
    w, h = width * scale, height * scale
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    color = _hex_to_rgba(PRIMARY_BLUE)
    blob_color = _hex_to_rgba(BG_BLOB)

    blob_cx, blob_cy = int(w * 0.65), int(h * 0.52)
    blob_r = int(h * 0.46)
    draw.ellipse([blob_cx - blob_r * 1.25, blob_cy - blob_r, blob_cx + blob_r * 1.25, blob_cy + blob_r], fill=blob_color)

    draw.ellipse([int(w * 0.22), int(h * 0.35), int(w * 0.22) + 24, int(h * 0.35) + 24], fill=blob_color)
    draw.ellipse([int(w * 0.92), int(h * 0.20), int(w * 0.92) + 18, int(h * 0.20) + 18], fill=blob_color)

    doc_x0, doc_y0 = int(w * 0.44), int(h * 0.16)
    doc_x1, doc_y1 = int(w * 0.65), int(h * 0.76)
    corner = int(w * 0.06)
    stroke = int(h * 0.038)
    r = int(h * 0.05)
    draw.rounded_rectangle([doc_x0, doc_y0, doc_x1, doc_y1], radius=r, outline=color, width=stroke, fill=(255, 255, 255, 240))
    draw.line([(doc_x1 - corner, doc_y0), (doc_x1 - corner, doc_y0 + corner)], fill=color, width=stroke)
    draw.line([(doc_x1 - corner, doc_y0 + corner), (doc_x1, doc_y0 + corner)], fill=color, width=stroke)
    for y_rel in [0.35, 0.48, 0.61]:
        y = int(h * y_rel)
        draw.line([(doc_x0 + 20, y), (doc_x1 - 24, y)], fill=color, width=stroke)

    db_x0, db_y0 = int(w * 0.68), int(h * 0.46)
    db_x1, db_y1 = int(w * 0.83), int(h * 0.86)
    db_w = db_x1 - db_x0
    draw.rounded_rectangle([db_x0, db_y0, db_x1, db_y1], radius=int(db_w * 0.2), fill=color)
    dw = int(h * 0.025)
    draw.line([(db_x0, int(h * 0.58)), (db_x1, int(h * 0.58))], fill=(255, 255, 255, 255), width=dw)
    draw.line([(db_x0, int(h * 0.70)), (db_x1, int(h * 0.70))], fill=(255, 255, 255, 255), width=dw)

    arrow_pts = []
    for t_i in range(30):
        t = t_i / 29.0
        p0 = (w * 0.72, h * 0.28)
        p1 = (w * 0.88, h * 0.22)
        p2 = (w * 0.89, h * 0.52)
        bx = (1 - t) ** 2 * p0[0] + 2 * (1 - t) * t * p1[0] + t ** 2 * p2[0]
        by = (1 - t) ** 2 * p0[1] + 2 * (1 - t) * t * p1[1] + t ** 2 * p2[1]
        arrow_pts.append((bx, by))

    for i in range(len(arrow_pts) - 1):
        draw.line([arrow_pts[i], arrow_pts[i + 1]], fill=color, width=int(stroke * 0.9))

    tip = arrow_pts[-1]
    draw.polygon([(tip[0], tip[1] + 12), (tip[0] - 14, tip[1] - 10), (tip[0] + 14, tip[1] - 10)], fill=color)

    resized = img.resize((width, height), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_badge_database(size: int = 32) -> ImageTk.PhotoImage:
    key = f"badge_db_{size}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    draw.ellipse([0, 0, s - 1, s - 1], fill=_hex_to_rgba(PRIMARY_BLUE))

    cx = s // 2
    w = int(s * 0.44)
    h_layer = int(s * 0.12)
    spacing = int(s * 0.14)
    y_start = int(s * 0.32)

    for i in range(3):
        y = y_start + i * spacing
        draw.rounded_rectangle([cx - w // 2, y - h_layer // 2, cx + w // 2, y + h_layer // 2], radius=h_layer // 2, fill=(255, 255, 255, 255))

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_badge_sliders(size: int = 32) -> ImageTk.PhotoImage:
    key = f"badge_sliders_{size}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    draw.ellipse([0, 0, s - 1, s - 1], fill=_hex_to_rgba(PRIMARY_BLUE))

    stroke = int(s * 0.05)
    line_x0 = int(s * 0.28)
    line_x1 = int(s * 0.72)
    knob_r = int(s * 0.08)

    configs = [(0.35, 0.42), (0.50, 0.60), (0.65, 0.45)]
    for y_rel, knob_rel in configs:
        y = int(s * y_rel)
        draw.line([(line_x0, y), (line_x1, y)], fill=(255, 255, 255, 255), width=stroke)
        kx = int(line_x0 + (line_x1 - line_x0) * knob_rel)
        draw.ellipse([kx - knob_r, y - knob_r, kx + knob_r, y + knob_r], fill=_hex_to_rgba(PRIMARY_BLUE), outline=(255, 255, 255, 255), width=stroke)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_cloud_upload(size: int = 42) -> ImageTk.PhotoImage:
    key = f"cloud_upload_{size}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    color = _hex_to_rgba(PRIMARY_BLUE)
    stroke = int(s * 0.065)

    draw.ellipse([int(s * 0.35), int(s * 0.25), int(s * 0.65), int(s * 0.58)], outline=color, width=stroke)
    draw.ellipse([int(s * 0.18), int(s * 0.42), int(s * 0.46), int(s * 0.72)], outline=color, width=stroke)
    draw.ellipse([int(s * 0.54), int(s * 0.44), int(s * 0.82), int(s * 0.72)], outline=color, width=stroke)

    bg_tint = (245, 249, 255, 255)
    draw.ellipse([int(s * 0.36), int(s * 0.27), int(s * 0.64), int(s * 0.56)], fill=bg_tint)
    draw.ellipse([int(s * 0.20), int(s * 0.44), int(s * 0.44), int(s * 0.70)], fill=bg_tint)
    draw.ellipse([int(s * 0.56), int(s * 0.46), int(s * 0.80), int(s * 0.70)], fill=bg_tint)
    draw.rectangle([int(s * 0.26), int(s * 0.46), int(s * 0.74), int(s * 0.70)], fill=bg_tint)

    ax = s // 2
    ay_top = int(s * 0.38)
    ay_bot = int(s * 0.66)
    draw.line([(ax, ay_top), (ax, ay_bot)], fill=color, width=stroke)
    draw.line([(ax - int(s * 0.10), ay_top + int(s * 0.10)), (ax, ay_top)], fill=color, width=stroke)
    draw.line([(ax + int(s * 0.10), ay_top + int(s * 0.10)), (ax, ay_top)], fill=color, width=stroke)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_excel_icon(size: int = 24) -> ImageTk.PhotoImage:
    key = f"excel_icon_{size}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.07)

    doc_x0, doc_y0 = int(s * 0.32), int(s * 0.12)
    doc_x1, doc_y1 = int(s * 0.88), int(s * 0.88)
    draw.rounded_rectangle([doc_x0, doc_y0, doc_x1, doc_y1], radius=int(s * 0.1), outline=(255, 255, 255, 240), width=stroke)
    for y_rel in [0.35, 0.52, 0.69]:
        y = int(s * y_rel)
        draw.line([(int(s * 0.44), y), (int(s * 0.78), y)], fill=(255, 255, 255, 200), width=int(stroke * 0.8))

    badge_x0, badge_y0 = int(s * 0.12), int(s * 0.28)
    badge_x1, badge_y1 = int(s * 0.52), int(s * 0.76)
    draw.rounded_rectangle([badge_x0, badge_y0, badge_x1, badge_y1], radius=int(s * 0.08), fill=(255, 255, 255, 255))
    draw.rounded_rectangle([badge_x0 + 2, badge_y0 + 2, badge_x1 - 2, badge_y1 - 2], radius=int(s * 0.07), fill=(16, 124, 65, 255))

    bx = (badge_x0 + badge_x1) // 2
    by = (badge_y0 + badge_y1) // 2
    arm = int(s * 0.11)
    draw.line([(bx - arm, by - arm), (bx + arm, by + arm)], fill=(255, 255, 255, 255), width=int(stroke * 1.1))
    draw.line([(bx + arm, by - arm), (bx - arm, by + arm)], fill=(255, 255, 255, 255), width=int(stroke * 1.1))

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_calendar_icon(size: int = 18, color_hex: str = PRIMARY_BLUE) -> ImageTk.PhotoImage:
    key = f"calendar_{size}_{color_hex}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.08)
    color = _hex_to_rgba(color_hex)

    x0, y0 = int(s * 0.15), int(s * 0.22)
    x1, y1 = int(s * 0.85), int(s * 0.88)
    draw.rounded_rectangle([x0, y0, x1, y1], radius=int(s * 0.1), outline=color, width=stroke)
    bar_y = int(s * 0.42)
    draw.line([(x0, bar_y), (x1, bar_y)], fill=color, width=stroke)
    loop_w = stroke
    draw.line([(int(s * 0.35), int(s * 0.10)), (int(s * 0.35), int(s * 0.26))], fill=color, width=loop_w)
    draw.line([(int(s * 0.65), int(s * 0.10)), (int(s * 0.65), int(s * 0.26))], fill=color, width=loop_w)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_refresh_icon(size: int = 18, color_hex: str = WHITE) -> ImageTk.PhotoImage:
    key = f"refresh_{size}_{color_hex}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.11)
    color = _hex_to_rgba(color_hex)

    cx, cy = s // 2, s // 2
    r = int(s * 0.34)
    draw.arc([cx - r, cy - r, cx + r, cy + r], start=45, end=300, fill=color, width=stroke)

    angle = math.radians(300)
    tx = cx + r * math.cos(angle)
    ty = cy + r * math.sin(angle)
    draw.polygon([(tx, ty - stroke * 1.2), (tx + stroke * 1.5, ty + stroke * 0.8), (tx - stroke * 0.8, ty + stroke * 1.2)], fill=color)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_folder_icon(size: int = 18, color_hex: str = PRIMARY_BLUE) -> ImageTk.PhotoImage:
    key = f"folder_{size}_{color_hex}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.08)
    color = _hex_to_rgba(color_hex)

    x0, y0 = int(s * 0.12), int(s * 0.24)
    x1, y1 = int(s * 0.88), int(s * 0.80)
    tab_w = int(s * 0.32)
    tab_h = int(s * 0.12)

    draw.rounded_rectangle([x0, y0 - tab_h, x0 + tab_w, y0 + 4], radius=stroke, outline=color, width=stroke)
    draw.rounded_rectangle([x0, y0, x1, y1], radius=int(s * 0.1), outline=color, width=stroke)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_chevron_down(size: int = 16, color_hex: str = PRIMARY_BLUE) -> ImageTk.PhotoImage:
    key = f"chevron_down_{size}_{color_hex}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.10)
    color = _hex_to_rgba(color_hex)

    cx, cy = s // 2, int(s * 0.55)
    w = int(s * 0.28)
    h = int(s * 0.18)
    draw.line([(cx - w, cy - h), (cx, cy + h)], fill=color, width=stroke)
    draw.line([(cx, cy + h), (cx + w, cy - h)], fill=color, width=stroke)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_chevron_up(size: int = 16, color_hex: str = PRIMARY_BLUE) -> ImageTk.PhotoImage:
    key = f"chevron_up_{size}_{color_hex}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.10)
    color = _hex_to_rgba(color_hex)

    cx, cy = s // 2, int(s * 0.45)
    w = int(s * 0.28)
    h = int(s * 0.18)
    draw.line([(cx - w, cy + h), (cx, cy - h)], fill=color, width=stroke)
    draw.line([(cx, cy - h), (cx + w, cy + h)], fill=color, width=stroke)

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))


def create_check_circle_icon(size: int = 20, color_hex: str = "#16a34a") -> ImageTk.PhotoImage:
    key = f"check_circle_{size}_{color_hex}"
    cached = IconCache.get(key)
    if cached:
        return cached

    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    stroke = int(s * 0.08)
    color = _hex_to_rgba(color_hex)

    cx, cy = s // 2, s // 2
    r = int(s * 0.40)
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=_hex_to_rgba("#eefcf3"), outline=color, width=stroke)

    pts = [
        (int(s * 0.32), int(s * 0.50)),
        (int(s * 0.44), int(s * 0.65)),
        (int(s * 0.70), int(s * 0.36)),
    ]
    draw.line([pts[0], pts[1]], fill=color, width=int(stroke * 1.3))
    draw.line([pts[1], pts[2]], fill=color, width=int(stroke * 1.3))

    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    return IconCache.store(key, ImageTk.PhotoImage(resized))

