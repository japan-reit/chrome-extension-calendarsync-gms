"""
Chrome拡張アイコン生成スクリプト
GMS Calendar Sync 用 16x16, 32x32, 48x48, 128x128 PNG を生成する
"""
from PIL import Image, ImageDraw, ImageFont
import math
import os

def draw_rounded_rect(draw, xy, radius, fill):
    x0, y0, x1, y1 = xy
    draw.rectangle([x0 + radius, y0, x1 - radius, y1], fill=fill)
    draw.rectangle([x0, y0 + radius, x1, y1 - radius], fill=fill)
    draw.ellipse([x0, y0, x0 + radius * 2, y0 + radius * 2], fill=fill)
    draw.ellipse([x1 - radius * 2, y0, x1, y0 + radius * 2], fill=fill)
    draw.ellipse([x0, y1 - radius * 2, x0 + radius * 2, y1], fill=fill)
    draw.ellipse([x1 - radius * 2, y1 - radius * 2, x1, y1], fill=fill)

def draw_sync_arrow(draw, cx, cy, r, start_angle, sweep, width, color):
    """円弧の矢印を描く"""
    steps = max(30, int(sweep))
    points = []
    for i in range(steps + 1):
        angle = math.radians(start_angle + sweep * i / steps)
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append((x, y))

    for i in range(len(points) - 1):
        draw.line([points[i], points[i + 1]], fill=color, width=width)

    # 矢印の先端
    tip = points[-1]
    angle_tip = math.radians(start_angle + sweep)
    # 矢印頭部（接線方向 + 90度）
    perp = math.radians(start_angle + sweep + 90)
    head_len = width * 2.5
    ax = tip[0] - head_len * math.cos(angle_tip)
    ay = tip[1] - head_len * math.sin(angle_tip)
    bx = ax + head_len * 0.8 * math.cos(perp)
    by = ay + head_len * 0.8 * math.sin(perp)
    cx2 = ax - head_len * 0.8 * math.cos(perp)
    cy2 = ay - head_len * 0.8 * math.sin(perp)
    draw.polygon([tip, (bx, by), (cx2, cy2)], fill=color)

def create_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    s = size
    pad = s * 0.06
    r_corner = s * 0.18

    # ---- 背景グラデーション風（2色の矩形で代用） ----
    top_color    = (52, 120, 246)   # Google Blue
    bottom_color = (30, 90, 200)
    for y in range(s):
        t = y / s
        rc = int(top_color[0] * (1 - t) + bottom_color[0] * t)
        gc = int(top_color[1] * (1 - t) + bottom_color[1] * t)
        bc = int(top_color[2] * (1 - t) + bottom_color[2] * t)
        draw.line([(0, y), (s - 1, y)], fill=(rc, gc, bc, 255))

    # 角丸マスク
    mask = Image.new("L", (s, s), 0)
    mask_draw = ImageDraw.Draw(mask)
    draw_rounded_rect(mask_draw,
                      (int(pad), int(pad), int(s - pad - 1), int(s - pad - 1)),
                      int(r_corner), 255)
    img.putalpha(mask)

    # ---- カレンダー本体 ----
    cal_x0 = s * 0.18
    cal_y0 = s * 0.20
    cal_x1 = s * 0.82
    cal_y1 = s * 0.80
    cal_r  = s * 0.07
    cal_color = (255, 255, 255, 230)

    # 外枠
    draw_rounded_rect(draw,
                      (cal_x0, cal_y0, cal_x1, cal_y1),
                      int(cal_r), cal_color)

    # ヘッダー（青帯）
    hdr_h = (cal_y1 - cal_y0) * 0.28
    hdr_color = (0, 80, 200, 240)
    draw_rounded_rect(draw,
                      (cal_x0, cal_y0, cal_x1, cal_y0 + hdr_h),
                      int(cal_r), hdr_color)
    # ヘッダー下部を角丸なしに
    draw.rectangle([cal_x0, cal_y0 + hdr_h * 0.5,
                    cal_x1, cal_y0 + hdr_h], fill=hdr_color)

    # ---- リング（上部の留め具） ----
    ring_y = cal_y0 - s * 0.04
    ring_r = s * 0.05
    ring_x1 = s * 0.33
    ring_x2 = s * 0.67
    ring_color = (200, 220, 255, 255)
    ring_w = max(1, int(s * 0.05))
    draw.ellipse([ring_x1 - ring_r, ring_y - ring_r,
                  ring_x1 + ring_r, ring_y + ring_r],
                 fill=ring_color)
    draw.ellipse([ring_x2 - ring_r, ring_y - ring_r,
                  ring_x2 + ring_r, ring_y + ring_r],
                 fill=ring_color)

    # ---- 同期矢印（カレンダー中央） ----
    arrow_cx = s * 0.50
    arrow_cy = cal_y0 + hdr_h + (cal_y1 - cal_y0 - hdr_h) * 0.50
    arrow_r  = s * 0.165
    arrow_w  = max(1, int(s * 0.07))
    arrow_color = (30, 100, 220, 255)

    # 上半円（時計回り）
    draw_sync_arrow(draw, arrow_cx, arrow_cy, arrow_r,
                    start_angle=200, sweep=160,
                    width=arrow_w, color=arrow_color)
    # 下半円（反時計回り、180度ずらし）
    draw_sync_arrow(draw, arrow_cx, arrow_cy, arrow_r,
                    start_angle=20, sweep=160,
                    width=arrow_w, color=arrow_color)

    return img


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    icons_dir = os.path.join(out_dir, "icons")
    os.makedirs(icons_dir, exist_ok=True)

    sizes = [16, 32, 48, 128]
    for sz in sizes:
        icon = create_icon(sz)
        path = os.path.join(icons_dir, f"icon{sz}.png")
        icon.save(path, "PNG")
        print(f"  生成: {path}")

    print("完了!")


if __name__ == "__main__":
    main()
