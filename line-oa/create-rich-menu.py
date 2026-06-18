from PIL import Image, ImageDraw, ImageFont
import os
import math

# =========================
# LINE OA Rich Menu Config
# =========================
W, H = 2500, 1686
CELL_W = [833, 833, 834]
CELL_H = 843

BG_COLOR = "#F0F4FF"
OUTPUT = os.path.join(os.path.dirname(__file__), "rich-menu.png")

buttons = [
    {"title": "ตรวจสถานะ", "subtitle": "ดูข้อมูลไอแพด",  "color": "#4F7EF7", "emoji": "🔎"},
    {"title": "แจ้งปัญหา",  "subtitle": "ไอแพดหาย/เสีย", "color": "#F59E0B", "emoji": "🔔"},
    {"title": "ยืม / คืน",  "subtitle": "ยืม-คืนไอแพด",  "color": "#8B5CF6", "emoji": "🔄"},
    {"title": "ขอแอพ",      "subtitle": "ขอติดตั้งแอพ",  "color": "#06B6D4", "emoji": "📲"},
    {"title": "คู่มือ",      "subtitle": "วิธีใช้งาน",    "color": "#10B981", "emoji": "📘"},
    {"title": "ติดตั้งแอป",  "subtitle": "Add to Home Screen", "color": "#EF4444", "emoji": "📥"},
]

# =========================
# Font Helpers
# =========================
SCRIPT_DIR = os.path.dirname(__file__)

def find_font(possible_paths):
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

thai_font_bold = find_font([
    os.path.join(SCRIPT_DIR, "NotoSansThai-Bold.ttf"),
    r"C:\Windows\Fonts\LeelaUIb.ttf",
    r"C:\Windows\Fonts\leelawdb.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf",
])
thai_font_regular = find_font([
    os.path.join(SCRIPT_DIR, "NotoSansThai-Regular.ttf"),
    r"C:\Windows\Fonts\LeelawUI.ttf",
    r"C:\Windows\Fonts\leelawad.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
])
emoji_font_path = find_font([
    os.path.join(SCRIPT_DIR, "NotoColorEmoji.ttf"),
    r"C:\Windows\Fonts\seguiemj.ttf",
    "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf",
])

if not thai_font_bold or not thai_font_regular:
    raise FileNotFoundError(
        "ไม่พบฟอนต์ Thai กรุณาวางไฟล์ NotoSansThai-Bold.ttf และ "
        "NotoSansThai-Regular.ttf ในโฟลเดอร์ line-oa/"
    )

title_font    = ImageFont.truetype(thai_font_bold,    96)
subtitle_font = ImageFont.truetype(thai_font_regular, 42)
try:
    emoji_font = ImageFont.truetype(emoji_font_path, 120) if emoji_font_path else ImageFont.truetype(thai_font_regular, 100)
except Exception:
    emoji_font = ImageFont.truetype(thai_font_regular, 100)

# =========================
# Helpers
# =========================
def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def draw_centered_text(draw, xy, text, font, fill):
    x, y = xy
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    draw.text((x - tw/2 - bbox[0], y - th/2 - bbox[1]), text, font=font, fill=fill)

def add_radial_gradient(base):
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    px = overlay.load()
    cx, cy = W//2, H//3
    max_d = math.sqrt(cx**2 + cy**2)
    gc = hex_to_rgb("#2563EB")
    for y in range(H):
        for x in range(W):
            d = math.sqrt((x-cx)**2 + (y-cy)**2)
            s = max(0, 1 - d/max_d)
            px[x, y] = (*gc, int(80 * s**1.8))
    return Image.alpha_composite(base, overlay)

def add_vertical_fade(img, box, color, max_alpha=90):
    x1, y1, x2, y2 = box
    ac = hex_to_rgb(color)
    fade = Image.new("RGBA", (x2-x1, y2-y1), (0,0,0,0))
    fp = fade.load()
    fh = y2-y1
    for y in range(fh):
        a = int(max_alpha * (1 - y/fh))
        for x in range(x2-x1):
            fp[x, y] = (*ac, a)
    img.alpha_composite(fade, (x1, y1))

# =========================
# Canvas
# =========================
img = Image.new("RGBA", (W, H), BG_COLOR)
img = add_radial_gradient(img)
draw = ImageDraw.Draw(img, "RGBA")

margin = 8
radius = 38
accent_h = 24
fade_h = 120
circle_d = 210
x_pos = [0, CELL_W[0], CELL_W[0]+CELL_W[1]]

for i, item in enumerate(buttons):
    row, col = i//3, i%3
    cx0 = x_pos[col]
    cy0 = row * CELL_H
    cw  = CELL_W[col]
    x1, y1 = cx0+margin, cy0+margin
    x2, y2 = cx0+cw-margin, cy0+CELL_H-margin
    ac  = item["color"]
    acr = hex_to_rgb(ac)

    # Card bg
    draw.rounded_rectangle((x1,y1,x2,y2), radius=radius, fill=(255,255,255,255), outline=(220,225,240,255), width=2)

    # Top accent strip
    draw.rounded_rectangle((x1,y1,x2,y1+accent_h+radius), radius=radius, fill=(*acr,255))
    draw.rectangle((x1,y1+accent_h,x2,y1+accent_h+radius), fill=(*acr,255))

    # Fade below strip
    add_vertical_fade(img, (x1, y1+accent_h, x2, y1+accent_h+fade_h), ac, max_alpha=90)
    draw = ImageDraw.Draw(img, "RGBA")

    # Icon circle
    icx = (x1+x2)//2
    icy = cy0 + 230
    ix1, iy1 = icx-circle_d//2, icy-circle_d//2
    ix2, iy2 = icx+circle_d//2, icy+circle_d//2
    draw.ellipse((ix1,iy1,ix2,iy2), fill=(*acr,60), outline=(*acr,120), width=3)

    # Emoji
    draw_centered_text(draw, (icx, icy-4), item["emoji"], emoji_font, (255,255,255,255))

    # Title
    ty = icy + 195
    draw_centered_text(draw, (icx, ty), item["title"], title_font, (15,23,42,255))

    # Subtitle
    draw_centered_text(draw, (icx, ty+88), item["subtitle"], subtitle_font, (100,116,139,255))

img.convert("RGBA").save(OUTPUT)
print(f"Saved: {OUTPUT}")
