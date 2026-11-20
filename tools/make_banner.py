"""Build the CleanSheet repository banner from scratch (artificial, simple)."""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "assets")
os.makedirs(OUT, exist_ok=True)

W, H = 1600, 440
BG = (248, 249, 250)          # near-white
PANEL = (233, 236, 239)       # light panel
ACCENT = (14, 165, 233)       # clean teal/cyan
ACCENT_DARK = (2, 132, 199)
TEXT = (17, 24, 39)           # near-black
MUTED = (107, 114, 128)       # gray
GRID = (209, 213, 219)        # light grid

canvas = Image.new("RGB", (W, H), BG)
draw = ImageDraw.Draw(canvas)

# --- left accent bar ---
draw.rectangle([(0, 0), (10, H)], fill=ACCENT)

# --- artificial spreadsheet grid motif (simple, no real data) ---
grid_left = 760
grid_top = 76
grid_cell_w = 110
grid_cell_h = 52
cols = 5
rows = 5

for r in range(rows):
    for c in range(cols):
        x = grid_left + c * grid_cell_w
        y = grid_top + r * grid_cell_h
        # header row shaded
        fill = PANEL if r == 0 else BG
        draw.rectangle([(x, y), (x + grid_cell_w - 4, y + grid_cell_h - 4)],
                       fill=fill, outline=GRID, width=1)

# --- typography ---
bold = r"C:\Windows\Fonts\segoeuib.ttf"
regular = r"C:\Windows\Fonts\segoeui.ttf"

title_font = ImageFont.truetype(bold, 72)
sub_font = ImageFont.truetype(regular, 28)
tag_font = ImageFont.truetype(regular, 20)

draw.text((60, 110), "CleanSheet", font=title_font, fill=TEXT)
draw.text((66, 202), "Clean messy Excel & CSV files offline", font=sub_font, fill=MUTED)
draw.text((66, 242),
          "Rule-based  ·  never guesses  ·  Arabic • French • English  ·  openpyxl",
          font=tag_font, fill=MUTED)

# --- small metadata chips (artificial) ---
chips = ["MIT", "88 tests", "Python 3.8+", "Windows exe"]
cx = 60
cy = 300
chip_font = ImageFont.truetype(regular, 20)
for chip in chips:
    bbox = draw.textbbox((0, 0), chip, font=chip_font)
    cw = bbox[2] - bbox[0] + 24
    ch = 34
    draw.rounded_rectangle([(cx, cy), (cx + cw, cy + ch)], radius=17,
                           fill=PANEL, outline=GRID, width=1)
    draw.text((cx + 12, cy + 7), chip, font=chip_font, fill=MUTED)
    cx += cw + 14

# --- thin frame ---
draw.rectangle([(0, 0), (W - 1, H - 1)], outline=GRID, width=1)

out_path = os.path.join(OUT, "banner.png")
canvas.convert("RGB").save(out_path, optimize=True)
print("saved", out_path, canvas.size)
