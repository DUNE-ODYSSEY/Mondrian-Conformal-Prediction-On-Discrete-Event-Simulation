"""
Single-spread wraparound cover (back cover | spine | front cover), matching
the print-cover-template convention the professor's reference image used -
built by compositing the existing front_cover.png/back_cover.png (generated
by generate_book_covers.py) side by side with a spine panel in between,
rather than regenerating cover art from scratch. Saved as its own PDF in
reports/assignments/, separate from the book itself (which keeps the front
and back covers as its literal first/last pages, not a single spread).

Re-run: .venv\\Scripts\\python.exe reports\\assignments\\generate_book_wraparound_cover.py
"""
import os
from PIL import Image, ImageDraw, ImageFont

FIG_DIR = "reports/assignments/figures"
OUT_DIR = "reports/assignments"

DPI = 300
INK = (16, 25, 43)
INK_3 = (12, 19, 33)
CORAL = (232, 112, 58)
WARM_WHITE = (244, 240, 231)
MUTED_ON_INK = (162, 173, 192)

FONT_DIR = r"C:\Windows\Fonts"


def font(name, size):
    return ImageFont.truetype(os.path.join(FONT_DIR, name), size)


TITLE = "BEYOND MARGINAL GUARANTEES"


def build_spine(height, width, title=TITLE):
    spine = Image.new("RGB", (width, height), INK_3)
    d = ImageDraw.Draw(spine)

    # Title, rendered horizontally then rotated 90 deg clockwise so it
    # reads top-to-bottom on the spine (standard convention: tilt your
    # head left, or rotate the book clockwise, to read it normally).
    tsize = int(width * 0.6)
    tf = font("georgiab.ttf", tsize)
    tmp = Image.new("RGB", (10, 10))
    td = ImageDraw.Draw(tmp)
    text_w = td.textlength(title, font=tf)
    text_h = tsize * 1.3

    label = Image.new("RGB", (int(text_w) + 40, int(text_h)), INK_3)
    ld = ImageDraw.Draw(label)
    ld.text((20, int(text_h * 0.12)), title, font=tf, fill=WARM_WHITE)
    rotated = label.rotate(-90, expand=True)

    rx = (width - rotated.width) // 2
    ry = (height - rotated.height) // 2
    spine.paste(rotated, (rx, ry))

    # thin coral accent ticks top and bottom of the spine, echoing the
    # front cover's own accent rule
    tick_h = int(height * 0.012)
    d.rectangle([0, int(height * 0.06), width, int(height * 0.06) + 4], fill=CORAL)
    d.rectangle([0, height - int(height * 0.06), width, height - int(height * 0.06) + 4], fill=CORAL)

    return spine


def dashed_vline(draw, x, y0, y1, fill, dash=14, gap=10, width=2):
    y = y0
    while y < y1:
        draw.line([(x, y), (x, min(y + dash, y1))], fill=fill, width=width)
        y += dash + gap


def build_wraparound(back_name, front_name, title, out_stem):
    back = Image.open(f"{FIG_DIR}/{back_name}")
    front = Image.open(f"{FIG_DIR}/{front_name}")
    assert back.size == front.size
    w, h = back.size

    spine_w = round(1.3 / 2.54 * DPI)  # 1.3cm spine, a plausible width for a book this length
    spine = build_spine(h, spine_w, title=title)

    total_w = w + spine_w + w
    spread = Image.new("RGB", (total_w, h), INK)
    spread.paste(back, (0, 0))
    spread.paste(spine, (w, 0))
    spread.paste(front, (w + spine_w, 0))

    d = ImageDraw.Draw(spread)
    dashed_vline(d, w, 0, h, MUTED_ON_INK)
    dashed_vline(d, w + spine_w, 0, h, MUTED_ON_INK)

    png_path = f"{FIG_DIR}/{out_stem}.png"
    spread.save(png_path, dpi=(DPI, DPI))
    print("Saved", png_path)

    pdf_path = f"{OUT_DIR}/{out_stem}.pdf"
    spread.save(pdf_path, "PDF", resolution=DPI)
    print("Saved", pdf_path)


def main():
    build_wraparound("back_cover.png", "front_cover.png", "BEYOND MARGINAL GUARANTEES",
                      "book_cover_wraparound")
    build_wraparound("back_cover2.png", "front_cover2.png", "WHEN EXTRAPOLATION FAILS",
                      "book_cover_wraparound2")


if __name__ == "__main__":
    main()
