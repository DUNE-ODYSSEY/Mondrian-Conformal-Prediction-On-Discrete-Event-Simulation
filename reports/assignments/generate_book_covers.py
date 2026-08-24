"""
Front and back cover pages for the book (reports/assignments/book1.docx),
styled after a reference Kindle-book cover the professor liked (title block
+ technical diagram on the front; "BOOK OVERVIEW" + blurb + a supporting
diagram on the back; a dark author/institution bar closing both pages).

Built with Pillow rather than the docx/matplotlib toolchain used elsewhere in
this project, since a full-bleed illustrated page needs pixel-level control
over gradients and layout that python-docx shapes can't give. Every number
and sentence used on the back cover is drawn from the book's own Preface/
Chapter 6-7 content (68.2% -> 90.9% worst-category coverage, the staffing x
arrival-rate Mondrian grid, "five further methods") - nothing invented.
"""
import math
import os
import textwrap
from PIL import Image, ImageDraw, ImageFont

OUT = "reports/assignments/figures"
os.makedirs(OUT, exist_ok=True)

DPI = 300
PAGE_W_IN = 17.6 / 2.54
PAGE_H_IN = 25.0 / 2.54
W = round(PAGE_W_IN * DPI)
H = round(PAGE_H_IN * DPI)
MARGIN = round(0.085 * W)

# ---------------------------------------------------------------------------
# Palette - grounded in the book's own subject: emergency-department triage
# (coral, the universal "alert/urgent" color) vs. clinical reliability (a
# cool teal), on an ink-navy ground rather than a generic black or the
# purple/blue gradient AI-generated covers default to.
# ---------------------------------------------------------------------------
INK = (16, 25, 43)
INK_2 = (23, 35, 57)
INK_3 = (12, 19, 33)
CORAL = (232, 112, 58)
TEAL = (58, 173, 165)
PARCHMENT = (244, 238, 226)
WARM_WHITE = (244, 240, 231)
MUTED_ON_INK = (162, 173, 192)
INK_TEXT = (24, 28, 22)
MUTED_ON_PARCHMENT = (110, 101, 84)

FONT_DIR = r"C:\Windows\Fonts"


def font(name, size):
    return ImageFont.truetype(os.path.join(FONT_DIR, name), size)


def title_font(s):
    return font("georgiab.ttf", s)


def italic_font(s):
    return font("georgiai.ttf", s)


def serif_font(s):
    return font("georgia.ttf", s)


def sans_font(s):
    return font("segoeui.ttf", s)


def sans_bold(s):
    return font("segoeuib.ttf", s)


def vgradient(w, h, top, bottom):
    col = Image.new("RGB", (1, h))
    px = col.load()
    for y in range(h):
        t = y / max(h - 1, 1)
        px[0, y] = tuple(round(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
    return col.resize((w, h))


def tracked_text(draw, pos, text, fnt, fill, tracking=0):
    x, y = pos
    for ch in text:
        draw.text((x, y), ch, font=fnt, fill=fill)
        x += draw.textlength(ch, font=fnt) + tracking
    return x


def wrap_to_width(draw, text, fnt, max_w):
    words = text.split()
    lines, cur = [], ""
    for w_ in words:
        trial = (cur + " " + w_).strip()
        if draw.textlength(trial, font=fnt) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w_
    if cur:
        lines.append(cur)
    return lines


def bottom_bar(draw, w, h, bar_h, line1, line2, line1_track=3, line2_track=1):
    y0 = h - bar_h
    draw.rectangle([0, y0, w, h], fill=INK_3)
    draw.rectangle([0, y0, w, y0 + 3], fill=CORAL)
    max_w = w - 2 * MARGIN
    tmp = Image.new("RGB", (10, 10))
    td = ImageDraw.Draw(tmp)

    def tracked_width(text, fnt, tracking):
        return sum(td.textlength(ch, font=fnt) + tracking for ch in text) - tracking

    size1 = int(bar_h * 0.20)
    f1 = sans_font(size1)
    while tracked_width(line1, f1, line1_track) > max_w and size1 > 10:
        size1 -= 1
        f1 = sans_font(size1)

    w1 = tracked_width(line1, f1, line1_track)
    x1 = (w - w1) / 2
    y1 = y0 + bar_h * 0.30
    tracked_text(draw, (x1, y1), line1, f1, WARM_WHITE, tracking=line1_track)

    size2 = int(bar_h * 0.165)
    f2 = italic_font(size2)
    while td.textlength(line2, font=f2) > max_w and size2 > 10:
        size2 -= 1
        f2 = italic_font(size2)

    w2 = draw.textlength(line2, font=f2)
    x2 = (w - w2) / 2
    y2 = y0 + bar_h * 0.60
    draw.text((x2, y2), line2, font=f2, fill=MUTED_ON_INK)


# ---------------------------------------------------------------------------
# Pipeline icons - flat, line-drawn glyphs (no photos/stock art available in
# this toolchain) for the 5-stage methodology row: ED -> DES -> surrogate ->
# UQ -> Mondrian CP, matching Chapters 4-6's actual pipeline.
# ---------------------------------------------------------------------------
def icon_hospital(d, cx, cy, s):
    bw, bh = s * 1.3, s * 1.05
    x0, y0 = cx - bw / 2, cy - bh / 2 + s * 0.12
    d.rounded_rectangle([x0, y0, x0 + bw, y0 + bh], radius=s * 0.07, outline=WARM_WHITE, width=3)
    arm = s * 0.16
    d.rectangle([cx - arm * 0.28, cy - arm, cx + arm * 0.28, cy + arm], fill=CORAL)
    d.rectangle([cx - arm, cy - arm * 0.28, cx + arm, cy + arm * 0.28], fill=CORAL)


def icon_queue_server(d, cx, cy, s):
    r = s * 0.09
    xs = cx - s * 0.55
    for dy in (-0.35, 0, 0.35):
        d.ellipse([xs - r, cy + s * dy - r, xs + r, cy + s * dy + r], fill=MUTED_ON_INK)
    d.line([(xs + r + 4, cy), (cx + s * 0.05, cy)], fill=MUTED_ON_INK, width=2)
    bw, bh = s * 0.55, s * 0.75
    d.rounded_rectangle([cx + s * 0.12, cy - bh / 2, cx + s * 0.12 + bw, cy + bh / 2],
                         radius=s * 0.06, outline=TEAL, width=3)


def icon_network(d, cx, cy, s):
    n = 6
    pts = [(cx + math.cos(2 * math.pi * i / n) * s * 0.5,
             cy + math.sin(2 * math.pi * i / n) * s * 0.5) for i in range(n)]
    center = (cx, cy)
    for p in pts:
        d.line([center, p], fill=(52, 68, 92), width=2)
    for i in range(n):
        d.line([pts[i], pts[(i + 1) % n]], fill=(40, 54, 74), width=1)
    for p in pts:
        rr = s * 0.06
        d.ellipse([p[0] - rr, p[1] - rr, p[0] + rr, p[1] + rr], fill=TEAL)
    rr = s * 0.075
    d.ellipse([cx - rr, cy - rr, cx + rr, cy + rr], fill=WARM_WHITE)


def icon_band(d, cx, cy, s):
    n = 24
    upper, lower, mid = [], [], []
    for i in range(n + 1):
        t = i / n
        x = cx - s * 0.6 + t * s * 1.2
        base = cy + math.sin(t * math.pi + 0.3) * -s * 0.22
        band = s * 0.10 + t * s * 0.16
        upper.append((x, base - band))
        lower.append((x, base + band))
        mid.append((x, base))
    d.polygon(upper + lower[::-1], fill=(30, 52, 60))
    d.line(mid, fill=TEAL, width=3)


def icon_grid_warning(d, cx, cy, s):
    cell = s * 0.32
    gp = s * 0.035
    x0 = cx - 1.5 * cell - gp
    y0 = cy - 1.5 * cell - gp
    warn_cell = (0, 2)
    for r in range(3):
        for c in range(3):
            xx = x0 + c * (cell + gp)
            yy = y0 + r * (cell + gp)
            d.rectangle([xx, yy, xx + cell, yy + cell], outline=MUTED_ON_INK, width=2)
            if (r, c) == warn_cell:
                tcx, tcy = xx + cell / 2, yy + cell / 2
                tsz = cell * 0.36
                d.polygon([(tcx, tcy - tsz), (tcx - tsz, tcy + tsz * 0.8), (tcx + tsz, tcy + tsz * 0.8)],
                          outline=CORAL, width=2)
                d.text((tcx - tsz * 0.09, tcy - tsz * 0.55), "!",
                       font=sans_bold(round(tsz * 1.15)), fill=CORAL)


PIPELINE_STEPS = [
    ("1", "EMERGENCY DEPARTMENT", "Stochastic arrivals & triage acuity", icon_hospital),
    ("2", "DISCRETE-EVENT SIMULATION", "Queueing dynamics & performance metrics", icon_queue_server),
    ("3", "SURROGATE MODEL", "Gradient-boosting regressor", icon_network),
    ("4", "UNCERTAINTY QUANTIFICATION", "Prediction intervals with a coverage guarantee", icon_band),
    ("5", "MONDRIAN CONFORMAL PREDICTION", "Conditional coverage within each regime", icon_grid_warning),
]


# ---------------------------------------------------------------------------
# FRONT COVER
# ---------------------------------------------------------------------------
def build_front_cover(title, subtitle, team_names, faculty_guide):
    img = vgradient(W, H, INK, INK_2)
    d = ImageDraw.Draw(img)

    bar_h = round(0.105 * H)
    content_w = W - 2 * MARGIN

    def centered(text, fnt, y, fill, tracking=0):
        if tracking:
            w = sum(d.textlength(ch, font=fnt) + tracking for ch in text) - tracking
            tracked_text(d, ((W - w) / 2, y), text, fnt, fill, tracking=tracking)
        else:
            w = d.textlength(text, font=fnt)
            d.text(((W - w) / 2, y), text, font=fnt, fill=fill)

    # Title block (centered poster treatment)
    ty = round(0.10 * H)
    tsize = round(H * 0.052)
    tf = title_font(tsize)
    for ln in wrap_to_width(d, title, tf, content_w * 0.92):
        centered(ln, tf, ty, WARM_WHITE)
        ty += round(tsize * 1.1)

    ty += round(H * 0.014)
    ssize = round(H * 0.0195)
    sf = italic_font(ssize)
    for ln in wrap_to_width(d, subtitle, sf, content_w * 0.72):
        centered(ln, sf, ty, MUTED_ON_INK)
        ty += round(ssize * 1.4)

    ty += round(H * 0.018)
    rule_w = round(content_w * 0.14)
    d.line([(W / 2 - rule_w / 2, ty), (W / 2 + rule_w / 2, ty)], fill=CORAL, width=4)

    # --- The 5-stage pipeline this book runs, start to finish - the actual
    # methodology (Chapters 4-6), not a spoiler of Chapter 7's result. Anchored
    # to a fixed fraction of the page (not chained off ty) so it sits centered
    # in the lower half regardless of how many lines the title/subtitle wrap to.
    dy = round(0.62 * H)
    centered("THE PIPELINE THIS BOOK BUILDS", sans_font(round(H * 0.0135)), dy, CORAL, tracking=2)
    dy += round(H * 0.06)

    n = len(PIPELINE_STEPS)
    col_w = content_w / n
    icon_s = col_w * 0.42
    icon_cy = dy + icon_s * 0.75

    name_f = sans_bold(round(H * 0.0102))
    cap_f = sans_font(round(H * 0.0094))
    name_line_h = round(H * 0.0158)
    cap_line_h = round(H * 0.0138)
    text_w = col_w * 0.88

    # Pass 1: wrap every column's title/caption to its own column width, so
    # a long phrase can never bleed into a neighboring column - then take
    # the tallest wrap across all columns as a shared reserved height, so
    # every column's caption starts on the same baseline regardless of how
    # many lines its own title happened to need.
    wrapped = []
    for num, name, caption, icon_fn in PIPELINE_STEPS:
        name_lines = wrap_to_width(d, name, name_f, text_w)
        cap_lines = wrap_to_width(d, caption, cap_f, text_w)
        wrapped.append((num, name_lines, cap_lines, icon_fn))
    max_name_lines = max(len(w[1]) for w in wrapped)
    max_cap_lines = max(len(w[2]) for w in wrapped)

    # Explicit, additive vertical stack (icon -> badge -> title) rather than
    # offsets measured backward from the title - the earlier version placed
    # the badge by subtracting from the title's y, which put it well inside
    # the icon's own footprint (icons extend up to ~0.7 * icon_s below their
    # center) and produced a visible number/diagram overlap on page 1.
    badge_r = round(H * 0.013)
    icon_bottom = icon_cy + icon_s * 0.72
    bcy = icon_bottom + badge_r + round(H * 0.012)
    label_y = bcy + badge_r + round(H * 0.016)

    for i, (num, name_lines, cap_lines, icon_fn) in enumerate(wrapped):
        ccx = MARGIN + col_w * (i + 0.5)
        icon_fn(d, ccx, icon_cy, icon_s)

        if i < n - 1:
            chev_y = icon_cy
            chev_x = MARGIN + col_w * (i + 1)
            cs = round(H * 0.008)
            d.line([(chev_x - cs, chev_y - cs), (chev_x, chev_y)], fill=(70, 84, 106), width=3)
            d.line([(chev_x - cs, chev_y + cs), (chev_x, chev_y)], fill=(70, 84, 106), width=3)

        # step-number badge, between the icon and its title
        d.ellipse([ccx - badge_r, bcy - badge_r, ccx + badge_r, bcy + badge_r], outline=CORAL, width=2)
        nf = sans_bold(round(badge_r * 1.15))
        nw = d.textlength(num, font=nf)
        d.text((ccx - nw / 2, bcy - badge_r * 0.72), num, font=nf, fill=CORAL)

        yy = label_y + (max_name_lines - len(name_lines)) * name_line_h / 2
        for ln in name_lines:
            w_ = d.textlength(ln, font=name_f)
            d.text((ccx - w_ / 2, yy), ln, font=name_f, fill=WARM_WHITE)
            yy += name_line_h
        yy = label_y + max_name_lines * name_line_h + round(H * 0.006)
        yy += (max_cap_lines - len(cap_lines)) * cap_line_h / 2
        for ln in cap_lines:
            w_ = d.textlength(ln, font=cap_f)
            d.text((ccx - w_ / 2, yy), ln, font=cap_f, fill=MUTED_ON_INK)
            yy += cap_line_h

    bottom_bar(d, W, H, bar_h,
               "  ·  ".join(nm.upper() for nm in team_names),
               f"Faculty Guide: {faculty_guide}")

    path = f"{OUT}/front_cover.png"
    img.save(path, dpi=(DPI, DPI))
    print("Saved", path)
    return path


# ---------------------------------------------------------------------------
# BACK COVER
# ---------------------------------------------------------------------------
def build_back_cover():
    img = Image.new("RGB", (W, H), PARCHMENT)
    d = ImageDraw.Draw(img)

    content_w = W - 2 * MARGIN

    ov_f = sans_font(round(H * 0.0125))
    y = round(0.052 * H)
    tracked_text(d, (MARGIN, y), "MONDRIAN CONFORMAL PREDICTION  ·  DISCRETE-EVENT SIMULATION",
                 ov_f, MUTED_ON_PARCHMENT, tracking=3)

    hy = round(0.088 * H)
    hf = title_font(round(H * 0.046))
    d.text((MARGIN, hy), "Book Overview", font=hf, fill=INK)
    hy += round(H * 0.058)
    d.line([(MARGIN, hy), (MARGIN + round(content_w * 0.14), hy)], fill=CORAL, width=5)
    hy += round(H * 0.026)

    body_f = serif_font(round(H * 0.0168))
    line_h = round(H * 0.0168 * 1.5)
    paragraphs = [
        "Conformal prediction offers a distribution-free way to attach a reliable "
        "uncertainty guarantee to a machine-learned predictor \u2013 but the standard "
        "guarantee holds only on average, with no promise for any specific operating "
        "condition. This book tests that limitation in a domain where discrete, "
        "queueing-driven dynamics differ structurally from where the technique was "
        "first validated: a discrete-event simulation of a hospital emergency "
        "department, calibrated on real arrival and triage-acuity data.",
        "A gradient-boosting surrogate stands in for the simulator across a grid of "
        "staffing and arrival-rate regimes. Standard conformal prediction covers only "
        "68.2% of the worst-affected regime against a 90% target. Mondrian conformal "
        "prediction \u2013 calibrating separately within each staffing \u00d7 arrival-rate "
        "category rather than pooling \u2013 closes that gap to 90.9%, without disturbing "
        "the calibration set's exchangeability guarantee.",
        "Five further extensions, from weighted calibration to a queueing-theoretic "
        "continuous alternative, follow \u2013 evaluated with the same honesty, including "
        "where they fall short.",
    ]
    for para in paragraphs:
        for ln in wrap_to_width(d, para, body_f, content_w * 0.94):
            d.text((MARGIN, hy), ln, font=body_f, fill=INK_TEXT)
            hy += line_h
        hy += round(line_h * 0.4)

    # --- Supporting diagram: the Mondrian staffing x arrival-rate grid ---
    hy += round(H * 0.02)
    gf = sans_font(round(H * 0.0128))
    d.text((MARGIN, hy), "THE MONDRIAN PARTITION", font=gf, fill=CORAL)
    hy += round(H * 0.034)

    axis_f = sans_font(round(H * 0.0108))
    grid_size = round(H * 0.026)
    gap = round(H * 0.004)
    grid_x0 = MARGIN
    row_label_x = grid_x0 + 3 * (grid_size + gap) + round(H * 0.014)

    d.text((MARGIN, hy), "Arrival-rate tercile \u2192", font=axis_f, fill=MUTED_ON_PARCHMENT)
    d.text((row_label_x, hy), "Staffing tercile \u2193", font=axis_f, fill=MUTED_ON_PARCHMENT)
    hy += round(H * 0.024)

    labels = ["Low", "Med", "High"]
    for c in range(3):
        x0 = grid_x0 + c * (grid_size + gap)
        d.text((x0 + grid_size * 0.18, hy), labels[c], font=axis_f, fill=INK_TEXT)
    hy += round(H * 0.022)

    grid_y0 = hy
    worst_cell = (0, 2)  # (staffing row, arrival col): Low staffing x High arrival
    for r in range(3):
        for c in range(3):
            x0 = grid_x0 + c * (grid_size + gap)
            y0 = grid_y0 + r * (grid_size + gap)
            fill = CORAL if (r, c) == worst_cell else (219, 210, 191)
            d.rectangle([x0, y0, x0 + grid_size, y0 + grid_size], fill=fill, outline=INK_TEXT, width=2)
    for r in range(3):
        y0 = grid_y0 + r * (grid_size + gap)
        d.text((row_label_x, y0 + grid_size * 0.22), labels[r], font=axis_f, fill=INK_TEXT)

    cap_f = sans_font(round(H * 0.0108))
    cap_y = grid_y0 + 3 * (grid_size + gap) + round(H * 0.016)
    d.rectangle([grid_x0, cap_y, grid_x0 + round(H * 0.01), cap_y + round(H * 0.01)], fill=CORAL)
    d.text((grid_x0 + round(H * 0.017), cap_y - round(H * 0.001)),
            "Understaffed + high-arrival: the worst-covered category (68.2% \u2192 90.9%)",
            font=cap_f, fill=MUTED_ON_PARCHMENT)

    content_bottom = cap_y + round(H * 0.02)
    if content_bottom > H - MARGIN:
        print(f"WARNING: back cover content bottom {content_bottom} exceeds page bottom margin {H - MARGIN}")

    path = f"{OUT}/back_cover.png"
    img.save(path, dpi=(DPI, DPI))
    print("Saved", path)
    return path


if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    import book_common as bc

    build_front_cover(
        "Beyond Marginal Guarantees",
        "Mondrian Conformal Prediction for High-Variance Discrete-Event Queueing Systems",
        [n for n, _ in bc.TEAM],
        bc.FACULTY_GUIDE,
    )
    build_back_cover()
