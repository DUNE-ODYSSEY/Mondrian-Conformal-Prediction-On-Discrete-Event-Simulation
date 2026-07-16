"""
Builds the mid-sem review PPT from the real results in results/tables/.
Re-run any time those tables change: .venv\\Scripts\\python.exe slides\\build_mid_sem_ppt.py

Uses PowerPoint's default theme and native placeholders/bullets rather than
custom-drawn boxes for every slide - a hand-built student deck, not a
templated pitch deck.

Slide 8 (literature review snapshot) is a placeholder - the literature
review is being done separately and isn't finished, so faking category or
citation counts here would misrepresent progress.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

OUT_PATH = "slides/mid_sem_presentation.pptx"

TEAM = [
    ("G Venugopalan", "CB.AI.U4AID25115"),
    ("Vipin Sudhakar", "CB.AI.U4AID25166"),
    ("Rithvik Arulprakash", "CB.AI.U4AID25148"),
    ("Harshith Kv", "CB.AI.U4AID25119"),
]
COURSE_CODE = "23AID201"
PROJECT_TITLE = "Mondrian Conformal Prediction for Uncertainty Quantification\nin ER Discrete-Event Simulation Surrogates"

GREY = RGBColor(0x40, 0x40, 0x40)


MARGIN = Inches(0.5)
CONTENT_W = Inches(9.0)  # matches the default template's placeholder width (10in slide)


def new_deck():
    return Presentation()


def set_title(slide, text, size=None):
    slide.shapes.title.text = text
    if size:
        for p in slide.shapes.title.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(size)


def body_placeholder(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx != 0:
            return ph
    return None


def fill_bullets(placeholder, items):
    """items: list of str, or (str, level) tuples."""
    tf = placeholder.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        text, level = item if isinstance(item, tuple) else (item, 0)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.level = level


def add_title_only_slide(prs, title_text):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(slide, title_text)
    return slide


def add_bulleted_slide(prs, title_text, items, size=None):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    set_title(slide, title_text)
    ph = body_placeholder(slide)
    fill_bullets(ph, items)
    if size:
        for p in ph.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(size)
    return slide


def style_table(table, header_bold=True, font_size=13):
    for j in range(len(table.columns)):
        cell = table.cell(0, j)
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.bold = header_bold
                r.font.size = Pt(font_size)
    for i in range(1, len(table.rows)):
        for j in range(len(table.columns)):
            cell = table.cell(i, j)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT
                for r in p.runs:
                    r.font.size = Pt(font_size)


def make_table(slide, x, y, w, h, headers, rows, col_widths=None, font_size=13):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    gframe = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = gframe.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    for j, hdr in enumerate(headers):
        table.cell(0, j).text = hdr
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            table.cell(i + 1, j).text = str(val)
    style_table(table, font_size=font_size)
    return table


def build():
    prs = new_deck()

    # ---------------- Slide 1: Title ----------------
    s = prs.slides.add_slide(prs.slide_layouts[0])
    title_ph = s.shapes.title
    title_ph.text = PROJECT_TITLE
    title_ph.left, title_ph.top = MARGIN, Inches(0.4)
    title_ph.width, title_ph.height = CONTENT_W, Inches(2.6)
    for p in title_ph.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(26)

    subtitle = s.placeholders[1]
    subtitle.left, subtitle.top = MARGIN, Inches(3.25)
    subtitle.width, subtitle.height = CONTENT_W, Inches(3.8)
    lines = ["Mid-Semester Review", ""]
    lines += [f"{n} ({i})" for n, i in TEAM]
    lines += ["", f"Course: {COURSE_CODE}"]
    subtitle.text_frame.text = lines[0]
    for line in lines[1:]:
        p = subtitle.text_frame.add_paragraph()
        p.text = line
    for p in subtitle.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(16)

    # ---------------- Slide 2: Problem Statement ----------------
    add_bulleted_slide(prs, "Problem Statement", [
        "Surrogate models are fast, learned approximations of expensive simulations. They're used more and "
        "more to guide decisions in stochastic service systems - ER staffing, queueing networks, etc.",
        "A point prediction from a surrogate isn't enough on its own for a decision like \"will adding 2 more "
        "doctors actually bring the wait time down?\" You need a calibrated interval, not just a number.",
        "Two common ways to get that interval, and where they fall short:",
        ("Gaussian Processes - solid theoretically, but expensive to train (cubic cost) and the coverage "
         "guarantee depends on distributional assumptions, not a hard finite-sample bound.", 1),
        ("Conformal Prediction - distribution-free, comes with a finite-sample coverage guarantee, cheap to "
         "calibrate. But it's only been validated for surrogate models in a narrow set of domains so far.", 1),
        "That's the question this project is built around: does CP still hold up when the surrogate is "
        "trained on a discrete-event / queueing simulation instead of the domains it's been tested on?",
    ], size=17)

    # ---------------- Slide 3: Research Gap ----------------
    add_bulleted_slide(prs, "Research Gap", [
        "Gopakumar et al. (2026) validate conformal prediction for surrogate-model UQ, but only in physics "
        "domains - PDEs, magnetohydrodynamics, weather, fusion.",
        "They call out two limitations explicitly:",
        ("Marginal coverage only - no guarantee that coverage holds within a specific subgroup or condition, "
         "only on average across everything.", 1),
        ("An exchangeability assumption between calibration and test data, which they don't test under "
         "distribution shift.", 1),
        "Neither of those has been checked outside physics simulation. Discrete-event / queueing systems are "
        "a genuinely different kind of domain - discrete stochastic arrivals and departures, resource "
        "contention, priority queues - not a continuous PDE field. That's the gap this project sits in.",
    ], size=17)

    # ---------------- Slide 4: Bridging Approach ----------------
    s = prs.slides.add_slide(prs.slide_layouts[3])  # Two Content
    set_title(s, "Our Bridging Approach")

    note = s.shapes.add_textbox(MARGIN, Inches(1.35), CONTENT_W, Inches(0.9))
    note.text_frame.word_wrap = True
    note.text_frame.text = ("Core idea: apply CP, specifically Mondrian CP, to an ER queueing surrogate, and "
                              "test directly whether Gopakumar et al.'s two limitations hold here.")
    for r in note.text_frame.paragraphs[0].runs:
        r.font.size = Pt(15)
        r.font.italic = True
        r.font.color.rgb = GREY

    left, right = s.placeholders[1], s.placeholders[2]
    col_w = Inches(4.4157)
    for ph, x in ((left, MARGIN), (right, Inches(5.083))):
        ph.left = x
        ph.top = Inches(2.45)
        ph.width = col_w
        ph.height = Inches(4.6)
    fill_bullets(left, [
        ("Addresses marginal coverage", 0),
        ("Mondrian CP partitions calibration by category - staffing level, shift, arrival-rate regime - "
         "instead of lumping everything into one marginal guarantee.", 1),
        ("Goal: coverage that holds up close to nominal within each category, not just on average.", 1),
    ])
    fill_bullets(right, [
        ("Addresses exchangeability", 0),
        ("Phase 2 includes a stress test: an out-of-distribution \"surge day\" scenario outside the normal "
         "training range.", 1),
        ("That's where we actually get to see whether standard CP and Mondrian CP hold up or break when "
         "exchangeability is violated.", 1),
    ])
    for ph in (left, right):
        for p in ph.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(15) if p.level == 0 else Pt(13)

    # ---------------- Slide 5: Methodology Overview ----------------
    s = add_title_only_slide(prs, "Methodology Overview")
    y = Inches(2.3)
    h = Inches(1.3)
    bw = Inches(2.6)
    gap = Inches(0.6)
    x_positions = [MARGIN, MARGIN + bw + gap, MARGIN + 2 * (bw + gap)]
    labels = [
        ("Discrete-Event\nSimulation", "SimPy, calibrated on\nreal ED data"),
        ("Surrogate Model", "Gradient boosting on the\nDES scenario sweep"),
        ("Uncertainty\nQuantification", "GP baseline vs.\nstandard CP vs. Mondrian CP"),
    ]
    boxes = []
    for x, (label, sub) in zip(x_positions, labels):
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, bw, h)
        box.fill.background()
        box.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
        box.line.width = Pt(1.25)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = label
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = sub
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = GREY
        boxes.append(box)

    for a, b in zip(boxes[:-1], boxes[1:]):
        connector = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, a.left + a.width, a.top + a.height // 2,
                                             b.left, b.top + b.height // 2)
        connector.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
        connector.line.width = Pt(1.5)

    note_box = s.shapes.add_textbox(MARGIN, Inches(4.0), CONTENT_W, Inches(2.9))
    tf = note_box.text_frame
    tf.word_wrap = True
    tf.text = "Each stage is checked against the previous one before moving on:"
    tf.paragraphs[0].runs[0].font.italic = True
    tf.paragraphs[0].runs[0].font.size = Pt(14)
    checks = [
        "- DES output checked against real aggregated ED stats (91.0% match on daily volume)",
        "- Surrogate accuracy checked against DES outputs held out from training (MAE / RMSE / R²)",
        "- All UQ methods use the same fixed test split and the same target coverage (alpha = 0.1), so "
        "GP, standard CP and Mondrian CP end up directly comparable",
    ]
    for c in checks:
        p = tf.add_paragraph()
        p.text = c
        p.level = 0
        for r in p.runs:
            r.font.size = Pt(15)

    # ---------------- Slide 6: Real-world data used ----------------
    add_bulleted_slide(prs, "Real-World Data Used", [
        "Hospital Triage and Patient History Data, Kaggle (maalona) - Yale New Haven Health System, "
        "retrospective, March 2014 to July 2017.",
        "560,486 ED visits total, 972 variables per visit, across 3 EDs (1 academic + 2 community). We "
        "calibrate on Department A only, the largest / academic site: 322,283 visits, ~258.2/day.",
        "Real vs. literature - kept separate on purpose, not blended silently:",
        ("From the real data: arrival pattern by 4-hour bin, day of week, and month; ESI acuity mix; "
         "department-level daily visit rate.", 1),
        ("From literature: service/treatment time per ESI level (log-normal). The dataset has no discharge "
         "timestamp - checked across all 972 columns to be sure before falling back to literature values.", 1),
    ], size=16)

    # ---------------- Slide 7: Work Completed ----------------
    s = add_title_only_slide(prs, "Work Completed (~50%)")
    make_table(s, MARGIN, Inches(1.4), Inches(4.3), Inches(1.1),
               ["DES validation (200 sim. days)", ""],
               [
                   ["Real visits/day (Dept. A)", "258.2"],
                   ["Simulated mean visits/day", "235.1"],
                   ["Match", "91.0%"],
               ],
               col_widths=[Inches(2.8), Inches(1.5)], font_size=12)

    make_table(s, Inches(5.2), Inches(1.4), Inches(4.3), Inches(1.55),
               ["Surrogate target", "MAE", "RMSE", "R²"],
               [
                   ["n_patients", "9.90", "12.59", "0.929"],
                   ["mean_wait_min", "8.86", "13.23", "0.787"],
                   ["mean_total_min", "9.88", "13.61", "0.762"],
                   ["p95_wait_min", "66.94", "102.47", "0.647"],
               ],
               col_widths=[Inches(1.6), Inches(0.9), Inches(0.9), Inches(0.9)], font_size=11)

    make_table(s, MARGIN, Inches(3.3), CONTENT_W, Inches(1.7),
               ["GP baseline (alpha = 0.1)", "Target coverage", "Empirical coverage", "Mean interval width"],
               [
                   ["n_patients", "90%", "88.5%", "39.3"],
                   ["mean_wait_minutes", "90%", "87.7%", "43.6"],
                   ["mean_total_minutes", "90%", "88.8%", "43.4"],
                   ["p95_wait_minutes", "90%", "90.1%", "350.3"],
               ],
               col_widths=[Inches(2.6), Inches(2.1), Inches(2.1), Inches(2.2)], font_size=12)

    note = s.shapes.add_textbox(MARGIN, Inches(5.3), CONTENT_W, Inches(1.4))
    note.text_frame.word_wrap = True
    note.text_frame.text = ("GP undercovers on 3 of 4 targets (87.7-88.8% vs. a 90% target) - not surprising, "
                              "since it relies on distributional assumptions rather than a finite-sample "
                              "guarantee. Whether CP / Mondrian CP close that gap is what Phase 2 is for.")
    note.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
    note.text_frame.paragraphs[0].runs[0].font.italic = True

    # ---------------- Slide 8: Literature review snapshot (placeholder) ----------------
    s = add_bulleted_slide(prs, "Literature Review Snapshot", [
        "Still in progress - counts below to be filled in before the actual presentation.",
        ("Conformal prediction foundations: [ ] papers", 1),
        ("Surrogate modeling: [ ] papers", 1),
        ("Queueing theory / ED simulation: [ ] papers", 1),
        "Target: 30 papers total, with 3 core papers (including Gopakumar et al. 2026) reviewed in depth.",
    ], size=17)

    # ---------------- Slide 9: Future scope ----------------
    s = add_title_only_slide(prs, "Future Scope - End-Sem")
    make_table(s, MARGIN, Inches(1.5), CONTENT_W, Inches(4.7),
               ["Weeks", "Plan"],
               [
                   ["9-10", "Standard conformal prediction on surrogate residuals; try a few different "
                             "nonconformity measures."],
                   ["11-12", "Mondrian CP - partition by staffing level / shift / arrival-rate category; "
                              "measure per-category coverage."],
                   ["13", "Stress-test exchangeability with an out-of-distribution \"surge day\" scenario; "
                           "see where standard CP and Mondrian CP hold or break."],
                   ["14-15", "Full comparison - GP vs. standard CP vs. Mondrian CP on coverage, interval "
                              "width, and computation time."],
                   ["16", "End-sem PPT and report: a quantified answer on whether Mondrian CP closes the "
                           "marginal-vs-conditional coverage gap in this domain."],
               ],
               col_widths=[Inches(1.1), Inches(7.9)], font_size=13)

    prs.save(OUT_PATH)
    print(f"Saved {OUT_PATH}")


if __name__ == "__main__":
    build()
