"""
Zeroth review PPT - short, simple: abstract, problem statement, literature
review (categories only, no fabricated citations - literature review is
still in progress), research gap, novelty, aim/objectives, methodology,
roadmap. Not the mid-sem or full-results deck - this is the earlier
proposal-stage review.

Re-run: .venv\\Scripts\\python.exe slides\\build_zeroth_review_ppt.py

Same polished custom-drawn design as build_full_results_ppt.py (user's
stated preference for this project's decks).
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUT_PATH = "slides/zeroth_review.pptx"

TEAM = [
    ("G Venugopalan", "CB.AI.U4AID25115"),
    ("Vipin Sudhakar", "CB.AI.U4AID25166"),
    ("Rithvik Arulprakash", "CB.AI.U4AID25148"),
    ("Harshith Kv", "CB.AI.U4AID25119"),
]
COURSE_CODE = "23AID201"
PROJECT_TITLE = "Mondrian Conformal Prediction for Uncertainty\nQuantification in ER Discrete-Event Simulation Surrogates"

DARK = RGBColor(0x1B, 0x22, 0x38)
ACCENT = RGBColor(0x3B, 0x82, 0xF6)
ACCENT_DARK = RGBColor(0x1E, 0x40, 0xAF)
TEXT_DARK = RGBColor(0x1F, 0x29, 0x37)
TEXT_MUTED = RGBColor(0x6B, 0x72, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BG = RGBColor(0xF8, 0xFA, 0xFC)
BORDER = RGBColor(0xE2, 0xE8, 0xF0)
GOOD = RGBColor(0x16, 0xA3, 0x4A)
WARN = RGBColor(0xD9, 0x77, 0x06)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
TOTAL_SLIDES = 9


def new_deck():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_rect(slide, x, y, w, h, color, line=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line:
        shape.line.color.rgb = BORDER
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, x, y, w, h, text, size=18, bold=False, color=TEXT_DARK,
             align=PP_ALIGN.LEFT, font="Calibri", anchor=MSO_ANCHOR.TOP, italic=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = font
    return box


def add_bullets(slide, x, y, w, h, items, size=16, color=TEXT_DARK, space_after=10, font="Calibri"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        text, level = item if isinstance(item, tuple) else (item, 0)
        p.level = level
        bullet = "•  " if level == 0 else "‒  "
        run = p.add_run()
        run.text = bullet + text
        run.font.size = Pt(size - level * 2)
        run.font.color.rgb = color
        run.font.name = font
    return box


def slide_header(slide, kicker, title, num, total=TOTAL_SLIDES, title_size=28):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.09), ACCENT)
    add_text(slide, Inches(0.6), Inches(0.3), Inches(11.5), Inches(0.35),
              kicker.upper(), size=13, bold=True, color=ACCENT, font="Calibri")
    add_text(slide, Inches(0.6), Inches(0.62), Inches(11.9), Inches(0.7),
              title, size=title_size, bold=True, color=DARK, font="Calibri")
    add_rect(slide, Inches(0.6), Inches(1.28), Inches(1.1), Pt(3), ACCENT)
    add_text(slide, Inches(12.2), Inches(7.05), Inches(0.9), Inches(0.35),
              f"{num} / {total}", size=11, color=TEXT_MUTED, align=PP_ALIGN.RIGHT)


def build():
    prs = new_deck()
    n = [0]

    def nn():
        n[0] += 1
        return n[0]

    # ---------------- Slide 1: Title ----------------
    s = blank_slide(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, DARK)
    add_rect(s, 0, Inches(6.55), SLIDE_W, Inches(0.12), ACCENT)
    add_text(s, Inches(1), Inches(1.3), Inches(11.3), Inches(2.0),
              PROJECT_TITLE, size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(3.35), Inches(11.3), Inches(0.5),
              "Zeroth Review", size=18, color=RGBColor(0x93, 0xC5, 0xFD), align=PP_ALIGN.CENTER, italic=True)
    team_text = "   |   ".join([f"{nm} ({i})" for nm, i in TEAM])
    add_text(s, Inches(1), Inches(5.35), Inches(11.3), Inches(0.8),
              team_text, size=13, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(5.85), Inches(11.3), Inches(0.4),
              f"Course: {COURSE_CODE}", size=13, color=RGBColor(0x93, 0xC5, 0xFD), align=PP_ALIGN.CENTER)
    add_text(s, Inches(12.2), Inches(7.05), Inches(0.9), Inches(0.35),
              f"{nn()} / {TOTAL_SLIDES}", size=11, color=RGBColor(0x93, 0xC5, 0xFD), align=PP_ALIGN.RIGHT)

    # ---------------- Slide 2: Abstract ----------------
    s = blank_slide(prs)
    slide_header(s, "Overview", "Abstract", nn())
    add_rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(4.8), LIGHT_BG, line=True)
    add_text(s, Inches(1.0), Inches(2.0), Inches(11.3), Inches(4.2),
              "Surrogate models are increasingly used to approximate expensive simulations and guide "
              "decisions in stochastic service systems such as emergency room staffing. However, reliable "
              "uncertainty quantification (UQ) for these surrogates remains an open problem: Gaussian "
              "Processes are costly and rely on distributional assumptions, while Conformal Prediction (CP) "
              "offers a distribution-free, finite-sample coverage guarantee but has only been validated for "
              "surrogate models in physics domains (PDEs, magnetohydrodynamics, weather, fusion).\n\n"
              "This project builds a discrete-event simulation (DES) of an emergency room, calibrated on "
              "real hospital data, trains a surrogate model on it, and applies standard and Mondrian "
              "Conformal Prediction to quantify uncertainty in the surrogate's predictions. The goal is to "
              "test whether two limitations explicitly noted in prior work — marginal-only coverage and an "
              "untested exchangeability assumption — hold up in a discrete-event / queueing domain that has "
              "not been studied before in this context.",
              size=16, color=TEXT_DARK)

    # ---------------- Slide 3: Problem Statement ----------------
    s = blank_slide(prs)
    slide_header(s, "Motivation", "Problem Statement", nn())
    add_bullets(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(5), [
        "Surrogate models are fast, learned approximations of expensive simulations, increasingly used "
        "to guide decisions in stochastic service systems - ER staffing, queueing networks, etc.",
        "A point prediction alone isn't enough for a decision like \"will adding 2 more doctors bring "
        "wait time down enough?\" - that needs a calibrated interval, not just a number.",
        "Gaussian Processes are the established way to get one, but are expensive to train and rely on "
        "distributional assumptions rather than a finite-sample guarantee.",
        "Conformal Prediction (CP) is distribution-free with a finite-sample coverage guarantee and is "
        "cheap to calibrate - but it's only been validated for surrogate models in a narrow set of "
        "domains so far.",
    ], size=18, space_after=16)

    # ---------------- Slide 4: Literature Review ----------------
    s = blank_slide(prs)
    slide_header(s, "Background", "Literature Review", nn())
    add_text(s, Inches(0.7), Inches(1.65), Inches(11.9), Inches(0.6),
              "Review is organized around the base paper and three supporting research areas:",
              size=16, color=TEXT_DARK)

    add_rect(s, Inches(0.7), Inches(2.3), Inches(11.9), Inches(1.0), LIGHT_BG, line=True)
    add_text(s, Inches(1.0), Inches(2.45), Inches(11.3), Inches(0.75),
              "Base paper: Gopakumar et al. (2026) validate Conformal Prediction for surrogate-model UQ "
              "in physics domains (PDEs, MHD, weather, fusion), and explicitly flag two untested "
              "limitations - marginal coverage and the exchangeability assumption.",
              size=14, color=TEXT_DARK)

    cats = [
        ("Conformal Prediction\nFoundations", "Split conformal prediction, coverage guarantees, and "
         "category-conditional (Mondrian) variants of CP."),
        ("Surrogate Modeling", "Learned approximations (gradient boosting, GP, neural nets) for "
         "expensive simulations, and their use in decision-making."),
        ("Queueing Theory /\nDiscrete-Event Simulation", "Discrete-event simulation of service systems, "
         "particularly emergency department queueing and staffing models."),
    ]
    cw = Inches(3.75)
    for i, (cat, desc) in enumerate(cats):
        cx = Inches(0.7) + i * (cw + Inches(0.15))
        add_rect(s, cx, Inches(3.55), cw, Inches(2.9), WHITE, line=True)
        add_rect(s, cx, Inches(3.55), cw, Inches(0.75), ACCENT_DARK)
        add_text(s, cx + Inches(0.15), Inches(3.65), cw - Inches(0.3), Inches(0.6),
                  cat, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, cx + Inches(0.2), Inches(4.5), cw - Inches(0.4), Inches(1.8),
                  desc, size=12.5, color=TEXT_DARK)

    add_text(s, Inches(0.7), Inches(6.65), Inches(11.9), Inches(0.5),
              "Target: 30 papers total, 3 core papers reviewed in depth. In progress.",
              size=13, italic=True, color=TEXT_MUTED)

    # ---------------- Slide 5: Research Gap ----------------
    s = blank_slide(prs)
    slide_header(s, "Literature", "Research Gap", nn())
    add_text(s, Inches(0.7), Inches(1.65), Inches(11.9), Inches(0.5),
              "Gopakumar et al. (2026) validate Conformal Prediction for surrogate-model UQ —", size=18, bold=True)
    add_text(s, Inches(0.7), Inches(2.1), Inches(11.9), Inches(0.5),
              "but only in physics domains: PDEs, magnetohydrodynamics (MHD), weather, fusion.", size=18, color=ACCENT_DARK, bold=True)
    add_rect(s, Inches(0.7), Inches(2.95), Inches(11.9), Inches(2.0), LIGHT_BG, line=True)
    add_text(s, Inches(1.0), Inches(3.15), Inches(11.3), Inches(0.4),
              "The paper states two explicit limitations:", size=15, bold=True, color=TEXT_DARK)
    add_bullets(s, Inches(1.0), Inches(3.65), Inches(11.2), Inches(1.2), [
        "Marginal coverage only — no guarantee of coverage within specific subgroups/conditions.",
        "Relies on an exchangeability assumption between calibration and test data — untested under "
        "distribution shift.",
    ], size=15, space_after=10)
    add_text(s, Inches(0.7), Inches(5.25), Inches(11.9), Inches(1.4),
              "Neither limitation has been tested outside physics simulation. Discrete-event / queueing "
              "systems are a fundamentally different domain — stochastic discrete arrivals and departures, "
              "resource contention, priority scheduling — not continuous PDE fields. This is the gap this "
              "project addresses.", size=16, color=TEXT_DARK)

    # ---------------- Slide 6: Bridging the Gap / Novelty ----------------
    s = blank_slide(prs)
    slide_header(s, "Contribution", "Bridging the Gap: Our Novelty", nn(), title_size=26)
    add_bullets(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(4.8), [
        "Apply Conformal Prediction — standard and Mondrian CP — to a surrogate model of a discrete-event "
        "ER queueing simulation, a domain never tested against these limitations before.",
        ("Marginal coverage limitation: addressed with Mondrian CP, which calibrates separately per "
         "category (staffing level x arrival-rate regime) instead of one pooled quantile — targets "
         "conditional coverage, not just an average.", 1),
        ("Exchangeability limitation: addressed with a stress test that pushes the evaluation data "
         "outside the training distribution (an out-of-distribution \"surge day\" scenario) to see where "
         "the coverage guarantee actually breaks down.", 1),
        "Novelty: this is the first empirical test of Conformal Prediction's stated physics-domain "
        "limitations in a discrete-event / queueing simulation setting — a direct bridge between the base "
        "paper's open questions and a new application domain.",
    ], size=17, space_after=14)

    # ---------------- Slide 7: Aim & Objectives ----------------
    s = blank_slide(prs)
    slide_header(s, "Goal", "Aim & Objectives", nn())
    add_rect(s, Inches(0.7), Inches(1.65), Inches(11.9), Inches(1.1), LIGHT_BG, line=True)
    add_text(s, Inches(1.0), Inches(1.8), Inches(11.3), Inches(0.85),
              "Aim: to evaluate whether Conformal Prediction — specifically Mondrian CP — provides "
              "reliable, well-calibrated uncertainty quantification for surrogate models of discrete-event "
              "simulations, and whether Gopakumar et al.'s stated physics-domain limitations hold in this "
              "new domain.", size=15, bold=True, color=TEXT_DARK)
    add_bullets(s, Inches(0.7), Inches(3.05), Inches(11.9), Inches(4), [
        "Build and validate a discrete-event ER simulation calibrated on real hospital arrival data.",
        "Train a surrogate model to approximate the simulation's outputs.",
        "Implement a Gaussian Process baseline and standard / Mondrian Conformal Prediction for UQ.",
        "Test marginal vs. conditional coverage — does Mondrian CP close the gap standard CP leaves open?",
        "Stress-test the exchangeability assumption under distribution shift.",
        "Compare all methods on coverage, interval width, and computational cost.",
    ], size=16, space_after=10)

    # ---------------- Slide 8: Methodology / Workflow ----------------
    s = blank_slide(prs)
    slide_header(s, "Approach", "Methodology / Workflow", nn())
    y = Inches(2.1)
    h = Inches(1.3)
    bw = Inches(3.6)
    gap = Inches(0.45)
    x_positions = [Inches(0.55), Inches(0.55) + bw + gap, Inches(0.55) + 2 * (bw + gap)]
    labels = [
        ("Discrete-Event\nSimulation", "SimPy, calibrated on\nreal ED data"),
        ("Surrogate Model", "Learned approximation of\nthe simulation"),
        ("Uncertainty\nQuantification", "GP baseline vs. standard CP\nvs. Mondrian CP"),
    ]
    for x, (label, sub) in zip(x_positions, labels):
        box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, bw, h)
        box.adjustments[0] = 0.08
        box.fill.solid()
        box.fill.fore_color.rgb = ACCENT_DARK
        box.line.fill.background()
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = label
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = WHITE
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = sub
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = WHITE
    for x in x_positions[:-1]:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + bw, y + h / 2 - Inches(0.2), gap, Inches(0.4))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = TEXT_MUTED
        arrow.line.fill.background()
        arrow.shadow.inherit = False

    add_text(s, Inches(0.55), Inches(4.0), Inches(12.2), Inches(0.4),
              "Each stage validated against the previous before moving on:", size=14, italic=True, color=TEXT_MUTED)
    add_bullets(s, Inches(0.7), Inches(4.45), Inches(11.7), Inches(2.4), [
        "DES output validated against real aggregated ED statistics before being trusted as a data source.",
        "Surrogate accuracy validated against DES outputs held out from training.",
        "Every UQ method evaluated on the same test data and the same target coverage level, so results "
        "are directly comparable across methods.",
    ], size=15, space_after=10)

    # ---------------- Slide 9: Roadmap ----------------
    s = blank_slide(prs)
    slide_header(s, "Plan", "Project Roadmap", nn())
    col_w = Inches(5.75)
    add_rect(s, Inches(0.7), Inches(1.65), col_w, Inches(5.0), WHITE, line=True)
    add_rect(s, Inches(0.7), Inches(1.65), col_w, Inches(0.55), ACCENT_DARK)
    add_text(s, Inches(0.95), Inches(1.77), col_w - Inches(0.5), Inches(0.4),
              "Phase 1 — Mid-Semester", size=15, bold=True, color=WHITE)
    add_bullets(s, Inches(0.95), Inches(2.4), col_w - Inches(0.5), Inches(4.1), [
        "Literature review + environment setup",
        "Explore real hospital dataset, extract calibration distributions",
        "Build and validate the discrete-event ER simulation",
        "Generate training data and train the surrogate model",
        "Implement Gaussian Process baseline for UQ",
    ], size=14, space_after=13)

    add_rect(s, Inches(6.9), Inches(1.65), col_w, Inches(5.0), WHITE, line=True)
    add_rect(s, Inches(6.9), Inches(1.65), col_w, Inches(0.55), ACCENT_DARK)
    add_text(s, Inches(7.15), Inches(1.77), col_w - Inches(0.5), Inches(0.4),
              "Phase 2 — End-Semester", size=15, bold=True, color=WHITE)
    add_bullets(s, Inches(7.15), Inches(2.4), col_w - Inches(0.5), Inches(4.1), [
        "Implement standard Conformal Prediction on surrogate residuals",
        "Implement Mondrian CP — per-category calibration and coverage",
        "Stress-test the exchangeability assumption (out-of-distribution scenario)",
        "Full comparison: GP vs. standard CP vs. Mondrian CP",
        "Final report and end-semester presentation",
    ], size=14, space_after=13)

    prs.save(OUT_PATH)
    print(f"Saved {OUT_PATH}")


if __name__ == "__main__":
    build()
