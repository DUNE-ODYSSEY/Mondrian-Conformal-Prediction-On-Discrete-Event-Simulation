"""
Shared content and formatting helpers for the single 200+ page "book" format
assignment (professor-requested expansion of the original 2-4 page report,
reformatted to the IEEE/Springer technical-textbook layout used by academic
engineering and robotics textbooks: numbered bracket citations, chapter-scoped
equation/figure/table numbering, decimal heading hierarchy, top-captioned
tables, bottom-captioned figures).

Chapters 1-5 (Introduction, Literature Review, Research Gap, Methodology,
Implementation) are written once here and imported by the single book build
script. Chapter 6 (Results & Discussion), the Conclusion, References, and the
code appendix are built in that script.

Formatting: Times New Roman 11pt body, 1.15 line spacing, 7in x 10in page
with a 0.85in binding gutter - a standard bound technical-textbook layout.
Every number/citation used here traces to results/tables/, PROJECT_LOG.md, or
literature/candidate_papers.md - nothing invented, consistent with this
project's standing practice.

Citation system: IEEE numbered bracket style. CITATION_DB below holds every
source's IEEE-formatted reference string, keyed by a short citation key. The
cite(key) function returns "[n]" for use inline in body text, assigning n on
first use in document-build order - so the reference list (built at the end
by iterating get_citation_order()) is automatically numbered by true order of
first appearance, with no manual bookkeeping required.
"""

import os
import re
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TEAM = [
    ("G Venugopalan", "CB.AI.U4AID25115"),
    ("Vipin Sudhakar", "CB.AI.U4AID25166"),
    ("Rithvik Arulprakash", "CB.AI.U4AID25148"),
    ("Harshith Kv", "CB.AI.U4AID25119"),
]
COURSE_CODE = "23AID201"
GROUP_NUMBER = "B9"

NAVY = RGBColor(0x1E, 0x3A, 0x5F)
GREY = RGBColor(0x40, 0x40, 0x40)
BLACK = RGBColor(0x00, 0x00, 0x00)
INK = RGBColor(0x0A, 0x0A, 0x0A)  # near-black, used for headings instead of a "deck" accent color

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))

# Content width used to place equation tab stops (page width minus margins
# and gutter); kept as a module constant since add_equation() needs it and
# has no direct handle on the live section object.
PAGE_WIDTH_IN = 7.0
PAGE_HEIGHT_IN = 10.0
MARGIN_IN = 0.75
GUTTER_IN = 0.85
CONTENT_WIDTH_IN = PAGE_WIDTH_IN - MARGIN_IN - GUTTER_IN

# Current chapter number, tracked so add_table/add_figure/add_equation can
# number "N.M" without every call site having to pass the chapter in.
_current_chapter = {"n": 0}


# --------------------------------------------------------------------------
# Low-level formatting helpers
# --------------------------------------------------------------------------

def set_document_defaults(doc):
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(8)
    for section in doc.sections:
        section.page_width = Inches(PAGE_WIDTH_IN)
        section.page_height = Inches(PAGE_HEIGHT_IN)
        section.left_margin = Inches(MARGIN_IN)
        section.right_margin = Inches(MARGIN_IN)
        section.top_margin = Inches(MARGIN_IN)
        section.bottom_margin = Inches(MARGIN_IN)
        section.gutter = Inches(GUTTER_IN)
        # True mirrored (odd/even) margins for bound-book layout - python-docx
        # has no direct property for this, so set the raw sectPr flag.
        sectPr = section._sectPr
        mirror = OxmlElement("w:mirrorMargins")
        sectPr.append(mirror)


def set_cell_shading(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def add_field(paragraph, field_code):
    """Insert a Word field code (e.g. TOC, PAGE) into a paragraph's run."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = field_code
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)
    return run


def add_page_break(doc):
    doc.add_page_break()


def add_chapter_heading(doc, chapter_no, title, new_page=True):
    """IEEE/textbook Level-1 heading: 'N. TITLE', bold, all caps, 14pt.
    Also resets this chapter's table/figure/equation counters and records
    the current chapter number, so add_table/add_figure/add_equation can
    number captions 'N.M' without the caller passing the chapter in."""
    if new_page:
        add_page_break(doc)
    _current_chapter["n"] = chapter_no
    _table_counter["n"] = 0
    _figure_counter["n"] = 0
    _equation_counter["n"] = 0

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = h.add_run(f"{chapter_no}. {title.upper()}")
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = INK
    r.font.name = "Times New Roman"
    doc.add_paragraph()
    return h


def add_section_heading(doc, text, level=None):
    """IEEE/textbook Level-2 ('N.M Title', bold Title Case, 12pt) or Level-3
    ('N.M.K Title', bold italic Title Case, 11pt) heading. Level is
    auto-detected from the leading decimal number's depth (e.g. '6.5 ...'
    -> level 2, '6.5.1 ...' -> level 3) unless given explicitly, so existing
    call sites that already embed 'chapter.section[.subsection]' numbers in
    their heading text don't need to be touched individually."""
    if level is None:
        m = re.match(r'^(\d+(?:\.\d+)*)\s', text)
        level = min((m.group(1).count(".") + 1), 3) if m else 2
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "Times New Roman"
        run.font.color.rgb = INK
        run.font.bold = True
        if level >= 3:
            run.font.size = Pt(11)
            run.font.italic = True
        else:
            run.font.size = Pt(12)
            run.font.italic = False
    return h


def add_body(doc, text_block, justify=True):
    """Split a triple-quoted text block on blank lines into paragraphs."""
    paras = [p.strip() for p in text_block.strip().split("\n\n") if p.strip()]
    for p_text in paras:
        p_text = " ".join(line.strip() for line in p_text.splitlines())
        p = doc.add_paragraph(p_text)
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0.3)


def add_native_caption(doc, kind, ref_no, text):
    """IEEE/textbook caption paragraph: 'Table N.M: text' (placed above a
    table) or 'Fig. N.M. text' (placed below a figure), where ref_no is a
    'chapter.number' string computed by the caller. Uses Word's 'Caption'
    paragraph style for consistent visual formatting; numbering is
    chapter-scoped and computed directly in Python (see add_table/add_figure
    below) rather than via a Word SEQ field, since SEQ's own auto-numbering
    has no built-in per-chapter reset without tying captions to true
    Heading-1-styled paragraphs."""
    p = doc.add_paragraph()
    try:
        p.style = doc.styles["Caption"]
    except KeyError:
        pass
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sep = ":" if kind == "Table" else "."
    r = p.add_run(f"{kind} {ref_no}{sep} {text}")
    r.font.italic = True
    r.font.size = Pt(10)
    r.font.color.rgb = GREY
    r.font.name = "Times New Roman"
    return p


_table_counter = {"n": 0}
_figure_counter = {"n": 0}
_equation_counter = {"n": 0}


def reset_counters():
    _table_counter["n"] = 0
    _figure_counter["n"] = 0
    _equation_counter["n"] = 0
    _current_chapter["n"] = 0


def add_table(doc, headers, rows, caption=None, col_widths=None):
    # IEEE/textbook style: the table caption is placed ABOVE the table,
    # chapter-scoped ("Table 6.3: ..."), so it is emitted before the table
    # itself rather than after.
    if caption:
        _table_counter["n"] += 1
        add_native_caption(doc, "Table", f"{_current_chapter['n']}.{_table_counter['n']}", caption)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for j, h in enumerate(headers):
        hdr_cells[j].text = str(h)
        set_cell_shading(hdr_cells[j], "1E3A5F")
        for p in hdr_cells[j].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(10)
                r.font.name = "Times New Roman"
    for row in rows:
        cells = table.add_row().cells
        for j, val in enumerate(row):
            # Handle embedded newlines correctly (multi-line-cell bug seen
            # three times before in this codebase: cell.text splits '\n'
            # into separate paragraphs, so every paragraph/run must be
            # styled, not just paragraphs[0].runs[0].
            cells[j].text = str(val)
            for p in cells[j].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.size = Pt(10)
                    r.font.name = "Times New Roman"
    if col_widths:
        for row in table.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = w
    doc.add_paragraph()
    return table


def add_figure(doc, image_path, caption, width_inches=5.1):
    # IEEE/textbook style: figure caption goes BELOW the figure, chapter-scoped
    # ("Fig. 6.3. ...").
    _figure_counter["n"] += 1
    if os.path.exists(image_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(image_path, width=Inches(width_inches))
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"[Figure not found: {image_path}]")
        r.font.italic = True
        r.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)
    add_native_caption(doc, "Fig.", f"{_current_chapter['n']}.{_figure_counter['n']}", caption)


def add_equation(doc, equation_text, note=None):
    """IEEE/textbook display equation: centered on its own line, with a
    right-aligned chapter-scoped number '(N.M)'. `equation_text` should use
    Unicode math notation directly (subscripts, Greek letters, operators
    such as ∑ √ ≤ ±) rather than LaTeX markup, since a
    plain docx run cannot render LaTeX; variables are italicized by
    convention by wrapping them in the caller's text with the understanding
    that the whole equation run is italic (matching 'variables are always
    italicized' - the equation run below is italic in full, which is the
    correct rendering for every equation in this report since each is
    composed entirely of variables, operators, and Greek letters, not
    upright prose). An optional short `note` is printed as a small line
    directly below the equation (e.g. defining a symbol used only here).
    """
    _equation_counter["n"] += 1
    tag = f"({_current_chapter['n']}.{_equation_counter['n']})"

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2 if note else 8)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(CONTENT_WIDTH_IN / 2), WD_TAB_ALIGNMENT.CENTER)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(CONTENT_WIDTH_IN), WD_TAB_ALIGNMENT.RIGHT)
    r1 = p.add_run("\t" + equation_text)
    r1.font.name = "Times New Roman"
    r1.font.italic = True
    r1.font.size = Pt(12)
    r2 = p.add_run("\t" + tag)
    r2.font.name = "Times New Roman"
    r2.font.italic = False
    r2.font.size = Pt(11)

    if note:
        pn = doc.add_paragraph()
        pn.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pn.paragraph_format.space_after = Pt(8)
        rn = pn.add_run(note)
        rn.font.size = Pt(9.5)
        rn.font.italic = True
        rn.font.color.rgb = GREY
        rn.font.name = "Times New Roman"
    return tag


# --------------------------------------------------------------------------
# IEEE numbered-citation system
# --------------------------------------------------------------------------
#
# CITATION_DB maps a short key to an IEEE-formatted reference string (author
# initials first, "Title," in quotes, venue italicized-by-convention in
# plain text since a dict value can't carry rich formatting - venue names
# are rendered in italics by add_reference_entry() in the builder script,
# which splits on the venue marker). cite(key) returns the bracket number
# "[n]" for inline use, assigning n on first use - so simply calling cite()
# in the order paragraphs are written to the document produces a reference
# list correctly numbered by true order of first appearance, with no manual
# renumbering ever required even if paragraphs are reordered later.

CITATION_DB = {
    "vovk2005": ("V. Vovk, A. Gammerman, and G. Shafer",
                 "Algorithmic Learning in a Random World,",
                 "Springer, New York, NY, USA, 2005."),
    "papadopoulos2002": ("H. Papadopoulos, K. Proedrou, V. Vovk, and A. Gammerman",
                 '"Inductive Confidence Machines for Regression,"',
                 "in Proc. 13th Eur. Conf. Machine Learning (ECML), Lecture Notes in Computer Science, vol. 2430, pp. 345-356, 2002."),
    "shafer2008": ("G. Shafer and V. Vovk",
                 '"A Tutorial on Conformal Prediction,"',
                 "Journal of Machine Learning Research, vol. 9, pp. 371-421, 2008."),
    "lei2018": ("J. Lei, M. G'Sell, A. Rinaldo, R. J. Tibshirani, and L. Wasserman",
                 '"Distribution-Free Predictive Inference for Regression,"',
                 "Journal of the American Statistical Association, vol. 113, no. 523, pp. 1094-1111, 2018."),
    "romano2019": ("Y. Romano, E. Patterson, and E. Candès",
                 '"Conformalized Quantile Regression,"',
                 "in Advances in Neural Information Processing Systems 32 (NeurIPS), 2019."),
    "angelopoulos2021": ("A. N. Angelopoulos and S. Bates",
                 '"A Gentle Introduction to Conformal Prediction and Distribution-Free Uncertainty Quantification,"',
                 "arXiv:2107.07511, 2021."),
    "tibshirani2019": ("R. J. Tibshirani, R. F. Barber, E. Candès, and A. Ramdas",
                 '"Conformal Prediction Under Covariate Shift,"',
                 "in Advances in Neural Information Processing Systems 32 (NeurIPS), 2019."),
    "barber2023": ("R. F. Barber, E. Candès, A. Ramdas, and R. J. Tibshirani",
                 '"Conformal Prediction Beyond Exchangeability,"',
                 "Annals of Statistics, vol. 51, no. 2, pp. 816-845, 2023."),
    "vovk2003": ("V. Vovk, D. Lindsay, I. Nouretdinov, and A. Gammerman",
                 '"Mondrian Confidence Machine,"',
                 "Technical Report, Royal Holloway, University of London, 2003."),
    "bostrom2020": ("H. Boström and U. Johansson",
                 '"Mondrian Conformal Regressors,"',
                 "Proceedings of Machine Learning Research (PMLR), vol. 128 (COPA 2020), pp. 114-133, 2020."),
    "bostrom2021": ("H. Boström, U. Johansson, and T. Löfström",
                 '"Mondrian Conformal Predictive Distributions,"',
                 "Proceedings of Machine Learning Research (PMLR), vol. 152 (COPA 2021), 2021."),
    "toccaceli2019": ("P. Toccaceli and A. Gammerman",
                 '"Combination of Inductive Mondrian Conformal Predictors,"',
                 "Machine Learning, vol. 108, pp. 489-510, 2019."),
    "gopakumar2026": ("V. Gopakumar et al.",
                 '"Uncertainty Quantification of Surrogate Models Using Conformal Prediction,"',
                 "Machine Learning: Science and Technology, 2026."),
    "kennedy2001": ("M. C. Kennedy and A. O'Hagan",
                 '"Bayesian Calibration of Computer Models,"',
                 "Journal of the Royal Statistical Society: Series B, vol. 63, no. 3, pp. 425-464, 2001."),
    "rasmussen2006": ("C. E. Rasmussen and C. K. I. Williams",
                 "Gaussian Processes for Machine Learning,",
                 "MIT Press, Cambridge, MA, USA, 2006."),
    "friedman2001": ("J. H. Friedman",
                 '"Greedy Function Approximation: A Gradient Boosting Machine,"',
                 "Annals of Statistics, vol. 29, no. 5, pp. 1189-1232, 2001."),
    "lakshminarayanan2017": ("B. Lakshminarayanan, A. Pritzel, and C. Blundell",
                 '"Simple and Scalable Predictive Uncertainty Estimation Using Deep Ensembles,"',
                 "in Advances in Neural Information Processing Systems 30 (NeurIPS), pp. 6402-6413, 2017."),
    "abdar2021": ("M. Abdar et al.",
                 '"A Review of Uncertainty Quantification in Deep Learning: Techniques, Applications and Challenges,"',
                 "Information Fusion, vol. 76, pp. 243-297, 2021."),
    "green2006": ("L. V. Green, J. Soares, J. F. Giglio, and R. A. Green",
                 '"Using Queueing Theory to Increase the Effectiveness of Emergency Department Provider Staffing,"',
                 "Academic Emergency Medicine, vol. 13, no. 1, pp. 61-68, 2006."),
    "green_book": ("L. V. Green",
                 '"Queueing Analysis in Healthcare,"',
                 "book chapter, Columbia Business School, New York, NY, USA."),
    "hu2018": ("X. Hu et al.",
                 '"Applying Queueing Theory to the Study of Emergency Department Operations: A Survey and a Discussion of Comparable Simulation Studies,"',
                 "International Transactions in Operational Research, 2018."),
    "mg1_queue": ("Anonymous",
                 '"Performance Evaluation of a M/G/1 Queue Model for Patient Flow in a Health Care System,"',
                 "Mathematical Modelling of Engineering Problems, IIETA."),
    "staffing_pmc": ("Anonymous",
                 '"Decision Support for the Optimization of Provider Staffing for Hospital Emergency Departments with a Queue-Based Approach,"',
                 "PMC6947400."),
    "des_calibration2021": ("Anonymous",
                 '"A Simulation-Based Optimization Approach for the Calibration of a Discrete Event Simulation Model of an Emergency Department,"',
                 "arXiv:2102.00945, 2021."),
    "des_review2022": ("Anonymous",
                 '"Discrete Event Simulation for Emergency Department Modelling: A Systematic Review of Validation Methods,"',
                 "ScienceDirect, 2022."),
    "des_kolding": ("Anonymous",
                 '"Discrete Event Simulation Modelling for an Improved Patient Flow at the Emergency Department, Sygehus Lillebælt, Kolding,"',
                 "PMC3327033."),
    "ambulance_diversion2021": ("Anonymous",
                 '"A Simulation-Based Optimization Approach for Analyzing the Ambulance Diversion Phenomenon in an Emergency Department Network,"',
                 "arXiv:2108.04162, 2021."),
    "los_prediction2023": ("Anonymous",
                 '"Machine Learning-Based Prediction of Hospital Prolonged Length of Stay Admission at Emergency Department: A Gradient Boosting Algorithm Analysis,"',
                 "Frontiers in Artificial Intelligence, 2023."),
    "overcrowding2025": ("Anonymous",
                 '"An Artificial Intelligence-Based Framework for Predicting Emergency Department Overcrowding: Development and Evaluation Study,"',
                 "arXiv:2504.18578, 2025."),
    "triage_ml_pmc": ("Anonymous",
                 '"Machine Learning-Based Triage to Identify Low-Severity Patients with a Short Discharge Length of Stay in Emergency Department,"',
                 "PMC9123815."),
    "kaggle_dataset": ("maalona (Kaggle username)",
                 '"Hospital Triage and Patient History Data,"',
                 "Kaggle dataset - Yale New Haven Health System, retrospective study, March 2014 - July 2017."),
    "hoot2008": ("N. R. Hoot, L. J. LeBlanc, I. Jones, S. R. Levin, C. Zhou, C. S. Gadd, and D. Aronsky",
                 '"Forecasting Emergency Department Crowding: A Discrete Event Simulation,"',
                 "Annals of Emergency Medicine, vol. 52, no. 2, pp. 116-125, 2008."),
    "mahmoodian2014": ("F. Mahmoodian, R. Eqtesadi, and A. Ghareghani",
                 '"Waiting Times in Emergency Department After Using the Emergency Severity Index Triage Tool,"',
                 "Archives of Trauma Research, vol. 3, no. 4, e19507, 2014."),
    "otto2022": ("R. Otto, S. Blaschke, W. Schirrmeister, et al.",
                 '"Length of Stay as Quality Indicator in Emergency Departments: Analysis of Determinants in the German Emergency Department Data Registry (AKTIN Registry),"',
                 "Internal and Emergency Medicine, vol. 17, no. 4, pp. 1199-1209, 2022."),
    "kim2021": ("T. Y. Kim, C. Ohmart, Z. Khan, M. Lance, and S. Kim",
                 '"The Effect on Length of Stay After Implementation of Discharging Low Acuity Patients From Triage,"',
                 "Cureus, vol. 13, no. 9, e17640, 2021."),
    "desantis2021": ("A. De Santis, T. Giovannelli, S. Lucidi, M. Messedaglia, and M. Roma",
                 '"Determining the Optimal Piecewise Constant Approximation for the Nonhomogeneous Poisson Process Rate of Emergency Department Patient Arrivals,"',
                 "arXiv:2101.11138, 2021."),
    "theiling2020": ("B. J. Theiling, K. V. Kennedy, A. T. Limkakeng Jr., P. Manandhar, A. Erkanli, and S. R. Pitts",
                 '"A Method for Grouping Emergency Department Visits by Severity and Complexity,"',
                 "Western Journal of Emergency Medicine, vol. 21, no. 5, pp. 1147-1155, 2020."),
    "laskowski2009": ("M. Laskowski, R. D. McLeod, M. R. Friesen, B. W. Podaima, and A. S. Alfa",
                 '"Models of Emergency Departments for Reducing Patient Waiting Times,"',
                 "PLoS ONE, vol. 4, no. 7, e6127, 2009."),
    "karaca2012": ("Z. Karaca, H. S. Wong, and R. L. Mutter",
                 '"Duration of Patients\' Visits to the Hospital Emergency Department,"',
                 "BMC Emergency Medicine, vol. 12, no. 15, 2012."),
    "locker2005": ("T. E. Locker and S. M. Mason",
                 '"Analysis of the Distribution of Time That Patients Spend in Emergency Departments,"',
                 "BMJ, vol. 330, no. 7501, pp. 1188-1189, 2005."),
    "kramer2020": ("A. Kramer, C. Dosi, M. Iori, and M. Vignoli",
                 '"Successful Implementation of Discrete Event Simulation: The Case of an Italian Emergency Department,"',
                 "arXiv:2006.13062, 2020."),
}

_citation_order = []
_citation_number = {}


def reset_citations():
    _citation_order.clear()
    _citation_number.clear()


def cite(key):
    """Return this citation's IEEE bracket number, e.g. '[7]', assigning the
    next number the first time `key` is used (so numbering matches true
    order of first appearance in the finished document)."""
    if key not in CITATION_DB:
        raise KeyError(f"Unknown citation key: {key!r} - add it to CITATION_DB first.")
    if key not in _citation_number:
        _citation_order.append(key)
        _citation_number[key] = len(_citation_order)
    return f"[{_citation_number[key]}]"


def get_citation_order():
    """Ordered list of (bracket_number, key) for every citation actually
    used, in order of first appearance - what the References chapter iterates."""
    return [(_citation_number[k], k) for k in _citation_order]


def add_code_listing(doc, rel_path, title=None):
    """Insert a real source file's contents as a monospace listing."""
    full_path = os.path.join(REPO_ROOT, rel_path)
    heading = doc.add_paragraph()
    r = heading.add_run(title or rel_path)
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.name = "Times New Roman"
    r.font.color.rgb = NAVY
    doc.add_paragraph()

    if not os.path.exists(full_path):
        p = doc.add_paragraph(f"[Source file not found: {rel_path}]")
        p.runs[0].font.italic = True
        return

    with open(full_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Group ~55 lines per "page block" paragraph for reasonable render perf;
    # Word will still paginate naturally within a monospace block.
    chunk = []
    for i, line in enumerate(lines, start=1):
        chunk.append(f"{i:4d}  {line.rstrip(chr(10))}")
    code_text = "\n".join(chunk)
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(code_text)
    run.font.name = "Consolas"
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x20, 0x20, 0x20)
    doc.add_paragraph()


# --------------------------------------------------------------------------
# Front matter
# --------------------------------------------------------------------------

def build_title_page(doc, title, subtitle):
    for _ in range(2):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.font.size = Pt(24)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r.font.name = "Times New Roman"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle)
    r.font.size = Pt(14)
    r.font.italic = True
    r.font.color.rgb = GREY
    r.font.name = "Times New Roman"

    for _ in range(2):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("A project-based report submitted in partial fulfilment of the "
                   f"requirements of course {COURSE_CODE}")
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Submitted by\nGroup {GROUP_NUMBER}")
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.name = "Times New Roman"

    doc.add_paragraph()
    for name, roll in TEAM:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"{name}  ({roll})")
        r.font.size = Pt(12)
        r.font.name = "Times New Roman"

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Department of Artificial Intelligence")
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"


def build_certificate(doc):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("BONAFIDE CERTIFICATE")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    add_body(doc, f"""
This is to certify that this project report is a bonafide record of work carried
out by Group {GROUP_NUMBER} - {", ".join(f"{n} ({i})" for n, i in TEAM)} - under my
supervision, submitted in partial fulfilment of the requirements for course
{COURSE_CODE}.

The results embodied in this report have not been submitted to any other
University or Institute for the award of any degree or diploma.
""")
    for _ in range(4):
        doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("[Guide's Name and Signature]")
    r.font.italic = True
    r.font.color.rgb = GREY
    p2 = doc.add_paragraph()
    r2 = p2.add_run("Project Guide")
    r2.font.bold = True
    p3 = doc.add_paragraph()
    r3 = p3.add_run("[Head of Department's Name and Signature]")
    r3.font.italic = True
    r3.font.color.rgb = GREY
    p4 = doc.add_paragraph()
    r4 = p4.add_run("Head of Department")
    r4.font.bold = True


def build_declaration(doc):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("DECLARATION")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    add_body(doc, f"""
We, the undersigned, declare that this project report titled above is our own
work, carried out under the supervision of our project guide, and that it has
not been submitted elsewhere for the award of any degree or diploma. All
sources of information, whether published or unpublished, have been
acknowledged in the References section.
""")
    doc.add_paragraph()
    for name, roll in TEAM:
        p = doc.add_paragraph()
        r = p.add_run(f"{name}  ({roll})")
        r.font.name = "Times New Roman"


def build_acknowledgement(doc):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("ACKNOWLEDGEMENT")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    add_body(doc, f"""
We express our sincere gratitude to our project guide for their continuous
guidance, technical insight, and encouragement throughout the course of this
project. We thank the Head of the Department and the faculty of the
Department of Artificial Intelligence for providing the resources and
environment necessary to carry out this work.

We also acknowledge the creators of the Hospital Triage and Patient History
Data dataset (Kaggle, user maalona), which made the real-data calibration
underlying this project's simulation possible, and the authors of the 30
research papers surveyed in Chapter 2, whose work forms the theoretical
foundation on which this project builds.

Finally, we thank our families and peers for their patience and support
throughout this project.
""")


def build_abstract(doc, abstract_text):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("ABSTRACT")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    add_body(doc, abstract_text)


def build_preface(doc, preface_text):
    """Plain book-style Preface page - matches the reference book's front
    matter (title page -> Preface -> Contents, no certificate/declaration/
    acknowledgement/abstract block), used in place of build_abstract for
    the single consolidated book report."""
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("Preface")
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = INK
    doc.add_paragraph()
    add_body(doc, preface_text)


def build_toc_page(doc):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("TABLE OF CONTENTS")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    p = doc.add_paragraph()
    add_field(p, 'TOC \\o "1-3" \\h \\z \\u')


def build_list_of_figures_tables(doc):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("LIST OF FIGURES")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    p = doc.add_paragraph()
    add_field(p, 'TOC \\c "Figure"')
    doc.add_paragraph()

    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("LIST OF TABLES")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    p = doc.add_paragraph()
    add_field(p, 'TOC \\c "Table"')


def build_abbreviations(doc):
    add_page_break(doc)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("LIST OF ABBREVIATIONS")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    doc.add_paragraph()
    abbrevs = [
        ("CP", "Conformal Prediction"),
        ("CQR", "Conformalized Quantile Regression"),
        ("DES", "Discrete-Event Simulation"),
        ("ED", "Emergency Department"),
        ("ER", "Emergency Room"),
        ("ESI", "Emergency Severity Index (1 = most acute, 5 = least acute)"),
        ("GBR", "Gradient Boosting Regressor (HistGradientBoostingRegressor)"),
        ("GP", "Gaussian Process"),
        ("LOS", "Length of Stay"),
        ("MAE", "Mean Absolute Error"),
        ("MLP", "Multi-Layer Perceptron"),
        ("OOD", "Out-of-Distribution"),
        ("R²", "Coefficient of Determination"),
        ("RMSE", "Root Mean Squared Error"),
        ("UQ", "Uncertainty Quantification"),
    ]
    add_table(doc, ["Abbreviation", "Full Form"], abbrevs)


# --------------------------------------------------------------------------
# Chapter 2: Literature Review (shared) - helper for one paper entry
# --------------------------------------------------------------------------

# The 30 reviewed papers, in the exact order add_paper_review() below is
# called for them (matches literature/candidate_papers.md's 5-category
# ordering) - maps each review's sequential 2.N section number to its
# CITATION_DB key, so IEEE bracket numbers are assigned in true
# order-of-first-appearance as the literature review is written, and every
# other add_paper_review() call site below is untouched (still just passes
# its section number; the old inline citation-string argument is ignored in
# favor of the IEEE-formatted CITATION_DB entry, kept only so existing call
# sites don't need editing).
PAPER_KEYS = [
    "vovk2005", "papadopoulos2002", "shafer2008", "lei2018", "romano2019",
    "angelopoulos2021", "tibshirani2019", "barber2023", "vovk2003", "bostrom2020",
    "bostrom2021", "toccaceli2019", "gopakumar2026", "kennedy2001", "rasmussen2006",
    "friedman2001", "lakshminarayanan2017", "abdar2021", "green2006", "green_book",
    "hu2018", "mg1_queue", "staffing_pmc", "des_calibration2021", "des_review2022",
    "des_kolding", "ambulance_diversion2021", "los_prediction2023", "overcrowding2025", "triage_ml_pmc",
]


def add_paper_review(doc, number, citation, summary, relevance):
    key = PAPER_KEYS[number - 1]
    n = cite(key)
    authors, title, venue = CITATION_DB[key]
    p = doc.add_paragraph()
    r = p.add_run(f"{n}  {authors}, {title}")
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"
    r.font.color.rgb = INK
    rv = p.add_run(f" {venue}")
    rv.font.bold = True
    rv.font.italic = True
    rv.font.size = Pt(12)
    rv.font.name = "Times New Roman"
    rv.font.color.rgb = INK
    add_body(doc, summary)
    rp = doc.add_paragraph()
    r = rp.add_run("Relevance to this project: ")
    r.font.bold = True
    r.font.italic = True
    r2 = rp.add_run(relevance)
    r2.font.italic = True
    rp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rp.paragraph_format.first_line_indent = Inches(0.3)
    doc.add_paragraph()


# --------------------------------------------------------------------------
# Chapter 1: Introduction (shared)
# --------------------------------------------------------------------------

def build_chapter1_introduction(doc):
    add_chapter_heading(doc, 1, "Introduction")

    add_section_heading(doc, "1.1 Background and Motivation")
    add_body(doc, """
Emergency departments (EDs) are among the most heavily studied stochastic service
systems in operations research, precisely because they combine three properties
that make good decision-making difficult: highly variable arrival patterns, a
service process whose duration depends on patient acuity in ways that are only
partially observable at intake, and severe, immediate consequences for getting
staffing or capacity decisions wrong. Overcrowding in an ED is not merely an
inconvenience; a substantial body of clinical and health-services literature
associates ED crowding with delayed treatment, increased rates of patients
leaving without being seen, and worse downstream outcomes. At the same time,
over-staffing relative to demand is costly and, given that clinical staff are a
scarce resource in most health systems, frequently not even feasible to correct
for by simply adding more capacity.

Because real EDs cannot be experimented on directly - a hospital cannot try five
different staffing policies on the same day to see which produces the shortest
waits - operations researchers and hospital administrators have long relied on
models: analytic queueing theory on one end of the spectrum, and discrete-event
simulation (DES) on the other. Queueing theory offers closed-form or
semi-closed-form expressions for quantities such as expected wait time under
simplifying assumptions (Poisson arrivals, exponential or otherwise well-behaved
service times, a fixed number of homogeneous servers). These assumptions are
mathematically convenient but often violated in a real ED, where service times
are heavy-tailed and acuity-dependent, and where the practical quantity of
interest - a full distribution of possible wait times under a given staffing
policy, not just its mean - is exactly what a closed-form queueing result is
least well suited to provide. DES relaxes these assumptions at the cost of
losing closed-form tractability: a DES model can simulate patient-by-patient
arrivals, triage, and service under essentially arbitrary distributional
assumptions, and its output for any one run is a sample from the same kind of
stochastic process a real ED experiences, rather than a mean-field
approximation of it.

The price of that flexibility is computational: a sufficiently detailed DES,
run enough times to characterize the distribution of outcomes under a given
scenario (a given staffing level, a given arrival-rate regime), can be
expensive to evaluate repeatedly - particularly if the goal is to explore many
scenarios, as it is in staffing and capacity planning, or to embed the
simulation inside an optimization loop. This computational cost is the
practical motivation for training a surrogate model: a fast, learned
approximation of the DES's input-output relationship, trained on a finite set
of DES runs and then queried in place of the DES itself wherever repeated,
low-latency evaluation is needed.
""")

    add_section_heading(doc, "1.2 The Role of Surrogate Models in Simulation-Based Decision Making")
    add_body(doc, """
A surrogate model, in the sense used throughout this report, is a supervised
machine learning model trained to predict a simulator's output metrics (here,
daily patient count, mean wait time, mean total time in system, and the 95th
percentile of wait time) directly from the simulator's scenario-level inputs
(here, staffing capacity and an arrival-rate multiplier), without re-running
the simulator itself. Once trained, a surrogate can be evaluated in
microseconds rather than the seconds-to-minutes a full DES run may take,
enabling the kind of dense scenario sweeps, sensitivity analyses, and
what-if staffing comparisons that would be computationally prohibitive if each
evaluation required a fresh simulation.

This efficiency is not free, however. A surrogate model is only as reliable as
the region of input space its training data actually covers, and it inherits
- indeed, it can amplify - two distinct sources of uncertainty: the epistemic
uncertainty of the learned function itself (the surrogate is a finite-sample
approximation, imperfectly fit even within its training domain) and the
aleatoric uncertainty of the underlying stochastic process it approximates
(a DES with fixed inputs still produces different outputs on different random
seeds, because arrivals, triage assignment, and service durations are all
drawn from probability distributions). A surrogate's point prediction alone -
"expected mean wait time is 42 minutes" - collapses both of these sources of
uncertainty into a single number and, critically, gives a decision-maker no way
to distinguish a scenario where 42 minutes is a confident, tightly-clustered
estimate from one where it is a rough central tendency over an output that
could plausibly range from 20 to 90 minutes. For an operational decision such
as ED staffing, where the cost of being wrong is asymmetric and often severe in
exactly the tail scenarios a point estimate is least informative about, this
gap between "central estimate" and "how much to trust the central estimate" is
not a cosmetic omission - it is the difference between a genuinely useful
decision-support tool and one that looks precise while quietly discarding the
information a decision-maker most needs.
""")

    add_section_heading(doc, "1.3 The Need for Uncertainty Quantification")
    add_body(doc, """
Uncertainty quantification (UQ) is the general term for methods that
accompany a point prediction with some characterization of its reliability -
most commonly, an interval or a full predictive distribution rather than a
single number. A wide family of UQ methods exists for machine learning models,
ranging from Bayesian approaches (Gaussian processes, Bayesian neural
networks, and their approximations) through frequentist resampling methods
(bootstrap-based intervals) to ensemble-based approaches (deep ensembles,
random-forest quantile estimates) and, most relevant to this project,
distribution-free methods based on conformal prediction. Each family makes a
different trade-off between the strength of its theoretical guarantee, its
computational cost, and the assumptions it requires about the underlying data
or model.

Gaussian process (GP) regression, used in this project as a UQ baseline (see
Chapter 4 and Chapter 6), is a Bayesian nonparametric method that produces a
full posterior predictive distribution - and therefore both a point estimate
and a principled uncertainty estimate - for any query point, under the
assumption that the underlying function is well described by a specified
covariance (kernel) structure. Its central practical limitation is
computational: exact GP inference scales cubically in the number of training
points, which makes it expensive to refit as new data arrives and effectively
rules out its use on large training sets without approximation. It also
provides no finite-sample coverage guarantee: its intervals are only as
trustworthy as the appropriateness of its modeling assumptions (kernel choice,
Gaussian noise, stationarity) for the data at hand, and there is no
distribution-free bound on how far its empirical coverage can drift from its
nominal target when those assumptions are violated.

Conformal prediction (CP), the family of methods this project is centrally
concerned with, offers a different trade-off. It is model-agnostic - it wraps
around any already-trained point predictor, including the gradient-boosting
and neural-network surrogates used in this project, without requiring the
underlying model to be retrained or to expose calibrated uncertainty itself -
and it comes with a finite-sample, distribution-free coverage guarantee: under
an exchangeability assumption between the calibration data and future test
points, a conformal prediction interval constructed at a nominal
(1 - alpha) confidence level is guaranteed to contain the true value with
probability at least (1 - alpha), regardless of the correctness of the
underlying point predictor's modeling assumptions. This guarantee is
attractive precisely because it does not depend on the point predictor being a
good model of reality, only on the exchangeability of calibration and test
data - a much weaker and more checkable assumption than "the model's
distributional assumptions are correct."
""")

    add_section_heading(doc, "1.4 Conformal Prediction as a Distribution-Free UQ Framework")
    add_body(doc, f"""
The theoretical foundation of conformal prediction was laid by Vovk, Gammerman,
and collaborators in the late 1990s and formalized in the book Algorithmic
Learning in a Random World {cite('vovk2005')}, with Shafer and
Vovk's tutorial {cite('shafer2008')} (surveyed in Chapter 2) later popularizing the method
outside the original algorithmic-learning-theory community. The core idea is
simple to state: given a trained point predictor and a held-out calibration
set, compute a nonconformity score for each calibration point (typically the
absolute residual between the predictor's output and the true value), and use
the (1 - alpha) empirical quantile of those scores as a fixed margin added to
and subtracted from any future point prediction. Because the calibration
residuals and a genuinely exchangeable test residual are, by construction,
draws from the same underlying distribution, the resulting interval is
guaranteed - via a straightforward rank-based argument, not an asymptotic
approximation - to cover the true value at least (1 - alpha) of the time,
in finite samples, without any assumption about the shape of the residual
distribution.

This guarantee, however, is explicitly marginal: it holds averaged over the
entire calibration and test distribution, and says nothing about whether
coverage holds within any specific subgroup of that distribution. If a
surrogate model is systematically less accurate under one operating regime
than another - for instance, understaffed and high-demand scenarios versus
comfortably staffed and low-demand ones - a single pooled conformal interval,
calibrated on average difficulty across all regimes, can silently undercover
in precisely the regime where a decision-maker most needs it to be reliable,
while simultaneously overcovering (wasting interval width) in the easy regime.
This marginal-versus-conditional coverage gap, and Mondrian conformal
prediction's remedy for it (calibrating a separate quantile per category
rather than one pooled quantile across the whole calibration set), forms the
central methodological axis of this project and is developed in full in
Chapters 2, 4, and 6.

Gopakumar et al. {cite('gopakumar2026')} - the base paper this project directly extends -
validate conformal prediction for surrogate-model uncertainty quantification
across several physics-simulation domains (partial differential equations,
magnetohydrodynamics, weather modeling, and fusion-plasma simulation) and
explicitly name two limitations of their own results as untested: the marginal
(rather than conditional) nature of CP's coverage guarantee, and the
exchangeability assumption's fragility under distribution shift between
calibration and test data. Both limitations are stated as open questions in a
physics-simulation context; neither had, to the best of the literature review
conducted for this project (Chapter 2), been tested in a fundamentally
different domain such as discrete-event queueing simulation. This report's
guiding question, developed fully in Chapter 3, is whether the first of
these - the marginal-versus-conditional coverage gap - holds, and whether
Mondrian conformal prediction remedies it, in exactly such a domain: an ED
discrete-event simulation calibrated on real hospital data. The second
limitation, exchangeability under distribution shift, was also stress-tested
as part of this project's broader implementation (Section 4.5.2) and is
referenced where relevant, but is not this report's central subject; it is
summarized as a secondary finding and flagged as a natural direction for
future work (Section 7.4).
""")

    add_section_heading(doc, "1.5 Overview of the Project")
    add_body(doc, """
This project builds a four-stage pipeline. First, a discrete-event simulation
of a single hospital emergency department is implemented in SimPy and
calibrated on real arrival-pattern and triage-acuity data extracted from a
large, de-identified Kaggle dataset of ED visits (Chapter 5 details the
dataset and the calibration procedure, including a deliberate and clearly
documented distinction between the arrival process, which is calibrated
directly on real data, and service-time distributions, which - because the
dataset contains no length-of-stay field - are calibrated on
literature-standard parameters by acuity level rather than presented as
data-derived). Second, the calibrated DES is run across thousands of randomly
sampled staffing-capacity and arrival-rate scenarios to generate a labeled
training set, from which a surrogate regression model is trained to predict
four scenario-level output metrics without needing to re-run the simulation.
Third, three uncertainty quantification approaches are applied on top of the
trained surrogate and compared on identical calibration and test data: a
Gaussian process baseline, standard (pooled) conformal prediction, and
Mondrian conformal prediction, later extended with conformalized quantile
regression (CQR) and a combined Mondrian-CQR variant as stronger baselines.
Fourth, the core comparison is stress-tested for robustness along two
independent axes central to this report - statistical significance across 30
repeated calibration/test draws (rather than trusting a single split), and a
second, independent hospital department (with materially different patient
volume and acuity mix) to check whether findings generalize beyond a single
site - with a third axis, a second surrogate architecture, used specifically
to test the exchangeability question summarized in Section 6.13 rather than
the report's central Mondrian CP finding.

This report's central finding is a genuine marginal-versus-conditional
coverage gap that Mondrian conformal prediction closes where it is real,
established in Chapter 6 across a single representative split, 30 repeated
draws with formal statistical significance testing, a stronger
width-adaptive baseline (conformalized quantile regression), and independent
replication at a second hospital department. Chapters 1 through 5 lay the
theoretical, methodological, and implementation groundwork this finding rests
on; Chapter 6 develops the result itself in full depth, and Chapter 7
concludes with its implications, limitations, and future scope.
""")

    add_section_heading(doc, "1.6 Organization of the Report")
    add_body(doc, """
Chapter 2 surveys thirty papers drawn from five areas directly relevant to
this project: conformal prediction foundations, Mondrian and
conditional-coverage methods, surrogate modeling and uncertainty
quantification more broadly, queueing theory and ED operations research, and
discrete-event simulation together with ED-specific machine learning
applications. Chapter 3 sharpens this survey into an explicit statement of the
research gap this project addresses and the problem statement that follows
from it. Chapter 4 describes the system design and methodology in full
technical detail: the DES's structure and calibration, the surrogate
architectures used, and the mathematics of standard conformal prediction,
Mondrian conformal prediction, and conformalized quantile regression. Chapter
5 documents the implementation - the dataset, the software stack, and a
module-by-module walkthrough of the codebase. Chapter 6 presents this report's
results and discussion in depth, and Chapter 7 concludes with a summary of
findings, an honest account of the project's limitations, and directions for
future work. References and appendices (source code listings and
supplementary result tables) follow.
""")


# --------------------------------------------------------------------------
# Chapter 2: Literature Review (shared) - full content
# --------------------------------------------------------------------------

def build_chapter2_literature_review(doc):
    add_chapter_heading(doc, 2, "Literature Review")

    add_section_heading(doc, "2.0 Scope and Method of This Review")
    add_body(doc, f"""
Thirty papers were selected across five areas directly relevant to this
project's methodology and domain: conformal prediction foundations (Section
2.1), Mondrian and conditional-coverage methods (Section 2.2), surrogate
modeling and uncertainty quantification more broadly (Section 2.3), queueing
theory and emergency department operations research (Section 2.4), and
discrete-event simulation together with ED-specific machine learning
applications (Section 2.5). Every citation in this chapter was verified
against a primary source - a publisher page, a DOI, or a preprint server -
before inclusion, rather than compiled from memory; the full sourcing record
is maintained in this project's repository at
literature/candidate_papers.md. This verification step caught at least one
concrete error during the review process itself: a paper on combining
Mondrian conformal predictors was initially attributed to the authors who
write most of the other Mondrian-CP literature (Boström, Johansson, and
Löfström), and was found, on checking the primary source, to actually be
authored by Toccaceli and Gammerman {cite('toccaceli2019')} - a reminder that even a
plausible-looking citation needs independent verification, and a small
illustration of why this project has, throughout, treated unverified
citations and results as unacceptable rather than as an acceptable
convenience.

Each entry below follows the same structure: the full citation, a summary of
what the paper actually establishes, and a short paragraph connecting it
specifically to this project - not a restatement of the abstract, but an
explanation of where the paper's result or method is actually used, tested,
extended, or contrasted with in the work described in later chapters.
""")

    # -------------------- Section 2.1: CP Foundations --------------------
    add_section_heading(doc, "2.1 Conformal Prediction: Foundations")
    add_body(doc, """
This section covers the theoretical lineage of conformal prediction from its
origin through the specific variants (split conformal, conformalized quantile
regression, and robustness results under distribution shift) that this
project builds on directly.
""")

    add_paper_review(doc, 1,
        "Vovk, V., Gammerman, A., Shafer, G. (2005). Algorithmic Learning in a "
        "Random World. Springer.",
        """
This book is the founding text of conformal prediction, consolidating work
Vovk, Gammerman, and collaborators had developed through the late 1990s into
a single coherent theory. It introduces the transductive (full) conformal
predictor, the notion of a nonconformity measure as the mechanism by which
any point predictor can be turned into a valid prediction region, and proves
the central finite-sample coverage guarantee under the exchangeability
assumption - the property that the joint distribution of the calibration and
test data is invariant to permutation, which is weaker than the more familiar
i.i.d. assumption and is what conformal prediction's guarantee actually
requires. The book develops the theory for both classification and
regression and situates conformal prediction within algorithmic learning
theory more broadly, including its relationship to Kolmogorov complexity and
online prediction with expert advice.
""",
        "This is the ultimate theoretical source for every coverage guarantee "
        "invoked in this project, from the standard CP implementation in "
        "src/uq/standard_cp.py through Mondrian CP. The exchangeability "
        "assumption defined here is precisely the assumption Chapter 6's "
        "exchangeability stress test is designed to violate deliberately and "
        "study the consequences of.")

    add_paper_review(doc, 2,
        "Papadopoulos, H., Proedrou, K., Vovk, V., Gammerman, A. (2002). "
        "Inductive Confidence Machines for Regression. ECML 2002, LNCS 2430, "
        "pp. 345-356.",
        """
This paper introduces split (inductive) conformal prediction as a
computationally tractable alternative to the full transductive conformal
predictor. Rather than refitting the underlying model once per candidate test
label - which the original transductive formulation requires and which is
prohibitively expensive for anything but the simplest models - the inductive
approach splits the available data into a proper training set and a separate
calibration set, fits the model once on the training set, and computes
nonconformity scores only on the calibration set. This sacrifices a small
amount of statistical efficiency (the calibration set is not used to fit the
model itself) in exchange for reducing the computational cost from
retraining per test point to fitting exactly once.
""",
        "This is the exact algorithmic template that src/uq/standard_cp.py "
        "and src/uq/mondrian_cp.py both implement: fit the surrogate once "
        "(already done in Week 6-7), hold out a separate calibration set "
        "(src/uq/generate_calibration_data.py, deliberately disjoint from "
        "both the surrogate's training data and the test set), and compute "
        "residual quantiles only on that calibration set.")

    add_paper_review(doc, 3,
        "Shafer, G., Vovk, V. (2008). A Tutorial on Conformal Prediction. "
        "JMLR, 9, 371-421.",
        """
This widely cited tutorial restates the conformal prediction framework in an
accessible form aimed at a broader machine learning audience than the
original algorithmic-learning-theory literature, and is frequently the first
paper researchers encounter when entering the field. It walks through the
nonconformity-measure formulation, proves the marginal coverage guarantee
via an exchangeability argument, and discusses both the transductive and
inductive (split) variants, along with worked examples in classification and
regression settings.
""",
        "This was the first paper adopted for this project (predating the "
        "systematic 30-paper search) and remains the primary reference used "
        "when explaining conformal prediction's guarantee to readers "
        "unfamiliar with the framework, including in this report's own "
        "Chapter 1 and Chapter 4.")

    add_paper_review(doc, 4,
        "Lei, J., G'Sell, M., Rinaldo, A., Tibshirani, R.J., Wasserman, L. "
        "(2018). Distribution-Free Predictive Inference for Regression. "
        "JASA, 113(523), 1094-1111.",
        """
This paper, from the statistics community's independent adoption and
extension of conformal prediction, establishes distribution-free finite-sample
coverage guarantees for several conformal regression variants, including
split conformal prediction, and analyzes the statistical efficiency
trade-offs between the full (transductive) and split approaches. It also
introduces jackknife+ as a cross-validation-based alternative that recovers
some of split conformal's lost statistical efficiency without the full
transductive method's computational cost.
""",
        "Provides the formal statistical grounding (efficiency loss from "
        "splitting, in exchange for tractability) for the choice made "
        "throughout this project to use split conformal prediction rather "
        "than the more expensive transductive form - a choice that only "
        "makes sense given the sample sizes actually used here (1,200 "
        "calibration points per repeat, well within the regime where split "
        "conformal's efficiency loss is acceptable).")

    add_paper_review(doc, 5,
        "Romano, Y., Patterson, E., Candès, E. (2019). Conformalized Quantile "
        "Regression. NeurIPS 32.",
        """
Conformalized quantile regression (CQR) combines conformal prediction's
finite-sample coverage guarantee with quantile regression's ability to adapt
interval width to local heteroscedasticity. Rather than conformalizing a
constant-width residual (as standard split conformal does with the symmetric
nonconformity measure), CQR trains two quantile regressors (targeting the
alpha/2 and 1 - alpha/2 conditional quantiles) and conformalizes the signed
distance from a test point's prediction to these estimated quantiles. The
result retains the exact finite-sample marginal coverage guarantee of
standard conformal prediction while typically producing substantially
narrower intervals in regions where the underlying quantile regressor
correctly identifies heteroscedasticity - regions of naturally higher or lower
variance in the residual.
""",
        "This is the direct source for src/surrogate/train_quantile_surrogates.py "
        "(the lower/upper quantile regressors) and src/uq/repeated_evaluation_cqr.py "
        "(the CQR and Mondrian-CQR nonconformity score, "
        "max(qlo(x) - y, y - qhi(x))). CQR is used in this report's "
        "Section 6.7 as a stronger baseline that achieves much of "
        "Mondrian CP's benefit through width-adaptivity rather than "
        "category-adaptivity - the two methods sit at different points on "
        "the same coverage/width frontier rather than one strictly "
        "dominating the other.")

    add_paper_review(doc, 6,
        "Angelopoulos, A., Bates, S. (2021). A Gentle Introduction to "
        "Conformal Prediction and Distribution-Free Uncertainty "
        "Quantification. arXiv:2107.07511.",
        """
This widely used tutorial-style survey collects and unifies conformal
prediction methods developed across the machine learning and statistics
communities, presenting split conformal prediction, full conformal
prediction, CQR, and conformal risk control in a single consistent notation,
with worked code examples. It is written specifically to be more accessible
to a machine-learning-practitioner audience than the original theoretical
literature, while still proving the key coverage guarantees rigorously.
""",
        "Used as the primary accessible reference for this report's own "
        "background exposition (Chapter 1, Chapter 4) and cross-checked "
        "against the more technical treatments (Shafer & Vovk, Lei et al.) "
        "to ensure the plain-language description of conformal prediction "
        "given to non-specialist readers of this report remains technically "
        "accurate.")

    add_paper_review(doc, 7,
        "Tibshirani, R.J., Barber, R.F., Candès, E., Ramdas, A. (2019). "
        "Conformal Prediction Under Covariate Shift. NeurIPS 32.",
        """
This paper studies what happens to conformal prediction's coverage guarantee
when the exchangeability assumption is violated specifically by covariate
shift - a change in the distribution of the input features between
calibration and test time, while the conditional relationship between inputs
and outputs stays fixed. It shows that if the likelihood ratio between the
test and calibration covariate distributions is known (or can be estimated),
a weighted conformal prediction procedure can restore valid coverage under
this specific, structured form of distribution shift, at the cost of
requiring that likelihood ratio.
""",
        "Directly relevant to this project's exchangeability stress test "
        "(src/uq/exchangeability_stress_test.py), which pushes the test "
        "distribution's arrival-rate multiplier progressively outside the "
        "calibration range - a covariate shift by exactly this paper's "
        "definition. This project's stress test does not attempt the "
        "likelihood-ratio correction this paper proposes (the shift is "
        "extreme enough, and the surrogate's own extrapolation failure "
        "severe enough, that the correction would not address the root "
        "cause identified in Section 6.13), but the "
        "paper's framing of covariate shift as a specific, structured "
        "violation of exchangeability - rather than an unstructured, "
        "unrecoverable one - is the correct lens for interpreting why the "
        "stress test's coverage collapse happens where and how it does.")

    add_paper_review(doc, 8,
        "Barber, R.F., Candès, E., Ramdas, A., Tibshirani, R.J. (2023). "
        "Conformal Prediction Beyond Exchangeability. Annals of Statistics, "
        "51(2), 816-845.",
        """
This paper generalizes conformal prediction to settings where exchangeability
fails entirely rather than failing in a specific, structured way (as in
covariate shift). It introduces weighted conformal prediction schemes that
can provide approximately valid coverage even without exchangeability, with
the degree of coverage degradation explicitly bounded in terms of a measure
of how far the actual data-generating process departs from exchangeability.
Crucially, the paper is explicit that no method can restore exact coverage
without exchangeability or a specific known structure to its violation - the
contribution is a principled characterization of how much coverage degrades,
and some partial mitigation strategies, not a full fix.
""",
        "This is the most directly relevant theoretical paper to this "
        "project's second major finding (the exchangeability stress test): "
        "it confirms, at a theoretical level, that this project's empirical "
        "result - coverage collapses once the test distribution moves "
        "sufficiently far outside the calibration distribution, and neither "
        "standard nor Mondrian CP can prevent this - is not a "
        "implementation flaw but an expected consequence of exchangeability "
        "violation that the theory itself predicts.")

    # -------------------- Section 2.2: Mondrian / Conditional Coverage --------------------
    add_section_heading(doc, "2.2 Mondrian and Conditional-Coverage Conformal Prediction")
    add_body(doc, """
Mondrian conformal prediction is a comparatively narrow subfield within the
broader conformal prediction literature - a handful of dedicated papers,
rather than the hundreds associated with, for instance, CQR or covariate-shift
robustness. This section covers all four papers identified as directly
relevant during the literature search. The relative thinness of this
subfield is itself worth noting explicitly: it supports this project's claim,
developed fully in Chapter 3, that testing Mondrian CP in a discrete-event
queueing domain is a genuine, previously unaddressed gap rather than a
crowded research area.
""")

    add_paper_review(doc, 9,
        "Vovk, V., Lindsay, D., Nouretdinov, I., Gammerman, A. (2003). "
        "Mondrian Confidence Machine. Technical Report, Royal Holloway, "
        "University of London.",
        """
This technical report introduces the Mondrian confidence machine, the origin
of the "Mondrian" name and the underlying idea used throughout this project:
rather than calibrating a single conformal quantile pooled across an entire
calibration set, partition the calibration set into categories (a
"Mondrian taxonomy," named for the resemblance of a partitioned plane to
Piet Mondrian's grid paintings) and calibrate a separate conformal quantile
per category. The guarantee this produces is stronger than standard
conformal prediction's marginal guarantee: it holds within each category
individually (group-conditional coverage), rather than only when averaged
across the whole population, at the cost of each category having a smaller
effective calibration sample than the pooled set would provide.
""",
        "This is the conceptual origin of every Mondrian CP result in this "
        "project. The nine-cell taxonomy used in src/uq/mondrian_cp.py "
        "(staffing tercile x arrival-rate tercile) is exactly this kind of "
        "Mondrian partition, chosen using only the two real covariates the "
        "DES output actually has, per the same principle this report's own "
        "Chapter 4 justifies: partition on real, available covariates, not "
        "invented ones.")

    add_paper_review(doc, 10,
        "Boström, H., Johansson, U. (2020). Mondrian Conformal Regressors. "
        "PMLR 128 (COPA 2020), pp. 114-133.",
        """
This paper extends Mondrian conformal prediction, originally developed
primarily in a classification context, to the regression setting in a
systematic way, addressing practical questions such as how to choose
category boundaries for continuous or near-continuous conditioning variables
(via binning) and how per-category calibration set size trades off against
the tightness of group-conditional coverage. It provides empirical results
across several regression benchmark datasets comparing pooled and Mondrian
calibration.
""",
        "This is the single most directly used paper in this project: "
        "src/uq/mondrian_cp.py implements exactly the regression-setting "
        "Mondrian conformal predictor this paper describes, including the "
        "same fundamental tradeoff this paper discusses - smaller "
        "per-category calibration sets in exchange for group-conditional "
        "rather than only marginal coverage. This project's own finding "
        "that Mondrian CP does not help n_patients (Chapter 6) is a direct, "
        "empirical instance of the finite-sample noise tradeoff this paper "
        "identifies as a known limitation of the method.")

    add_paper_review(doc, 11,
        "Boström, H., Johansson, U., Löfström, T. (2021). Mondrian "
        "Conformal Predictive Distributions. PMLR 152 (COPA 2021).",
        """
This follow-up paper extends Mondrian conformal prediction from producing
fixed-coverage intervals to producing full conformal predictive
distributions - a distribution over possible outcomes, from which any
interval or quantile can be read off after the fact, rather than committing
to a single alpha level at calibration time. The Mondrian partitioning
principle is unchanged; what changes is the object being calibrated, from an
interval to a distribution function.
""",
        "Not directly implemented in this project - src/uq/mondrian_cp.py "
        "produces intervals at a single, fixed alpha = 0.1, matching the "
        "GP baseline and standard CP for a clean comparison - but this "
        "paper is the natural next step flagged in this report's own "
        "Chapter 7 (Future Scope): a full predictive distribution would let "
        "an ER staffing decision-maker read off any confidence level of "
        "interest after the fact, rather than committing to 90% coverage in "
        "advance as this project does throughout.")

    add_paper_review(doc, 12,
        "Toccaceli, P., Gammerman, A. (2019). Combination of Inductive "
        "Mondrian Conformal Predictors. Machine Learning, 108, 489-510.",
        """
This paper addresses a practical problem that arises when multiple Mondrian
conformal predictors are trained - for instance, on different subsets of
available training data, or with different taxonomies - and their
predictions need to be combined into a single prediction region without
losing the individual predictors' validity guarantees. It develops
combination rules that preserve conformal validity under this kind of
model aggregation.
""",
        "Relevant background for the nine-cell Mondrian taxonomy used in "
        "src/uq/mondrian_cp.py, which is a single taxonomy rather than a "
        "combination of several, but this paper's treatment of how "
        "different category definitions interact with calibration validity "
        "informed the choice, documented in Chapter 4, to use only the two "
        "real covariates the DES actually outputs (staffing, arrival rate) "
        "rather than combining multiple partially-overlapping taxonomies.")

    # -------------------- Section 2.3: Surrogate Modeling & UQ --------------------
    add_section_heading(doc, "2.3 Surrogate Modeling and Uncertainty Quantification")
    add_body(doc, """
This section covers the base paper this project directly extends, together
with the theoretical foundations of the surrogate architectures (Gaussian
processes, gradient boosting, multi-layer perceptrons) and alternative UQ
approaches (deep ensembles) used or discussed as points of comparison
throughout this project.
""")

    add_paper_review(doc, 13,
        "Gopakumar, V. et al. (2026). Uncertainty Quantification of "
        "Surrogate Models Using Conformal Prediction. Machine Learning: "
        "Science and Technology.",
        """
This is the base paper this entire project is built to extend. Gopakumar et
al. validate conformal prediction as a practical, distribution-free
uncertainty quantification method for machine-learned surrogate models across
several physics-simulation domains - partial differential equation solvers,
magnetohydrodynamics simulations, weather models, and fusion-plasma
simulations - demonstrating that conformal prediction intervals achieve their
nominal coverage across these domains despite the underlying surrogates being
architecturally diverse (neural operators, among others) and the physical
systems being simulated being highly nonlinear. The paper explicitly and
prominently flags two limitations of its own results as untested rather than
resolved: first, that the coverage guarantee demonstrated is marginal, not
conditional, and could in principle mask conditional miscalibration within
subpopulations of the test distribution the paper did not specifically probe
for; second, that all validation was performed within the training
distribution, leaving conformal prediction's behavior under exchangeability
violation (distribution shift between calibration and deployment) as an open
question the paper does not address.
""",
        "This paper is this report's entire reason for existing. Its two "
        "explicitly flagged limitations frame this project's broader "
        "empirical work: Chapter 6 examines the marginal coverage question "
        "in depth as this report's central finding, and Section 6.13 "
        "summarizes a second, secondary stress test of the exchangeability "
        "question run as part of the same project, in both cases in a "
        "domain - discrete-event queueing simulation of an ER - "
        "categorically different from the physics-simulation domains "
        "Gopakumar et al. tested.")

    add_paper_review(doc, 14,
        "Kennedy, M.C., O'Hagan, A. (2001). Bayesian Calibration of "
        "Computer Models. JRSS-B, 63(3), 425-464.",
        """
This is a foundational paper in the computer-model calibration literature,
introducing a Bayesian framework for combining a computationally expensive
simulator's output with real observational data, explicitly separating model
discrepancy (systematic error between the simulator and reality) from
observation noise and parameter uncertainty. Although developed before the
modern surrogate-modeling and conformal-prediction literature, it establishes
the conceptual vocabulary - simulator, emulator (an early term for what this
project calls a surrogate), calibration, and discrepancy - that much of the
later surrogate-modeling literature, including this project, builds on.
""",
        "Provides the conceptual and historical grounding for this "
        "project's own DES-to-surrogate pipeline: the distinction this "
        "paper draws between a simulator's systematic discrepancy from "
        "reality and pure statistical noise maps directly onto this "
        "project's own careful separation (Chapter 5) between the DES's "
        "real-data-calibrated arrival process and its "
        "literature-calibrated service-time distributions - a discrepancy "
        "this project discloses explicitly rather than presenting both as "
        "equally data-derived.")

    add_paper_review(doc, 15,
        "Rasmussen, C.E., Williams, C.K.I. (2006). Gaussian Processes for "
        "Machine Learning. MIT Press.",
        """
This is the standard reference textbook for Gaussian process (GP) regression
and classification, covering the theory of GPs as distributions over
functions, the role of the covariance (kernel) function in encoding
assumptions about function smoothness, exact inference and its O(n^3)
computational cost, and sparse/approximate methods developed to mitigate that
cost for larger datasets.
""",
        "Direct theoretical basis for src/uq/gp_baseline.py, this project's "
        "GP uncertainty quantification baseline. The O(n^3) scaling this "
        "book documents is exactly why the GP baseline in this project is "
        "trained on a 1,000-point subsample rather than the full "
        "calibration set (Chapter 4), and why the GP-versus-CP computation "
        "time comparison in this project's full_comparison.py finds "
        "conformal prediction's calibration cost roughly 650-1000 times "
        "faster than a GP refit - a direct empirical demonstration of the "
        "scaling behavior this textbook derives theoretically.")

    add_paper_review(doc, 16,
        "Friedman, J.H. (2001). Greedy Function Approximation: A Gradient "
        "Boosting Machine. Annals of Statistics, 29(5), 1189-1232.",
        """
This paper introduces gradient boosting as a general framework for additive
function approximation, recasting boosting (originally developed as an
ensemble classification technique) as gradient descent in function space
rather than parameter space. It derives specific gradient-boosting algorithms
for least-squares, least-absolute-deviation, and Huber-loss regression, and
for multiclass classification, and shows that when the individual additive
components are shallow regression trees, the resulting method is
competitive, robust to outliers and irrelevant features, and computationally
efficient relative to comparably accurate alternatives.
""",
        "This is the theoretical basis for scikit-learn's "
        "HistGradientBoostingRegressor, the primary surrogate architecture "
        "used throughout this project (src/surrogate/train_surrogate.py "
        "and, for the quantile variant used in CQR, "
        "src/surrogate/train_quantile_surrogates.py). The tree-based "
        "structure this paper describes is also the direct explanation for "
        "this report's Section 6.13 finding that the "
        "gradient-boosting surrogate's predictions freeze at a constant "
        "value once queried outside its training range - a well-known "
        "property of tree ensembles that this project traces empirically "
        "to a specific, verified mechanism rather than treating as an "
        "unexplained black-box failure.")

    add_paper_review(doc, 17,
        "Lakshminarayanan, B., Pritzel, A., Blundell, C. (2017). Simple and "
        "Scalable Predictive Uncertainty Estimation using Deep Ensembles. "
        "NeurIPS 30, 6402-6413.",
        """
This paper proposes deep ensembles - training several neural networks
independently with different random initializations and treating the spread
of their predictions as an uncertainty estimate - as a simple, scalable
alternative to Bayesian neural network approximations for uncertainty
quantification in deep learning. It shows empirically that this simple
approach often matches or exceeds more sophisticated Bayesian approximation
methods on predictive uncertainty benchmarks, while being substantially
easier to implement and parallelize.
""",
        "Provides an important point of contrast for this report's "
        "framing of why conformal prediction was chosen over alternative "
        "UQ approaches (Chapter 1, Chapter 4): unlike deep ensembles, which "
        "provide no finite-sample coverage guarantee and require training "
        "multiple independent models, conformal prediction wraps a single "
        "already-trained surrogate and provides a distribution-free "
        "guarantee - a meaningfully different trade-off that this report "
        "makes explicit rather than leaving deep ensembles as an "
        "unexamined alternative.")

    add_paper_review(doc, 18,
        "Abdar, M. et al. (2021). A Review of Uncertainty Quantification in "
        "Deep Learning: Techniques, Applications and Challenges. "
        "Information Fusion, 76, 243-297.",
        """
This is a broad, widely cited survey of uncertainty quantification methods
for deep learning, covering Bayesian approaches, ensemble methods, and
distribution-free methods (including conformal prediction), organized around
the distinction between aleatoric uncertainty (irreducible noise in the data)
and epistemic uncertainty (uncertainty about the model itself, reducible with
more data). It also surveys applications across computer vision, natural
language processing, and safety-critical domains such as medical diagnosis
and autonomous driving, and discusses open challenges including
computational cost, calibration under distribution shift, and the lack of
standardized evaluation protocols across the UQ literature.
""",
        "Used to situate this project's specific choice of conformal "
        "prediction within the much broader UQ landscape this survey maps "
        "out (Chapter 1), and to frame this project's own aleatoric/"
        "epistemic distinction: the DES's own stochasticity (arrival "
        "randomness, service-time variance) is aleatoric uncertainty that "
        "any UQ method must characterize, while the surrogate's imperfect "
        "fit to the DES is an additional epistemic component - a "
        "distinction made explicit in this project's early design "
        "decisions (documented in this project's PROJECT_LOG.md, Week "
        "6-7 entry) even before this survey was formally reviewed.")

    # -------------------- Section 2.4: Queueing Theory & ED Ops --------------------
    add_section_heading(doc, "2.4 Queueing Theory and Emergency Department Operations Research")
    add_body(doc, """
This section covers the queueing-theory and health-operations-research
literature that motivates and contextualizes this project's choice of a
discrete-event simulation, rather than a closed-form analytic queueing model,
as the underlying data-generating process the surrogate approximates.
""")

    add_paper_review(doc, 19,
        "Green, L.V., Soares, J., Giglio, J.F., Green, R.A. (2006). Using "
        "Queueing Theory to Increase the Effectiveness of Emergency "
        "Department Provider Staffing. Academic Emergency Medicine, 13(1), "
        "61-68.",
        """
This paper applies a Lag-SIPP (Stationary Independent Period-by-Period)
queueing analysis to real ED arrival data from an urban hospital to identify
provider staffing patterns that reduce the fraction of patients who leave
without being seen. It demonstrates that relatively modest, well-targeted
increases in provider staffing hours during identified high-demand periods
produce disproportionately large reductions in patients leaving without
being seen, compared to uniformly increasing staffing across all hours.
""",
        "Directly parallels this project's own staffing-scenario design: "
        "the staffing-capacity and arrival-rate-multiplier scenario grid "
        "used to generate surrogate training data (src/surrogate/"
        "generate_training_data.py) and the staffing-tercile x "
        "arrival-tercile Mondrian taxonomy (src/uq/mondrian_cp.py) both "
        "encode the same real-world insight this paper demonstrates "
        "empirically: staffing effectiveness is highly dependent on the "
        "interaction between staffing level and demand, not staffing level "
        "alone, which is exactly why this project's core finding - "
        "coverage failure concentrated specifically in the "
        "understaffed/high-demand category - has real operational meaning "
        "rather than being an arbitrary partition choice.")

    add_paper_review(doc, 20,
        "Green, L.V. Queueing Analysis in Healthcare. Book chapter, "
        "Columbia Business School.",
        """
This chapter provides an accessible overview of queueing-theoretic models
applied to healthcare capacity and staffing problems, covering the basic
M/M/s and M/G/1-type models, the Erlang loss and delay formulas, and their
use in estimating offered load and required server (staff or bed) counts
under target service-level constraints.
""",
        "This is the direct theoretical basis for the offered-load (Erlang) "
        "capacity calculation documented in this project's PROJECT_LOG.md "
        "(Week 4-5 entry) and inline in src/des/er_simulation.py: department "
        "A's default capacity of 30 and department B's default capacity of "
        "14 were both derived from an Erlang-style offered-load calculation "
        "using each department's real calibrated arrival rate and "
        "literature-standard service times, following exactly the "
        "methodology this chapter describes, rather than chosen arbitrarily.")

    add_paper_review(doc, 21,
        "Hu, X. et al. (2018). Applying Queueing Theory to the Study of "
        "Emergency Department Operations: A Survey and a Discussion of "
        "Comparable Simulation Studies. International Transactions in "
        "Operational Research.",
        """
This survey compares analytic queueing-theory approaches and discrete-event
simulation approaches to modeling ED operations, discussing the
circumstances under which each is preferable: analytic queueing models offer
speed and closed-form insight but require restrictive distributional
assumptions that often do not hold for real ED service-time distributions,
while DES relaxes those assumptions at higher computational cost and lower
interpretability of any single closed-form result.
""",
        "Directly supports this project's methodological choice (Chapter 4) "
        "to build a discrete-event simulation rather than rely on a "
        "closed-form queueing model as the underlying data-generating "
        "process: the acuity-dependent, non-exponential service-time "
        "structure this project calibrates from literature-standard "
        "log-normal parameters by ESI level is exactly the kind of "
        "structure this survey identifies as poorly suited to closed-form "
        "analytic queueing treatment.")

    add_paper_review(doc, 22,
        "Performance Evaluation of a M/G/1 Queue Model for Patient Flow in "
        "a Health Care System. Mathematical Modelling of Engineering "
        "Problems (IIETA).",
        """
This paper develops and analyzes an M/G/1 queueing model (Poisson arrivals,
general service-time distribution, a single server) for patient flow in a
healthcare setting, deriving performance measures such as expected queue
length and waiting time under this general-service-time relaxation of the
more restrictive M/M/1 model.
""",
        "Provides a useful analytic comparison point for this project's own "
        "DES resource model: the DES's simplification from two nested "
        "resource pools (doctors and beds) to a single combined capacity "
        "resource (documented in PROJECT_LOG.md, Week 4-5 entry) makes it "
        "structurally closer to an M/G/c queue (a multi-server generalization "
        "of the M/G/1 model this paper studies) than to a more complex "
        "network-of-queues representation, a simplification this project "
        "justifies on the grounds of having no real data to calibrate a "
        "doctor-to-bed ratio.")

    add_paper_review(doc, 23,
        "Decision Support for the Optimization of Provider Staffing for "
        "Hospital Emergency Departments with a Queue-Based Approach. "
        "PMC6947400.",
        """
This paper develops a queueing-based decision-support tool for ED provider
staffing optimization, using queueing performance measures as the objective
that a staffing schedule is optimized against, rather than relying purely on
historical volume-matching heuristics.
""",
        "Relevant to the staffing dimension of this project's Mondrian CP "
        "taxonomy: this paper's framing of staffing as a decision variable "
        "to be optimized against a queueing-performance objective is "
        "conceptually the same framing underlying this project's own "
        "staffing-capacity scenario sweep, though this project's "
        "contribution is specifically about quantifying uncertainty in the "
        "resulting performance predictions, not about optimizing the "
        "staffing decision itself.")

    # -------------------- Section 2.5: DES & ED-specific ML --------------------
    add_section_heading(doc, "2.5 Discrete-Event Simulation and ED-Specific Machine Learning")
    add_body(doc, """
This final section covers discrete-event simulation studies of emergency
departments directly comparable to this project's own DES, together with
recent (2023-2025) machine-learning applications to ED operations that
situate this project within current, rather than only historical, literature.
""")

    add_paper_review(doc, 24,
        "A Simulation-Based Optimization Approach for the Calibration of a "
        "Discrete Event Simulation Model of an Emergency Department. "
        "arXiv:2102.00945 (2021).",
        """
This paper develops a simulation-based optimization procedure for
calibrating a discrete-event simulation model of an ED against real
operational data, framing calibration as an optimization problem where the
objective function is the deviation between simulated and real performance
measures, and using this procedure to recover parameters (such as resource
counts) that are not directly observable in the available data.
""",
        "Directly parallel to this project's own DES calibration "
        "methodology (Chapter 5): both this paper and this project face the "
        "same core problem of an ED DES needing calibration against real "
        "but incomplete data, and both adopt a validation-against-real-"
        "aggregate-statistics approach (this project validates against real "
        "daily patient volume, achieving 91.0% agreement, documented in "
        "PROJECT_LOG.md's Week 4-5 entry and src/des/validate.py) rather "
        "than assuming a hand-specified parameterization is correct without "
        "checking it against data.")

    add_paper_review(doc, 25,
        "Discrete Event Simulation for Emergency Department Modelling: A "
        "Systematic Review of Validation Methods. ScienceDirect (2022).",
        """
This systematic review surveys how published ED discrete-event simulation
studies validate their models against real-world data, finding substantial
variation in validation rigor across the literature - from no formal
validation at all to detailed statistical comparison against held-out real
performance data - and argues for more consistent, transparent validation
reporting as a methodological standard for the field.
""",
        "Directly relevant to justifying this project's own validation "
        "approach: the explicit 91.0% daily-volume match reported in "
        "PROJECT_LOG.md (Week 4-5) and the corresponding validation for "
        "Department B (88.6% match) represent the kind of quantitative, "
        "transparently reported validation this review argues the field "
        "should adopt as standard practice, rather than treating DES "
        "calibration as self-evidently correct without independent "
        "verification.")

    add_paper_review(doc, 26,
        "Discrete Event Simulation Modelling for an Improved Patient Flow "
        "at the Emergency Department, Sygehus Lillebælt, Kolding. "
        "PMC3327033.",
        """
This case study applies discrete-event simulation to a real Danish hospital
ED to identify patient-flow bottlenecks and evaluate proposed operational
changes before implementation, demonstrating DES's practical use as a
decision-support tool for hospital administrators rather than purely an
academic modeling exercise.
""",
        "A real-world precedent for the SimPy-based ED modeling approach "
        "used throughout this project (src/des/er_simulation.py), "
        "demonstrating that the DES-based methodology this project applies "
        "to uncertainty quantification is built on a modeling approach "
        "already established as practically useful for real hospital "
        "decision-making, not a purely synthetic academic exercise.")

    add_paper_review(doc, 27,
        "A Simulation-Based Optimization Approach for Analyzing the "
        "Ambulance Diversion Phenomenon in an Emergency Department Network. "
        "arXiv:2108.04162.",
        """
This paper uses discrete-event simulation to study ambulance diversion - the
practice of redirecting incoming ambulances away from an overloaded ED to a
neighboring facility - across a network of interconnected emergency
departments, modeling how capacity constraints at one site propagate to
neighboring sites.
""",
        "Relevant to the surge-scenario framing of this project's "
        "exchangeability stress test (src/uq/exchangeability_stress_test.py "
        "and its MLP counterpart): both this paper and this project's "
        "stress test are concerned with ED behavior under extreme demand "
        "surge, though this project's specific contribution is about "
        "surrogate and uncertainty-quantification behavior under that "
        "surge, not about the ambulance-diversion operational response "
        "this paper studies.")

    add_paper_review(doc, 28,
        "Machine Learning-Based Prediction of Hospital Prolonged Length of "
        "Stay Admission at Emergency Department: A Gradient Boosting "
        "Algorithm Analysis. Frontiers in Artificial Intelligence (2023).",
        """
This paper applies gradient boosting to predict prolonged length-of-stay
admissions at an ED from patient-level features available at triage,
evaluating the model's predictive performance and discussing its potential
use in early identification of patients likely to require extended ED stays.
""",
        "Uses the same model family (gradient boosting) as this project's "
        "primary surrogate architecture (src/surrogate/train_surrogate.py), "
        "applied to a related but distinct prediction task - patient-level "
        "length-of-stay classification, versus this project's "
        "scenario-level regression of aggregate daily performance metrics. "
        "The shared architecture choice reflects gradient boosting's "
        "broader status, evident across the recent ED-ML literature "
        "surveyed here, as a strong default choice for structured, "
        "tabular healthcare prediction problems.")

    add_paper_review(doc, 29,
        "An Artificial Intelligence-Based Framework for Predicting "
        "Emergency Department Overcrowding: Development and Evaluation "
        "Study. arXiv:2504.18578 (2025).",
        """
This recent paper develops a machine-learning framework for predicting ED
waiting-room occupancy at hourly and daily time scales, intended to support
proactive staffing decisions and early intervention before overcrowding
occurs, evaluated against real ED occupancy data.
""",
        "Positions this project within current (2025-2026) literature "
        "rather than only older, established queueing-theory work: this "
        "paper's goal - using predictive modeling to support proactive ED "
        "staffing decisions - is the same broad motivation underlying this "
        "project's own surrogate-plus-uncertainty-quantification pipeline, "
        "though this project's specific contribution (testing conformal "
        "prediction's marginal-coverage and exchangeability assumptions in "
        "this domain) is a methodological question this paper does not "
        "address.")

    add_paper_review(doc, 30,
        "Machine Learning-Based Triage to Identify Low-Severity Patients "
        "with a Short Discharge Length of Stay in Emergency Department. "
        "PMC9123815.",
        """
This paper applies machine learning to triage-stage data to identify
low-severity patients likely to have a short ED length of stay, with the
goal of supporting fast-track routing decisions at intake.
""",
        "Relevant given this project's DES also explicitly models Emergency "
        "Severity Index (ESI) acuity mix as part of its arrival-process "
        "calibration (src/utils/extract_distributions.py): both this paper "
        "and this project treat ESI-level acuity as a first-class variable "
        "affecting downstream flow and outcomes, reinforcing the "
        "methodological choice to calibrate the DES's ESI mix directly "
        "from real data rather than assuming a uniform acuity distribution.")

    add_section_heading(doc, "2.6 Synthesis and Positioning of This Project")
    add_body(doc, """
Read together, these thirty papers trace two literatures that, prior to this
project, had not been directly connected: the conformal prediction and
Mondrian conformal prediction literature (Sections 2.1-2.2), developed and
validated almost entirely in general machine learning benchmark settings or,
in the specific case of the base paper (Gopakumar et al., 2026), in physics
simulation domains; and the queueing-theory, discrete-event simulation, and
ED operations research literature (Sections 2.4-2.5), which has extensively
studied ED capacity and staffing problems but has not, to the extent this
review was able to establish, applied conformal prediction's finite-sample
coverage guarantees to surrogate models of discrete-event queueing
simulations specifically. The surrogate-modeling and uncertainty-
quantification literature reviewed in Section 2.3 supplies the theoretical
and architectural vocabulary (Gaussian processes, gradient boosting, deep
ensembles) that bridges the two.

This project's contribution, developed in full in Chapter 3, is precisely
this connection: applying conformal prediction and Mondrian conformal
prediction to a discrete-event simulation surrogate calibrated on real
hospital data, and testing both of the base paper's explicitly flagged,
previously untested limitations - marginal-only coverage and exchangeability
fragility - in this new domain.
""")

    add_section_heading(doc, "2.7 Critical Assessment of the Reviewed Literature")
    add_body(doc, f"""
A literature review that only summarizes is incomplete without an honest
assessment of the reviewed work's own limitations, and of how those
limitations shaped this project's methodological choices. This section
provides that assessment across the five reviewed areas.

The conformal prediction foundations literature (Section 2.1) is
theoretically mature and its core coverage guarantee ({cite('vovk2005')}; {cite('shafer2008')}) is not in serious dispute - the proofs
are elementary rank-based arguments that do not depend on contested
modeling assumptions. Its main limitation, acknowledged within the
literature itself rather than only by this report, is that the
exchangeability assumption underlying every coverage guarantee in this
family is a property of the data-generating process that can be stated
precisely but not directly tested from a finite sample - a calibration set
and test set can look exchangeable in every measurable respect and still
fail to be, if the true underlying shift is small enough to escape
detection at the available sample size. This project's own exchangeability
stress test (Section 4.5.2, summarized in Section 6.13)
sidesteps this untestability problem by constructing a shift large and
deliberate enough that its effect on coverage is unambiguous, rather than
attempting to detect a subtle, naturally occurring shift - a pragmatic
choice given this project's scope, but one that leaves open the harder
question of how large a naturally occurring, undetected shift would need to
be before it meaningfully degraded real-world coverage.

The Mondrian and conditional-coverage literature (Section 2.2) is, as
already noted in Section 2.2's introduction, comparatively thin - a
consequence, plausibly, of Mondrian CP's own central limitation (smaller
per-category calibration samples, directly observed in this project's own
n_patients exception, Section 6.5.6) making it a
less immediately attractive default than pooled calibration for
practitioners without a specific, known source of conditional
miscalibration to target. This project's own results (Chapter 6) arguably
strengthen the case for this literature's continued development: a
practitioner without prior knowledge of where a pooled quantile might be
conditionally miscalibrated (as would genuinely be the case in an
unfamiliar deployment) would have no principled way to know whether
Mondrian calibration is worth its finite-sample cost, absent exactly the
kind of exploratory per-category analysis this project performs in Section
6.5.

The surrogate-modeling and uncertainty-quantification literature (Section
2.3) is broad enough that this review necessarily samples rather than
exhaustively covers it - the Abdar et al. {cite('abdar2021')} survey alone cites several
hundred papers, of which this review's six entries in Section 2.3 represent
a deliberately narrow, project-relevant slice (the base paper, and the
specific architectures and alternative UQ methods this project directly
uses or contrasts against) rather than a claim to comprehensive coverage of
the UQ field. A specific limitation worth naming: this project did not
implement or compare against Bayesian neural networks or deep ensembles
{cite('lakshminarayanan2017')} empirically, only discussed
them as points of theoretical contrast (Section 4.4, Section 1.3) - a
direct empirical comparison against a deep-ensemble baseline, alongside the
GP baseline this project does implement, would have strengthened the
"why conformal prediction specifically" argument beyond the theoretical
case made in Chapter 1.

The queueing-theory and ED operations research literature (Section 2.4) is
overwhelmingly oriented toward staffing optimization and operational
decision support (Green et al., 2006; the PMC6947400 queue-based staffing
paper) rather than toward the specific methodological question this project
investigates (uncertainty quantification for a surrogate of a queueing
simulation). This is not a gap in that literature's own quality - it
reflects a genuinely different research question - but it does mean this
project's own queueing-theoretic grounding (the Erlang-load capacity
derivation in Section 4.2.1) is comparatively simple relative to the more
sophisticated queueing models (time-varying arrival-rate models, network-
of-queues formulations) that literature has developed for staffing
optimization specifically. A more sophisticated queueing-theoretic
foundation was judged unnecessary for this project's own research question
(Chapter 3), which concerns UQ methodology rather than staffing
optimization itself, but this is a deliberate scoping choice worth stating
rather than an oversight.

The discrete-event simulation and ED-specific machine learning literature
(Section 2.5) is the most heterogeneous of the five areas reviewed, mixing
DES calibration methodology papers, DES case studies, and ED-specific ML
prediction papers that do not share a common methodology or evaluation
standard - the systematic review of DES validation methods (Section 2.5,
entry 25) explicitly documents this heterogeneity as a field-wide issue,
finding substantial variation in how rigorously published ED DES models are
validated against real data. This project's own validation approach
(Section 6.1: an explicit, quantified 91.0 percent match to real daily
volume, with the specific mechanism behind the residual 9 percent gap
identified and explained rather than left unexplained) was deliberately
designed to sit at the more rigorous end of the range this systematic
review documents, rather than assumed to be adequate by default.
""")


# --------------------------------------------------------------------------
# Chapter 3: Research Gap and Problem Statement (shared)
# --------------------------------------------------------------------------

def build_chapter3_research_gap(doc):
    add_chapter_heading(doc, 3, "Research Gap and Problem Statement")

    add_section_heading(doc, "3.1 Summary of the Research Gap")
    add_body(doc, f"""
Chapter 2's review identifies a specific, previously unaddressed intersection
between two established literatures. Conformal prediction and, more
specifically, Mondrian conformal prediction (Sections 2.1-2.2) have been
developed and validated primarily on general machine learning benchmarks and,
in the case of this project's base paper (Gopakumar et al. {cite('gopakumar2026')}), on
physics-simulation surrogate models spanning partial differential equations,
magnetohydrodynamics, weather, and fusion-plasma domains. Separately,
queueing theory, discrete-event simulation, and machine learning have all
been applied extensively to emergency department operations (Sections
2.4-2.5), but - to the extent this project's literature search was able to
establish - none of that body of work has applied conformal prediction's
finite-sample coverage guarantees to a discrete-event simulation surrogate
specifically, nor tested whether the guarantee's known limitations survive
the transition from a physics-simulation setting to a queueing-simulation
one.

The base paper motivating this project, Gopakumar et al. {cite('gopakumar2026')}, is explicit
that its own validation leaves two limitations untested: first, that its
demonstrated coverage guarantee is marginal rather than conditional, and a
pooled conformal interval could in principle mask systematic
miscalibration within subpopulations of the physics-simulation test
distributions the paper studied without the paper's own evaluation being
positioned to detect it; second, that all of the paper's validation was
performed within the training distribution of its surrogate models, leaving
open the question of how conformal prediction's coverage guarantee behaves
under distribution shift between calibration and deployment - a realistic
concern for any surrogate deployed beyond the exact scenarios it was trained
on. Both limitations are framed by the base paper as open questions for
future work, not as resolved or dismissed concerns.

This project's research gap, stated directly, is this: neither of these two
explicitly flagged limitations of conformal prediction for surrogate-model
uncertainty quantification had been empirically tested in a discrete-event,
queueing-based simulation domain prior to this work, despite such domains
(hospital operations, call centers, manufacturing systems, and other
stochastic service systems) being exactly the kind of setting where
operational decisions - staffing levels, capacity investments - depend on
uncertainty-aware predictions, and where the consequences of an unreliable
coverage guarantee are practically, not just theoretically, significant.
""")

    add_section_heading(doc, "3.2 Why Discrete-Event Queueing Simulation Is a Meaningfully Different Test")
    add_body(doc, """
It is worth being explicit about why testing conformal prediction's
limitations in a discrete-event queueing domain constitutes a genuine
extension of Gopakumar et al.'s results, rather than a mechanical
repetition of their experiments on a different dataset. The physics
simulation domains the base paper studies - PDE solvers, magnetohydrodynamics,
weather, and fusion-plasma simulations - are governed by continuous,
typically smooth, deterministic-given-initial-conditions dynamics, and the
uncertainty a surrogate model must characterize in that setting is
predominantly epistemic: uncertainty about how well the surrogate
approximates a fixed, in-principle-fully-determined physical process. A
discrete-event queueing simulation is a categorically different kind of
system: outcomes are driven by explicitly stochastic processes (random
arrival times, random triage assignment, random service durations) even
when every scenario parameter is held fixed, meaning the DES itself - not
just the surrogate approximating it - is a source of irreducible aleatoric
uncertainty that has no analogue in a deterministic PDE solve. A surrogate
trained to predict DES outputs must therefore have its residual variance
absorb both the surrogate's own epistemic approximation error and the DES's
aleatoric stochasticity, in a proportion that is itself not constant across
the input space (a lightly loaded, well-staffed scenario produces less
variable outcomes than an overloaded, understaffed one). This
heteroscedasticity - directly evidenced in this project by the systematic
gap between pooled and Mondrian CP coverage documented in Chapter 6 - is
plausibly more severe and more structurally different in a queueing
simulation than in the physics domains previously tested, which is precisely
why the marginal-coverage limitation is worth testing here specifically
rather than assumed to generalize automatically from a physics-domain
result.

The exchangeability question is similarly domain-specific in its practical
stakes. A physics simulation's operating envelope (e.g., a range of plasma
confinement parameters) is typically defined by the physical or engineering
constraints of the system being modeled and may change relatively slowly.
An ED's operating envelope, by contrast, can shift abruptly and
unpredictably - a mass-casualty event, a disease outbreak, a public health
emergency - in ways that push real-world conditions outside whatever range a
surrogate was trained and calibrated on, precisely when reliable uncertainty
quantification matters most for triage and staffing decisions. Testing
conformal prediction's exchangeability assumption under exactly this kind of
demand-surge scenario (Section 6.13) is
therefore not an arbitrary stress test but a probe of the specific failure
mode most operationally relevant to this project's own application domain.
""")

    add_section_heading(doc, "3.3 Problem Statement")
    add_body(doc, """
Given a discrete-event simulation of an emergency department, calibrated on
real arrival and triage-acuity data, and a machine-learned surrogate model
trained to approximate that simulation's scenario-level outputs, this
project addresses the following problem: does conformal prediction, applied
to the surrogate's residuals, retain a valid finite-sample coverage
guarantee (i) uniformly across operationally meaningful subgroups of the
scenario space, and (ii) when the deployment scenario distribution shifts
outside the range the calibration data was drawn from? And, where the answer
to (i) is negative for standard (pooled) conformal prediction, does Mondrian
conformal prediction - calibrating separately per subgroup rather than
pooling - restore valid coverage within the subgroups where a genuine
conditional miscalibration exists, without an unacceptable cost in
statistical efficiency from smaller per-subgroup calibration samples?
""")

    add_section_heading(doc, "3.4 Objectives")
    add_body(doc, """
This project's objectives, carried through in the methodology (Chapter 4),
implementation (Chapter 5), and results (Chapter 6) that follow, are to:

1. Build and validate a discrete-event simulation of a hospital emergency
department, calibrated on real arrival-pattern and triage-acuity data,
against real aggregate daily-volume statistics.

2. Train a machine-learned surrogate model to approximate the calibrated
simulation's scenario-level output metrics, evaluated on standard regression
accuracy metrics (MAE, RMSE, R-squared).

3. Implement and compare three uncertainty quantification approaches on top
of the trained surrogate - a Gaussian process baseline, standard conformal
prediction, and Mondrian conformal prediction - evaluated on identical
calibration and test data at a common nominal coverage target.

4. Test whether standard conformal prediction's marginal coverage guarantee
masks conditional miscalibration within operationally meaningful subgroups
of the scenario space (staffing level, arrival-rate regime), and whether
Mondrian conformal prediction corrects any such miscalibration found.

5. Test conformal prediction's behavior under a controlled violation of the
exchangeability assumption, by evaluating coverage as the test distribution's
arrival-rate regime is pushed progressively outside the calibration range.

6. Establish the statistical robustness of all findings through repeated
evaluation across independent calibration/test draws with formal
significance testing, rather than relying on a single data split.

7. Test the generality of all findings along two further axes: a second
surrogate architecture, to check whether results are specific to one model
family, and a second, independent hospital department, to check whether
results generalize beyond a single site.
""")

    add_section_heading(doc, "3.5 Scope and Delimitations")
    add_body(doc, """
This project's scope is deliberately bounded in several respects that are
stated here explicitly rather than left implicit. The discrete-event
simulation models a single emergency department's arrival, triage, and
combined bed-and-provider service process; it does not model inter-hospital
transfer networks, ambulance diversion, or downstream inpatient admission
processes. Service-time distributions are calibrated from literature-standard
parameters by triage acuity level, not derived from the real dataset used in
this project, because that dataset (documented fully in Chapter 5) contains
no length-of-stay or treatment-duration field - this distinction is
maintained explicitly throughout this report rather than presented as
uniformly data-derived. The surrogate models studied are a gradient-boosting
regressor (primary) and a multi-layer perceptron (robustness check); other
architectures (Gaussian process surrogates used as the point predictor
itself, rather than as a UQ baseline; graph neural networks; other
tree-ensemble variants) are outside this project's scope. The uncertainty
quantification methods compared are a Gaussian process baseline, standard
conformal prediction, Mondrian conformal prediction, conformalized quantile
regression, and Mondrian-CQR; Bayesian neural networks and deep ensembles are
discussed in the literature review (Chapter 2) as points of comparison but
are not implemented or evaluated empirically in this project. Generalization
testing covers two of the three emergency departments present in the
underlying dataset; the third department is not evaluated, for reasons of
project scope and time rather than any expectation that it would behave
differently from the two that were tested.
""")


# --------------------------------------------------------------------------
# Chapter 4: System Design and Methodology (shared)
# --------------------------------------------------------------------------

def build_chapter4_methodology(doc):
    add_chapter_heading(doc, 4, "System Design and Methodology")

    add_section_heading(doc, "4.1 Pipeline Overview")
    add_body(doc, """
The system built for this project consists of four stages, each depending
on the outputs of the previous one: a calibrated discrete-event simulation
(Section 4.2) generates labeled scenario data; a surrogate regression model
(Section 4.3) is trained on that data to approximate the simulation's
outputs directly from scenario parameters; three (later five) uncertainty
quantification methods (Section 4.4) are applied on top of the trained,
frozen surrogate to produce calibrated prediction intervals; and a set of
robustness checks (Section 4.5) - repeated evaluation, a second surrogate
architecture, a second hospital department - establish whether the
resulting findings are stable and general rather than artifacts of one
particular random split, model, or site. Every stage after the DES itself
depends only on the DES's output distribution, not on its internal
mechanics, which is precisely what makes the surrogate-plus-UQ approach
practical: the expensive stochastic simulator is queried many times up
front to generate training and calibration data, and every subsequent use
(scenario sweeps, staffing comparisons, the statistical robustness checks in
Section 4.5) queries the cheap surrogate instead.
""")

    add_section_heading(doc, "4.2 Discrete-Event Simulation Design")
    add_body(doc, """
The emergency department is modeled as a SimPy discrete-event simulation
with three sequential stages per patient: arrival, triage/acuity assignment,
and service, the last of which consumes a single shared capacity resource
representing a combined bed-and-provider slot for the full duration of a
patient's visit. An earlier design used two nested resource pools (a
doctor pool and a bed pool); this was deliberately simplified to a single
combined resource because the dataset used to calibrate this project (Section
4.2.3) contains no information from which a realistic doctor-to-bed ratio
could be derived, and introducing a second, uncalibrated resource pool would
add an unverified assumption without adding insight relevant to this
project's actual research question, which concerns uncertainty
quantification methodology rather than fine-grained hospital operations
modeling. This simplification is also the standard reduction used in the
M/G/c queueing literature (Section 2.4) when a detailed multi-resource model
cannot be justified by available data.

Patient arrivals follow a non-homogeneous process whose hourly and daily
rate is calibrated directly from real data (Section 4.2.3): rather than a
single constant arrival rate, the simulation draws arrivals according to
hour-of-day and day-of-week rate multipliers extracted from the dataset,
reflecting the well-documented reality that ED arrival volume is far from
uniform across a 24-hour cycle. Each simulated day runs for a fixed 24-hour
window; patients still queued or in service when that window ends are
right-censored out of that day's completed-visit statistics rather than
carried forward into a following day, since each simulated day is intended
as one independent sample of a scenario for surrogate training purposes
(Section 4.3), not as a continuous multi-day rollout. This right-censoring
mechanism is a deliberate, documented modeling choice, not an oversight, and
it is the direct explanation for two findings elsewhere in this report:
the DES's simulated daily patient count running slightly
below the real calibrated rate even at matched capacity (Section 4.2.4), and
the non-monotonic behavior of the 95th-percentile wait time under extreme
demand surge documented in Section 6.13.

Upon arrival, each patient is assigned an Emergency Severity Index (ESI)
acuity level (1, most acute, through 5, least acute) drawn according to the
real, calibrated ESI mix for the department being simulated, and a service
duration drawn from a log-normal distribution whose parameters depend on
that acuity level. The log-normal parameters used are literature-standard
values by ESI level (Section 4.2.3), not derived from the dataset used in
this project, since that dataset contains no length-of-stay field.

Two scenario-level parameters govern each simulated day and are the only two
inputs the downstream surrogate model (Section 4.3) ever sees: staffing
capacity (the number of concurrent combined bed-and-provider slots
available) and an arrival-rate multiplier (a scalar applied uniformly to the
calibrated real arrival rate, allowing the simulation to be run under
higher- or lower-than-observed demand). Four scenario-level output metrics
are recorded per simulated day: the total number of patients completing
service (n_patients), the mean wait time before service begins
(mean_wait_minutes), the mean total time in the system, from arrival to
service completion (mean_total_minutes), and the 95th percentile of wait
time (p95_wait_minutes) - included specifically because a tail statistic
computed from one stochastic simulated day is expected to be harder to
predict than a mean, making it a useful stress case for comparing
uncertainty quantification methods' ability to produce informatively wide,
rather than falsely narrow, intervals.
""")

    add_section_heading(doc, "4.2.1 Default Capacity via Offered-Load (Erlang) Calculation")
    add_body(doc, """
Rather than choosing a default staffing capacity arbitrarily, this project
derives it from an offered-load calculation in the style of the Erlang
queueing formulas surveyed in Section 2.4. Offered load a, in erlangs, is
the product of the mean arrival rate λ and the mean service duration E[S]:
""")
    add_equation(doc, "a = λ · E[S]",
        note="a dimensionless quantity: the mean number of servers that would be simultaneously busy under the given arrival rate and service-time distribution, independent of the number of servers actually provisioned.")
    add_body(doc, """
This project evaluates a twice per department - once at λ = the real
calibrated average arrival rate, once at λ = the real calibrated
highest-demand hourly bin's rate - giving a range [a_avg, a_peak] the true
required capacity should plausibly fall within, with the chosen default
capacity n_capacity set strictly between the two, closer to a_peak, so that
the simulation is neither under-provisioned relative to typical demand nor
implausibly over-provisioned relative to any plausible real staffing level:
""")
    add_equation(doc, "a_avg < n_capacity < a_peak")
    add_body(doc, """
For the primary department studied in this project,
this calculation yields approximately 22 erlangs of average load and
approximately 34 erlangs at peak (the 11:00-14:00 arrival bin); a default
capacity of 30 was chosen as a value between these two figures, closer to
the peak, deliberately neither under-provisioned relative to typical demand
nor implausibly over-provisioned relative to any plausible real staffing
level. For the second, independent department used in the generalization
check (Chapter 6), the same method applied to that department's own real
arrival rate and acuity mix yields approximately 10.4 erlangs average and
16.1 erlangs peak load, giving a default capacity of 14 at the same relative
position between average and peak as the first department's capacity of 30
- capacity was recalibrated to the second department's own offered load, not
copied or linearly rescaled from the first department's absolute numbers.
""")

    add_section_heading(doc, "Derivation: Offered Load as an Instance of Little's Law", level=3)
    add_body(doc, """
The offered-load formula a = λ·E[S] used above is not an independent
queueing-theoretic assumption; it is a direct consequence of Little's Law,
one of the few genuinely distribution-free results in queueing theory - it
requires no assumption about the arrival process, the service-time
distribution, or the queue discipline beyond long-run stability, which is
what makes it applicable here regardless of the log-normal service-time
model's specific shape (Section 4.2.3).
""")
    add_body(doc, """
Little's Law states that for any stable queueing system observed over a
long time horizon, the long-run average number of customers in the system,
L, the long-run average arrival rate, λ, and the long-run average time a
customer spends in the system, W, satisfy
""")
    add_equation(doc, "L = λ · W")
    add_body(doc, """
The standard proof is a direct accounting argument rather than anything
probabilistic, which is exactly why the result needs no distributional
assumptions. Let N(t) denote the number of customers in the system at time
t, and consider a long observation window [0, T]. The total customer-time
accumulated in the system over this window can be computed two different
ways, and Little's Law is simply the statement that these two computations
must agree. First, integrating the instantaneous occupancy directly,
""")
    add_equation(doc, "∫₀ᵀ N(t) dt = (average occupancy over [0,T]) × T ≈ L · T,  for T large")
    add_body(doc, """
Second, the same integral equals the sum, over every customer who arrived
during [0, T], of that customer's own individual time in the system - each
customer i contributes exactly Wᵢ (their own sojourn time) to the area
under N(t), since they are counted as "in the system" for precisely that
duration:
""")
    add_equation(doc, "∫₀ᵀ N(t) dt = Σᵢ Wᵢ ≈ (λT) · W,  for T large",
        note="λT customers arrive over the window on average, each contributing its own sojourn time; averaging these sojourn times gives W by definition.")
    add_body(doc, """
Equating the two expressions for the same integral and dividing both sides
by T gives L = λW directly, with no assumption used beyond the system
being observed long enough, and stably enough, for both long-run averages
to be well defined.
""")
    add_body(doc, """
Offered load is obtained by applying this identical argument to a
deliberately restricted subsystem: instead of the whole queueing system
(waiting line plus servers), apply Little's Law to the service facility
alone, treating "in service" as the state being counted rather than
"present in the system." The long-run average number of customers
simultaneously in service is then, by the same L = λW logic with W replaced
by the mean service duration E[S] rather than the full sojourn time,
exactly the offered load:
""")
    add_equation(doc, "a = λ · E[S]")
    add_body(doc, """
which is precisely the equation stated above - now derived, rather than
introduced as a definition. Because this argument used nothing but the
accounting identity above, a is a valid measure of the mean simultaneous
service demand regardless of how many servers c actually handle it, which
is exactly why it is a sound quantity to compare against a candidate
staffing capacity independent of any assumption about the arrival process
being Poisson or service times being log-normal (Section 4.2.3).
""")

    add_section_heading(doc, "Derivation: Why Wait Times Grow Sharply Near Saturation", level=3)
    add_body(doc, f"""
Section 6.5.4 explains this project's central empirical finding - that
conditional miscalibration concentrates specifically in the
understaffed/high-demand category - by appeal to the qualitative,
well-documented fact that queueing systems' wait-time variance grows
sharply as utilization approaches capacity. That claim is derived here
explicitly, via the birth-death process underlying the M/M/c queue
{cite('green_book')}, rather than left as an assertion. The
derivation is included for the M/M/c idealization specifically because it
is the simplest queueing model for which the relevant closed-form result
(the Erlang C formula below) exists in exact form; this project's own DES
uses non-exponential, ESI-level-dependent service times (Section 4.2.3),
so the M/M/c result below is used qualitatively, as the standard
theoretical explanation for the mechanism observed, not as a claim that
this project's simulation is itself an M/M/c queue.
""")
    add_body(doc, """
Model the number of customers in an M/M/c system (Poisson arrivals at rate
λ, c identical servers each working at rate μ, so offered load a = λ/μ) as
a continuous-time birth-death Markov chain on states n = 0, 1, 2, …, with
arrival ("birth") rate λ_n = λ at every state, and service ("death") rate
""")
    add_equation(doc, "μₙ = { n·μ  for n ≤ c;   c·μ  for n ≥ c }",
        note="for n ≤ c, only n of the c servers are occupied, so the pooled service rate is nμ; once n ≥ c, all c servers are busy and the pooled service rate saturates at cμ regardless of how many additional customers are queued.")
    add_body(doc, """
For a birth-death chain, detailed balance (the stationary flow from n − 1
to n must equal the flow from n back to n − 1) gives πₙμₙ = πₙ₋₁λ, so each
stationary probability is fixed recursively in terms of the previous one,
πₙ = πₙ₋₁ · λ/μₙ. Unrolling this recursion from π₀ gives, in the two
regimes,
""")
    add_equation(doc, "πₙ = π₀ · aⁿ/n!,  for n ≤ c")
    add_equation(doc, "πₙ = π₀ · (aᶜ/c!) · ρⁿ⁻ᶜ,  for n ≥ c,  with ρ = a/c",
        note="ρ is the per-server utilization; the system is stable, and a stationary distribution exists at all, only if ρ < 1 - the queue-length recursion for n ≥ c is a plain geometric series with ratio ρ, which diverges to an unbounded expected queue length as ρ → 1.")
    add_body(doc, """
Normalizing so that Σₙ πₙ = 1 (using Σₖ₌₀^∞ ρᵏ = 1/(1 − ρ) for the
geometric tail, valid precisely when ρ < 1) fixes π₀, and summing πₙ over
every state with n ≥ c gives the probability that an arriving customer
finds all servers busy and must wait - the Erlang C formula:
""")
    add_equation(doc, "P(wait > 0) = C(c, a) = [ (aᶜ/c!) · 1/(1 − ρ) ] / [ Σₖ₌₀^{c−1} aᵏ/k!  +  (aᶜ/c!) · 1/(1 − ρ) ]")
    add_body(doc, """
and, via the memoryless property of exponential service applied to
whichever customers are ahead in queue, the mean additional wait time
experienced by a customer who does have to wait is 1/(cμ − λ), so the
unconditional mean wait time is
""")
    add_equation(doc, "W_q = C(c, a) / (cμ − λ)")
    add_body(doc, """
The specific mechanism behind Section 6.5.4's claim is now visible directly
in this last equation rather than only asserted: as offered load a
approaches the server count c from below, ρ = a/c → 1, the denominator
(cμ − λ) → 0⁺, and W_q → ∞ - not gradually, but as a term whose
denominator is heading to zero, so the rate of growth accelerates sharply
in the vicinity of ρ = 1 rather than increasing at a roughly constant rate
as utilization rises from, say, 50 to 70 percent. This is the precise,
quantitative form of the "variance grows sharply near saturation" claim
Section 6.5.4 relies on qualitatively: the staff = Low, arrival = High
category (Section 6.5.1) is, by construction, the cell of this project's
sampled scenario grid operating at the highest ρ = a/n_capacity of any
category, placing it closest to the region where this derivation shows
W_q's sensitivity to small changes in a or c is most extreme - exactly
where a pooled quantile computed mostly from lower-ρ, lower-variance
calibration points is least equipped to characterize the residual spread
correctly.
""")

    add_section_heading(doc, "4.2.2 Scenario Sampling for Surrogate Training and Calibration")
    add_body(doc, """
To generate data for surrogate training, the calibrated DES is run across
thousands of randomly sampled scenarios, with staffing capacity sampled
from a range around the Erlang-derived default and the arrival-rate
multiplier sampled from a range covering both below- and above-average
demand, each scenario run with its own independent random seed so that the
DES's own stochasticity (arrival randomness, acuity assignment, service-time
variance) is preserved in the resulting labels - this residual noise is
exactly what the downstream uncertainty quantification methods (Section 4.4)
are being asked to characterize, and pre-averaging it away by running each
scenario multiple times and reporting only a mean would defeat the purpose
of the entire UQ comparison. A separate, disjoint pool of scenarios, drawn
with a different sampling seed and a large seed offset from the training
data, is generated specifically for conformal prediction calibration
(Section 4.4.2); this separation is methodologically necessary rather than
a formality, because calibrating on residuals from data the surrogate was
trained on would systematically understate the true residual spread (the
surrogate fits its own training points more closely than it fits genuinely
unseen points) and would invalidate the resulting coverage guarantee.
""")

    add_section_heading(doc, "4.2.3 Data Calibration: Real Arrivals, Literature Service Times")
    add_body(doc, """
The DES's arrival process and ESI acuity mix are calibrated directly from a
large, de-identified, publicly available dataset of real emergency
department visits (described fully in Chapter 5), specifically the hourly,
daily, and monthly arrival rate distributions and the ESI mix extracted for
a single department within that dataset. Service-time distributions are not
derived from this dataset, because the dataset - despite containing several
hundred columns of vitals, diagnosis flags, and laboratory results - contains
no length-of-stay or treatment-duration field of any kind, a limitation
confirmed by an exhaustive search across every column in the dataset rather
than assumed. In place of data-derived service times, this project uses
literature-standard log-normal service-time parameters by ESI acuity level.

This distinction - real-data-calibrated arrivals versus literature-calibrated
service times - is maintained explicitly and consistently throughout this
report and every presentation material produced over
the course of this project, rather than allowing both to be presented as
equally data-derived. It is a genuine limitation of what this project's DES
can claim to represent, and it is disclosed as such rather than obscured.
""")

    add_body(doc, """
For completeness and to give the cross-check that follows a concrete
target to check against, Table 4.1 states the literature-standard
log-normal service-time parameters actually used by the DES (Section 4.2,
implemented in src/utils/extract_distributions.py, Appendix A), by ESI
acuity level.
""")
    add_table(doc,
        ["ESI level", "Mean (minutes)", "SD (minutes)"],
        [
            ["1 (most acute)", "180", "90"],
            ["2", "150", "75"],
            ["3", "120", "60"],
            ["4", "75", "40"],
            ["5 (least acute)", "45", "25"],
        ],
        caption="Literature-standard log-normal service-time parameters by ESI acuity level, as used by this project's DES.")

    add_section_heading(doc, "4.2.3.1 Cross-Checking Service-Time Parameters Against Independently Reported ED Data", level=3)
    add_body(doc, f"""
Because the dataset used throughout this project (Section 5.1) cannot
supply service-time information, the literature-standard log-normal
parameters used in its place (Table 4.1 above) are checked here against
ten independently published emergency department studies, each drawn from
a real hospital dataset of its own and none overlapping with the Kaggle
dataset this project otherwise uses, rather than left as a single
unattributed "literature-typical" assumption. Table 4.2 summarizes what
each study actually reports and how it bears on this project's own
calibration.
""")
    add_table(doc,
        ["Study", "Real dataset", "What it reports"],
        [
            [f"Hoot et al. {cite('hoot2008')}", "1 US academic ED, DES model", "Log-normal per-ESI evaluation/treatment duration; median 4.0/4.6/3.1/1.7/1.2 h for ESI 1-5; ESI mix 0.7/37.8/44.2/15.8/1.4%; nonstationary Poisson arrivals, 1.6-10.3/h"],
            [f"Otto et al. {cite('otto2022')}", "AKTIN registry, 12 German EDs, n=304,606 (2019)", "Mean +/- SD length of stay by ESI/MTS triage level, admitted and non-admitted separately"],
            [f"Theiling et al. {cite('theiling2020')}", "US NHAMCS, ~805.7M weighted visits (2010-2015)", "Median length of stay by 5-level ESI-NHAMCS severity tier"],
            [f"Karaca et al. {cite('karaca2012')}", "AZ/MA/UT state databases, n=4,955,590 (2008)", "Overall mean/median ED duration (195.7 / 130.2 min); confirms right-skew"],
            [f"Kim et al. {cite('kim2021')}", "1 US ED, n=2,107", "Mean LOS for ESI 4 and ESI 5 specifically, pre/post a triage-discharge intervention"],
            [f"Mahmoodian et al. {cite('mahmoodian2014')}", "2 hospitals, n=900", "Median time-to-first-physician-visit by ESI level 1-5"],
            [f"Laskowski et al. {cite('laskowski2009')}", "6 Winnipeg hospitals, n=185,659", "Priority-queue ED model calibrated on CTAS triage classes"],
            [f"Locker and Mason {cite('locker2005')}", "UK NHS EDs", "Documents the right-skewed shape of ED time-in-department distributions"],
            [f"De Santis et al. {cite('desantis2021')}", "1 large Italian ED", "Nonhomogeneous-Poisson-process methodology for hourly arrival-rate calibration"],
            [f"Kramer et al. {cite('kramer2020')}", "1 Italian ED, ~7,000 visits/month", "DES implementation case study for a real, operating ED"],
        ],
        caption="Ten independently published ED studies used to cross-check this project's literature-calibrated service-time parameters and arrival-process methodology.")
    add_body(doc, """
Two of these ten bear directly enough on this project's own numeric choices
to warrant a specific quantitative comparison rather than only a
qualitative citation, and both comparisons are reported honestly - neither
is a perfect match, and where they diverge, that divergence is stated
rather than smoothed over.
""")
    add_body(doc, f"""
Hoot et al. {cite('hoot2008')} is the closest structural match available: like
this project, it fits a separate log-normal distribution to evaluation and
treatment duration within each ESI level. Table 4.3 places their reported
per-ESI medians alongside this project's own literature-calibrated means
(Table 4.1).
""")
    add_table(doc,
        ["ESI level", "This project's mean (min)", "Hoot et al. median (min)", "Direction of difference"],
        [
            ["1 (most acute)", "180", "240", "Hoot's is higher"],
            ["2", "150", "276", "Hoot's is higher"],
            ["3", "120", "186", "Hoot's is higher"],
            ["4", "75", "102", "Hoot's is higher"],
            ["5 (least acute)", "45", "72", "Hoot's is higher"],
        ],
        caption="This project's literature-calibrated service-time means vs. Hoot et al.'s independently reported per-ESI medians.")
    add_body(doc, """
Every level in Hoot et al.'s data runs higher than this project's own
parameter - consistent with the general direction expected, since a
log-normal distribution's median sits below its mean, so a same-valued
comparison would already understate Hoot et al.'s corresponding mean
somewhat; the gap is nonetheless larger than that alone would explain,
and this project's parameters should be read as being toward the shorter,
more conservative end of what published ED studies report rather than a
precise match to any one of them. A second, genuinely informative
discrepancy is the ordering itself: Hoot et al.'s ESI-2 duration (276 min)
exceeds their ESI-1 duration (240 min), a real, documented non-monotonicity
- the most acute patients are evaluated fastest in some systems precisely
because they are triaged and moved with the highest urgency, while ESI-2
patients may undergo more extensive diagnostic workups once stabilized.
This project's own parameters are strictly monotonic by construction (180
> 150 > 120 > 75 > 45), a simplification that Hoot et al.'s independently
reported data shows is not always how real EDs behave, disclosed here as a
specific, named limitation rather than left for a reader to discover
unaided.
""")
    add_body(doc, f"""
Otto et al. {cite('otto2022')}'s AKTIN registry figures, while reporting total
length of stay rather than service time alone and therefore not directly
comparable in absolute minutes, are informative on a different, unitless
dimension: the ratio of standard deviation to mean. Across Otto et al.'s
ten reported ESI/MTS category-outcome combinations, this ratio ranges
from roughly 0.88 to 0.96 - standard deviation nearly as large as the mean
itself. This project's own log-normal parameters (Table 4.1) use a
substantially tighter ratio, close to 0.5 at every ESI level (for example,
ESI-1's 90-minute SD against a 180-minute mean). Some of this difference
is structural and expected - total length of stay accumulates wait-time
variability on top of service-time variability, so it should be more
variable than service time alone - but not necessarily all of it, and this
project's SD choices should accordingly be read as comparatively
conservative (narrower) relative to the variability real, independently
collected ED duration data exhibits, a second specific, quantified
limitation of the literature-calibrated service-time model disclosed here
directly rather than left implicit in Section 7.3's more general
limitations discussion.
""")
    add_body(doc, f"""
The remaining papers in Table 4.2 support this project's methodology in
ways that do not reduce to a single numeric comparison. Theiling et al.
{cite('theiling2020')} and Karaca et al. {cite('karaca2012')}, drawing on two of the
largest publicly documented ED datasets available (805.7 million weighted
NHAMCS visits and 4.96 million treat-and-release state-database visits
respectively), both confirm the general shape this project assumes -
duration decreasing with decreasing acuity, and a right-skewed
distribution consistent with a log-normal or similar heavy-tailed family -
without this project's own dataset being able to confirm either fact
directly. Kim et al. {cite('kim2021')} and Mahmoodian et al. {cite('mahmoodian2014')}
each provide an independent, smaller-scale data point in the same
direction at the lower-acuity (Kim et al.) and full-range (Mahmoodian et
al.) ends of the ESI scale. Laskowski et al. {cite('laskowski2009')} and
Kramer et al. {cite('kramer2020')} demonstrate that priority-queue and
discrete-event simulation approaches structurally similar to this
project's own (Section 4.2) have been successfully calibrated against real
ED operations at other sites, lending independent, cross-site support to
the DES-based methodology itself, not merely to its specific numeric
inputs. De Santis et al. {cite('desantis2021')} validates this project's choice
of a nonhomogeneous Poisson process for hourly arrival-rate calibration
(Section 4.2) as an established, actively refined methodology in its own
right rather than an ad hoc modeling convenience. Locker and Mason
{cite('locker2005')} is cited as further, independent confirmation that the
right-skewed shape this project's log-normal assumption relies on is a
property documented across health systems (here, the UK NHS) rather than
one specific to the US-based studies this section otherwise draws on.
""")

    add_section_heading(doc, "4.2.4 Validation Against Real Aggregate Statistics")
    add_body(doc, """
The calibrated DES is validated by comparing its simulated mean daily patient
volume, averaged over a large number of simulated days at the default
scenario configuration, against the dataset's real calibrated daily volume
for the same department. This validation is deliberately restricted to
volume (a statistic the real dataset can actually provide, given the
arrival-process calibration in Section 4.2.3) rather than extended to wait
times or service durations, which the real dataset cannot validate against
at all, for the reasons given in Section 4.2.3. Results of this validation,
for both departments studied in this project, are reported in full in
Chapter 6.
""")

    add_section_heading(doc, "4.3 Surrogate Model Architectures")
    add_body(doc, """
A surrogate model is trained independently for each of the four DES output
metrics (Section 4.2), taking the two scenario parameters (staffing
capacity, arrival-rate multiplier) as input. Two architectures are used in
this project. In-distribution point-prediction accuracy for both is reported
throughout Chapter 6 using three standard regression metrics, computed over
a held-out test set of size m:
""")
    add_equation(doc, "MAE = (1/m) Σᵢ₌₁ᵐ |yᵢ − ŷᵢ|")
    add_equation(doc, "RMSE = √[ (1/m) Σᵢ₌₁ᵐ (yᵢ − ŷᵢ)² ]")
    add_equation(doc, "R² = 1 − [ Σᵢ₌₁ᵐ (yᵢ − ŷᵢ)² / Σᵢ₌₁ᵐ (yᵢ − ȳ)² ]",
        note="ȳ the test-set mean of y; R² = 1 indicates a perfect fit, R² = 0 indicates the model performs no better than always predicting ȳ.")

    add_section_heading(doc, "4.3.1 Gradient-Boosting Regressor (Primary Architecture)")
    add_body(doc, """
The primary surrogate architecture is a histogram-based gradient boosting
regressor, an ensemble of shallow regression trees fit sequentially via
gradient descent in function space, each new tree fit to the negative
gradient (for squared-error loss, equivalently the residual) of the current
ensemble's predictions. This architecture was chosen over alternatives
(linear models, a single deep neural network) as the primary surrogate
because it is a strong, standard default for small-to-moderate-sized tabular
regression problems with few input features, requires comparatively little
hyperparameter tuning to perform well, and is fast enough to fit and query
that it does not become a computational bottleneck relative to the DES data
generation step it approximates. Section 4.3.3 discusses a specific
structural property of tree-based models - their inability to extrapolate
predictions outside the range of their training data - that becomes directly
relevant to the exchangeability stress test summarized in Section 6.13.
""")

    add_section_heading(doc, "Derivation: Gradient Boosting as Functional Gradient Descent", level=3)
    add_body(doc, f"""
The claim that "each new tree is fit to the negative gradient, equivalently
the residual" is stated above as a fact about how this project's surrogate
is trained; this section derives it, following {cite('friedman2001')}, since the
equivalence between "fit a tree to the residual" (an operational recipe)
and "gradient descent in function space" (the theoretical justification for
why that recipe is a sensible thing to do) is not obvious from the recipe
alone.
""")
    add_body(doc, """
The learning problem is to find a function F: 𝒳 → ℝ minimizing the
expected squared-error loss over the training distribution:
""")
    add_equation(doc, "F* = argminF  𝔼[ (y − F(x))² ]")
    add_body(doc, """
Gradient boosting builds F* as an additive expansion, F_M(x) = Σₘ₌₁ᴹ hₘ(x),
constructed greedily: at stage m, having already built the ensemble
F_{m-1} from the previous m − 1 trees, the next weak learner hₘ is chosen
to approximate the direction of steepest descent of the loss, treating the
ensemble's prediction at each training point as a free parameter to be
adjusted. Concretely, the functional gradient of the squared-error loss
L(y, F) = ½(y − F)² with respect to the prediction F, evaluated at each
training point xᵢ under the current ensemble F_{m-1}, is
""")
    add_equation(doc, "−∂L(yᵢ, F(xᵢ)) / ∂F(xᵢ) |_{F = F_{m-1}} = −∂[ ½(yᵢ − F(xᵢ))² ] / ∂F(xᵢ) |_{F = F_{m-1}} = yᵢ − F_{m-1}(xᵢ)")
    add_body(doc, """
that is, for squared-error loss specifically, the negative functional
gradient at each training point is exactly that point's current residual,
rᵢ = yᵢ − F_{m-1}(xᵢ). The next tree hₘ is then fit, by ordinary
least-squares regression-tree training, to predict these residuals as its
target:
""")
    add_equation(doc, "hₘ = argmin_h  Σᵢ₌₁ⁿ ( rᵢ − h(xᵢ) )²,  rᵢ = yᵢ − F_{m-1}(xᵢ)")
    add_body(doc, """
and the ensemble is updated by taking a small step in this direction,
F_m(x) = F_{m-1}(x) + ν·hₘ(x), with a learning rate ν ∈ (0, 1] controlling
step size (shrinkage) exactly as in ordinary gradient descent on a
parameter vector, except that the "parameter" being descended on here is
the ensemble's prediction at every training point simultaneously, and each
descent step is itself constrained to be representable by a shallow
regression tree rather than an arbitrary vector update. "Fit a tree to the
residual" is therefore not a heuristic alternative to gradient descent; for
squared-error loss it is gradient descent in function space, with the
residual arising as the specific, closed form the negative functional
gradient happens to take for this one loss function. (For other losses -
the quantile ("pinball") loss used by the CQR quantile regressors in
Section 4.4.4, for instance - the same functional-gradient-descent
recipe applies with a different negative-gradient target: piecewise
constant rather than the raw residual, which is why quantile regression
trees are trained by the same underlying algorithm but do not simply fit
raw residuals.)
""")

    add_section_heading(doc, "4.3.2 Multi-Layer Perceptron (Robustness-Check Architecture)")
    add_body(doc, """
A second surrogate architecture, a multi-layer perceptron with two hidden
layers trained with early stopping and preceded by standard feature scaling,
is trained on the identical train/test split as the gradient-boosting
surrogate, specifically to test whether findings obtained with the primary
architecture are specific to tree ensembles or generalize across a
structurally different model family. Unlike a tree ensemble, a neural
network has no inherent floor or ceiling on its output as inputs move
outside its training range - it will continue to extrapolate, in whichever
direction its learned weights imply, rather than saturating at a boundary
value. This structural difference is the reason this architecture is
included as a specific, targeted robustness check on the exchangeability
findings summarized in Section 6.13, rather than as a general
"try another model and see" exercise.
""")

    add_section_heading(doc, "4.3.3 Tree-Based Extrapolation Limits")
    add_body(doc, """
A specific property of tree-based ensembles, directly relevant to
interpreting results in Chapter 6 and central to Section 6.13,
is worth stating precisely here. A regression tree partitions its
input space into axis-aligned regions (leaves) during training and predicts
a constant value - the mean of the training targets falling into that leaf -
for any query point landing in that region. For an input that falls outside
the entire range spanned by the training data, the tree does not
extrapolate a trend; it simply returns whatever constant value the nearest
boundary leaf holds, because that boundary leaf is where every such
out-of-range query is routed by the tree's learned split thresholds. A
gradient-boosting ensemble, being a sum of many such trees, inherits this
property: its prediction for an input arbitrarily far outside the training
range is frozen at the value implied by the boundary leaves of its
constituent trees, however far outside that range the query point actually
lies. This is a well-known, textbook property of tree ensembles, not a bug
specific to this project's implementation, and it is verified directly
(rather than merely asserted) in Section 6.13 by
observing gradient-boosting surrogate predictions that are numerically
frozen across a wide range of out-of-distribution query points.
""")

    add_section_heading(doc, "4.4 Uncertainty Quantification Methods")
    add_body(doc, """
Four uncertainty quantification approaches are implemented and compared in
this project, all evaluated at a common nominal miscoverage rate alpha =
0.1 (that is, targeting 90 percent coverage) on identical calibration and
test data wherever a direct comparison is drawn, so that any difference
found reflects the UQ method itself rather than a difference in evaluation
conditions.
""")

    add_section_heading(doc, "4.4.1 Gaussian Process Baseline")
    add_body(doc, """
The Gaussian process (GP) baseline treats each of the four surrogate targets
as an independent Gaussian process regression problem over the
two-dimensional scenario-parameter input space. A GP prior is placed
directly over the unknown function f:
""")
    add_equation(doc, "f(x) ~ 𝒢𝒫( m(x), k(x, x′) )",
        note="m(x) the prior mean function (taken as zero here) and k(x, x′) the covariance (kernel) function encoding assumed smoothness across the input space.")
    add_body(doc, """
This project's implementation (src/uq/gp_baseline.py) uses the squared-
exponential (radial basis function, RBF) kernel, the standard default
choice for a smoothly varying, continuous scenario-parameter input space
with no known periodicity or discontinuity to encode structurally:
""")
    add_equation(doc, "k(x, x′) = σf² exp( − ‖x − x′‖² / (2ℓ²) )",
        note="σf² the signal variance (the kernel's value at x = x′, an overall scale on how much f is expected to vary) and ℓ the length scale (how quickly correlation between f(x) and f(x′) decays with distance ‖x − x′‖); both are hyperparameters fit by maximizing the marginal likelihood during training, together with σₙ² below, rather than set by hand.")
    add_body(doc, """
Conditioning on the training inputs X and targets y yields a closed-form
Gaussian posterior predictive distribution at any query point x*:
""")
    add_equation(doc, "f(x*) | X, y ~ 𝒩( μ(x*), σ²(x*) )")
    add_equation(doc, "μ(x*) = k*ᵀ (K + σₙ²I)⁻¹ y")
    add_equation(doc, "σ²(x*) = k(x*, x*) − k*ᵀ (K + σₙ²I)⁻¹ k*",
        note="K the training-set kernel matrix, k* the vector of kernel values between x* and each training point, and σₙ² the observation-noise variance.")

    add_section_heading(doc, "Derivation of the Posterior Predictive Distribution", level=3)
    add_body(doc, f"""
The three equations above are stated as the standard result throughout the
Gaussian process literature (Section 2.3, {cite('rasmussen2006')}); the derivation is
given here because it is what makes precise exactly which matrix inversion
the O(n³) cost discussed below is paying for, and because it is a direct,
mechanical consequence of one identity - conditioning a multivariate
Gaussian on part of itself - rather than something specific to GP
regression that must be taken on faith.
""")
    add_body(doc, """
By the GP prior's definition, the observed training targets y (assumed to
carry independent Gaussian observation noise of variance σₙ², so that
y = f(X) + ε with ε ~ 𝒩(0, σₙ²I)) and the unknown function value f(x*) at
a new query point are jointly Gaussian, since any finite collection of
values from a Gaussian process is itself jointly Gaussian by definition:
""")
    add_equation(doc, "[ y ; f(x*) ]  ~  𝒩( 0,  [[ K + σₙ²I,  k* ],  [ k*ᵀ,  k(x*, x*) ]] )",
        note="a (n+1)-dimensional joint Gaussian, block-partitioned into the n training targets and the single query-point value; K is the n×n matrix with Kᵢⱼ = k(xᵢ, xⱼ), and k* is the n-vector with (k*)ᵢ = k(xᵢ, x*).")
    add_body(doc, """
The result then follows from the standard identity for conditioning a
jointly Gaussian vector on a subset of its own components (equivalently,
computing the Schur complement of the joint covariance's training-training
block). For a general partitioned Gaussian
[a; b] ~ 𝒩([μₐ; μ_b], [[Σₐₐ, Σₐᵦ],[Σᵦₐ, Σᵦᵦ]]), the conditional distribution
of b given a is itself Gaussian, with
""")
    add_equation(doc, "b | a  ~  𝒩( μ_b + Σᵦₐ Σₐₐ⁻¹ (a − μₐ),   Σᵦᵦ − Σᵦₐ Σₐₐ⁻¹ Σₐᵦ )")
    add_body(doc, """
Substituting a = y, b = f(x*), μₐ = μ_b = 0 (the zero prior mean assumed
above), Σₐₐ = K + σₙ²I, Σₐᵦ = Σᵦₐᵀ = k*, and Σᵦᵦ = k(x*, x*) directly into
this general identity reproduces μ(x*) and σ²(x*) exactly as stated:
substituting into the mean expression gives
μ_b + Σᵦₐ Σₐₐ⁻¹(a − μₐ) = 0 + k*ᵀ (K + σₙ²I)⁻¹ (y − 0) = k*ᵀ (K + σₙ²I)⁻¹ y,
and substituting into the variance expression gives
Σᵦᵦ − Σᵦₐ Σₐₐ⁻¹ Σₐᵦ = k(x*, x*) − k*ᵀ (K + σₙ²I)⁻¹ k* - term for term the two
equations already stated above. Nothing beyond this one conditioning
identity, applied to the specific block structure a Gaussian process
prior induces, is needed to obtain the closed form; the entire derivation
is linear algebra once the joint-Gaussian claim itself is granted, which
is precisely the defining property of a Gaussian process.
""")
    add_body(doc, """
The (K + σₙ²I)⁻¹ term appearing in both μ(x*) and σ²(x*) is exactly the
n × n matrix inverse whose Cholesky factorization costs O(n³) - Section
4.4.1's motivation for subsampling to n = 1,000 training points is, in
light of this derivation, a direct consequence of this one matrix
appearing in both closed-form results, not an incidental implementation
detail of this project's specific code.
""")
    add_body(doc, """
The (1 − α) prediction interval is then constructed directly from this
posterior:
""")
    add_equation(doc, "C(x*) = [ μ(x*) − z·σ(x*),  μ(x*) + z·σ(x*) ]",
        note="z the standard normal quantile at the target confidence level (z ≈ 1.645 for the 90% interval used throughout this project).")
    add_body(doc, """
Exact GP inference requires inverting (in practice, Cholesky-factorizing)
the n × n matrix (K + σₙ²I), an O(n³) computational cost in the number of
training points n, which motivates training the GP on a representative
subsample (1,000 points) of the available calibration data rather than the
full set - a choice that is also directly relevant to this project's own
narrative, since the GP's comparatively poor computational scalability
relative to conformal prediction's near-constant calibration cost is one of
the quantities this project's results (Chapter 6) explicitly measure and
report, not simply an implementation convenience footnoted away.
""")

    add_section_heading(doc, "4.4.2 Standard (Split) Conformal Prediction")
    add_body(doc, f"""
Standard conformal prediction, in the split (inductive) form introduced by
Papadopoulos et al. {cite('papadopoulos2002')} and reviewed in Section 2.1, computes a
nonconformity score for each point in a held-out calibration set - disjoint
from both the surrogate's training data and the final test set, per Section
4.2.2 - and uses the ceiling of (n + 1)(1 - alpha) divided by n-th empirical
quantile of those scores (the standard finite-sample-corrected quantile
formula that gives conformal prediction its exact coverage guarantee) as a
fixed margin applied to every future prediction. Two nonconformity measures
are implemented: a symmetric measure, the absolute residual between the
surrogate's prediction and the true value, giving a fixed-width interval
around each prediction; and an asymmetric measure, in which upper and lower
residual quantiles are calibrated separately (each at 1 - alpha/2) and
combined via a union bound, a valid but slightly conservative route to at
least (1 - alpha) coverage that does not require assuming a symmetric
residual distribution around the point prediction.
""")

    add_section_heading(doc, "4.4.3 Mondrian Conformal Prediction")
    add_body(doc, """
Mondrian conformal prediction (Section 2.2) replaces standard conformal
prediction's single pooled quantile with a separate quantile calibrated
within each category of a partition (a "Mondrian taxonomy") of the
calibration set. This project's taxonomy is the cross of staffing-capacity
tercile (low, medium, high) and arrival-rate-multiplier tercile (low,
medium, high) - nine categories in total - using only these two real,
available scenario covariates rather than an invented additional dimension,
since the DES produces no other scenario-level covariate (such as a "shift"
or "day of week" label) that could be used instead. Category bin edges are
computed from the calibration set's own quantiles and never from the test
set, preserving the same calibration/test separation principle as standard
CP. For a fair, isolated comparison of pooling versus partitioning, Mondrian
CP in this project's core comparison uses the same symmetric nonconformity
measure as standard CP's symmetric variant, differing only in whether one
pooled quantile or nine per-category quantiles are computed and applied.

Evaluation of Mondrian CP's effect proceeds in two complementary ways on the
same test points: applying standard CP's single pooled quantile broken down
by category (revealing where the pooled, marginal guarantee actually breaks
down conditionally, even though it is not designed to be evaluated this
way), and applying each category's own Mondrian quantile (showing whether
per-category calibration corrects any breakdown found). Both a per-category
view (Chapter 6, following the same structure as the core result) and a
marginal, whole-test-set view (obtained by re-aggregating the per-category
predictions across the full test set, for a like-for-like comparison against
the GP baseline and standard CP's own marginal coverage) are reported.
""")

    add_section_heading(doc, "4.4.4 Conformalized Quantile Regression and Mondrian-CQR")
    add_body(doc, """
Conformalized quantile regression (CQR, Section 2.1) is implemented as a
stronger baseline against which Mondrian CP's benefit can be compared. Two
quantile regressors - gradient-boosting models trained with a quantile loss
function, targeting the alpha/2 and 1 - alpha/2 conditional quantiles
respectively - are trained on the identical train/test split as the primary
mean-regression surrogate. The CQR nonconformity score for a calibration
point is the greater of (the lower quantile prediction minus the true value)
and (the true value minus the upper quantile prediction); conformalizing
this score - taking its (1 - alpha) empirical quantile on the calibration
set and using it to inflate or deflate the raw quantile regressors' interval
- restores an exact finite-sample coverage guarantee even though the raw,
uncalibrated quantile regressors alone do not reliably achieve nominal
coverage. Mondrian-CQR combines both ideas: the same CQR nonconformity score,
calibrated separately within each of the nine Mondrian categories rather
than pooled, testing whether Mondrian's category-conditional calibration and
CQR's width-adaptivity address the same or different sources of
miscalibration.
""")

    add_section_heading(doc, "4.4.5 Formal Algorithm Statements")
    add_body(doc, """
For precision and reproducibility, this section states the three core
calibration procedures implemented in this project (standard split
conformal prediction, Mondrian conformal prediction, and conformalized
quantile regression) as explicit, numbered algorithms with their governing
equations, using notation consistent with Chapter 2's literature review.
Let f̂ denote the already-trained surrogate point predictor (Section 4.3),
Dcal the calibration set of (x, y) pairs disjoint from both the surrogate's
training data and the final test set (Section 4.2.2), n the size of Dcal,
and α the target miscoverage rate (α = 0.1 throughout this project, Section
4.4), so that the target coverage level is 1 − α = 0.9.
""")

    add_section_heading(doc, "Algorithm 1: Standard Split Conformal Prediction")
    add_body(doc, """
Step 1. For each calibration point (xᵢ, yᵢ) in Dcal, compute the symmetric
nonconformity score:
""")
    add_equation(doc, "sᵢ = |yᵢ − f̂(xᵢ)|,  i = 1, …, n")
    add_body(doc, """
(For the asymmetric measure, positive and negative residuals are instead
collected separately: sᵢ⁺ = max(0, yᵢ − f̂(xᵢ)) and sᵢ⁻ = max(0, f̂(xᵢ) − yᵢ).)

Step 2. Compute q̂, the finite-sample-corrected empirical quantile of the
calibration scores:
""")
    add_equation(doc, "q̂ = the ⌈(n + 1)(1 − α)⌉-th smallest value of {s₁, …, sₙ}",
        note="⌈·⌉ denotes the ceiling function; this rank-based correction (rather than the simpler ⌊n(1 − α)⌋-th value) is what gives the resulting interval its exact, non-asymptotic coverage guarantee.")
    add_body(doc, f"""
- the specific correction whose finite-sample validity is established by the
theory in {cite('vovk2005')} and {cite('papadopoulos2002')}, reviewed in Section 2.1. For the
asymmetric measure, this quantile is instead computed separately for {{sᵢ⁺}}
and {{sᵢ⁻}}, each at level 1 − α/2.

Step 3. For a new test point x, construct the prediction interval:
""")
    add_equation(doc, "C(x) = [ f̂(x) − q̂,  f̂(x) + q̂ ]")
    add_body(doc, """
(or [f̂(x) − q̂⁻, f̂(x) + q̂⁺] for the asymmetric measure, using each side's
own quantile).

This procedure, implemented in src/uq/standard_cp.py, is guaranteed - under
the assumption that (x, y) pairs in Dcal and the test point are
exchangeable - to satisfy the marginal coverage guarantee:
""")
    add_equation(doc, "P( Y ∈ C(X) ) ≥ 1 − α",
        note="regardless of f̂'s accuracy, because the rank-based correction in Step 2 accounts exactly for the probability that a genuinely exchangeable test score ranks among the top ⌈(n + 1)α⌉ scores out of n + 1 total (the n calibration scores plus the test score itself).")

    add_section_heading(doc, "Proof of the Coverage Guarantee", level=3)
    add_body(doc, """
The guarantee stated above is not an approximation or an asymptotic result
- it holds exactly, for every finite n, under nothing more than the
exchangeability assumption. The full argument is given here rather than
only cited, since it is short, and since understanding exactly what
exchangeability buys makes the exchangeability stress test in Section 6.13
interpretable as a violation of a specific, named assumption rather than an
unexplained empirical failure.
""")
    add_section_heading(doc, "Setup", level=3)
    add_body(doc, """
Let Z₁ = (x₁, y₁), …, Zₙ = (xₙ, yₙ) denote the n calibration points and
Zₙ₊₁ = (x, y) the test point. f̂ is trained on a data set disjoint from
both (Section 4.2.2) and is therefore held fixed throughout this argument -
it is not re-estimated as calibration or test points are permuted, so it
introduces no dependence that could break exchangeability among
Z₁, …, Zₙ₊₁. Define the nonconformity score at every one of these n + 1
points using the same fixed f̂:
""")
    add_equation(doc, "sᵢ = |yᵢ − f̂(xᵢ)|,  i = 1, …, n + 1")
    add_body(doc, """
Because Z₁, …, Zₙ₊₁ are assumed exchangeable (their joint distribution is
invariant under any permutation of the n + 1 indices) and each sᵢ is
computed from Zᵢ by the same fixed function, the scores s₁, …, sₙ₊₁ are
themselves exchangeable random variables.
""")
    add_section_heading(doc, "Step 1: The Rank of an Exchangeable Score Is Uniform", level=3)
    add_body(doc, """
This is the single fact the whole guarantee rests on. For n + 1
exchangeable real-valued random variables with no ties (ties are broken by
an independent random perturbation if the underlying scores can repeat, so
this is without loss of generality), every one of the (n + 1)! orderings of
s₁, …, sₙ₊₁ is equally likely - permuting the underlying (xᵢ, yᵢ) pairs
permutes the scores identically, and exchangeability states precisely that
every such permutation of the data has the same joint distribution. Under
a uniform distribution over orderings, the rank of any single specified
element among the n + 1 - here, the test score sₙ₊₁ - is itself uniform
on {1, …, n + 1}:
""")
    add_equation(doc, "P( Rank(sₙ₊₁) = k ) = 1 / (n + 1),  for every k = 1, …, n + 1",
        note="Rank(sₙ₊₁) counts how many of s₁, …, sₙ₊₁ (including sₙ₊₁ itself) are ≤ sₙ₊₁.")
    add_section_heading(doc, "Step 2: Relating the Rank to the Calibration Quantile", level=3)
    add_body(doc, """
q̂ (Step 2 of Algorithm 1) is defined purely from the n calibration
scores, as their ⌈(n + 1)(1 − α)⌉-th smallest value. The test score sₙ₊₁
exceeds q̂ exactly when sₙ₊₁ is large enough to rank above at least
⌈(n + 1)(1 − α)⌉ of the n + 1 scores - that is, exactly when
""")
    add_equation(doc, "sₙ₊₁ > q̂  ⟺  Rank(sₙ₊₁) > ⌈(n + 1)(1 − α)⌉")
    add_body(doc, """
Combining this equivalence with Step 1's uniform-rank result gives the
miscoverage probability directly, as a simple count of how many of the
n + 1 equally likely ranks exceed the threshold:
""")
    add_equation(doc, "P( sₙ₊₁ > q̂ ) = P( Rank(sₙ₊₁) > ⌈(n + 1)(1 − α)⌉ ) = 1 − ⌈(n + 1)(1 − α)⌉ / (n + 1)")
    add_body(doc, """
Since ⌈(n + 1)(1 − α)⌉ ≥ (n + 1)(1 − α) by definition of the ceiling
function, dividing by (n + 1) and substituting into the line above gives
the one-line inequality this whole algorithm exists to guarantee:
""")
    add_equation(doc, "P( sₙ₊₁ > q̂ ) ≤ 1 − (n + 1)(1 − α) / (n + 1) = α")
    add_body(doc, """
Step 3: Translating back from the score to the interval. Because
sₙ₊₁ = |y − f̂(x)|, the event {sₙ₊₁ ≤ q̂} is, by construction, exactly the
event {y ∈ [f̂(x) − q̂, f̂(x) + q̂]} = {Y ∈ C(X)}. Taking the complement of
the inequality just derived,
""")
    add_equation(doc, "P( Y ∈ C(X) ) = P( sₙ₊₁ ≤ q̂ ) = 1 − P( sₙ₊₁ > q̂ ) ≥ 1 − α",
        note="which is exactly Algorithm 1's stated guarantee, obtained here from first principles rather than merely cited. Nowhere in this argument is any assumption made about f̂'s accuracy, the distribution of Y given X, or the dimensionality of X - the guarantee is genuinely distribution-free, and its only load-bearing assumption is the exchangeability of Z₁, …, Zₙ₊₁ used in Step 1.")
    add_body(doc, """
A second, less frequently stated consequence of the same argument is worth
recording because it bounds the guarantee from above as well as below,
making explicit that split conformal prediction is not simply
"conservative by an unknown margin": under the additional assumption that
ties among the scores occur with probability zero (satisfied here since
f̂'s residuals are continuous-valued),
""")
    add_equation(doc, "P( Y ∈ C(X) ) ≤ 1 − α + 1 / (n + 1)",
        note="obtained by the same rank argument applied to ⌈(n + 1)(1 − α)⌉ − 1 in place of ⌈(n + 1)(1 − α)⌉. Together, the two bounds sandwich the true coverage within 1/(n + 1) of the 1 − α target - at this project's calibration size (n = 1,200, Section 4.2.2), a band of well under 0.1 percentage points, negligible next to the several-percentage-point effects this report's results (Chapter 6) actually report.")
    add_body(doc, """
Exchangeability is the assumption doing all the work in Step 1, and
nowhere else in the argument. Section 6.13's exchangeability stress test is
therefore, precisely, a test of whether Step 1 continues to hold once the
test distribution's arrival-rate multiplier is pushed outside the range the
calibration data was drawn from - and Section 6.13's coverage collapse is
exactly what this proof predicts happens the moment that one assumption
stops being true: no other step in the argument offers any fallback once
exchangeability itself is violated.
""")

    add_section_heading(doc, "Algorithm 2: Mondrian Conformal Prediction")
    add_body(doc, """
Step 1. Define a partition function c(x) mapping each scenario x to one of
K categories (K = 9 in this project: the cross of staffing tercile and
arrival-rate tercile, Section 4.4.3):
""")
    add_equation(doc, "c : 𝒳 → {1, 2, …, K}")
    add_body(doc, """
Tercile boundaries used by c(x) are computed from Dcal's own marginal
distribution of staffing capacity and arrival-rate multiplier - never from
the test set, preserving the same calibration/test separation as Algorithm
1.

Step 2. Partition Dcal into K disjoint subsets by category:
""")
    add_equation(doc, "Dcal⁽ᵏ⁾ = { (xᵢ, yᵢ) ∈ Dcal : c(xᵢ) = k },  k = 1, …, K")
    add_body(doc, """
Step 3. For each category k, apply Algorithm 1's Steps 1-2 using only
Dcal⁽ᵏ⁾, producing a category-specific quantile:
""")
    add_equation(doc, "q̂⁽ᵏ⁾ = the ⌈(nₖ + 1)(1 − α)⌉-th smallest value of {sᵢ : (xᵢ, yᵢ) ∈ Dcal⁽ᵏ⁾}",
        note="nₖ = |Dcal⁽ᵏ⁾|, the number of calibration points in category k.")
    add_body(doc, """
Step 4. For a new test point x, determine its category k = c(x) and
construct its prediction interval:
""")
    add_equation(doc, "C(x) = [ f̂(x) − q̂⁽ᶜ⁽ˣ⁾⁾,  f̂(x) + q̂⁽ᶜ⁽ˣ⁾⁾ ]")
    add_body(doc, """
This procedure, implemented in src/uq/mondrian_cp.py, is guaranteed to
satisfy group-conditional coverage:
""")
    add_equation(doc, "P( Y ∈ C(X) | c(X) = k ) ≥ 1 − α,  for every k = 1, …, K",
        note="a strictly stronger guarantee than Algorithm 1's marginal one, at the cost of each q̂⁽ᵏ⁾ being estimated from only nₖ calibration points rather than the full n - the direct source of the finite-sample-noise tradeoff discussed at length in Sections 6.5.6 and 2.2.")

    add_section_heading(doc, "Proof of the Group-Conditional Guarantee", level=3)
    add_body(doc, """
The category partition c(x) is computed once, from Dcal's own marginal
covariate distribution (Step 1), before any label information is used and
identically for calibration and test points - it is a fixed, deterministic
function of x alone, not re-estimated per category and not dependent on
which points happen to fall into which category. Conditioning the
exchangeable sequence Z₁, …, Zₙ₊₁ on the event {c(X₁) = ⋯ = c(Xₙ₊₁) = k}
(that is, restricting attention to the subsequence of points - calibration
and test alike - that happen to fall in category k) preserves
exchangeability among that subsequence, because exchangeability is a
property of the joint distribution of the Zᵢ's that a fixed, label-blind
selection rule cannot disturb: relabelling which of the category-k points
is "the test point" is still equally likely for each of the nₖ + 1 members
of the subsequence. Algorithm 1's proof (this section, above) therefore
applies verbatim to this subsequence, with n replaced by nₖ and D_cal
replaced by Dcal⁽ᵏ⁾ throughout, yielding
""")
    add_equation(doc, "P( Y ∈ C(X) | c(X) = k ) ≥ 1 − α")
    add_body(doc, """
for each k independently. This is a genuine corollary, not merely an
analogy: no new argument is required beyond noting that the conditioning
event is determined by x alone and applied identically before any
calibration or test label is observed, so Step 1's uniform-rank result
transfers unchanged to each category's own subsequence. The price of this
stronger, per-category statement is visible directly in the equation for
q̂⁽ᵏ⁾ above - nₖ is necessarily smaller than n (nₖ ≈ n / K on average across
K = 9 balanced categories), so the same finite-sample gap between the
lower and upper coverage bounds derived for Algorithm 1
(1/(n + 1) there becomes 1/(nₖ + 1) per category) widens roughly ninefold,
the precise, quantifiable source of the added estimation noise documented
empirically in Sections 6.5.6 and 6.6.1.
""")

    add_section_heading(doc, "Algorithm 3: Conformalized Quantile Regression (CQR)")
    add_body(doc, """
Step 1. Train two auxiliary quantile regression models on the same training
data as the primary surrogate f̂ (but using a pinball/quantile, rather than
squared-error, loss function):
""")
    add_equation(doc, "q̂lo targets the α/2 conditional quantile of Y | X;  q̂hi targets the (1 − α/2) conditional quantile of Y | X")
    add_body(doc, """
Step 2. For each calibration point (xᵢ, yᵢ) in Dcal, compute the CQR
nonconformity score:
""")
    add_equation(doc, "sᵢ = max( q̂lo(xᵢ) − yᵢ,  yᵢ − q̂hi(xᵢ) )",
        note="positive when the true value falls outside the raw [q̂lo(x), q̂hi(x)] band in either direction, negative when it falls inside - directly measuring how much the raw, uncalibrated quantile band under- or overshoots.")
    add_body(doc, """
Step 3. Compute the same finite-sample-corrected empirical quantile as in
Algorithm 1, Step 2, applied to the CQR scores:
""")
    add_equation(doc, "q̂ = the ⌈(n + 1)(1 − α)⌉-th smallest value of {s₁, …, sₙ}")
    add_body(doc, """
Step 4. For a new test point x, construct the prediction interval:
""")
    add_equation(doc, "C(x) = [ q̂lo(x) − q̂,  q̂hi(x) + q̂ ]")
    add_body(doc, """
This procedure, implemented across src/surrogate/train_quantile_surrogates.py
and src/uq/repeated_evaluation_cqr.py, restores an exact marginal coverage
guarantee, P(Y ∈ C(X)) ≥ 1 − α (by the same exchangeability argument as
Algorithm 1, applied to the CQR score rather than the raw residual) even
when the raw quantile regressors q̂lo and q̂hi do not themselves
achieve nominal coverage - Section 6.7 of this report
reports that this project's raw, uncalibrated quantile regressors achieve
only 81-92 percent coverage before this calibration step, exactly the gap
Step 3's conformalization closes. Mondrian-CQR combines Algorithm 3 with
Algorithm 2's partitioning: Steps 2-3 are repeated separately within each of
the K Mondrian categories, using each category's own calibration subset, to
produce a category-specific q̂⁽ᵏ⁾ applied to that category's test points in
Step 4.
""")

    add_section_heading(doc, "Proof That CQR Restores the Guarantee", level=3)
    add_body(doc, """
The claim that Step 3's conformalization restores exact coverage
regardless of how poorly q̂lo and q̂hi themselves perform is the single
most important property of CQR, and it follows from Algorithm 1's proof
with only one substitution, made precise here rather than left implicit.
q̂lo and q̂hi are trained on data disjoint from Dcal and the test point
(the same disjointness Section 4.2.2 establishes for f̂), so, exactly as in
Algorithm 1's setup, they are fixed functions with respect to the
exchangeable sequence Z₁, …, Zₙ₊₁. Define the CQR score at every one of
the n + 1 points using these same fixed, held-out quantile functions:
""")
    add_equation(doc, "sᵢ = max( q̂lo(xᵢ) − yᵢ,  yᵢ − q̂hi(xᵢ) ),  i = 1, …, n + 1")
    add_body(doc, """
Because q̂lo and q̂hi are held fixed, s₁, …, sₙ₊₁ inherit exchangeability
from Z₁, …, Zₙ₊₁ by the identical reasoning given for the residual score
in Algorithm 1's setup - nothing about that argument used the specific
functional form sᵢ = |yᵢ − f̂(xᵢ)|, only that each sᵢ is computed from Zᵢ
by one common, fixed function. Step 1's uniform-rank result, Step 2's
translation into a quantile bound, and the final inequality
P(sₙ₊₁ ≤ q̂) ≥ 1 − α therefore all transfer without modification. What
changes is only the final translation from the score back to an interval:
{sₙ₊₁ ≤ q̂} is, by this score's definition, equivalent to
{q̂lo(x) − q̂ ≤ y ≤ q̂hi(x) + q̂} = {Y ∈ C(X)}, giving
""")
    add_equation(doc, "P( Y ∈ C(X) ) = P( sₙ₊₁ ≤ q̂ ) ≥ 1 − α")
    add_body(doc, """
exactly as claimed, and for exactly the same reason as Algorithm 1's
guarantee: the argument never used any property of q̂lo or q̂hi beyond
their being fixed functions of x, so it is entirely indifferent to whether
the raw quantile band [q̂lo(x), q̂hi(x)] is itself well-calibrated. This is
what makes q̂ in Step 3 a genuine correction rather than a redundant
formality when the raw band under- or overshoots (Section 6.7's 81-92
percent raw-coverage finding): a poorly calibrated raw band simply shows up
as a q̂ systematically far from zero - large and positive when the band is
too narrow, as observed here, or negative when it is wastefully wide - and
the conformalization step absorbs exactly that miscalibration into q̂
without requiring q̂lo or q̂hi to be refit or even known to be imperfect.
""")

    add_section_heading(doc, "4.5 Robustness and Generality Checks")
    add_body(doc, """
Three checks establish whether this project's findings are stable and
general rather than artifacts of a single random split, a single surrogate
architecture, or a single hospital site.
""")

    add_section_heading(doc, "4.5.1 Repeated Evaluation with Statistical Significance Testing")
    add_body(doc, """
Rather than relying on a single calibration/test split, the full GP /
standard-CP / Mondrian-CP (and, in the extended version, CQR / Mondrian-CQR)
pipeline is repeated across R = 30 independent (calibration, test) draws, with
both calibration and test data freshly generated from the DES on each
repeat using disjoint seed ranges (never overlapping with each other, with
the original training data, or with any other project stage). Because the
same R draws underlie every method compared, per-repeat coverage across
methods forms a paired sample: for repeat r, let dᵣ = coverageᵣ(A) −
coverageᵣ(B) be the coverage difference between two methods A and B on
identical calibration/test data. A paired t-test - rather than a test
appropriate only for independent samples - is used to test whether the mean
difference d̄ is statistically distinguishable from zero, via the test
statistic:
""")
    add_equation(doc, "t = d̄ / (sd / √R)",
        note="d̄ the sample mean of {d₁, …, d_R} and sd their sample standard deviation, compared against a Student's t-distribution with R − 1 degrees of freedom to obtain a p-value.")
    add_body(doc, """
together with 95 percent confidence intervals on each method's own mean
coverage and width, computed as:
""")
    add_equation(doc, "x̄ ± t₀.₀₂₅,ᵣ₋₁ · (s / √R)",
        note="x̄ and s the sample mean and standard deviation of the quantity (coverage or width) across the R repeats, and t₀.₀₂₅,ᵣ₋₁ the two-sided 97.5th-percentile critical value of the t-distribution with R − 1 degrees of freedom.")

    add_section_heading(doc, "Derivation: Why the Test Statistic Follows a t-Distribution", level=3)
    add_body(doc, """
The paired differences d₁, …, d_R (one per independent calibration/test
draw, Section 4.5.1 above) are modeled as independent and identically
distributed draws from a normal distribution with unknown mean μ_d and
unknown variance σ_d² - a standard, and here reasonable, working
assumption given that each dᵣ is itself already an average-type quantity
(a difference of two empirical coverage rates, each computed over
hundreds of test points within repeat r) and therefore approximately
normal by the central limit theorem even before any normality is assumed
at the level of the R repeats themselves. Under this model, two classical
facts about sampling from a normal distribution - both consequences of
Cochran's theorem, not assumed outright - combine to produce the
t-distribution used above.
""")
    add_body(doc, """
First, the sample mean d̄ = (1/R) Σᵣ dᵣ, being a linear combination of
independent normal random variables, is itself exactly normal:
""")
    add_equation(doc, "d̄ ~ 𝒩( μ_d,  σ_d² / R )")
    add_body(doc, """
so that standardizing gives a standard normal pivot,
Z = (d̄ − μ_d) / (σ_d / √R) ~ 𝒩(0, 1) - but this pivot cannot be used
directly to test H₀: μ_d = 0, because it requires the true σ_d, which is
unknown and must be estimated from the same R repeats via the sample
standard deviation sd. Second, Cochran's theorem establishes that, for
data drawn from a normal distribution, the sample variance and sample mean
are independent random variables, and specifically that the scaled sample
variance follows a chi-squared distribution with R − 1 degrees of freedom:
""")
    add_equation(doc, "(R − 1)·sd² / σ_d²  ~  χ²_{R−1},   independent of d̄")
    add_body(doc, """
The test statistic actually used is the ratio of the standard normal pivot
Z to the square root of this independent chi-squared quantity, divided by
its own degrees of freedom:
""")
    add_equation(doc, "t = Z / √( [(R−1)sd²/σ_d²] / (R−1) ) = (d̄ − μ_d) / (σ_d/√R) · (σ_d / sd) = (d̄ − μ_d) / (sd/√R)")
    add_body(doc, """
By definition, the ratio of a standard normal random variable to the
square root of an independent chi-squared random variable divided by its
degrees of freedom follows, exactly (not merely approximately), a
Student's t-distribution with that many degrees of freedom - which is
precisely the quantity just derived, with R − 1 degrees of freedom because
one degree of freedom is used up estimating μ_d via d̄ before sd can be
computed from the residuals dᵣ − d̄. Substituting the null-hypothesis value
μ_d = 0 (no true difference in coverage between the two methods being
compared) gives exactly the test statistic t = d̄/(sd/√R) used above; the
σ_d that appeared in the derivation cancels out algebraically and never
needs to be known, which is the entire point of using a t-distribution
rather than a normal one here - it is what correctly accounts for the
extra uncertainty introduced by estimating σ_d itself from only R = 30
repeats rather than treating it as known in advance.
""")
    add_body(doc, """
The confidence-interval formula stated above follows by the standard
pivotal-quantity argument: since t = (x̄ − μ)/(s/√R) follows a
t-distribution with R − 1 degrees of freedom for any true mean μ,
regardless of its unknown value,
""")
    add_equation(doc, "P( −t₀.₀₂₅,ᵣ₋₁ ≤ (x̄ − μ)/(s/√R) ≤ t₀.₀₂₅,ᵣ₋₁ ) = 0.95")
    add_body(doc, """
Rearranging the inequality inside the probability statement to isolate μ -
purely algebraic manipulation, valid for any fixed x̄ and s - gives
P(x̄ − t₀.₀₂₅,ᵣ₋₁·s/√R ≤ μ ≤ x̄ + t₀.₀₂₅,ᵣ₋₁·s/√R) = 0.95, which is exactly
the stated 95 percent confidence interval: an interval, computed from the
observed R repeats, that would contain the true mean μ in 95 percent of
hypothetical repetitions of this entire 30-repeat procedure.
""")

    add_section_heading(doc, "4.5.2 Exchangeability Stress Test")
    add_body(doc, """
As a secondary robustness check (summarized in Section 6.13), this project
also develops a controlled violation of the
exchangeability assumption: the calibration data is held fixed exactly as
generated for the core comparison, while the test distribution's
arrival-rate multiplier is pushed progressively beyond the calibration
range's upper bound, up to several multiples of it, with staffing capacity
still drawn from its normal range so that the shift is isolated to the
arrival-rate dimension specifically. Coverage for both standard and Mondrian
CP is evaluated at each severity level on freshly generated DES scenarios,
and the same procedure is repeated with the MLP surrogate (Section 4.3.2) in
place of the gradient-boosting surrogate, to test whether the direction and
severity of any coverage breakdown depends on the surrogate's own
extrapolation behavior.
""")

    add_section_heading(doc, "4.5.3 Cross-Site Generalization")
    add_body(doc, """
The entire pipeline - DES calibration, surrogate training, conformal
prediction calibration, and the core Mondrian-versus-pooled comparison - is
independently repeated for a second hospital department present in the same
underlying dataset, with its own real arrival rate, real ESI acuity mix, and
its own Erlang-derived default capacity (Section 4.2.1), without reusing or
modifying any of the first department's data, models, or results. This
tests whether this project's core finding reflects a genuine, general
property of pooled-versus-conditional conformal calibration in a queueing
domain, or is instead an artifact specific to one department's particular
volume and acuity characteristics.
""")

    add_section_heading(doc, "4.6 Evaluation Metrics")
    add_body(doc, """
Every uncertainty quantification method compared in this project is
evaluated on the same three metrics wherever applicable: empirical coverage
(the fraction of test points whose true value falls within the constructed
interval, compared against the nominal 90 percent target), mean interval
width (the average size of the constructed interval, where narrower is
better at matched coverage), and computation time (measured as the method's
own calibration or fitting cost - a GP's model-fitting time, or a conformal
method's calibration-quantile computation time - on a like-for-like basis,
with a deliberate warm-up prediction call performed before timing begins to
exclude one-off process-startup costs such as thread-pool initialization
from the measurement).
""")


# --------------------------------------------------------------------------
# Chapter 5: Implementation (shared)
# --------------------------------------------------------------------------

def build_chapter5_implementation(doc):
    add_chapter_heading(doc, 5, "Implementation")

    add_section_heading(doc, "5.1 Dataset")
    add_body(doc, f"""
This project uses the Hospital Triage and Patient History Data dataset
{cite('kaggle_dataset')},
publicly available on Kaggle, a de-identified, MIMIC-style dataset of
560,486 emergency department visit records spanning 972 columns - vital
signs, on the order of 250 diagnosis flags, roughly 300 laboratory-result
columns, approximately 180 chief-complaint flags, ESI acuity level, and
disposition, among others. The dataset covers three distinct emergency
departments (referred to in the data as departments A, B, and C - one
academic-affiliated site and two community sites) over a period of 1,248
days (March 2014 through July 2017), a fact confirmed directly from the
dataset's own documentation rather than assumed; an earlier version of this
project's DES calibration incorrectly assumed a single year of data at a
single site, producing a daily arrival rate roughly 28 times too low before
this was caught and corrected (documented further in Section 5.3).

Department A, the largest site (322,283 of the 560,486 total visits) and
the most likely academic site by volume, is used as this project's primary
department throughout Chapters 4 and 6. Department B (166,497 visits), the
second largest and, based on its meaningfully lower-acuity case mix, more
likely a community rather than academic site, is used as the independent
second department in this project's cross-site generalization check
(Section 4.5.3). Department C is not used in this project, for reasons of
scope and time rather than any expectation it would behave differently.

An exhaustive search across all 972 columns confirmed that the dataset
contains only coarse, four-hour-bucketed arrival-hour information
(arrivalhour_bin) alongside arrival month and day fields - no
length-of-stay, treatment-duration, or discharge-timestamp field of any
kind exists anywhere in the dataset. This is a property of the dataset
itself (consistent with its de-identified, MIMIC-style construction, which
often coarsens or removes exact timestamps to reduce re-identification
risk), not something recoverable by further data processing, and it is the
direct reason this project's DES service-time distributions are
literature-calibrated rather than data-derived (Section 4.2.3).
""")

    add_section_heading(doc, "5.2 Software Stack")
    add_body(doc, """
The project is implemented in Python 3.13. The discrete-event simulation
uses SimPy for process-based discrete-event modeling. Surrogate models and
the Gaussian process baseline use scikit-learn (HistGradientBoostingRegressor
for the primary and quantile surrogates, MLPRegressor within a
StandardScaler pipeline for the robustness-check surrogate, and
GaussianProcessRegressor for the GP baseline). Conformal prediction, Mondrian
conformal prediction, and conformalized quantile regression are implemented
directly in project code (not via a third-party conformal prediction
library) so that every step of the calibration procedure - nonconformity
score computation, quantile computation with the correct finite-sample
correction, category partitioning - is fully transparent and auditable
rather than hidden inside a dependency; the mapie library was evaluated
early in the project and confirmed compatible with the Python 3.13
environment, but a direct implementation was chosen for this reason.
Data handling uses pandas, numpy, and pyarrow/pyreadr (the latter for
converting the dataset's original .rdata format to parquet for faster
loading). Statistical significance testing uses scipy. Result visualization
uses matplotlib and seaborn. Reports and presentations are generated
programmatically: slide decks via python-pptx, and both the original short
written assignments and this book-format expansion via python-docx, in both
cases with Microsoft Word or PowerPoint COM automation used to render the
generated file and visually verify it before considering the work complete
(Section 5.4).
""")

    add_section_heading(doc, "5.3 Module-by-Module Walkthrough")
    add_body(doc, """
The codebase is organized under src/ into four packages, plus supporting
scripts for reporting. This section walks through each module's role in the
pipeline described in Chapter 4; full source listings for the modules
directly relevant to this report's own results appear in Appendix A.
""")

    add_section_heading(doc, "5.3.1 src/utils/ - Distribution Extraction")
    add_body(doc, """
extract_distributions.py reads the raw dataset, filters to a single
department, and computes the real arrival-rate distributions (by hour bin,
day, and month) and ESI mix used to calibrate the DES's arrival process
(Section 4.2.3), writing them to results/tables/ as CSV files the DES reads
at runtime. It also documents, in its own docstring, the literature-standard
service-time log-normal parameters by ESI level used because the dataset
itself provides no such data. A specific bug fixed during this module's
development is worth noting for implementation correctness: the dataset's
"23-02" arrival-hour bin wraps around midnight (covering hours 23, 0, 1, and
2), but a naive hour-integer-divided-by-four binning scheme places it at
hours 0-3 instead, silently misattributing arrivals in the last hour before
midnight to the wrong bin; this was corrected with an explicit hour-to-bin
lookup table rather than an arithmetic formula.
""")

    add_section_heading(doc, "5.3.2 src/des/ - Discrete-Event Simulation")
    add_body(doc, """
er_simulation.py implements the SimPy-based ED simulation described in
Section 4.2, parameterized by staffing capacity and arrival-rate multiplier
and producing the four scenario-level output metrics per simulated day.
validate.py runs the calibrated simulation across a large number of
simulated days at its default configuration and compares mean simulated
daily patient volume against the real calibrated rate, producing the
validation result reported in Section 6.1.
""")

    add_section_heading(doc, "5.3.3 src/surrogate/ - Surrogate Training")
    add_body(doc, """
generate_training_data.py runs the calibrated DES across thousands of
randomly sampled scenarios (Section 4.2.2) to produce the labeled dataset
surrogate models are trained on. train_surrogate.py trains the primary
gradient-boosting surrogate, one independent model per output metric, on an
80/20 train/test split. train_mlp_surrogate.py trains the second,
robustness-check architecture (Section 4.3.2) on the identical split.
train_quantile_surrogates.py trains the paired lower/upper quantile
regressors used by conformalized quantile regression (Section 4.4.4), also
on the identical split, so that all surrogate variants are directly
comparable rather than trained on different data.
""")

    add_section_heading(doc, "5.3.4 src/uq/ - Uncertainty Quantification")
    add_body(doc, """
generate_calibration_data.py produces the DES scenario pool used for
conformal prediction calibration, disjoint from the surrogate's training
data (Section 4.2.2). gp_baseline.py implements the Gaussian process
baseline (Section 4.4.1). standard_cp.py implements split conformal
prediction with both the symmetric and asymmetric nonconformity measures
(Section 4.4.2). mondrian_cp.py implements the nine-category Mondrian
conformal predictor (Section 4.4.3), including both the per-category and
re-aggregated marginal evaluation views. repeated_evaluation.py implements
the 30-repeat statistical robustness check (Section 4.5.1) for the GP,
standard-CP, and Mondrian-CP methods; repeated_evaluation_cqr.py extends
this to CQR and Mondrian-CQR (Section 4.4.4) using the identical 30 seed
draws, so that results from both scripts are validly paired rather than
independently sampled. exchangeability_stress_test.py and
exchangeability_stress_test_mlp.py implement the exchangeability stress test
(Section 4.5.2) for the gradient-boosting and MLP surrogates respectively.
full_comparison.py and publication_comparison_chart.py assemble
already-computed results from the scripts above into the summary tables and
comparison figures presented in Chapter 6.
""")

    add_section_heading(doc, "5.3.5 src/generalization/ - Cross-Site Generalization")
    add_body(doc, """
This package mirrors the core pipeline (distribution extraction, DES data
generation, surrogate training, DES validation, and the standard-versus-
Mondrian CP comparison) for the second, independent hospital department used
in the cross-site generalization check (Section 4.5.3), writing all outputs
to department-B-specific subdirectories that never overwrite or share files
with the primary department's results, so that the two departments' results
can be directly compared without either one having been able to influence
the other's computation.
""")

    add_section_heading(doc, "5.4 Reproducibility and Verification Practice")
    add_body(doc, """
Every script in this codebase is re-runnable independently via the project's
virtual environment (documented in this project's README.md and
requirements.txt), and every reported number in this report traces to a
specific CSV file under results/tables/ or a
specific script under src/ - no number in this report was entered by hand
without a corresponding generation script. Random seeds are fixed and, where
multiple stages must not share data (training versus calibration versus
test; repeat r's calibration versus repeat r's test versus repeat r+1's
data), deliberately offset into disjoint ranges rather than left to
incidental non-overlap, so that the exchangeability and independence
assumptions this project's own methodology (Chapter 4) depends on are
actually satisfied by construction rather than merely assumed.

Every generated report or presentation artifact in this project - including
this one - is verified by rendering it and visually inspecting the result,
rather than trusting the generation code to have produced correct output
merely because it ran without raising an exception. This practice caught
three separate, non-obvious formatting bugs across this project's earlier
slide decks (invisible white-on-white text from an inherited shape style, a
zero-width table column from a partially-specified transform, and a
multi-line table cell rendering only its first line at the correct font size
because only the first paragraph of a multi-paragraph cell was styled) -
none of which raised any error and all of which would have shipped silently
otherwise. The same verification discipline was applied while building this
report itself: an early draft of the front matter placed a single line of
text alone on an otherwise blank page due to excess spacing, and the initial
table/figure-numbering implementation used plain styled text rather than
Word's native caption mechanism, silently breaking the automatic List of
Figures and List of Tables pages; both were caught by rendering the document
and inspecting it page by page before treating this chapter as complete,
consistent with the practice established across every earlier deliverable
in this project.
""")
