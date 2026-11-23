"""'Attendance Statement' PDF format (Times, boxed page, navy tables, signature lines).

build_statement_pdf()        one student, every subject starts on its own page (student export, share, email, WhatsApp link)
build_class_statement_pdf()  admin / teacher class or subject report in the same look

Both functions are defensive about data types (None, Decimal, int-as-text) so database differences never crash a report.
"""
import io
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (CondPageBreak, Flowable, Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None

NAVY = colors.HexColor("#12203f")
LINE = colors.HexColor("#1f2a44")
SOFT = colors.HexColor("#c9d1e0")
LABEL_BG = colors.HexColor("#eceff4")
ZEBRA = colors.HexColor("#f3f5fa")
GREEN = colors.HexColor("#1e7b34")
RED = colors.HexColor("#c4161c")
INK = colors.HexColor("#111827")

PAGE_W, PAGE_H = A4
LEFT = 51.0                       # points (18 mm)
CONTENT_W = PAGE_W - 2 * LEFT     # 493.3 pt = 174 mm


# ---------------------------------------------------------------- safe converters
def _i(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _f(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _s(v, default="-"):
    text = "" if v is None else str(v)
    return text if text.strip() else default


def fmt_date(value):
    """'2026-09-29' -> '29 Sep 2026, Tue'."""
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").strftime("%d %b %Y, %a")
    except ValueError:
        return _s(value)


def _style(name, **kw):
    base = dict(fontName="Times-Roman", fontSize=9.5, leading=12, textColor=INK, alignment=TA_LEFT)
    base.update(kw)
    return ParagraphStyle(name, **base)


S_LABEL = _style("lab", fontName="Times-Bold", fontSize=9.5)
S_VALUE = _style("val", fontSize=9.5)
S_HEAD9 = _style("h9", fontName="Times-Bold", fontSize=9.5, textColor=colors.white, alignment=TA_CENTER)
S_SUM = _style("sum", fontSize=10.5, leading=13, alignment=TA_CENTER)
S_SUM_B = _style("sumb", fontName="Times-Bold", fontSize=10.5, leading=13, alignment=TA_CENTER)
S_HEAD8 = _style("h8", fontName="Times-Bold", fontSize=8, leading=10, textColor=colors.white, alignment=TA_CENTER)
S_C8 = _style("c8", fontSize=8, leading=10, alignment=TA_CENTER)
S_L8 = _style("l8", fontSize=8, leading=10)
S_NOTE = _style("note", fontName="Times-Italic", fontSize=8.5, leading=11, textColor=colors.HexColor("#374151"))
S_SECTION = _style("sec", fontName="Times-Bold", fontSize=11.5, leading=14, spaceBefore=6, spaceAfter=5)


def _status_style(good):
    return _style("st%s" % int(good), fontName="Times-Bold", fontSize=8, leading=10, alignment=TA_CENTER, textColor=GREEN if good else RED)


def _remark(pct, total, threshold):
    if not total:
        return "-", INK
    return ("Satisfactory", GREEN) if pct >= threshold else ("Shortage", RED)


# ---------------------------------------------------------------- building blocks
class HeaderBlock(Flowable):
    """Logo, college name, tagline and the double rule (first page of every statement)."""

    def __init__(self, college, tagline, logo_path):
        super().__init__()
        self.college, self.tagline, self.logo_path = college, tagline, logo_path
        self.height = 66

    def wrap(self, aw, ah):
        return CONTENT_W, self.height

    def draw(self):
        c = self.canv
        top = self.height
        logo = Path(self.logo_path) if self.logo_path else None
        if logo and logo.exists():
            try:
                c.drawImage(str(logo), 0, top - 54, 54, 54, mask="auto", preserveAspectRatio=True)
            except Exception:
                pass
        c.setFillColor(NAVY)
        c.setFont("Times-Bold", 18)
        c.drawCentredString(CONTENT_W / 2.0 + 6, top - 20, str(self.college))
        c.setFont("Times-Italic", 9)
        c.setFillColor(colors.HexColor("#374151"))
        c.drawCentredString(CONTENT_W / 2.0 + 6, top - 35, str(self.tagline))
        c.setStrokeColor(NAVY)
        c.setLineWidth(2.4)
        c.line(0, 8, CONTENT_W, 8)
        c.setLineWidth(0.8)
        c.line(0, 3.5, CONTENT_W, 3.5)


def _spaced(text):
    words = [" ".join(w) for w in text.upper().split()]
    return "&nbsp;&nbsp;&nbsp;".join(words)


def _title(text):
    return Paragraph(_spaced(text), _style("ttl", fontName="Times-Bold", fontSize=12.5, leading=15, alignment=TA_CENTER, textColor=NAVY))


def _info_table(pairs):
    """pairs: list of 4-tuples (label1, value1, label2, value2)."""
    data = [[Paragraph(escape(_s(a)), S_LABEL), Paragraph(escape(_s(b)), S_VALUE),
             Paragraph(escape(_s(c)), S_LABEL), Paragraph(escape(_s(d)), S_VALUE)] for a, b, c, d in pairs]
    t = Table(data, colWidths=[90.7, 190.0, 72.0, CONTENT_W - 90.7 - 190.0 - 72.0], rowHeights=[23] * len(data))
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, LINE), ("BACKGROUND", (0, 0), (0, -1), LABEL_BG),
                           ("BACKGROUND", (2, 0), (2, -1), LABEL_BG), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
    return t


def _summary_table(heads, values, remark_col=None, remark_color=INK, bold_col=None):
    n = len(heads)
    head = [Paragraph(escape(h), S_HEAD9) for h in heads]
    row = []
    for i, v in enumerate(values):
        if i == remark_col:
            row.append(Paragraph(escape(_s(v)), _style("rm", fontName="Times-Bold", fontSize=10.5, leading=13, alignment=TA_CENTER, textColor=remark_color)))
        else:
            row.append(Paragraph(escape(_s(v)), S_SUM_B if i == bold_col else S_SUM))
    t = Table([head, row], colWidths=[CONTENT_W / n] * n, rowHeights=[25, 25])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("GRID", (0, 1), (-1, 1), 0.6, LINE), ("BOX", (0, 0), (-1, -1), 0.8, LINE),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    return t


def _history_table(history):
    heads = ["#", "Date", "Lecture", "Time", "Session", "Status"]
    widths = [10 * mm, 44 * mm, 20 * mm, 38 * mm, 28 * mm]
    widths.append(CONTENT_W - sum(widths))
    rows = [[Paragraph(h, S_HEAD8) for h in heads]]
    style = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("BOX", (0, 0), (-1, -1), 0.8, LINE), ("INNERGRID", (0, 1), (-1, -1), 0.4, SOFT),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]
    ordered = sorted(history, key=lambda r: (_s(r.get("attendance_date"), ""), _i(r.get("lecture_no"))))
    for n, r in enumerate(ordered, start=1):
        ok = _s(r.get("status"), "") == "Present"
        rows.append([Paragraph(str(n), S_C8), Paragraph(escape(fmt_date(r.get("attendance_date"))), S_L8),
                     Paragraph("L%s" % _i(r.get("lecture_no"), 1), S_C8), Paragraph(escape(_s(r.get("lecture_time"))), S_C8),
                     Paragraph(escape(_s(r.get("session_type"), "Lecture")), S_C8), Paragraph(escape(_s(r.get("status"))), _status_style(ok))])
        if n % 2 == 0:
            style.append(("BACKGROUND", (0, n), (-1, n), ZEBRA))
    if len(rows) == 1:
        rows.append([Paragraph("", S_C8), Paragraph("No attendance records.", S_L8)] + [Paragraph("", S_C8)] * 4)
    t = Table(rows, colWidths=widths, rowHeights=[20] * len(rows), repeatRows=1)
    t.setStyle(TableStyle(style))
    return t


def _signatures():
    cells = []
    for label in ("Subject Incharge", "Head of Department", "Principal"):
        cells.append([Paragraph("_" * 17, _style("sg", fontSize=10, leading=12, alignment=TA_CENTER)),
                      Paragraph(label, _style("sl", fontName="Times-Bold", fontSize=10, leading=12, alignment=TA_CENTER))])
    t = Table([cells], colWidths=[CONTENT_W / 3.0] * 3)
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    return KeepTogether([Spacer(1, 16), t])


class _NumberedCanvas(rl_canvas.Canvas):
    """Draws the page border, footer line and 'Page x of y' on every page."""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._saved_pages = []
        self.stamp = ""

    def showPage(self):
        self._saved_pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_pages)
        for state in self._saved_pages:
            self.__dict__.update(state)
            self.setStrokeColor(LINE)
            self.setLineWidth(0.8)
            self.rect(22.7, 18, PAGE_W - 45.4, PAGE_H - 40.7, stroke=1, fill=0)
            self.setLineWidth(0.6)
            self.line(LEFT, 34.3, PAGE_W - LEFT, 34.3)
            self.setFont("Helvetica", 7)
            self.setFillColor(colors.HexColor("#374151"))
            self.drawString(LEFT, 26, "Generated on %s | Computer-generated report" % self.stamp)
            self.drawRightString(PAGE_W - LEFT, 26, "Page %s of %s" % (self._pageNumber, total))
            super().showPage()
        super().save()


def _stamp():
    now = datetime.now(ZoneInfo("Asia/Kolkata")) if ZoneInfo else datetime.utcnow()
    return now.strftime("%d %b %Y, %I:%M %p")


def _build(story, college):
    stamp = _stamp()

    class _Canvas(_NumberedCanvas):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self.stamp = stamp

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=LEFT, rightMargin=LEFT, topMargin=50, bottomMargin=69,
                            title="Attendance Statement", author=str(college))
    doc.build(story, canvasmaker=_Canvas)
    return buf.getvalue()


# ---------------------------------------------------------------- public builders
def build_statement_pdf(student, report, *, college, tagline="Student Attendance Management System", logo_path=None, threshold=75.0):
    """student: mapping (name, prn, course, year). report: list of dicts
    (semester, code, name, total, present, absent, pct, history[list of lecture rows])."""
    thr = _f(threshold, 75.0)
    story = []
    for idx, sub in enumerate(report):
        if idx:
            story.append(PageBreak())
        total, present, absent = _i(sub.get("total")), _i(sub.get("present")), _i(sub.get("absent"))
        pct = _f(sub.get("pct"), (100.0 * present / total) if total else 0.0)
        label, color = _remark(pct, total, thr)
        story += [HeaderBlock(college, tagline, logo_path), Spacer(1, 14), _title("Attendance Statement"), Spacer(1, 17),
                  _info_table([("Name of Student", student.get("name"), "PRN", student.get("prn")),
                               ("Course", student.get("course") or "B.Pharm", "Year", student.get("year")),
                               ("Subject", "%s - %s" % (_s(sub.get("code"), ""), _s(sub.get("name"), "")), "Semester", sub.get("semester") or student.get("semester"))]),
                  Spacer(1, 11),
                  _summary_table(["Total Lectures", "Present", "Absent", "Attendance %", "Remark"],
                                 [total, present, absent, "%.1f%%" % pct, label], remark_col=4, remark_color=color, bold_col=3),
                  Spacer(1, 8), Paragraph("Minimum required attendance: %g%%." % thr, S_NOTE), Spacer(1, 12),
                  Paragraph("Lecture-wise Record", S_SECTION), _history_table(sub.get("history") or []), _signatures()]
    if not report:
        story += [HeaderBlock(college, tagline, logo_path), Spacer(1, 14), _title("Attendance Statement"), Spacer(1, 17),
                  _info_table([("Name of Student", student.get("name"), "PRN", student.get("prn")),
                               ("Course", student.get("course") or "B.Pharm", "Year", student.get("year"))]),
                  Spacer(1, 11), Paragraph("No attendance records found.", S_NOTE)]
    return _build(story, college)


def build_class_statement_pdf(rows, *, college, tagline="Student Attendance Management System", logo_path=None, threshold=75.0,
                              meta=None, student_history=None, history_title=""):
    """rows: list of dicts (prn, name, total, present, absent, pct). meta: list of (label, value) pairs."""
    thr = _f(threshold, 75.0)
    clean = []
    for r in rows:
        total, present, absent = _i(r.get("total")), _i(r.get("present")), _i(r.get("absent"))
        clean.append({"prn": r.get("prn"), "name": r.get("name"), "total": total, "present": present, "absent": absent,
                      "pct": _f(r.get("pct"), (100.0 * present / total) if total else 0.0)})
    counted = [r for r in clean if r["total"] > 0]
    avg = (sum(r["pct"] for r in counted) / len(counted)) if counted else 0.0
    meeting = sum(1 for r in counted if r["pct"] >= thr)

    pairs = []
    items = list(meta or [])
    for i in range(0, len(items), 2):
        a = items[i]
        b = items[i + 1] if i + 1 < len(items) else ("", "")
        pairs.append((a[0], a[1], b[0], b[1]))

    heads = ["#", "PRN", "Student Name", "Total", "Present", "Absent", "Attendance %", "Remark"]
    widths = [8 * mm, 34 * mm, 42 * mm, 15 * mm, 15 * mm, 15 * mm, 21 * mm]
    widths.append(CONTENT_W - sum(widths))
    table = [[Paragraph(h, S_HEAD8) for h in heads]]
    style = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("BOX", (0, 0), (-1, -1), 0.8, LINE), ("INNERGRID", (0, 1), (-1, -1), 0.4, SOFT),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]
    for n, r in enumerate(clean, start=1):
        label, color = _remark(r["pct"], r["total"], thr)
        table.append([Paragraph(str(n), S_C8), Paragraph(escape(_s(r["prn"])), S_L8), Paragraph(escape(_s(r["name"])), S_L8),
                      Paragraph(str(r["total"]), S_C8), Paragraph(str(r["present"]), S_C8), Paragraph(str(r["absent"]), S_C8),
                      Paragraph("%.1f%%" % r["pct"] if r["total"] else "-", _style("pc", fontName="Times-Bold", fontSize=8, leading=10, alignment=TA_CENTER)),
                      Paragraph(label, _style("rk", fontName="Times-Bold", fontSize=7.5, leading=9, alignment=TA_CENTER, textColor=color))])
        if n % 2 == 0:
            style.append(("BACKGROUND", (0, n), (-1, n), ZEBRA))
    if len(table) == 1:
        table.append([Paragraph("", S_C8), Paragraph("No students found for the selected filters.", S_L8)] + [Paragraph("", S_C8)] * 6)
    student_table = Table(table, colWidths=widths, rowHeights=[20] * len(table), repeatRows=1)
    student_table.setStyle(TableStyle(style))

    story = [HeaderBlock(college, tagline, logo_path), Spacer(1, 14), _title("Attendance Report"), Spacer(1, 17)]
    if pairs:
        story += [_info_table(pairs), Spacer(1, 11)]
    story += [_summary_table(["Students", "Average Attendance", "Meeting %g%% Limit" % thr, "Below %g%% Limit" % thr],
                             [len(clean), "%.1f%%" % avg, meeting, len(counted) - meeting], bold_col=1),
              Spacer(1, 8), Paragraph("Minimum required attendance: %g%%." % thr, S_NOTE), Spacer(1, 12),
              Paragraph("Student-wise Attendance", S_SECTION), student_table]
    if student_history is not None:
        story += [Spacer(1, 14), CondPageBreak(120), Paragraph(escape(_s(history_title, "Lecture-wise Record")), S_SECTION),
                  _history_table(list(student_history))]
    story.append(_signatures())
    return _build(story, college)
