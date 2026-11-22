"""
4 professional attendance PDF designs. Same data in, different look out.

    build_pdf(student, report, style="classic" | "dashboard" | "reportcard" | "hero" |
              "maroon" | "mono" | "sky" | "dark",
              threshold=75, college="Personal Attendance Management", logo_path=None)

`student` : dict-like with name, prn, course, year
`report`  : list of subjects, each {semester, code, name, total, present, absent, pct, history:[...]}
            (exactly what app.py already passes to build_subject_wise_pdf)

To use inside app.py, keep this file next to app.py and replace the body of
build_subject_wise_pdf with:

    from attendance_pdf_designs import build_pdf
    return build_pdf(student, report, style="dashboard", threshold=attendance_threshold_value(db),
                     college=COLLEGE_NAME, logo_path=BASE_DIR / "static" / LOGO_FILE)
"""
import io
from datetime import datetime
from collections import OrderedDict
from pathlib import Path
from xml.sax.saxutils import escape as esc
from zoneinfo import ZoneInfo

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
                                TableStyle, Image, PageBreak, Flowable, KeepTogether)
from reportlab.platypus.flowables import HRFlowable

H = colors.HexColor
GREEN, GREEN_BG = H("#15803d"), H("#e3f5ea")
RED, RED_BG = H("#b91c1c"), H("#fdeaea")
AMBER, AMBER_BG = H("#b45309"), H("#fff3dc")
_BASE = getSampleStyleSheet()["Normal"]


# ----------------------------------------------------------------- helpers
def hx(c):
    return "#" + c.hexval()[2:]


def P(name, **kw):
    kw.setdefault("fontName", "Helvetica")
    return ParagraphStyle(name, parent=_BASE, **kw)


def verdict(pct, thr):
    if pct >= thr:
        return "GOOD STANDING", GREEN, GREEN_BG
    if pct >= thr - 10:
        return "NEEDS ATTENTION", AMBER, AMBER_BG
    return "LOW ATTENDANCE", RED, RED_BG


def fdate(v):
    try:
        return datetime.strptime(str(v), "%Y-%m-%d").strftime("%d %b %Y, %a")
    except Exception:
        return str(v or "-")


def month_summary(history):
    """[(label, present, total)] ordered by month."""
    m = OrderedDict()
    for r in sorted(history, key=lambda r: str(r["attendance_date"])):
        try:
            d = datetime.strptime(str(r["attendance_date"]), "%Y-%m-%d")
        except Exception:
            continue
        k = d.strftime("%b %y")
        p, t = m.get(k, (0, 0))
        m[k] = (p + (str(r["status"]) == "Present"), t + 1)
    return [(k, p, t) for k, (p, t) in m.items()]


def numbered_canvas(footer):
    class NC(rl_canvas.Canvas):
        def __init__(s, *a, **k):
            super().__init__(*a, **k)
            s._saved = []

        def showPage(s):
            s._saved.append(dict(s.__dict__))
            s._startPage()

        def save(s):
            n = len(s._saved)
            for st in s._saved:
                s.__dict__.update(st)
                footer(s, s._pageNumber, n)
                super().showPage()
            super().save()
    return NC


def make_footer(line_col, text_col, font="Helvetica", left_pad=14*mm, right_pad=14*mm, lift=0, page_w=None):
    def footer(c, n, total):
        w = page_w or A4[0]
        c.setStrokeColor(line_col); c.setLineWidth(0.6)
        c.line(left_pad, 14*mm + lift, w - right_pad, 14*mm + lift)
        c.setFont(font, 7.5); c.setFillColor(text_col)
        stamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d %b %Y, %I:%M %p")
        c.drawString(left_pad, 9.5*mm + lift, f"Generated on {stamp}  |  Computer-generated report")
        c.drawRightString(w - right_pad, 9.5*mm + lift, f"Page {n} of {total}")
    return footer


# -------------------------------------------------------------- flowables
class LogoMark(Flowable):
    """Real logo if the file exists, otherwise a clean initials badge."""
    def __init__(self, size, name, path, fg, bg):
        super().__init__()
        self.size, self.name, self.path, self.fg, self.bg = size, name, path, fg, bg

    def wrap(self, *_):
        return self.size, self.size

    def draw(self):
        c, s = self.canv, self.size
        if self.path and Path(str(self.path)).exists():
            c.drawImage(str(self.path), 0, 0, s, s, preserveAspectRatio=True, mask="auto")
            return
        c.setFillColor(self.bg); c.circle(s/2, s/2, s/2, fill=1, stroke=0)
        initials = "".join(w[0] for w in str(self.name).split()[:3]).upper() or "A"
        c.setFillColor(self.fg); c.setFont("Helvetica-Bold", s*0.34)
        c.drawCentredString(s/2, s/2 - s*0.12, initials)


class Donut(Flowable):
    def __init__(self, pct, color, size, thick, track, text_col, sub="ATTENDANCE"):
        super().__init__()
        self.pct, self.color, self.size, self.thick = pct, color, size, thick
        self.track, self.text_col, self.sub = track, text_col, sub

    def wrap(self, *_):
        return self.size, self.size

    def draw(self):
        c, s = self.canv, self.size
        cx = cy = s / 2
        r = s/2 - self.thick/2
        c.setLineWidth(self.thick); c.setStrokeColor(self.track); c.circle(cx, cy, r, fill=0, stroke=1)
        if self.pct > 0:
            c.setStrokeColor(self.color); c.setLineCap(1)
            c.arc(cx-r, cy-r, cx+r, cy+r, 90, -3.6 * min(self.pct, 99.99))
        c.setFillColor(self.text_col); c.setFont("Helvetica-Bold", s*0.2)
        c.drawCentredString(cx, cy - s*0.03, f"{self.pct:.1f}%")
        c.setFillColor(H("#64748b")); c.setFont("Helvetica", s*0.075)
        c.drawCentredString(cx, cy - s*0.16, self.sub)


class MonthBars(Flowable):
    def __init__(self, data, width, height, thr, accent):
        super().__init__()
        self.data, self.width, self.height, self.thr, self.accent = data, width, height, thr, accent

    def wrap(self, *_):
        return self.width, self.height

    def draw(self):
        c, w, h = self.canv, self.width, self.height
        base, top = 6*mm, h - 5*mm
        usable = top - base
        n = max(1, len(self.data))
        slot = w / n
        bw = min(16*mm, slot * 0.5)
        ty = base + usable * self.thr / 100
        c.setStrokeColor(H("#94a3b8")); c.setDash(2, 2); c.setLineWidth(0.6)
        c.line(0, ty, w, ty); c.setDash()
        c.setFillColor(H("#94a3b8")); c.setFont("Helvetica", 6.5)
        c.drawRightString(w, ty + 1.2, f"{self.thr:g}% required")
        for i, (lab, p, t) in enumerate(self.data):
            pct = 100.0 * p / t if t else 0
            x = i*slot + (slot - bw)/2
            c.setFillColor(H("#e8eef5")); c.roundRect(x, base, bw, usable, 2, fill=1, stroke=0)
            col = GREEN if pct >= self.thr else (AMBER if pct >= self.thr - 10 else RED)
            c.setFillColor(col)
            c.roundRect(x, base, bw, max(2, usable*pct/100), 2, fill=1, stroke=0)
            c.setFillColor(H("#1e2b3a")); c.setFont("Helvetica-Bold", 7.5)
            c.drawCentredString(x + bw/2, base + usable*pct/100 + 1.5, f"{pct:.0f}%")
            c.setFillColor(H("#64748b")); c.setFont("Helvetica", 7.5)
            c.drawCentredString(x + bw/2, base - 4*mm, lab)


class ThinBar(Flowable):
    def __init__(self, pct, thr, color, width, track="#e5ebf3", marker="#0b2f5b"):
        super().__init__()
        self.pct, self.thr, self.color, self.width = pct, thr, color, width
        self.track, self.marker = H(track), H(marker)

    def wrap(self, *_):
        return self.width, 9*mm

    def draw(self):
        c, w, h, y = self.canv, self.width, 4*mm, 4*mm
        c.setFillColor(self.track); c.roundRect(0, y, w, h, h/2, fill=1, stroke=0)
        if self.pct > 0:
            c.setFillColor(self.color); c.roundRect(0, y, max(h, w*min(self.pct, 100)/100), h, h/2, fill=1, stroke=0)
        mx = w * self.thr / 100
        c.setStrokeColor(self.marker); c.setLineWidth(1); c.line(mx, y-1*mm, mx, y+h+1*mm)
        c.setFillColor(H("#64748b")); c.setFont("Helvetica", 6.8)
        c.drawCentredString(mx, 0.3*mm, f"Required {self.thr:g}%")


# -------------------------------------------------------- shared table bits
def history_table(sub, W, *, head_bg, head_fg=colors.white, zebra=("#ffffff", "#f5f8fc"), line="#d5dfeb",
                  font="Helvetica", status="fill", strip=False, pad=4.5):
    fb = font + "-Bold" if font != "Times-Roman" else "Times-Bold"
    cell = P("hc", fontName=font, fontSize=8.5, leading=11, textColor=H("#1e2b3a"))
    cc = P("hcc", fontName=font, fontSize=8.5, leading=11, textColor=H("#1e2b3a"), alignment=TA_CENTER)
    hd = P("hh", fontName=fb, fontSize=8.5, leading=11, textColor=head_fg, alignment=TA_CENTER)
    first = "" if strip else "#"
    data = [[Paragraph(h, hd) for h in (first, "Date", "Lecture", "Time", "Session", "Status")]]
    extra = []
    for i, r in enumerate(sub["history"], start=1):
        ok = str(r["status"]) == "Present"
        col = GREEN if ok else RED
        word = "Present" if ok else "Absent"
        data.append([Paragraph("" if strip else str(i), cc), Paragraph(fdate(r["attendance_date"]), cell),
                     Paragraph(f"L{r['lecture_no']}", cc), Paragraph(esc(str(r["lecture_time"] or "-")), cc),
                     Paragraph(esc(str(r["session_type"] or "Lecture")), cc),
                     Paragraph(f'<font color="{hx(col)}" name="{fb}">{word}</font>', cc)])
        if status == "fill":
            extra.append(("BACKGROUND", (5, i), (5, i), GREEN_BG if ok else RED_BG))
        if strip:
            extra.append(("BACKGROUND", (0, i), (0, i), col))
    if len(data) == 1:
        data.append([Paragraph("", cell), Paragraph("No attendance records yet", cell)] + [Paragraph("", cell)] * 4)
    first_w = 4*mm if strip else 10*mm
    widths = [first_w, 44*mm, 20*mm, 38*mm, 28*mm]
    widths.append(W - sum(widths))
    t = Table(data, colWidths=widths, repeatRows=1)
    if strip:
        extra += [("LEFTPADDING", (0, 0), (0, -1), 0), ("RIGHTPADDING", (0, 0), (0, -1), 0)]
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), head_bg), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [H(zebra[0]), H(zebra[1])]),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, H(line)), ("BOX", (0, 0), (-1, -1), 0.6, H(line)),
        ("TOPPADDING", (0, 0), (-1, -1), pad), ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
    ] + extra))
    return t


def pill(text, bg, w=38*mm):
    t = Table([[Paragraph(text, P("pl", fontName="Helvetica-Bold", fontSize=8.5, leading=11,
                                  textColor=colors.white, alignment=TA_CENTER))]], colWidths=[w])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), bg),
                           ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    return t


# =========================================================== 1. CLASSIC
def design_classic(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    NAVY = H("#14213d")
    pct = float(sub["pct"])
    name = P("n", fontName="Times-Bold", fontSize=19, leading=22, textColor=NAVY, alignment=TA_CENTER)
    sub1 = P("s", fontName="Times-Italic", fontSize=9.5, leading=12, textColor=H("#475569"), alignment=TA_CENTER)
    title = P("t", fontName="Times-Bold", fontSize=13, leading=16, textColor=NAVY, alignment=TA_CENTER)
    lab = P("l", fontName="Times-Bold", fontSize=10, leading=13, textColor=NAVY)
    val = P("v", fontName="Times-Roman", fontSize=10, leading=13, textColor=H("#111827"))
    hdr = Table([[LogoMark(19*mm, college, ctx["logo"], colors.white, NAVY),
                  [Paragraph(esc(college), name), Paragraph(esc(ctx.get("tagline") or "Student Attendance Management System"), sub1)],
                  ""]], colWidths=[24*mm, W - 48*mm, 24*mm])
    hdr.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    out = [hdr, Spacer(1, 3*mm),
           HRFlowable(width="100%", thickness=2, color=NAVY, spaceAfter=1.5),
           HRFlowable(width="100%", thickness=0.6, color=NAVY, spaceAfter=5*mm),
           Paragraph("A T T E N D A N C E &nbsp; S T A T E M E N T", title), Spacer(1, 5*mm)]
    rows = [[Paragraph("Name of Student", lab), Paragraph(esc(str(student["name"] or "-")), val),
             Paragraph("PRN", lab), Paragraph(esc(str(student["prn"] or "-")), val)],
            [Paragraph("Course", lab), Paragraph(esc(str(student["course"] or "B.Pharm")), val),
             Paragraph("Year", lab), Paragraph(esc(str(student["year"] or "-")), val)],
            [Paragraph("Subject", lab), Paragraph(esc(f"{sub['code']} - {sub['name']}"), val),
             Paragraph("Semester", lab), Paragraph(str(sub["semester"]), val)]]
    t = Table(rows, colWidths=[32*mm, W/2 - 32*mm + 12*mm, 24*mm, W - (W/2 + 12*mm) - 24*mm + 32*mm - 32*mm + 0])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, NAVY), ("BACKGROUND", (0, 0), (0, -1), H("#eef1f6")),
                           ("BACKGROUND", (2, 0), (2, -1), H("#eef1f6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    out += [t, Spacer(1, 5*mm)]
    hd = P("hd", fontName="Times-Bold", fontSize=10, leading=13, textColor=colors.white, alignment=TA_CENTER)
    cv = P("cv", fontName="Times-Roman", fontSize=11, leading=14, alignment=TA_CENTER)
    ok = pct >= thr
    summ = Table([[Paragraph(h, hd) for h in ("Total Lectures", "Present", "Absent", "Attendance %", "Remark")],
                  [Paragraph(str(sub["total"]), cv), Paragraph(str(sub["present"]), cv),
                   Paragraph(str(sub["absent"]), cv),
                   Paragraph(f"<b>{pct:.1f}%</b>", cv),
                   Paragraph(f'<font color="{hx(GREEN if ok else RED)}"><b>{"Satisfactory" if ok else "Shortage"}</b></font>', cv)]],
                 colWidths=[W*0.2]*5)
    summ.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("GRID", (0, 0), (-1, -1), 0.6, NAVY),
                              ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    out += [summ, Spacer(1, 2*mm),
            Paragraph(f"<i>Minimum required attendance: {thr:g}%.</i>", P("note", fontName="Times-Italic", fontSize=9, textColor=H("#475569"))),
            Spacer(1, 5*mm),
            Paragraph("<b>Lecture-wise Record</b>", P("rh", fontName="Times-Bold", fontSize=12, leading=15, textColor=NAVY, spaceAfter=4)),
            history_table(sub, W, head_bg=NAVY, zebra=("#ffffff", "#f6f7fa"), line="#9aa5b8", font="Times-Roman", status="text"),
            Spacer(1, 14*mm)]
    sig = Table([["_______________", "_______________", "_______________"],
                 ["Subject Incharge", "Head of Department", "Principal"]], colWidths=[W/3]*3)
    sig.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTNAME", (0, 1), (-1, 1), "Times-Bold"),
                             ("FONTSIZE", (0, 1), (-1, 1), 9.5), ("TEXTCOLOR", (0, 0), (-1, -1), NAVY)]))
    out.append(KeepTogether([sig]))
    return out


# ======================================================== 2. DASHBOARD
def design_dashboard(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    TEAL, MINT = H("#0f766e"), H("#ecfdf5")
    pct = float(sub["pct"])
    label, tone, tone_bg = verdict(pct, thr)
    big = P("b", fontName="Helvetica-Bold", fontSize=15, leading=18, textColor=colors.white)
    small = P("sm", fontSize=8, leading=11, textColor=H("#a7f3d0"))
    stamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d %b %Y")
    hdr = Table([[LogoMark(15*mm, college, ctx["logo"], TEAL, colors.white),
                  [Paragraph(esc(college), big), Paragraph("ATTENDANCE REPORT", small)],
                  Paragraph(f"<b>SEM {sub['semester']}</b><br/>{stamp}", P("r", fontSize=8.5, leading=12, textColor=colors.white, alignment=TA_RIGHT))]],
                colWidths=[21*mm, W - 21*mm - 38*mm, 38*mm])
    hdr.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), TEAL), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                             ("LEFTPADDING", (0, 0), (0, 0), 6*mm), ("RIGHTPADDING", (-1, 0), (-1, 0), 6*mm),
                             ("TOPPADDING", (0, 0), (-1, -1), 5*mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 5*mm),
                             ("LINEBELOW", (0, 0), (-1, -1), 3, H("#2dd4bf"))]))
    nm = P("nm", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=H("#0f172a"))
    meta = P("mt", fontSize=9, leading=13, textColor=H("#475569"))
    who = [Paragraph(esc(str(student["name"] or "-")), nm),
           Paragraph(esc(f"PRN {student['prn'] or '-'}   |   {student['course'] or 'B.Pharm'}   |   {student['year'] or '-'}"), meta)]
    def stat(num, cap, col):
        return [Paragraph(f'<font size="19" color="{hx(col)}"><b>{num}</b></font>', P("sn", leading=23, alignment=TA_CENTER)),
                Paragraph(cap, P("sc", fontSize=7.5, leading=10, textColor=H("#64748b"), alignment=TA_CENTER))]
    gw = (W - 62*mm) / 2
    cards = Table([[stat(sub["total"], "TOTAL LECTURES", H("#0f172a")), stat(sub["present"], "PRESENT", GREEN)],
                   [stat(sub["absent"], "ABSENT", RED), stat(f"{max(0, thr - pct):.1f}%" if pct < thr else "OK",
                                                           "SHORT OF REQUIRED" if pct < thr else "REQUIREMENT MET", tone)]],
                  colWidths=[gw, gw], rowHeights=[19*mm, 19*mm])
    cards.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), MINT), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("INNERGRID", (0, 0), (-1, -1), 2, colors.white), ("BOX", (0, 0), (-1, -1), 2, colors.white)]))
    hero = Table([[Donut(pct, tone, 56*mm, 9, H("#e2e8f0"), H("#0f172a")), cards]], colWidths=[62*mm, W - 62*mm])
    hero.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    st = P("st", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=H("#0f172a"))
    ss = P("ss", fontSize=8.5, leading=11, textColor=H("#64748b"))
    trow = Table([[[Paragraph(esc(sub["name"]), st), Paragraph(esc(f"Subject code: {sub['code']}"), ss)], pill(label, tone)]],
                 colWidths=[W - 40*mm, 40*mm])
    trow.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    sect = P("se", fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=TEAL, spaceBefore=4, spaceAfter=3)
    out = [hdr, Spacer(1, 6*mm)] + who + [Spacer(1, 4*mm), trow, Spacer(1, 3*mm), hero, Spacer(1, 3*mm)]
    months = month_summary(sub["history"])
    if months:
        out += [Paragraph("MONTH-WISE ATTENDANCE", sect), MonthBars(months, W, 38*mm, thr, TEAL), Spacer(1, 3*mm)]
    out += [Paragraph("LECTURE-WISE HISTORY", sect),
            history_table(sub, W, head_bg=TEAL, zebra=("#ffffff", "#f0fdfa"), line="#cfe8e3")]
    return out


# ======================================================== 3. REPORT CARD
def design_reportcard(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    NAVY, GOLD, CREAM = H("#1b2a49"), H("#b8892b"), H("#fbf8f1")
    pct = float(sub["pct"])
    grade, gcol = ("EXCELLENT", GREEN) if pct >= 90 else (("SATISFACTORY", H("#1d4ed8")) if pct >= thr else ("SHORTAGE", RED))
    cn = P("cn", fontName="Times-Bold", fontSize=20, leading=24, textColor=NAVY, alignment=TA_CENTER)
    out = [Table([[LogoMark(18*mm, college, ctx["logo"], colors.white, NAVY)]], colWidths=[W], style=[("ALIGN", (0, 0), (-1, -1), "CENTER")]),
           Spacer(1, 2*mm), Paragraph(esc(college), cn), HRFlowable(width="40%", thickness=1.2, color=GOLD, spaceBefore=2, spaceAfter=4*mm)]
    rb = Table([[Paragraph("ATTENDANCE  REPORT  CARD", P("rb", fontName="Times-Bold", fontSize=14, leading=18, textColor=H("#f3d98b"), alignment=TA_CENTER))]],
               colWidths=[W])
    rb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("TOPPADDING", (0, 0), (-1, -1), 7),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("BOX", (0, 0), (-1, -1), 1.2, GOLD)]))
    out += [rb, Spacer(1, 5*mm),
            Paragraph(esc(str(student["name"] or "-")), P("sn", fontName="Times-Bold", fontSize=18, leading=22, textColor=NAVY, alignment=TA_CENTER)),
            Paragraph(esc(f"PRN {student['prn'] or '-'}   |   {student['course'] or 'B.Pharm'}   |   {student['year'] or '-'}   |   Semester {sub['semester']}"),
                      P("sm", fontName="Times-Roman", fontSize=10.5, leading=14, textColor=H("#475569"), alignment=TA_CENTER)),
            Spacer(1, 4*mm)]
    num = P("num", fontName="Times-Bold", fontSize=22, leading=26, alignment=TA_CENTER)
    cap = P("cap", fontName="Times-Roman", fontSize=8.5, leading=11, textColor=H("#64748b"), alignment=TA_CENTER)
    def m(v, c, col):
        return [Paragraph(f'<font color="{hx(col)}">{v}</font>', num), Paragraph(c, cap)]
    box = Table([[Paragraph(f"<b>{esc(sub['code'])}</b> &nbsp;-&nbsp; {esc(sub['name'])}",
                            P("sb", fontName="Times-Roman", fontSize=12, leading=15, textColor=NAVY, alignment=TA_CENTER))],
                 [Table([[m(sub["total"], "TOTAL LECTURES", NAVY), m(sub["present"], "PRESENT", GREEN),
                          m(sub["absent"], "ABSENT", RED), m(f"{pct:.1f}%", grade, gcol)]], colWidths=[(W - 8)/4]*4,
                        style=[("LINEAFTER", (0, 0), (2, 0), 0.6, GOLD), ("ALIGN", (0, 0), (-1, -1), "CENTER")])]],
                colWidths=[W])
    box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 1.2, GOLD), ("BACKGROUND", (0, 0), (-1, -1), CREAM),
                             ("LINEBELOW", (0, 0), (-1, 0), 0.6, GOLD), ("TOPPADDING", (0, 0), (-1, -1), 8),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    out += [box, Spacer(1, 5*mm)]
    months = month_summary(sub["history"])
    if months:
        hd = P("mh", fontName="Times-Bold", fontSize=9.5, leading=12, textColor=colors.white, alignment=TA_CENTER)
        cc = P("mc", fontName="Times-Roman", fontSize=9.5, leading=12, alignment=TA_CENTER)
        rows = [[Paragraph(h, hd) for h in ("Month", "Lectures", "Present", "Absent", "Attendance")]]
        for lab, p, t in months:
            mp = 100.0 * p / t
            col = GREEN if mp >= thr else RED
            rows.append([Paragraph(lab, cc), Paragraph(str(t), cc), Paragraph(str(p), cc), Paragraph(str(t - p), cc),
                         Paragraph(f'<font color="{hx(col)}"><b>{mp:.1f}%</b></font>', cc)])
        mt = Table(rows, colWidths=[W/5]*5)
        mt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CREAM]),
                                ("BOX", (0, 0), (-1, -1), 0.8, GOLD), ("LINEBELOW", (0, 1), (-1, -1), 0.4, H("#e6dcc2")),
                                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        out += [Paragraph("<b>Month-wise Summary</b>", P("h2", fontName="Times-Bold", fontSize=12, leading=15, textColor=NAVY, spaceAfter=3)), mt, Spacer(1, 5*mm)]
    out += [Paragraph("<b>Lecture-wise Record</b>", P("h3", fontName="Times-Bold", fontSize=12, leading=15, textColor=NAVY, spaceAfter=3)),
            history_table(sub, W, head_bg=NAVY, zebra=("#ffffff", "#fbf8f1"), line="#d9cba3", font="Times-Roman", status="text")]
    return out


# ============================================================ 4. HERO
def design_hero(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    INDIGO, LAV = H("#312e81"), H("#eef2ff")
    pct = float(sub["pct"])
    label, tone, tone_bg = verdict(pct, thr)
    top = Table([[LogoMark(11*mm, college, ctx["logo"], colors.white, INDIGO),
                  Paragraph(f"<b>{esc(college)}</b>", P("c", fontSize=12, leading=15, textColor=INDIGO)),
                  Paragraph("ATTENDANCE REPORT", P("ar", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=H("#6366f1"), alignment=TA_RIGHT))]],
                colWidths=[15*mm, W - 55*mm, 40*mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (-1, 0), (-1, 0), 0)]))
    left = [Paragraph("ATTENDANCE", P("l1", fontSize=8, leading=10, textColor=H("#c7d2fe"))),
            Paragraph(f"<b>{pct:.1f}%</b>", P("l2", fontName="Helvetica-Bold", fontSize=34, leading=40, textColor=colors.white)),
            Paragraph(label, P("l3", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=H("#fde68a") if tone != GREEN else H("#86efac"))),
            Spacer(1, 3*mm),
            Paragraph(f"{sub['present']} of {sub['total']} lectures attended", P("l4", fontSize=8.5, leading=11, textColor=H("#c7d2fe")))]
    k = P("k", fontSize=7.5, leading=9, textColor=H("#6366f1"))
    v = P("v", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=H("#1e1b4b"))
    right = [Paragraph(esc(str(student["name"] or "-")), P("nm", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=INDIGO)),
             Spacer(1, 2*mm),
             Table([[[Paragraph("PRN", k), Paragraph(esc(str(student["prn"] or "-")), v)],
                     [Paragraph("COURSE", k), Paragraph(esc(str(student["course"] or "B.Pharm")), v)]],
                    [[Paragraph("YEAR", k), Paragraph(esc(str(student["year"] or "-")), v)],
                     [Paragraph("SEMESTER", k), Paragraph(str(sub["semester"]), v)]]],
                   colWidths=[(W*0.62 - 14*mm)/2]*2, style=[("LEFTPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]),
             Paragraph(f"<b>{esc(sub['code'])}</b> - {esc(sub['name'])}", P("sj", fontSize=9.5, leading=12, textColor=H("#3730a3")))]
    hero = Table([[left, right]], colWidths=[W*0.38, W*0.62])
    hero.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), INDIGO), ("BACKGROUND", (1, 0), (1, 0), LAV),
                              ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 7*mm),
                              ("TOPPADDING", (0, 0), (-1, -1), 7*mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 7*mm),
                              ("LINEBELOW", (0, 0), (-1, -1), 3, H("#f59e0b"))]))
    def stat(n, c, col):
        return [Paragraph(f'<font color="{hx(col)}"><b>{n}</b></font>', P("sn", fontName="Helvetica-Bold", fontSize=18, leading=22)),
                Paragraph(c, P("sc", fontSize=7.5, leading=10, textColor=H("#64748b")))]
    strip = Table([[stat(sub["total"], "TOTAL LECTURES", INDIGO), stat(sub["present"], "PRESENT", GREEN), stat(sub["absent"], "ABSENT", RED)]],
                  colWidths=[W/3]*3)
    strip.setStyle(TableStyle([("LINEBEFORE", (0, 0), (-1, -1), 3, LAV), ("LEFTPADDING", (0, 0), (-1, -1), 6*mm),
                               ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    sect = P("se", fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=INDIGO, spaceBefore=2, spaceAfter=3)
    return [top, Spacer(1, 5*mm), hero, Spacer(1, 3*mm), strip, ThinBar(pct, thr, tone, W, marker="#312e81"), Spacer(1, 2*mm),
            Paragraph("LECTURE-WISE HISTORY", sect),
            history_table(sub, W, head_bg=INDIGO, zebra=("#ffffff", "#f7f8ff"), line="#dfe3fa", status="text", strip=True)]



# ------------------------------------------------ extra flowables (5-8)
class SplitBar(Flowable):
    """Present vs absent proportion bar with labels inside."""
    def __init__(self, present, absent, width, height=8*mm):
        super().__init__()
        self.present, self.absent, self.width, self.height = present, absent, width, height

    def wrap(self, *_):
        return self.width, self.height

    def draw(self):
        c, w, h = self.canv, self.width, self.height
        tot = max(1, self.present + self.absent)
        pw = w * self.present / tot
        c.setFillColor(H("#ef4444")); c.roundRect(0, 0, w, h, h/2, fill=1, stroke=0)
        if self.present:
            c.setFillColor(H("#22c55e")); c.roundRect(0, 0, max(h, pw), h, h/2, fill=1, stroke=0)
        c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 8)
        if pw > 30*mm:
            c.drawString(4*mm, h/2 - 2.8, f"Present  {self.present}")
        if w - pw > 28*mm:
            c.drawRightString(w - 4*mm, h/2 - 2.8, f"Absent  {self.absent}")


class Heatmap(Flowable):
    """One square per lecture (in order): green = present, red = absent."""
    def __init__(self, statuses, width, cell=6.2*mm, gap=1.6*mm):
        super().__init__()
        self.st, self.width, self.cell, self.gap = statuses, width, cell, gap
        self.cols = max(1, int((width + gap) // (cell + gap)))
        self.rows = max(1, -(-len(statuses) // self.cols))
        self.height = self.rows * (cell + gap)

    def wrap(self, *_):
        return self.width, self.height

    def draw(self):
        c = self.canv
        for i, ok in enumerate(self.st):
            r, col = divmod(i, self.cols)
            x = col * (self.cell + self.gap)
            y = self.height - (r + 1) * (self.cell + self.gap) + self.gap
            c.setFillColor(H("#16a34a") if ok else H("#dc2626"))
            c.roundRect(x, y, self.cell, self.cell, 1.6, fill=1, stroke=0)
            c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 6.3)
            c.drawCentredString(x + self.cell/2, y + self.cell/2 - 2.2, str(i + 1))


def _streaks(history):
    best = cur = 0
    for r in history:
        if str(r["status"]) == "Absent":
            cur += 1; best = max(best, cur)
        else:
            cur = 0
    run = 0
    for r in reversed(history):
        if str(r["status"]) == "Present":
            run += 1
        else:
            break
    return best, run


def _sign_row(W, labels, font="Times-Bold", col="#222222"):
    t = Table([["_" * 22] * len(labels), list(labels)], colWidths=[W / len(labels)] * len(labels))
    t.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTNAME", (0, 1), (-1, 1), font),
                           ("FONTSIZE", (0, 1), (-1, 1), 9.5), ("TEXTCOLOR", (0, 0), (-1, -1), H(col))]))
    return t


# =================================================== 5. CLASSIC MAROON
def design_maroon(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    MAR, GOLD = H("#7a1f2b"), H("#a67c52")
    pct = float(sub["pct"]); ok = pct >= thr
    name = P("n", fontName="Times-Bold", fontSize=18, leading=21, textColor=MAR)
    sub1 = P("s", fontName="Times-Italic", fontSize=10.5, leading=13, textColor=H("#57534e"))
    right = P("r", fontName="Times-Roman", fontSize=10, leading=13, textColor=H("#57534e"), alignment=TA_RIGHT)
    top = Table([[LogoMark(18*mm, college, ctx["logo"], colors.white, MAR),
                  [Paragraph(esc(college), name), Paragraph("Student Attendance Report", sub1)],
                  Paragraph(f"<b>Semester {sub['semester']}</b><br/>{datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%d %B %Y')}", right)]],
                colWidths=[23*mm, W - 23*mm - 42*mm, 42*mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (-1, 0), (-1, 0), 0)]))
    lab = P("l", fontName="Times-Bold", fontSize=10.5, leading=14, textColor=MAR)
    val = P("v", fontName="Times-Roman", fontSize=10.5, leading=14, textColor=H("#1c1917"))
    det = Table([[Paragraph("Student", lab), Paragraph(esc(str(student["name"] or "-")), val)],
                 [Paragraph("PRN", lab), Paragraph(esc(str(student["prn"] or "-")), val)],
                 [Paragraph("Course / Year", lab), Paragraph(esc(f"{student['course'] or 'B.Pharm'} / {student['year'] or '-'}"), val)],
                 [Paragraph("Subject", lab), Paragraph(esc(f"{sub['code']} - {sub['name']}"), val)]],
                colWidths=[28*mm, W*0.56 - 28*mm])
    det.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.4, H("#d6cfc7")), ("TOPPADDING", (0, 0), (-1, -1), 4),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("LEFTPADDING", (0, 0), (0, -1), 0)]))
    big = P("b", fontName="Times-Bold", fontSize=30, leading=34, alignment=TA_CENTER, textColor=GREEN if ok else RED)
    cap = P("c", fontName="Times-Roman", fontSize=9, leading=12, alignment=TA_CENTER, textColor=H("#57534e"))
    box = Table([[Paragraph("ATTENDANCE", P("bt", fontName="Times-Bold", fontSize=9, leading=12, alignment=TA_CENTER, textColor=MAR))],
                 [Paragraph(f"{pct:.1f}%", big)],
                 [Paragraph(f"{sub['present']} present  |  {sub['absent']} absent  |  {sub['total']} total", cap)],
                 [Paragraph(f'<b>{"Meets the required " if ok else "Below the required "}{thr:g}%</b>', P("m", fontName="Times-Roman", fontSize=9, leading=12, alignment=TA_CENTER, textColor=GREEN if ok else RED))]],
                colWidths=[W*0.40])
    box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 1.2, MAR), ("BACKGROUND", (0, 0), (-1, 0), H("#f6ece9")),
                             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    both = Table([[det, box]], colWidths=[W*0.58, W*0.42])
    both.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    t = history_table(sub, W, head_bg=colors.white, head_fg=MAR, zebra=("#ffffff", "#faf6f3"), line="#d6cfc7", font="Times-Roman", status="text")
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, 0), 1.6, MAR), ("LINEABOVE", (0, 0), (-1, 0), 1.6, MAR), ("BOX", (0, 0), (-1, -1), 0, colors.white)]))
    return [top, Spacer(1, 3*mm), HRFlowable(width="100%", thickness=2.2, color=MAR, spaceAfter=1.2),
            HRFlowable(width="100%", thickness=0.6, color=GOLD, spaceAfter=6*mm),
            Paragraph("Attendance Statement", P("tt", fontName="Times-Bold", fontSize=15, leading=19, textColor=MAR, spaceAfter=4)),
            both, Spacer(1, 6*mm),
            Paragraph("Lecture-wise Record", P("h", fontName="Times-Bold", fontSize=12, leading=15, textColor=MAR, spaceAfter=3)),
            t, Spacer(1, 14*mm), KeepTogether([_sign_row(W, ("Subject Incharge", "Principal"), col="#7a1f2b")])]


# ===================================================== 6. CLASSIC MONO
def design_mono(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    INK = H("#1f1f1f")
    pct = float(sub["pct"]); ok = pct >= thr
    out = [Paragraph(esc(college).upper(), P("n", fontName="Times-Bold", fontSize=17, leading=21, textColor=INK, alignment=TA_CENTER)),
           Paragraph("Student Attendance Management System", P("s", fontName="Times-Italic", fontSize=10, leading=13, textColor=H("#555555"), alignment=TA_CENTER)),
           Spacer(1, 3*mm), HRFlowable(width="100%", thickness=1.4, color=INK, spaceAfter=1.5),
           HRFlowable(width="100%", thickness=0.4, color=INK, spaceAfter=5*mm),
           Paragraph("STATEMENT OF ATTENDANCE", P("t", fontName="Times-Bold", fontSize=12.5, leading=16, textColor=INK, alignment=TA_CENTER)),
           Paragraph(f"Semester {sub['semester']}  |  {esc(sub['code'])} - {esc(sub['name'])}", P("t2", fontName="Times-Roman", fontSize=10.5, leading=14, textColor=H("#444444"), alignment=TA_CENTER)),
           Spacer(1, 5*mm)]
    lab = P("l", fontName="Times-Bold", fontSize=10.5, leading=14, textColor=INK)
    val = P("v", fontName="Times-Roman", fontSize=10.5, leading=14, textColor=INK)
    def row(a, b):
        return [Paragraph(a, lab), Paragraph(":", val), Paragraph(esc(str(b or "-")), val)]
    kv = Table([row("Name of the Student", student["name"]), row("PRN", student["prn"]),
                row("Course", student["course"] or "B.Pharm"), row("Year", student["year"])],
               colWidths=[42*mm, 6*mm, W - 48*mm])
    kv.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5), ("LEFTPADDING", (0, 0), (0, -1), 0)]))
    hd = P("hd", fontName="Times-Bold", fontSize=10, leading=13, textColor=INK, alignment=TA_CENTER)
    cv = P("cv", fontName="Times-Roman", fontSize=11, leading=14, textColor=INK, alignment=TA_CENTER)
    summ = Table([[Paragraph(h, hd) for h in ("Total Lectures", "Lectures Attended", "Lectures Missed", "Attendance %")],
                  [Paragraph(str(sub["total"]), cv), Paragraph(str(sub["present"]), cv), Paragraph(str(sub["absent"]), cv), Paragraph(f"<b>{pct:.2f}%</b>", cv)]],
                 colWidths=[W/4]*4)
    summ.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), H("#e6e6e6")), ("GRID", (0, 0), (-1, -1), 0.7, INK),
                              ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    remark = ("The student satisfies the minimum attendance requirement of %g%%." % thr) if ok else \
             ("The student does NOT satisfy the minimum attendance requirement of %g%%." % thr)
    t = history_table(sub, W, head_bg=H("#e6e6e6"), head_fg=INK, zebra=("#ffffff", "#f7f7f7"), line="#9b9b9b", font="Times-Roman", status="text")
    out += [kv, Spacer(1, 4*mm), summ, Spacer(1, 3*mm),
            Paragraph(f"<i>Remark: {remark}</i>", P("rm", fontName="Times-Roman", fontSize=10, leading=13, textColor=INK)),
            Spacer(1, 6*mm), Paragraph("<b>Details of Lectures</b>", P("h", fontName="Times-Bold", fontSize=11.5, leading=14, textColor=INK, spaceAfter=3)),
            t, Spacer(1, 6*mm),
            Paragraph("<i>Certified that the above attendance record is correct as per the college register.</i>", P("c", fontName="Times-Roman", fontSize=9.5, leading=12, textColor=H("#444444"))),
            Spacer(1, 14*mm), KeepTogether([_sign_row(W, ("Subject Teacher", "Head of Department", "Principal"))])]
    return out


# ==================================================== 7. MODERN SKY
def design_sky(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    BLUE, INK = H("#2563eb"), H("#0f172a")
    pct = float(sub["pct"])
    label, tone, tone_bg = verdict(pct, thr)
    top = Table([[LogoMark(11*mm, college, ctx["logo"], colors.white, BLUE),
                  Paragraph(f"<b>{esc(college)}</b>", P("c", fontSize=13, leading=16, textColor=INK)),
                  pill("ATTENDANCE REPORT", BLUE, 42*mm)]], colWidths=[15*mm, W - 57*mm, 42*mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (-1, 0), (-1, 0), 0)]))
    left = [Paragraph(esc(str(student["name"] or "-")), P("nm", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.white)),
            Paragraph(esc(f"PRN {student['prn'] or '-'}  |  {student['course'] or 'B.Pharm'}  |  {student['year'] or '-'}"), P("mt", fontSize=9, leading=13, textColor=H("#bfdbfe"))),
            Spacer(1, 1.5*mm),
            Paragraph(esc(f"{sub['code']} - {sub['name']}  (Semester {sub['semester']})"), P("sj", fontName="Helvetica-Bold", fontSize=9, leading=12, textColor=colors.white))]
    rt = [Paragraph(f"<b>{pct:.1f}%</b>", P("pc", fontName="Helvetica-Bold", fontSize=30, leading=34, textColor=colors.white, alignment=TA_CENTER)),
          Paragraph(label, P("lb", fontName="Helvetica-Bold", fontSize=7.5, leading=10, textColor=H("#fde68a") if tone != GREEN else H("#bbf7d0"), alignment=TA_CENTER))]
    hero = Table([[left, rt]], colWidths=[W - 54*mm, 54*mm])
    hero.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), BLUE), ("ROUNDEDCORNERS", [10, 10, 10, 10]),
                              ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, 0), 7*mm),
                              ("TOPPADDING", (0, 0), (-1, -1), 6*mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 6*mm)]))
    cw = (W - 8*mm) / 3
    def card(n, cap, col, bg):
        t = Table([[Paragraph(f'<font color="{hx(col)}"><b>{n}</b></font>', P("cn", fontName="Helvetica-Bold", fontSize=20, leading=24, alignment=TA_CENTER))],
                   [Paragraph(cap, P("cc", fontSize=7.5, leading=10, textColor=H("#475569"), alignment=TA_CENTER))]], colWidths=[cw])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), H(bg)), ("ROUNDEDCORNERS", [8, 8, 8, 8]),
                               ("TOPPADDING", (0, 0), (-1, 0), 8), ("BOTTOMPADDING", (0, -1), (-1, -1), 8)]))
        return t
    cards = Table([[card(sub["total"], "TOTAL LECTURES", H("#1d4ed8"), "#eff6ff"), card(sub["present"], "PRESENT", GREEN, "#f0fdf4"),
                    card(sub["absent"], "ABSENT", RED, "#fef2f2")]], colWidths=[cw + 4*mm, cw + 4*mm, cw])
    cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (1, 0), 4*mm), ("RIGHTPADDING", (2, 0), (2, 0), 0),
                               ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    sect = P("se", fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=BLUE, spaceBefore=3, spaceAfter=3)
    out = [top, Spacer(1, 5*mm), hero, Spacer(1, 4*mm), cards, Spacer(1, 5*mm),
           Paragraph("PRESENT vs ABSENT", sect), SplitBar(sub["present"], sub["absent"], W), Spacer(1, 4*mm)]
    months = month_summary(sub["history"])
    if len(months) > 1:
        out += [Paragraph("MONTH-WISE", sect), MonthBars(months, W, 34*mm, thr, BLUE), Spacer(1, 3*mm)]
    out += [Paragraph("LECTURE-WISE HISTORY", sect),
            history_table(sub, W, head_bg=H("#dbeafe"), head_fg=H("#1e40af"), zebra=("#ffffff", "#f8fafc"), line="#e2e8f0")]
    return out


# ==================================================== 8. MODERN DARK
def design_dark(student, sub, ctx):
    W, thr, college = ctx["W"], ctx["thr"], ctx["college"]
    SL, OR = H("#0f172a"), H("#f97316")
    pct = float(sub["pct"])
    label, tone, tone_bg = verdict(pct, thr)
    hdr = Table([[LogoMark(14*mm, college, ctx["logo"], colors.white, OR),
                  [Paragraph(esc(college), P("c", fontName="Helvetica-Bold", fontSize=14, leading=17, textColor=colors.white)),
                   Paragraph("ATTENDANCE REPORT  |  SEMESTER %s" % sub["semester"], P("a", fontSize=7.5, leading=10, textColor=OR))],
                  [Paragraph(f"<b>{pct:.1f}%</b>", P("p", fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=colors.white, alignment=TA_CENTER)),
                   Paragraph("ATTENDANCE", P("pp", fontSize=7, leading=9, textColor=colors.white, alignment=TA_CENTER))]]],
                colWidths=[22*mm, W - 22*mm - 42*mm, 42*mm])
    hdr.setStyle(TableStyle([("BACKGROUND", (0, 0), (1, 0), SL), ("BACKGROUND", (2, 0), (2, 0), OR), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                             ("LEFTPADDING", (0, 0), (0, 0), 5*mm), ("TOPPADDING", (0, 0), (-1, -1), 5*mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 5*mm)]))
    nm = P("nm", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=SL)
    mt = P("mt", fontSize=9, leading=13, textColor=H("#475569"))
    who = Table([[[Paragraph(esc(str(student["name"] or "-")), nm),
                   Paragraph(esc(f"PRN {student['prn'] or '-'}  |  {student['course'] or 'B.Pharm'}  |  {student['year'] or '-'}"), mt),
                   Paragraph(esc(f"{sub['code']} - {sub['name']}"), P("sj", fontName="Helvetica-Bold", fontSize=9.5, leading=13, textColor=H("#c2410c")))],
                  pill(label, tone, 38*mm)]], colWidths=[W - 40*mm, 40*mm])
    who.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    def stat(n, c):
        return [Paragraph(f"<b>{n}</b>", P("sn", fontName="Helvetica-Bold", fontSize=19, leading=23, textColor=colors.white, alignment=TA_CENTER)),
                Paragraph(c, P("sc", fontSize=7, leading=9, textColor=OR, alignment=TA_CENTER))]
    cells = Table([[stat(sub["total"], "TOTAL LECTURES"), stat(sub["present"], "PRESENT"), stat(sub["absent"], "ABSENT"),
                    stat(f"{max(0, thr - pct):.1f}%" if pct < thr else "OK", "SHORT OF %g%%" % thr if pct < thr else "REQUIREMENT MET")]],
                  colWidths=[W/4]*4)
    cells.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), SL), ("INNERGRID", (0, 0), (-1, -1), 3, colors.white),
                               ("BOX", (0, 0), (-1, -1), 3, colors.white), ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    statuses = [str(r["status"]) == "Present" for r in sub["history"]]
    best, run = _streaks(sub["history"])
    sect = P("se", fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=H("#c2410c"), spaceBefore=3, spaceAfter=3)
    out = [hdr, Spacer(1, 5*mm), who, Spacer(1, 4*mm), cells, Spacer(1, 5*mm)]
    if statuses:
        out += [Paragraph("LECTURE MAP  (green = present, red = absent)", sect), Heatmap(statuses, W), Spacer(1, 2*mm),
                Paragraph(f"Longest absent streak: <b>{best}</b> lecture(s)   |   Present streak right now: <b>{run}</b> lecture(s)",
                          P("sk", fontSize=8.5, leading=12, textColor=H("#475569"))), Spacer(1, 4*mm)]
    out += [Paragraph("LECTURE-WISE HISTORY", sect),
            history_table(sub, W, head_bg=SL, zebra=("#ffffff", "#f8fafc"), line="#e2e8f0")]
    return out


# =============================================================== frames
def _page_classic(c, d):
    c.setStrokeColor(H("#14213d")); c.setLineWidth(0.5); c.rect(8*mm, 8*mm, A4[0]-16*mm, A4[1]-16*mm)


def _page_dashboard(c, d):
    c.setFillColor(H("#0f766e")); c.rect(0, A4[1]-3*mm, A4[0], 3*mm, fill=1, stroke=0)


def _page_reportcard(c, d):
    c.setStrokeColor(H("#1b2a49")); c.setLineWidth(1.6); c.rect(8*mm, 8*mm, A4[0]-16*mm, A4[1]-16*mm)
    c.setStrokeColor(H("#b8892b")); c.setLineWidth(0.6); c.rect(10.5*mm, 10.5*mm, A4[0]-21*mm, A4[1]-21*mm)


def _page_hero(c, d):
    c.setFillColor(H("#312e81")); c.rect(0, 0, 7*mm, A4[1], fill=1, stroke=0)
    c.setFillColor(H("#f59e0b")); c.rect(7*mm, 0, 1.2*mm, A4[1], fill=1, stroke=0)


def _page_maroon(c, d):
    c.setFillColor(H("#7a1f2b")); c.rect(0, A4[1]-5*mm, A4[0], 5*mm, fill=1, stroke=0)
    c.setFillColor(H("#a67c52")); c.rect(0, A4[1]-6*mm, A4[0], 1*mm, fill=1, stroke=0)


def _page_mono(c, d):
    pass


def _page_sky(c, d):
    c.setFillColor(H("#2563eb")); c.rect(0, A4[1]-3*mm, A4[0], 3*mm, fill=1, stroke=0)
    c.setFillColor(H("#93c5fd")); c.rect(0, A4[1]-4*mm, A4[0]*0.35, 1*mm, fill=1, stroke=0)


def _page_dark(c, d):
    c.setFillColor(H("#f97316")); c.rect(0, A4[1]-3*mm, A4[0], 3*mm, fill=1, stroke=0)


STYLES = {
    # name: (design fn, page bg fn, margins L,R,T,B in mm, footer)
    "classic": (design_classic, _page_classic, (18, 18, 16, 22), make_footer(H("#14213d"), H("#475569"), "Times-Roman", 18*mm, 18*mm)),
    "dashboard": (design_dashboard, _page_dashboard, (14, 14, 12, 20), make_footer(H("#99d5cb"), H("#64748b"))),
    "reportcard": (design_reportcard, _page_reportcard, (20, 20, 18, 28), make_footer(H("#b8892b"), H("#64748b"), "Times-Roman", 20*mm, 20*mm, lift=5*mm)),
    "maroon": (design_maroon, _page_maroon, (18, 18, 14, 22), make_footer(H("#a67c52"), H("#57534e"), "Times-Roman", 18*mm, 18*mm)),
    "mono": (design_mono, _page_mono, (20, 20, 16, 22), make_footer(H("#9b9b9b"), H("#555555"), "Times-Roman", 20*mm, 20*mm)),
    "sky": (design_sky, _page_sky, (14, 14, 12, 20), make_footer(H("#bfdbfe"), H("#64748b"))),
    "dark": (design_dark, _page_dark, (14, 14, 12, 20), make_footer(H("#fdba74"), H("#64748b"))),
    "hero": (design_hero, _page_hero, (20, 14, 14, 20), make_footer(H("#c7d2fe"), H("#64748b"), "Helvetica", 20*mm, 14*mm)),
}


def build_pdf(student, report, style="classic", threshold=75, college="Personal Attendance Management", logo_path=None, tagline=None):
    design, page_bg, (ml, mr, mt, mb), footer = STYLES[style]
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=A4, title="Attendance Report", author=str(college))
    frame = Frame(ml*mm, mb*mm, A4[0] - (ml + mr)*mm, A4[1] - (mt + mb)*mm,
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=page_bg)])
    ctx = {"W": A4[0] - (ml + mr)*mm, "thr": float(threshold), "college": str(college), "logo": logo_path, "tagline": tagline}
    story = []
    for i, sub in enumerate(report):
        if i:
            story.append(PageBreak())
        sub = dict(sub)  # oldest lecture first, like a register
        sub["history"] = sorted(sub.get("history") or [], key=lambda r: (str(r["attendance_date"]), int(r["lecture_no"] or 0)))
        story += design(student, sub, ctx)
    doc.build(story, canvasmaker=numbered_canvas(footer))
    return buf.getvalue()



# ====================================================== class / subject summary
def build_class_report_pdf(rows, *, college, logo_path=None, threshold=75, tagline="Student Attendance Management System",
                           meta=None, student_history=None, history_title="", signature="Teacher Signature|(Subject Incharge)"):
    """Admin / teacher report: one row per student, in the same 'dashboard' design as the student reports.

    rows            : list of dicts  prn, name, course, year, total, present, absent, pct
    meta            : list of (label, value) pairs shown in the information card
    student_history : optional list of lecture rows (attendance_date, lecture_no, lecture_time, session_type, status)
    """
    TEAL, MINT = H("#0f766e"), H("#ecfdf5")
    page = landscape(A4)
    ml = mr = 12 * mm
    W = page[0] - ml - mr
    thr = float(threshold)
    stamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d %b %Y")
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=page, title="Attendance Report", author=str(college))
    frame = Frame(ml, 20 * mm, W, page[1] - 12 * mm - 20 * mm, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    def top_strip(c, d):
        c.setFillColor(TEAL); c.rect(0, page[1] - 3 * mm, page[0], 3 * mm, fill=1, stroke=0)
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=top_strip)])

    big = P("cb", fontName="Helvetica-Bold", fontSize=16, leading=19, textColor=colors.white)
    small = P("cs", fontSize=9, leading=12, textColor=H("#a7f3d0"))
    hdr = Table([[LogoMark(16 * mm, college, logo_path, TEAL, colors.white),
                  [Paragraph(esc(str(college)), big), Paragraph(esc(str(tagline)).upper() + "  |  ATTENDANCE REPORT", small)],
                  Paragraph(f"<b>CLASS REPORT</b><br/>{stamp}", P("cr", fontSize=9, leading=13, textColor=colors.white, alignment=TA_RIGHT))]],
                colWidths=[23 * mm, W - 23 * mm - 50 * mm, 50 * mm])
    hdr.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), TEAL), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                             ("LEFTPADDING", (0, 0), (0, 0), 6 * mm), ("RIGHTPADDING", (-1, 0), (-1, 0), 6 * mm),
                             ("TOPPADDING", (0, 0), (-1, -1), 4.5 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5 * mm),
                             ("LINEBELOW", (0, 0), (-1, -1), 3, H("#2dd4bf"))]))
    story = [hdr, Spacer(1, 5 * mm)]

    # information card (3 columns)
    lab = P("ml", fontName="Helvetica-Bold", fontSize=7, leading=9, textColor=H("#0f766e"))
    val = P("mv", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=H("#0f172a"))
    cells = [[Paragraph(esc(str(a)).upper(), lab), Paragraph(esc(str(b)), val)] for a, b in (meta or [])]
    while len(cells) % 3:
        cells.append("")
    if cells:
        grid = [cells[i:i + 3] for i in range(0, len(cells), 3)]
        card = Table(grid, colWidths=[W / 3.0] * 3)
        card.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), MINT), ("BOX", (0, 0), (-1, -1), 0.6, H("#cfe8e3")),
                                  ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 9),
                                  ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        story += [card, Spacer(1, 4 * mm)]

    # summary tiles
    counted = [r for r in rows if int(r["total"] or 0) > 0]
    avg = (sum(float(r["pct"]) for r in counted) / len(counted)) if counted else 0.0
    meeting = sum(1 for r in counted if float(r["pct"]) >= thr)
    below = len(counted) - meeting

    def tile(num, cap, col):
        return [Paragraph(f'<font size="20" color="{hx(col)}"><b>{num}</b></font>', P("tn", leading=24, alignment=TA_CENTER)),
                Paragraph(cap, P("tc", fontSize=7.5, leading=10, textColor=H("#64748b"), alignment=TA_CENTER))]
    tiles = Table([[tile(len(rows), "STUDENTS", H("#0f172a")), tile(f"{avg:.1f}%", "AVERAGE ATTENDANCE", TEAL),
                    tile(meeting, f"MEETING {thr:g}% LIMIT", GREEN), tile(below, f"BELOW {thr:g}% LIMIT", RED)]],
                  colWidths=[W / 4.0] * 4, rowHeights=[19 * mm])
    tiles.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), MINT), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("INNERGRID", (0, 0), (-1, -1), 2, colors.white), ("BOX", (0, 0), (-1, -1), 2, colors.white)]))
    story += [tiles, Spacer(1, 5 * mm)]

    # student table
    hd = P("th", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.white, alignment=TA_CENTER)
    hl = P("thl", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.white)
    c1 = P("c1", fontSize=8.8, leading=11, textColor=H("#1e2b3a"))
    cc = P("cc", fontSize=8.8, leading=11, textColor=H("#1e2b3a"), alignment=TA_CENTER)
    widths = [10 * mm, 36 * mm, 56 * mm, 22 * mm, 22 * mm, 22 * mm, 58 * mm]
    widths.append(W - sum(widths))
    data = [[Paragraph(t, hl if i in (1, 2) else hd) for i, t in enumerate(["#", "PRN", "Student name", "Total", "Present", "Absent", "Attendance", "Status"])]]
    for i, r in enumerate(rows, start=1):
        total, pct = int(r["total"] or 0), float(r["pct"] or 0)
        if total:
            label, tone, _bg = verdict(pct, thr)
        else:
            label, tone = "NO DATA", H("#64748b")
        bar = Table([[ThinBar(pct, thr, tone, 34 * mm) if total else "",
                      Paragraph(f'<font color="{hx(tone)}"><b>{pct:.1f}%</b></font>', P("pp", fontSize=8.8, leading=11))]],
                    colWidths=[38 * mm, 16 * mm])
        bar.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                 ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        data.append([Paragraph(str(i), cc), Paragraph(esc(str(r["prn"] or "-")), c1), Paragraph(esc(str(r["name"] or "")), c1),
                     Paragraph(str(total), cc), Paragraph(str(r["present"]), cc), Paragraph(str(r["absent"]), cc), bar, pill(label, tone, 40 * mm)])
    if len(data) == 1:
        data.append([Paragraph("", cc), Paragraph("No students found for the selected filters.", c1)] + [Paragraph("", cc)] * 6)
    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), TEAL), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("ROWBACKGROUNDS", (0, 1), (-1, -1), [H("#ffffff"), H("#f0fdfa")]),
                               ("LINEBELOW", (0, 1), (-1, -1), 0.4, H("#cfe8e3")), ("BOX", (0, 0), (-1, -1), 0.6, H("#cfe8e3")),
                               ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5)]))
    story.append(table)

    if student_history is not None:
        sect = P("se", fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=TEAL, spaceBefore=4, spaceAfter=3)
        story += [Spacer(1, 6 * mm), Paragraph(esc(history_title or "LECTURE-WISE HISTORY").upper(), sect)]
        hist = sorted(student_history, key=lambda r: (str(r["attendance_date"]), int(r["lecture_no"] or 0)))
        story.append(history_table({"history": hist}, W, head_bg=TEAL, zebra=("#ffffff", "#f0fdfa"), line="#cfe8e3"))

    top, bottom = (signature.split("|") + [""])[:2]
    sig = Table([["", [Paragraph("______________________________", P("sg", fontSize=9, leading=12, alignment=TA_CENTER)),
                       Paragraph(f"<b>{esc(top)}</b><br/>{esc(bottom)}", P("sg2", fontSize=8.5, leading=11, alignment=TA_CENTER, textColor=H("#334155")))]]],
                colWidths=[W - 90 * mm, 90 * mm])
    sig.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story += [Spacer(1, 10 * mm), KeepTogether(sig)]
    footer = make_footer(H("#99d5cb"), H("#64748b"), page_w=page[0], left_pad=ml, right_pad=mr)
    doc.build(story, canvasmaker=numbered_canvas(footer))
    return buf.getvalue()
