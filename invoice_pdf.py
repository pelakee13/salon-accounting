"""
invoice_pdf.py — Generate a Persian (RTL) PDF invoice for salon transactions.

Uses reportlab + arabic_reshaper + python-bidi for proper Persian text rendering.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable)
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

FONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts", "Vazirmatn.ttf")

# Register font once
_pdf_font_registered = False
def _register_font():
    global _pdf_font_registered
    if not _pdf_font_registered:
        pdfmetrics.registerFont(TTFont("Vazirmatn", FONT_PATH))
        _pdf_font_registered = True

_register_font()
FONT_NAME = "Vazirmatn"
PRIMARY = colors.HexColor("#81c784")
PINK = colors.HexColor("#e8f5e9")
DARK = colors.HexColor("#1a1a2e")
GREY = colors.HexColor("#5f6368")
LIGHT = colors.HexColor("#f5f5f5")


def _fa(text):
    """Reshape + bidi-fix Persian text for reportlab."""
    if text is None:
        return ""
    return get_display(arabic_reshaper.reshape(str(text)))


def _styles():
    ss = getSampleStyleSheet()
    title = ParagraphStyle("title", fontName=FONT_NAME, fontSize=18, leading=24,
                            alignment=TA_CENTER, textColor=PRIMARY)
    sub = ParagraphStyle("sub", fontName=FONT_NAME, fontSize=10, leading=14,
                         alignment=TA_CENTER, textColor=GREY)
    hdr = ParagraphStyle("hdr", fontName=FONT_NAME, fontSize=10, leading=14,
                         alignment=TA_RIGHT, textColor=colors.white)
    cell = ParagraphStyle("cell", fontName=FONT_NAME, fontSize=9, leading=13,
                          alignment=TA_RIGHT, textColor=DARK)
    cell_c = ParagraphStyle("cell_c", fontName=FONT_NAME, fontSize=9, leading=13,
                            alignment=TA_CENTER, textColor=DARK)
    cell_r = ParagraphStyle("cell_r", fontName=FONT_NAME, fontSize=9, leading=13,
                            alignment=TA_RIGHT, textColor=DARK)
    label = ParagraphStyle("label", fontName=FONT_NAME, fontSize=9, leading=13,
                            alignment=TA_RIGHT, textColor=GREY)
    meta_r = ParagraphStyle("meta_r", fontName=FONT_NAME, fontSize=10, leading=16,
                            alignment=TA_RIGHT, textColor=DARK)
    total = ParagraphStyle("total", fontName=FONT_NAME, fontSize=13, leading=18,
                           alignment=TA_RIGHT, textColor=PRIMARY)
    return dict(title=title, sub=sub, hdr=hdr, cell=cell, cell_c=cell_c, cell_r=cell_r, label=label, meta_r=meta_r, total=total)


def generate_invoice(invoice, filepath):
    """
    invoice: dict with keys:
      salon_name, date (Jalali str), invoice_no, customer, phone,
      items: [ {service, employee, amount, commission}, ... ],
      subtotal, discount, final_amount, payment_method, tip, employee_commissions (str)
    """
    S = _styles()
    doc = SimpleDocTemplate(filepath, pagesize=A4,
                             rightMargin=18*mm, leftMargin=18*mm,
                             topMargin=16*mm, bottomMargin=16*mm,
                             title="فاکتور")
    elems = []

    # Header
    elems.append(Paragraph(_fa(invoice.get("salon_name", "آکادمی هلیا")), S["title"]))
    elems.append(Spacer(1, 3))
    elems.append(Paragraph(_fa("سیستم مدیریت سالن زیبایی"), S["sub"]))
    elems.append(Spacer(1, 8))
    elems.append(HRFlowable(width="100%", thickness=1.2, color=PRIMARY))
    elems.append(Spacer(1, 8))

    # Invoice meta — right-aligned paragraphs (guaranteed RTL for Persian)
    elems.append(Paragraph(_fa(f"شماره فاکتور: {invoice.get('invoice_no', '—')}"), S["meta_r"]))
    elems.append(Paragraph(_fa(f"تاریخ: {invoice.get('date', '—')}"), S["meta_r"]))
    elems.append(Paragraph(_fa(f"مشتری: {invoice.get('customer', '—')}"), S["meta_r"]))
    elems.append(Paragraph(_fa(f"تلفن: {invoice.get('phone', '—')}"), S["meta_r"]))
    elems.append(Spacer(1, 6))

    # Items table
    head = [Paragraph(_fa(h), S["hdr"]) for h in ["ردیف", "خدمت", "کارمند", "مبلغ (تومان)"]]
    data = [head]
    for i, it in enumerate(invoice["items"], 1):
        data.append([
            Paragraph(_fa(str(i)), S["cell_r"]),
            Paragraph(_fa(it["service"]), S["cell"]),
            Paragraph(_fa(it["employee"]), S["cell"]),
            Paragraph(_fa(f"{int(it['amount']):,}"), S["cell_c"]),
        ])
    itab = Table(data, colWidths=[14*mm, 62*mm, 40*mm, 32*mm], repeatRows=1)
    itab.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), PRIMARY),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, PINK]),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#e0e0e0")),
        ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    elems.append(itab)
    elems.append(Spacer(1, 12))

    # Totals
    totals = [
        [Paragraph(_fa("جمع خدمات"), S["label"]), Paragraph(_fa(f"{int(invoice.get('subtotal',0)):,} تومان"), S["cell"])],
    ]
    if invoice.get("discount"):
        totals.append([Paragraph(_fa("تخفیف"), S["label"]), Paragraph(_fa(f"{int(invoice['discount']):,} تومان"), S["cell"])])
    if invoice.get("tip"):
        totals.append([Paragraph(_fa("انعام"), S["label"]), Paragraph(_fa(f"{int(invoice['tip']):,} تومان"), S["cell"])])
    if invoice.get("deposit"):
        totals.append([Paragraph(_fa("بیعانه پرداخت‌شده (قبلی)"), S["label"]), Paragraph(_fa(f"{int(invoice['deposit']):,} تومان"), S["cell"])])
        totals.append([Paragraph(_fa("مانده قابل پرداخت"), S["total"]), Paragraph(_fa(f"{int(invoice.get('payable_amount', invoice.get('final_amount',0))):,} تومان"), S["total"])])
    totals.append([Paragraph(_fa("روش پرداخت"), S["label"]), Paragraph(_fa(invoice.get("payment_method", "—")), S["cell"])])
    totals.append([Paragraph(_fa("مبلغ نهایی"), S["total"]), Paragraph(_fa(f"{int(invoice.get('final_amount',0)):,} تومان"), S["total"])])
    tt = Table(totals, colWidths=[40*mm, 108*mm])
    tt.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#e0e0e0")),
        ("INNERGRID", (0,0), (-1,-1), 0.5, colors.HexColor("#e0e0e0")),
        ("BACKGROUND", (0,0), (-1,-2), LIGHT),
        ("BACKGROUND", (0,-1), (-1,-1), PINK),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    elems.append(tt)
    elems.append(Spacer(1, 14))

    # Footer
    elems.append(HRFlowable(width="100%", thickness=0.6, color=colors.HexColor("#e0e0e0")))
    elems.append(Spacer(1, 6))
    elems.append(Paragraph(_fa("از اعتماد شما سپاسگزاریم 🌸"), S["sub"]))

    doc.build(elems)
    return filepath
