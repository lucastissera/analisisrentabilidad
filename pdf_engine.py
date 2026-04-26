"""
pdf_engine.py — Reporte ejecutivo en PDF (ReportLab)
"""

import io
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ALLOC_LABELS = {
    "ventas": "Por % de ventas",
    "empleados": "Por Nº de empleados",
    "m2": "Por M² del local",
    "igualitario": "Igualitario",
    "transacciones": "Por transacciones",
    "manual": "Manual / mixto",
}


def _fmt_money(x):
    if x is None:
        return "—"
    return f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(x):
    if x is None:
        return "—"
    return f"{100 * x:.1f} %".replace(".", ",")


def _story_title(st_title, cfg):
    t = f"ANÁLISIS DE RENTABILIDAD — {cfg.get('empresa', 'Empresa')}"
    return Paragraph(t, st_title)


def build_pdf_report(cfg, monthly, metrics) -> io.BytesIO:
    """Genera un PDF A4 a partir de config, datos mensuales y dict de métricas (misma forma que /api/metrics)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="Rentabilidad",
    )
    styles = getSampleStyleSheet()
    st_title = ParagraphStyle(
        name="rent_title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=16,
        spaceAfter=14,
        textColor=colors.HexColor("#1E3A5F"),
    )
    st_h2 = ParagraphStyle(
        name="rent_h2",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        spaceBefore=12,
        spaceAfter=8,
        textColor=colors.HexColor("#1A2744"),
    )
    st_small = ParagraphStyle(
        name="rent_small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#444444"),
    )

    story = []
    story.append(_story_title(st_title, cfg))
    per = cfg.get("period", "")
    ptype = "Mensual" if cfg.get("periodType", "monthly") == "monthly" else "Anual"
    story.append(Paragraph(f"<b>Período de referencia:</b> {per or '—'}", st_small))
    story.append(Paragraph(f"<b>Tipo de análisis:</b> {ptype}", st_small))
    alloc = ALLOC_LABELS.get(cfg.get("allocMethod", "ventas"), cfg.get("allocMethod", "ventas"))
    story.append(Paragraph(f"<b>Método de prorrateo corporativo:</b> {alloc}", st_small))
    suc = ", ".join(cfg.get("branchNames", [])) or "—"
    story.append(Paragraph(f"<b>Sucursales:</b> {suc}", st_small))
    corp_total = sum((cfg.get("corpCosts") or {}).values())
    story.append(Paragraph(f"<b>Gastos corporativos (configurados):</b> $ {_fmt_money(corp_total)}", st_small))
    story.append(Paragraph(f"<b>Generado:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", st_small))
    story.append(Spacer(1, 0.4 * cm))

    periods = sorted(p for p in monthly.keys() if p)
    if not periods or not cfg.get("branchNames"):
        story.append(Paragraph("No hay datos suficientes para el detalle (períodos o sucursales).", styles["Normal"]))
        doc.build(story)
        buf.seek(0)
        return buf

    # Tabla 1: consolidado por período
    story.append(Paragraph("Resumen consolidado por período", st_h2))
    h1 = ["Período", "Ventas netas", "Margen bruto", "EBITDA directo", "EBITDA neto", "% EBITDA neto / VN"]
    data = [h1]
    for p in periods:
        m = metrics.get(p) or {}
        t = m.get("totals") or {}
        vn = t.get("vn", 0)
        row = [
            p,
            _fmt_money(t.get("vn")),
            _fmt_money(t.get("mb")),
            _fmt_money(t.get("ebitdaDirecto")),
            _fmt_money(t.get("ebitdaNeto")),
            _fmt_pct(t.get("pctEbitdaNeto") if vn else 0),
        ]
        data.append(row)
    if "__cumulative__" in metrics and len(periods) > 1:
        m = metrics["__cumulative__"]
        t = m.get("totals") or {}
        vn = t.get("vn", 0)
        data.append(
            [
                "ACUMULADO",
                _fmt_money(t.get("vn")),
                _fmt_money(t.get("mb")),
                _fmt_money(t.get("ebitdaDirecto")),
                _fmt_money(t.get("ebitdaNeto")),
                _fmt_pct(t.get("pctEbitdaNeto") if vn else 0),
            ]
        )

    t1 = Table(data, colWidths=[2.4 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, 2.3 * cm])
    t1.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A5F")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F5F9")]),
            ]
        )
    )
    story.append(t1)
    story.append(Spacer(1, 0.5 * cm))

    # Tabla 2: por sucursal (vista acumulada o único período)
    if len(periods) > 1:
        key_m = "__cumulative__"
        sub = "Vista acumulada (todos los períodos cargados)"
    else:
        key_m = periods[0]
        sub = f"Período: {periods[0]}"
    m = metrics.get(key_m) or {}
    branches = m.get("branches") or []
    names = list(cfg.get("branchNames", []))
    story.append(Paragraph("Detalle por sucursal", st_h2))
    story.append(Paragraph(sub, st_small))
    h2 = ["Sucursal", "Ventas netas", "% MB", "EBITDA neto", "% EBITDA neto", "ROA"]
    data2 = [h2]
    for i, b in enumerate(branches):
        if i >= len(names):
            break
        vn = b.get("vn", 0)
        data2.append(
            [
                names[i],
                _fmt_money(vn),
                _fmt_pct(b.get("pctMb", 0)),
                _fmt_money(b.get("ebitdaNeto", 0)),
                _fmt_pct(b.get("pctEbitdaNeto", 0)) if vn else "—",
                _fmt_pct(b.get("roa", 0)),
            ]
        )
    t2 = Table(data2, colWidths=[3.2 * cm, 2.4 * cm, 2.2 * cm, 2.4 * cm, 2.4 * cm, 1.8 * cm])
    t2.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A5F")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F5F9")]),
            ]
        )
    )
    story.append(t2)
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("Los importes se expresan en moneda de la empresa (sin ajuste inflacionario).", st_small))

    doc.build(story)
    buf.seek(0)
    return buf
