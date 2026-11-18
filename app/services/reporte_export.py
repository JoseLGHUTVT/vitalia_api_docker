"""Exportación de reportes a PDF y Excel para planes Profesional y Empresarial."""
import os
from datetime import datetime
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    HRFlowable,
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

VERDE_VITALIA = "0F4C45"
VERDE_SECUNDARIO = "2DD4BF"
GRIS_BORDE = "CBD5E1"
GRIS_ZEBRA = "F8FAFC"
GRIS_TEXTO = "1E293B"

COLOR_PRIMARIO = colors.HexColor(f"#{VERDE_VITALIA}")
COLOR_SECUNDARIO = colors.HexColor(f"#{VERDE_SECUNDARIO}")
COLOR_BORDE = colors.HexColor(f"#{GRIS_BORDE}")
COLOR_ZEBRA = colors.HexColor(f"#{GRIS_ZEBRA}")
COLOR_TEXTO = colors.HexColor(f"#{GRIS_TEXTO}")

COLOR_CRITICO = colors.HexColor("#DC2626")
COLOR_OPORTUNIDAD = colors.HexColor("#D97706")
COLOR_OPTIMO = colors.HexColor("#16A34A")


class NumberedCanvas(canvas.Canvas):
    """Genera pie de página profesional con conteo exacto de páginas 'Página X de Y'."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[Dict[str, Any]] = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, page_count: int):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))
        width, height = self._pagesize

        # Línea superior de pie
        self.setStrokeColor(COLOR_BORDE)
        self.setLineWidth(0.6)
        self.line(40, 32, width - 40, 32)

        self.drawString(40, 20, "VITALIA · Sistema de Gestión Clínica Inteligente · Confidencial")
        texto_pag = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(width - 40, 20, texto_pag)

        # Encabezado para páginas 2 en adelante
        if self._pageNumber > 1:
            self.line(40, height - 30, width - 40, height - 30)
            self.drawString(40, height - 24, "VITALIA · Reporte de Operación Clínica")
            self.drawRightString(width - 40, height - 24, datetime.now().strftime("%d/%m/%Y"))

        self.restoreState()


def _buscar_logo_disco() -> Optional[str]:
    base_dir = os.path.dirname(__file__)
    rutas_candidatas = [
        os.path.abspath(os.path.join(base_dir, "..", "..", "..", "logo.png")),
        os.path.abspath(os.path.join(base_dir, "..", "..", "assets", "logo.png")),
        os.path.abspath(os.path.join(base_dir, "..", "..", "..", "App", "assets", "images", "logo.png")),
    ]
    for r in rutas_candidatas:
        if os.path.isfile(r):
            return r
    return None


def _doctor(x: dict) -> str:
    if x.get("doctor"):
        return x["doctor"]
    return f"Doctor #{x['doctor_id']}" if x.get("doctor_id") is not None else "Sin doctor"


def _periodo(data: dict) -> str:
    d, h = data.get("desde"), data.get("hasta")
    if not d and not h:
        return "Todo el historial registrado"
    return f"{d or 'Inicio'} al {h or 'Presente'}"


def preparar(tipo: str, data: dict) -> dict:
    if tipo == "ocupacion":
        return {
            "titulo": "Ocupación de salas",
            "subtitulo": _periodo(data),
            "resumen": [
                ("Salas registradas", data["total_salas"]),
                ("Citas en salas", sum(x.get("total_citas", 0) for x in data["detalle"])),
                ("Horas ocupadas totales", round(sum(x.get("horas_ocupadas", 0) for x in data["detalle"]), 1)),
            ] + [(f"Estado «{k}»", v) for k, v in data.get("por_estado", {}).items()],
            "columnas": ["Sala", "Tipo", "Piso", "Estado", "Capacidad", "Citas", "Horas ocupadas"],
            "filas": [
                [x["nombre"], x["tipo"] or "", x["piso"] or "", x["estado"] or "",
                 x["capacidad"] or "", x["total_citas"], x["horas_ocupadas"]]
                for x in data["detalle"]
            ],
            "grafico": ("Horas ocupadas", [x["nombre"] for x in data["detalle"]],
                        [x["horas_ocupadas"] for x in data["detalle"]]),
            "col_grafico": 6,
        }
    if tipo == "citas":
        estados = sorted({k for x in data["detalle"] for k in x.get("por_estado", {})})
        return {
            "titulo": "Citas por doctor",
            "subtitulo": _periodo(data),
            "resumen": [("Total de citas", data["total_citas"]), ("Doctores activos", len(data["detalle"]))],
            "columnas": ["Doctor", "Total"] + [e.replace("_", " ").capitalize() for e in estados],
            "filas": [[_doctor(x), x["total"]] + [x.get("por_estado", {}).get(e, 0) for e in estados]
                      for x in data["detalle"]],
            "grafico": ("Citas", [_doctor(x) for x in data["detalle"]], [x["total"] for x in data["detalle"]]),
            "col_grafico": 1,
        }
    if tipo == "espera":
        return {
            "titulo": "Tiempos de espera en recepción",
            "subtitulo": _periodo(data),
            "resumen": [
                ("Espera promedio (min)", data["promedio_minutos"]),
                ("Pacientes atendidos", data["total_atenciones"]),
            ],
            "columnas": ["Doctor", "Atenciones", "Promedio (min)", "Máximo (min)"],
            "filas": [[_doctor(x), x["atenciones"], x["promedio_minutos"], x["maximo_minutos"]]
                      for x in data["detalle"]],
            "grafico": ("Promedio (min)", [_doctor(x) for x in data["detalle"]],
                        [x["promedio_minutos"] for x in data["detalle"]]),
            "col_grafico": 2,
        }
    if tipo == "inventario":
        return {
            "titulo": "Inventario bajo de medicamentos",
            "subtitulo": f"Medicamentos con existencias iguales o menores a {data.get('umbral', 10)} unidades",
            "resumen": [("Medicamentos en riesgo", data["total"])],
            "columnas": ["Medicamento", "Presentación", "Categoría", "Lotes", "Próx. vencimiento", "Stock"],
            "filas": [[x["nombre"], x["presentacion"] or "", x["categoria"] or "", x["lotes"],
                       x["proximo_vencimiento"] or "", x["stock_total"]] for x in data["detalle"]],
            "grafico": ("Stock", [x["nombre"] for x in data["detalle"]],
                        [x["stock_total"] for x in data["detalle"]]),
            "col_grafico": 5,
        }
    if tipo == "recetas":
        return {
            "titulo": "Recetas médicas emitidas",
            "subtitulo": _periodo(data),
            "resumen": [("Recetas emitidas", data["total"])]
            + [(f"Estado «{k}»", v) for k, v in data.get("por_estado", {}).items()],
            "columnas": ["Doctor", "Recetas emitidas"],
            "filas": [[_doctor(x), x["total"]] for x in data["detalle"]],
            "grafico": ("Recetas", [_doctor(x) for x in data["detalle"]], [x["total"] for x in data["detalle"]]),
            "col_grafico": 1,
        }
    raise ValueError(f"Tipo de reporte desconocido: {tipo}")


# ============================================================
# EXPORTACIÓN PROFESIONAL
# ============================================================

def a_excel(rep: dict, clinica_nombre: str = "Clínica VITALIA") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte Operativo"

    thin_border = Border(
        left=Side(style="thin", color=GRIS_BORDE),
        right=Side(style="thin", color=GRIS_BORDE),
        top=Side(style="thin", color=GRIS_BORDE),
        bottom=Side(style="thin", color=GRIS_BORDE),
    )
    fill_header = PatternFill("solid", fgColor=VERDE_VITALIA)
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10)

    # Encabezado corporativo
    ws["A1"] = f"VITALIA · {clinica_nombre.upper()}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=VERDE_VITALIA)
    ws["A2"] = f"Reporte: {rep['titulo']} | Periodo: {rep['subtitulo']}"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="475569")
    ws["A3"] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A3"].font = Font(name="Calibri", size=9, color="64748B")

    # Resumen de KPIs
    fila = 5
    ws.cell(row=fila, column=1, value="RESUMEN DE INDICADORES").font = Font(name="Calibri", size=10, bold=True, color=VERDE_VITALIA)
    fila += 1
    for k, v in rep["resumen"]:
        c1 = ws.cell(row=fila, column=1, value=k)
        c1.font = font_bold
        c1.fill = PatternFill("solid", fgColor=GRIS_ZEBRA)
        c2 = ws.cell(row=fila, column=2, value=v)
        c2.font = font_normal
        c2.alignment = Alignment(horizontal="right")
        fila += 1

    fila += 1
    cab = fila
    # Cabecera de la tabla
    for c, nombre in enumerate(rep["columnas"], 1):
        cel = ws.cell(row=cab, column=c, value=nombre)
        cel.font = font_header
        cel.fill = fill_header
        cel.alignment = Alignment(horizontal="center", vertical="center")
        cel.border = thin_border
    ws.row_dimensions[cab].height = 24

    # Filas de datos con zebra striping
    for idx, r in enumerate(rep["filas"]):
        fila += 1
        color_fila = GRIS_ZEBRA if idx % 2 == 1 else "FFFFFF"
        fill_fila = PatternFill("solid", fgColor=color_fila)
        for c, v in enumerate(r, 1):
            cel = ws.cell(row=fila, column=c, value=v)
            cel.font = font_normal
            cel.fill = fill_fila
            cel.border = thin_border
            if isinstance(v, (int, float)):
                cel.alignment = Alignment(horizontal="right")
                cel.number_format = "#,##0.0" if isinstance(v, float) else "#,##0"
            else:
                cel.alignment = Alignment(horizontal="left")

    # Filtros automáticos y ajuste de anchos
    if rep["filas"]:
        col_final_letra = get_column_letter(len(rep["columnas"]))
        ws.auto_filter.ref = f"A{cab}:{col_final_letra}{fila}"

    for c in range(1, len(rep["columnas"]) + 1):
        ancho_max = max(len(str(ws.cell(row=r, column=c).value or "")) for r in range(cab, fila + 1))
        ws.column_dimensions[get_column_letter(c)].width = min(max(ancho_max + 4, 12), 40)

    # Gráfica openpyxl nativa
    if rep["filas"]:
        ch = BarChart()
        ch.type = "col"
        ch.style = 10
        ch.title = f"{rep['titulo']} - {rep['grafico'][0]}"
        ch.y_axis.title = rep["grafico"][0]
        col_val = rep["col_grafico"] + 1
        ch.add_data(Reference(ws, min_col=col_val, min_row=cab, max_row=fila), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=cab + 1, max_row=fila))
        ch.legend = None
        ch.height = 10
        ch.width = 20
        ws.add_chart(ch, f"A{fila + 3}")

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def a_pdf(rep: dict, clinica_nombre: str = "Clínica VITALIA") -> bytes:
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.8 * cm,
        title=f"VITALIA - {rep['titulo']}",
    )
    est = getSampleStyleSheet()

    est_titulo = ParagraphStyle(
        "TituloCorp",
        parent=est["Title"],
        fontSize=18,
        leading=22,
        textColor=COLOR_PRIMARIO,
        alignment=0,
    )
    est_sub = ParagraphStyle(
        "SubCorp",
        parent=est["Normal"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#475569"),
    )
    est_meta = ParagraphStyle(
        "MetaCorp",
        parent=est["Normal"],
        fontSize=8.5,
        textColor=colors.HexColor("#64748B"),
        alignment=2,
    )
    est_th = ParagraphStyle(
        "TH",
        parent=est["Normal"],
        fontSize=8.5,
        leading=11,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        alignment=1,
    )
    est_td = ParagraphStyle(
        "TD",
        parent=est["Normal"],
        fontSize=8,
        leading=10,
        textColor=COLOR_TEXTO,
    )

    el = []

    # Encabezado corporativo
    logo_path = _buscar_logo_disco()
    col_izq = []
    if logo_path:
        col_izq.append(Image(logo_path, width=32, height=32))
    col_izq.append(Paragraph(f"<b>VITALIA</b> · {clinica_nombre}", est_titulo))
    col_izq.append(Paragraph(f"<b>Reporte:</b> {rep['titulo']} &nbsp;|&nbsp; <b>Periodo:</b> {rep['subtitulo']}", est_sub))

    col_der = [
        Paragraph(f"<b>Fecha de emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", est_meta),
        Paragraph("<b>Plan:</b> Profesional", est_meta),
        Paragraph("Documento Oficial de Operación", est_meta),
    ]

    header_table = Table([[col_izq, col_der]], colWidths=[520, 240])
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    el.append(header_table)
    el.append(HRFlowable(width="100%", thickness=1.5, color=COLOR_PRIMARIO, spaceAfter=12, spaceBefore=6))

    # Cajas de resumen de KPIs
    resumen_cajas = []
    items_kpis = rep["resumen"]
    fila_kpis = []
    for k, v in items_kpis[:6]:
        box_text = f"<font size='7.5' color='#64748B'>{k.upper()}</font><br/><font size='13' color='#{VERDE_VITALIA}'><b>{v}</b></font>"
        fila_kpis.append(Paragraph(box_text, est["Normal"]))

    if fila_kpis:
        col_w = 760 / len(fila_kpis)
        t_kpis = Table([fila_kpis], colWidths=[col_w] * len(fila_kpis))
        t_kpis.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("BOX", (0, 0), (-1, -1), 1, COLOR_BORDE),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        el.append(t_kpis)
        el.append(Spacer(1, 14))

    # Gráfica corregida
    if rep["filas"]:
        nombre_g, etiquetas_g, valores_g = rep["grafico"]
        etiquetas_g = etiquetas_g[:12]
        valores_float = [float(v) for v in valores_g[:12]]
        todos_cero = all(v == 0 for v in valores_float) or len(valores_float) == 0

        el.append(Paragraph(f"<b>Gráfico: {nombre_g}</b>", est["Heading4"]))
        el.append(Spacer(1, 4))

        if todos_cero:
            aviso_data = [[
                Paragraph("<font color='#64748B'><b>Sin actividad registrada en el periodo (0 eventos acumulados).</b></font>", est["Normal"])
            ]]
            tabla_aviso = Table(aviso_data, colWidths=[760])
            tabla_aviso.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                ("BOX", (0, 0), (-1, -1), 1, COLOR_BORDE),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 18),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 18),
            ]))
            el.append(tabla_aviso)
            el.append(Spacer(1, 14))
        else:
            max_val = max(valores_float)
            d = Drawing(760, 160)
            bc = VerticalBarChart()
            bc.x = 45
            bc.y = 35
            bc.width = 690
            bc.height = 110
            bc.data = [valores_float]
            bc.categoryAxis.categoryNames = [(str(e) or "")[:18] for e in etiquetas_g]
            bc.categoryAxis.labels.angle = 25
            bc.categoryAxis.labels.boxAnchor = "ne"
            bc.categoryAxis.labels.fontSize = 7
            bc.categoryAxis.labels.dy = -4
            bc.valueAxis.valueMin = 0
            val_max_axis = max(max_val, 5)
            bc.valueAxis.valueMax = val_max_axis * 1.15
            bc.valueAxis.valueStep = max(1, int(val_max_axis / 5))
            bc.valueAxis.labelTextFormat = '%d'
            bc.valueAxis.labels.fontSize = 7.5
            bc.bars[0].fillColor = COLOR_PRIMARIO
            d.add(bc)
            el.append(d)
            el.append(Spacer(1, 12))

        # Tabla de datos detallados con zebra striping
        tabla_datos = []
        tabla_datos.append([Paragraph(c, est_th) for c in rep["columnas"]])
        for r in rep["filas"]:
            tabla_datos.append([Paragraph(str(c), est_td) for c in r])

        num_cols = len(rep["columnas"])
        col_w_tabla = 760 / num_cols
        tabla = Table(tabla_datos, colWidths=[col_w_tabla] * num_cols, repeatRows=1)
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_ZEBRA]),
            ("GRID", (0, 0), (-1, -1), 0.4, COLOR_BORDE),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        el.append(tabla)
    else:
        el.append(Paragraph("Sin datos para el periodo seleccionado.", est["Normal"]))

    doc.build(el, canvasmaker=NumberedCanvas)
    return buf.getvalue()


# ============================================================
# EXPORTACIÓN EMPRESARIAL CON IA
# ============================================================

def a_excel_empresarial(
    rep: dict,
    comparativa: List[dict],
    insights_ia: dict,
    clinica_nombre: str = "Clínica VITALIA",
) -> bytes:
    """Genera Excel Multi-hoja corporativo VITALIA para el Plan Empresarial."""
    wb = Workbook()

    thin_border = Border(
        left=Side(style="thin", color=GRIS_BORDE),
        right=Side(style="thin", color=GRIS_BORDE),
        top=Side(style="thin", color=GRIS_BORDE),
        bottom=Side(style="thin", color=GRIS_BORDE),
    )
    fill_header = PatternFill("solid", fgColor=VERDE_VITALIA)
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Calibri", size=14, bold=True, color=VERDE_VITALIA)
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10)

    # ------------------------------------------------------------
    # HOJA 1: Resumen de la Clínica
    # ------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Resumen de la Clínica"

    ws1["A1"] = f"VITALIA · REPORTE GENERAL DE LA CLÍNICA"
    ws1["A1"].font = font_title
    ws1["A2"] = f"Clínica: {clinica_nombre} | Reporte: {rep['titulo']}"
    ws1["A2"].font = Font(name="Calibri", size=11, bold=True, color="334155")
    ws1["A3"] = f"Periodo: {rep['subtitulo']} | Fecha emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws1["A3"].font = Font(name="Calibri", size=9, color="64748B")

    # Nivel de actividad clínica
    score = insights_ia.get("score_eficiencia", 85)
    ws1["A5"] = "NIVEL DE ACTIVIDAD CLÍNICA"
    ws1["A5"].font = Font(name="Calibri", size=10, bold=True, color=VERDE_VITALIA)
    ws1["B5"] = f"{score} / 100"
    ws1["B5"].font = Font(name="Calibri", size=12, bold=True, color="047857" if score >= 80 else "B45309")

    # Resumen general
    ws1["A7"] = "RESUMEN Y OBSERVACIONES DE LA CLÍNICA"
    ws1["A7"].font = Font(name="Calibri", size=11, bold=True, color=VERDE_VITALIA)
    ws1.merge_cells("A8:G9")
    cel_res = ws1["A8"]
    cel_res.value = insights_ia.get("resumen_ejecutivo", "Sin resumen disponible.")
    cel_res.font = font_normal
    cel_res.alignment = Alignment(wrap_text=True, vertical="top")

    # Puntos de atención
    fila_a = 11
    ws1.cell(row=fila_a, column=1, value="PUNTOS DE ATENCIÓN DETECTADOS").font = font_bold
    fila_a += 1
    cols_alertas = ["Prioridad", "Observación", "Detalle", "Cifra"]
    for c, n in enumerate(cols_alertas, 1):
        cel = ws1.cell(row=fila_a, column=c, value=n)
        cel.font = font_header
        cel.fill = fill_header
        cel.border = thin_border
    for a in insights_ia.get("alertas_anomalias", []):
        fila_a += 1
        ws1.cell(row=fila_a, column=1, value=str(a.get("prioridad", "")).upper()).font = font_bold
        ws1.cell(row=fila_a, column=2, value=a.get("titulo", "")).font = font_bold
        ws1.cell(row=fila_a, column=3, value=a.get("descripcion", "")).font = font_normal
        ws1.cell(row=fila_a, column=4, value=a.get("metrica", "")).font = font_normal
        for c in range(1, 5):
            ws1.cell(row=fila_a, column=c).border = thin_border

    # Recomendaciones
    fila_r = fila_a + 2
    ws1.cell(row=fila_r, column=1, value="RECOMENDACIONES PARA LA CLÍNICA").font = font_bold
    fila_r += 1
    cols_recs = ["Área", "Acción Sugerida", "Beneficio Esperado"]
    for c, n in enumerate(cols_recs, 1):
        cel = ws1.cell(row=fila_r, column=c, value=n)
        cel.font = font_header
        cel.fill = fill_header
        cel.border = thin_border
    for r in insights_ia.get("recomendaciones", []):
        fila_r += 1
        ws1.cell(row=fila_r, column=1, value=r.get("area", "")).font = font_bold
        ws1.cell(row=fila_r, column=2, value=f"{r.get('titulo', '')}: {r.get('accion', '')}").font = font_normal
        ws1.cell(row=fila_r, column=3, value=r.get("impacto_esperado", "")).font = font_normal
        for c in range(1, 4):
            ws1.cell(row=fila_r, column=c).border = thin_border

    for col in ("A", "B", "C", "D"):
        ws1.column_dimensions[col].width = 28

    # ------------------------------------------------------------
    # HOJA 2: Detalle del Reporte
    # ------------------------------------------------------------
    ws2 = wb.create_sheet(title="Detalle del Reporte")
    ws2["A1"] = f"DETALLE DE DATOS: {rep['titulo'].upper()}"
    ws2["A1"].font = font_title
    ws2["A2"] = f"Periodo analizado: {rep['subtitulo']}"
    ws2["A2"].font = Font(name="Calibri", size=10, italic=True)

    cab2 = 4
    for c, n in enumerate(rep["columnas"], 1):
        cel = ws2.cell(row=cab2, column=c, value=n)
        cel.font = font_header
        cel.fill = fill_header
        cel.alignment = Alignment(horizontal="center")
        cel.border = thin_border

    fila2 = cab2
    for idx, r in enumerate(rep["filas"]):
        fila2 += 1
        color_fila = GRIS_ZEBRA if idx % 2 == 1 else "FFFFFF"
        for c, v in enumerate(r, 1):
            cel = ws2.cell(row=fila2, column=c, value=v)
            cel.font = font_normal
            cel.fill = PatternFill("solid", fgColor=color_fila)
            cel.border = thin_border
            if isinstance(v, (int, float)):
                cel.alignment = Alignment(horizontal="right")
                cel.number_format = "#,##0.0" if isinstance(v, float) else "#,##0"

    if rep["filas"]:
        col_final_letra = get_column_letter(len(rep["columnas"]))
        ws2.auto_filter.ref = f"A{cab2}:{col_final_letra}{fila2}"

    for c in range(1, len(rep["columnas"]) + 1):
        ancho_max = max(len(str(ws2.cell(row=r, column=c).value or "")) for r in range(cab2, fila2 + 1))
        ws2.column_dimensions[get_column_letter(c)].width = min(max(ancho_max + 4, 12), 35)

    if rep["filas"]:
        ch = BarChart()
        ch.type = "col"
        ch.title = f"{rep['titulo']} - {rep['grafico'][0]}"
        ch.y_axis.title = rep["grafico"][0]
        col_val = rep["col_grafico"] + 1
        ch.add_data(Reference(ws2, min_col=col_val, min_row=cab2, max_row=fila2), titles_from_data=True)
        ch.set_categories(Reference(ws2, min_col=1, min_row=cab2 + 1, max_row=fila2))
        ch.legend = None
        ch.height = 9
        ch.width = 18
        ws2.add_chart(ch, f"A{fila2 + 3}")

    # ------------------------------------------------------------
    # HOJA 3: Comparativa de Periodos
    # ------------------------------------------------------------
    ws3 = wb.create_sheet(title="Comparativa de Periodos")
    ws3["A1"] = "COMPARATIVA CON EL PERIODO ANTERIOR"
    ws3["A1"].font = font_title
    ws3["A2"] = "Comparativa del Periodo Actual contra el Periodo Anterior Equivalente"
    ws3["A2"].font = Font(name="Calibri", size=10, italic=True)

    cab3 = 4
    cols_comp = ["Indicador Operativo", "Unidad", "Periodo Anterior", "Periodo Actual", "Variación (%)", "Tendencia"]
    for c, n in enumerate(cols_comp, 1):
        cel = ws3.cell(row=cab3, column=c, value=n)
        cel.font = font_header
        cel.fill = fill_header
        cel.alignment = Alignment(horizontal="center")
        cel.border = thin_border

    fila3 = cab3
    for idx, comp in enumerate(comparativa):
        fila3 += 1
        color_fila = GRIS_ZEBRA if idx % 2 == 1 else "FFFFFF"
        ws3.cell(row=fila3, column=1, value=comp.get("indicador", "")).font = font_bold
        ws3.cell(row=fila3, column=2, value=comp.get("unidad", "")).font = font_normal
        c_ant = ws3.cell(row=fila3, column=3, value=comp.get("valor_anterior", 0))
        c_ant.font = font_normal
        c_ant.number_format = "#,##0.0"
        c_act = ws3.cell(row=fila3, column=4, value=comp.get("valor_actual", 0))
        c_act.font = font_normal
        c_act.number_format = "#,##0.0"

        v_pct = comp.get("variacion_pct", 0)
        c_pct = ws3.cell(row=fila3, column=5, value=f"{'+' if v_pct > 0 else ''}{v_pct}%")
        color_pct = "16A34A" if comp.get("favorable", True) else "DC2626"
        c_pct.font = Font(name="Calibri", size=10, bold=True, color=color_pct)
        c_pct.alignment = Alignment(horizontal="right")

        ws3.cell(row=fila3, column=6, value="▲ Favorable" if comp.get("favorable") else "▼ A observar").font = font_normal

        for c in range(1, 7):
            ws3.cell(row=fila3, column=c).fill = PatternFill("solid", fgColor=color_fila)
            ws3.cell(row=fila3, column=c).border = thin_border

    for c in range(1, 7):
        ws3.column_dimensions[get_column_letter(c)].width = 22

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def a_pdf_empresarial(
    rep: dict,
    comparativa: List[dict],
    insights_ia: dict,
    clinica_nombre: str = "Clínica VITALIA",
) -> bytes:
    """Genera PDF Ejecutivo Empresarial con análisis de IA Groq y comparativas."""
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.8 * cm,
        title=f"VITALIA Empresarial - {rep['titulo']}",
    )
    est = getSampleStyleSheet()

    est_titulo = ParagraphStyle(
        "TituloEmp",
        parent=est["Title"],
        fontSize=18,
        leading=22,
        textColor=COLOR_PRIMARIO,
        alignment=0,
    )
    est_sub = ParagraphStyle(
        "SubEmp",
        parent=est["Normal"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#475569"),
    )
    est_meta = ParagraphStyle(
        "MetaEmp",
        parent=est["Normal"],
        fontSize=8.5,
        textColor=colors.HexColor("#64748B"),
        alignment=2,
    )
    est_th = ParagraphStyle(
        "THEmp",
        parent=est["Normal"],
        fontSize=8.5,
        leading=11,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        alignment=1,
    )
    est_td = ParagraphStyle(
        "TDEmp",
        parent=est["Normal"],
        fontSize=8,
        leading=10,
        textColor=COLOR_TEXTO,
    )

    el = []

    # Encabezado
    logo_path = _buscar_logo_disco()
    col_izq = []
    if logo_path:
        col_izq.append(Image(logo_path, width=32, height=32))
    col_izq.append(Paragraph(f"<b>VITALIA</b> · {clinica_nombre}", est_titulo))
    col_izq.append(Paragraph(f"<b>Reporte:</b> {rep['titulo']} &nbsp;|&nbsp; <b>Periodo:</b> {rep['subtitulo']}", est_sub))

    score = insights_ia.get("score_eficiencia", 85)
    color_score_hex = "#16A34A" if score >= 80 else "#D97706"
    col_der = [
        Paragraph("<font color='#D97706'><b>★ REPORTE GENERAL · PLAN EMPRESARIAL</b></font>", est_meta),
        Paragraph(f"<b>Nivel de actividad clínica:</b> <font color='{color_score_hex}' size='11'><b>{score} / 100</b></font>", est_meta),
        Paragraph(f"<b>Fecha de emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", est_meta),
    ]

    header_table = Table([[col_izq, col_der]], colWidths=[520, 240])
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    el.append(header_table)
    el.append(HRFlowable(width="100%", thickness=1.5, color=COLOR_PRIMARIO, spaceAfter=10, spaceBefore=4))

    # SECCIÓN DE RESUMEN Y OBSERVACIONES DE LA CLÍNICA
    resumen_ia = insights_ia.get("resumen_ejecutivo", "Diagnóstico no disponible.")
    caja_ia_data = [
        [Paragraph(f"<b>RESUMEN Y OBSERVACIONES DE LA CLÍNICA</b>", ParagraphStyle("H_IA", parent=est["Normal"], fontSize=9, fontName="Helvetica-Bold", textColor=COLOR_PRIMARIO))],
        [Paragraph(resumen_ia, ParagraphStyle("Txt_IA", parent=est["Normal"], fontSize=8.5, leading=12, textColor=COLOR_TEXTO))],
    ]
    t_caja_ia = Table(caja_ia_data, colWidths=[760])
    t_caja_ia.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F0FDFA")),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#99F6E4")),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    el.append(t_caja_ia)
    el.append(Spacer(1, 10))

    # Observaciones y Recomendaciones en 2 columnas
    alertas = insights_ia.get("alertas_anomalias", [])
    recs = insights_ia.get("recomendaciones", [])

    tabla_alertas_recs = []
    txt_alertas = "<b>PUNTOS DE ATENCIÓN DETECTADOS</b><br/>"
    for a in alertas[:3]:
        prio = str(a.get("prioridad", "oportunidad")).upper()
        color_p = "#DC2626" if prio == "CRITICO" else ("#D97706" if prio == "OPORTUNIDAD" else "#16A34A")
        txt_alertas += f"• <font color='{color_p}'><b>[{prio}]</b></font> <b>{a.get('titulo', '')}</b>: {a.get('descripcion', '')} ({a.get('metrica', '')})<br/>"

    txt_recs = "<b>RECOMENDACIONES PARA EL EQUIPO MÉDICO</b><br/>"
    for idx, r in enumerate(recs[:3], 1):
        txt_recs += f"<b>{idx}. {r.get('titulo', '')}</b> ({r.get('area', 'General')}): {r.get('accion', '')} <font color='#0F766E'><i>→ {r.get('impacto_esperado', '')}</i></font><br/>"

    p_al = Paragraph(txt_alertas, ParagraphStyle("Al", parent=est["Normal"], fontSize=8, leading=11))
    p_rc = Paragraph(txt_recs, ParagraphStyle("Rc", parent=est["Normal"], fontSize=8, leading=11))

    t_ar = Table([[p_al, p_rc]], colWidths=[375, 375])
    t_ar.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("BOX", (0, 0), (-1, -1), 0.8, COLOR_BORDE),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    el.append(t_ar)
    el.append(Spacer(1, 10))

    # Tabla de métricas comparativas con periodo anterior
    if comparativa:
        el.append(Paragraph("<b>Métricas Comparativas con el Periodo Anterior Equivalente</b>", est["Heading4"]))
        el.append(Spacer(1, 3))
        comp_rows = [[
            Paragraph("Indicador", est_th),
            Paragraph("Periodo Anterior", est_th),
            Paragraph("Periodo Actual", est_th),
            Paragraph("Variación", est_th),
            Paragraph("Evaluación", est_th),
        ]]
        for c in comparativa:
            v_pct = c.get("variacion_pct", 0)
            signo = "+" if v_pct > 0 else ""
            color_v = "#16A34A" if c.get("favorable") else "#DC2626"
            comp_rows.append([
                Paragraph(f"<b>{c.get('indicador', '')}</b>", est_td),
                Paragraph(f"{c.get('valor_anterior', 0)} {c.get('unidad', '')}", est_td),
                Paragraph(f"<b>{c.get('valor_actual', 0)} {c.get('unidad', '')}</b>", est_td),
                Paragraph(f"<font color='{color_v}'><b>{signo}{v_pct}%</b></font>", est_td),
                Paragraph("Óptimo / Favorable" if c.get("favorable") else "Requiere Atención", est_td),
            ])
        t_comp = Table(comp_rows, colWidths=[200, 140, 140, 120, 160])
        t_comp.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_ZEBRA]),
            ("GRID", (0, 0), (-1, -1), 0.4, COLOR_BORDE),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        el.append(t_comp)
        el.append(Spacer(1, 12))

    # Gráfica corregida
    if rep["filas"]:
        nombre_g, etiquetas_g, valores_g = rep["grafico"]
        etiquetas_g = etiquetas_g[:12]
        valores_float = [float(v) for v in valores_g[:12]]
        todos_cero = all(v == 0 for v in valores_float) or len(valores_float) == 0

        el.append(Paragraph(f"<b>Detalle y Tendencia: {nombre_g}</b>", est["Heading4"]))
        el.append(Spacer(1, 4))

        if todos_cero:
            aviso_data = [[
                Paragraph("<font color='#64748B'><b>Sin actividad registrada en el periodo (0 eventos acumulados).</b></font>", est["Normal"])
            ]]
            tabla_aviso = Table(aviso_data, colWidths=[760])
            tabla_aviso.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                ("BOX", (0, 0), (-1, -1), 1, COLOR_BORDE),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 14),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
            ]))
            el.append(tabla_aviso)
            el.append(Spacer(1, 10))
        else:
            max_val = max(valores_float)
            d = Drawing(760, 145)
            bc = VerticalBarChart()
            bc.x = 45
            bc.y = 35
            bc.width = 690
            bc.height = 100
            bc.data = [valores_float]
            bc.categoryAxis.categoryNames = [(str(e) or "")[:18] for e in etiquetas_g]
            bc.categoryAxis.labels.angle = 25
            bc.categoryAxis.labels.boxAnchor = "ne"
            bc.categoryAxis.labels.fontSize = 7
            bc.categoryAxis.labels.dy = -4
            bc.valueAxis.valueMin = 0
            val_max_axis = max(max_val, 5)
            bc.valueAxis.valueMax = val_max_axis * 1.15
            bc.valueAxis.valueStep = max(1, int(val_max_axis / 5))
            bc.valueAxis.labelTextFormat = '%d'
            bc.valueAxis.labels.fontSize = 7.5
            bc.bars[0].fillColor = COLOR_PRIMARIO
            d.add(bc)
            el.append(d)
            el.append(Spacer(1, 10))

        # Tabla detallada con zebra striping
        tabla_datos = [[Paragraph(c, est_th) for c in rep["columnas"]]]
        for r in rep["filas"]:
            tabla_datos.append([Paragraph(str(c), est_td) for c in r])

        num_cols = len(rep["columnas"])
        col_w_tabla = 760 / num_cols
        tabla = Table(tabla_datos, colWidths=[col_w_tabla] * num_cols, repeatRows=1)
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_ZEBRA]),
            ("GRID", (0, 0), (-1, -1), 0.4, COLOR_BORDE),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        el.append(tabla)

    doc.build(el, canvasmaker=NumberedCanvas)
    return buf.getvalue()
