from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

pdf_filename = "proforma_ejemplo.pdf"
doc = SimpleDocTemplate(
    pdf_filename,
    pagesize=A4,
    leftMargin=1.5*cm, rightMargin=1.5*cm,
    topMargin=3*cm, bottomMargin=2.5*cm
)

styles = getSampleStyleSheet()
style_normal = ParagraphStyle("Normal", fontName="Helvetica", fontSize=10, leading=13)
style_bold = ParagraphStyle("Bold", fontName="Helvetica-Bold", fontSize=10, leading=13)
style_cabecera = ParagraphStyle("Cabecera", fontName="Helvetica-Bold", fontSize=13, alignment=1, textColor=colors.HexColor("#14477a"))

tabla_data = []

# 1. PROFORMA (cabecera)
tabla_data.append([
    Paragraph("<b>PROFORMA 12345</b>", style_cabecera), '', '', '', '', '', ''
])

# 2. Campos principales (label-respuesta + 5 vacías para merge)
campos = [
    ('CLIENTE', 'FISCALIA DE BOLIVAR'),
    ('RUC', '0260018030001'),
    ('DIRECCIÓN:', 'Av. Candido Rada y S/N'),
    ('FECHA', '2025-08-03'),
    ('TELEFONO', '022345678'),
    ('NECESIDAD', 'Compra de útiles'),
    ('FUNCIONARIO ENCARGADO:', 'Lic. María Pérez'),
    ('CORREO', 'maría.perez@correo.com'),
    ('HORA MAXIMA:', '12:00')
]
for label, valor in campos:
    tabla_data.append([
        Paragraph(label, style_bold),
        Paragraph(valor, style_normal),
        '', '', '', '', ''
    ])

# 3. OBJETO DE COMPRA
tabla_data.append([
    Paragraph('<b>OBJETO DE COMPRA</b>', style_cabecera), '', '', '', '', '', ''
])
tabla_data.append([
    Paragraph("ADQUISICIÓN SUMINISTROS DE OFICINA PARA STOCK BODEGA DE LA FISCALIA DE BOLIVAR", style_normal), '', '', '', '', '', ''
])

# 4. CABECERAS DE ÍTEMS
headers = ['No.', 'CPC', 'UNIDAD', 'ESPECIFICACIONES', 'CANTIDAD', 'P. UNIT', 'P.TOTAL']
tabla_data.append([Paragraph(f"<b>{h}</b>", style_normal) for h in headers])

# 5. FILAS DE ÍTEMS
items = [
    (1, '12345', 'CAJA', 'Lápiz HB', 10, 2.5, 25),
    (2, '54321', 'UNIDAD', 'Borrador', 20, 0.5, 10),
]
for item in items:
    tabla_data.append([Paragraph(str(v), style_normal) for v in item])

# 6. TOTAL
tabla_data.append([
    Paragraph("<b>TOTAL</b>", style_bold), '', '', '', '', '', Paragraph("$35.00", style_bold)
])

# 7. NO GRAVAMOS IVA
tabla_data.append([
    Paragraph('<font color="#DC2626"><b>NO GRAVAMOS IVA - SOMOS REGIMEN RIMPE - NEGOCIO POPULAR</b></font>', style_normal),
    '', '', '', '', '', ''
])

# 8. CAMPOS FINALES
finales = [
    ('PLAZO DE ENTREGA', '7 días'),
    ('VIGENCIA DE LA OFERTA', '15 días'),
    ('FORMA DE PAGO:', 'Transferencia'),
    ('METODOLOGIA DE TRABAJO', 'Entrega a bodega'),
    ('GARANTIA', '1 año'),
    ('ENLACE', 'https://ejemplo.com')
]
for label, valor in finales:
    tabla_data.append([
        Paragraph(f"<b>{label}</b>", style_bold),
        Paragraph(valor, style_normal),
        '', '', '', '', ''
    ])

# --- ColWidths debe sumar <= 18cm
col_widths = [3*cm, 15*cm/6, 15*cm/6, 15*cm/6, 15*cm/6, 15*cm/6, 15*cm/6]
col_widths[1] = 15*cm  # La columna 1 ocupa casi todo el ancho cuando haces SPAN

table = Table(tabla_data, colWidths=col_widths)

span_cmds = []
# Campos principales
for idx in range(1, 10):
    span_cmds.append(('SPAN', (1, idx), (6, idx)))
# OBJETO DE COMPRA y NO GRAVAMOS IVA
span_cmds += [
    ('SPAN', (0,0), (6,0)),       # PROFORMA
    ('SPAN', (0,10), (6,10)),     # OBJETO DE COMPRA
    ('SPAN', (0,11), (6,11)),     # OBJETO DE COMPRA RESPUESTA
    ('SPAN', (0,17), (6,17)),     # NO GRAVAMOS IVA
]
# Campos finales
for idx in range(18, 24):
    span_cmds.append(('SPAN', (1, idx), (6, idx)))

table.setStyle(TableStyle([
    ('GRID', (0,0), (-1,-1), 0.7, colors.black),
    *span_cmds
]))

elements = [table, Spacer(1,1*cm)]
firma = Paragraph(
    '<para align="center"><b>DAYANA LISBETH ZAMBRANO MACIAS</b><br/>REPRESENTANTE LEGAL<br/>RUC: 2350621211001</para>',
    ParagraphStyle("firma", fontSize=10, fontName="Helvetica")
)
elements.append(firma)

doc.build(elements)
print(f"PDF generado: {pdf_filename}")
