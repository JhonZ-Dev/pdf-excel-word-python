from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle

# 1. Parámetros de página y márgenes
ancho_hoja, alto_hoja = A4
margen_izq = 1.5 * cm
margen_der = 1.5 * cm
margen_sup = 3 * cm
margen_inf = 2.5 * cm

# 2. Calcula el ancho útil para la tabla
ancho_util = ancho_hoja - margen_izq - margen_der

# Proporciones de columnas (deben sumar 1.0)
proporciones = [0.18, 0.18, 0.10, 0.28, 0.09, 0.085, 0.085]
colWidths = [ancho_util * p for p in proporciones]

# 3. Datos de la tabla
styles = getSampleStyleSheet()
styleN = styles['Normal']
styleBH = styles['Heading4']
# Centrado, tamaño 12, rojo, negrita
style_rojo_centrado = ParagraphStyle(
    'RojoCentrado',
    parent=styles['Normal'],
    alignment=0,         # 0=left, 1=center, 2=right, 4=justify
    fontSize=10,
    textColor=colors.red,
    fontName='Helvetica-Bold'
)
style_rojo_centrados = ParagraphStyle(
    'RojoCentrado',
    parent=styles['Normal'],
    alignment=1,         # 0=left, 1=center, 2=right, 4=justify
    fontSize=10,
    textColor=colors.red,
    fontName='Helvetica-Bold'
)

style_bold = ParagraphStyle(
    'NegritaCentrado',
    parent=styles['Normal'],
    alignment=0,         # 0=left, 1=center, 2=right, 4=justify
    fontSize=10,
    textColor=colors.black,
    fontName='Helvetica-Bold'
)
style_bold_centrado = ParagraphStyle(
    'NegritaCentrado',
    parent=styles['Normal'],
    alignment=1,         # 0=left, 1=center, 2=right, 4=justify
    fontSize=10,
    textColor=colors.black,
    fontName='Helvetica-Bold'
)
data = []

# Fila: Proforma (toda la fila)
data.append([
    Paragraph('<b>PROFORMA N° 001-2024</b>', styleBH), '', '', '', '', '', ''
])

# Datos generales
data.append([
    Paragraph('<b>CLIENTE</b>', style_bold_centrado),
    Paragraph('FISCALIA DE BOLIVAR', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>RUC</b>', style_bold_centrado),
    Paragraph('0260018030001', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>DIRECCIÓN</b>', style_bold_centrado),
    Paragraph('AV. CANDIDO RADA Y SALINAS', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>FECHA</b>', style_bold_centrado),
    Paragraph('06/08/2025', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>TELÉFONO</b>', style_bold_centrado),
    Paragraph('0981234567', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>NECESIDAD</b>', style_bold_centrado),
    Paragraph('SUMINISTROS DE OFICINA', style_rojo_centrado), '', '', '', '', ''
])
data.append([
    Paragraph('<b>FUNCIONARIO ENCARGADO</b>', style_bold_centrado),
    Paragraph('JUAN PÉREZ', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>CORREO</b>', style_bold_centrado),
    Paragraph('juan.perez@fiscalia.gob.ec', style_bold), '', '', '', '', ''
])
data.append([
    Paragraph('<b>HORA MÁXIMA</b>', style_bold_centrado),
    Paragraph('17:00', style_bold), '', '', '', '', ''
])

# Fila: OBJETO DE COMPRA
data.append([
    Paragraph('<b>OBJETO DE COMPRA</b>', style_bold_centrado), '', '', '', '', '', ''
])
# Fila: Respuesta objeto de compra
data.append([
    Paragraph('ADQUISICION SUMINISTROS DE OFICINA PARA STOCK BODEGA DE LA FISCALIA DE BOLIVAR', styleN), '', '', '', '', '', ''
])

# Fila: CABECERA DE ITEMS
data.append([
    Paragraph('<b>No.</b>', styleN),
    Paragraph('<b>CPC</b>', styleN),
    Paragraph('<b>UNIDAD</b>', styleN),
    Paragraph('<b>ESPECIFICACIONES</b>', styleN),
    Paragraph('<b>CANTIDAD</b>', styleN),
    Paragraph('<b>P. UNIT</b>', styleN),
    Paragraph('<b>P. TOTAL</b>', styleN),
])

# Ítems de la tabla
items = [
    ['1', '1234', 'UND', 'Lápiz negro', '50', '0.40', '20.00'],
    ['2', '5678', 'UND', 'Borrador', '30', '0.60', '18.00'],
    ['3', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['4', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['5', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['6', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['7', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],
    ['8', '9123', 'UND', 'Tijera', '10', '1.50', '15.00'],     

]
for item in items:
    # Asegura que siempre sean 7 columnas
    while len(item) < 7:
        item.append('---')
    fila = [str(x) if x not in (None, '') else '---' for x in item]
    data.append(fila)

# Fila: TOTAL (ocupa columnas 0-5, respuesta en 6)
data.append([
    Paragraph('<b>TOTAL</b>', style_bold_centrado), '', '', '', '', '', Paragraph('<b>53.00</b>', styleN)
])
index_total = len(data) - 1  # posición de la fila TOTAL

# Fila: NO GRAVAMOS IVA
data.append([
    Paragraph('NO GRAVAMOS IVA', style_rojo_centrados), '', '', '', '', '', ''
])
index_no_gravamos_iva = len(data) - 1  # posición de la fila NO GRAVAMOS IVA

# Nuevos campos después de NO GRAVAMOS IVA
nuevos_campos = [
    ('PLAZO DE ENTREGA', '5 días hábiles'),
    ('VIGENCIA DE LA OFERTA', '15 días'),
    ('FORMA DE PAGO', 'Transferencia bancaria'),
    ('METODOLOGIA DE TRABAJO', 'Entrega en bodega del cliente'),
    ('GARANTIA', '1 año contra defectos de fábrica'),
    ('ENLACE', 'www.fiscalia.gob.ec/proformas'),
]
index_primero_nuevo = len(data)
for etiqueta, valor in nuevos_campos:
    data.append([
        Paragraph(f'<b>{etiqueta}</b>', styleN),
        Paragraph(valor, styleN), '', '', '', '', ''
    ])

# Agrega SPANs para las filas de nuevos campos
span_cmds = [
    ('SPAN', (0,0), (6,0)),    # PROFORMA
    ('SPAN', (1,1), (6,1)),    # CLIENTE
    ('SPAN', (1,2), (6,2)),    # RUC
    ('SPAN', (1,3), (6,3)),    # DIRECCIÓN
    ('SPAN', (1,4), (6,4)),    # FECHA
    ('SPAN', (1,5), (6,5)),    # TELÉFONO
    ('SPAN', (1,6), (6,6)),    # NECESIDAD
    ('SPAN', (1,7), (6,7)),    # FUNCIONARIO ENCARGADO
    ('SPAN', (1,8), (6,8)),    # CORREO
    ('SPAN', (1,9), (6,9)),    # HORA MÁXIMA
    ('SPAN', (0,10), (6,10)),  # OBJETO DE COMPRA
    ('SPAN', (0,11), (6,11)),  # Respuesta OBJETO DE COMPRA
    ('SPAN', (0,index_total), (5,index_total)),  # TOTAL ocupa columnas 0-5
    ('SPAN', (0,index_no_gravamos_iva), (6,index_no_gravamos_iva)),  # NO GRAVAMOS IVA toda la fila
]

# SPAN para los nuevos campos
for i in range(index_primero_nuevo, index_primero_nuevo + len(nuevos_campos)):
    span_cmds.append(('SPAN', (1, i), (6, i)))  # Respuesta ocupa columnas 1-6

# 5. Crear la tabla ajustada a los márgenes
for idx, row in enumerate(data):
    if hasattr(row[0], 'getPlainText'):
        print(idx, row[0].getPlainText())
    else:
        print(idx, row[0])
print("index_total:", index_total)
tabla = Table(
    data,
    colWidths=colWidths,
    repeatRows=1
)
tabla.setStyle(TableStyle([
    ('GRID', (0,0), (-1,-1), 0.4, colors.black),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('BACKGROUND', (0,0), (6,0), colors.HexColor('#DAE9F7')),
    ('BACKGROUND', (0,12), (6,12), colors.HexColor('#DAE9F7')),  # Cabecera de ítems
     # Fondo azul claro para columna 0 de CLIENTE a HORA MÁXIMA (filas 1 a 9)
    ('BACKGROUND', (0,1), (0,9), colors.HexColor('#DAE9F7')),

    # Fondo azul claro para columna 0 de los nuevos campos (ajusta el rango si cambias)
    ('BACKGROUND', (0, index_primero_nuevo), (0, index_primero_nuevo + len(nuevos_campos) - 1), colors.HexColor('#DAE9F7')),
    ('BACKGROUND', (0, index_total), (6, index_total), colors.HexColor('#DAE9F7')),
    ('BACKGROUND', (0,10), (6,10), colors.HexColor('#DAE9F7')),

    *span_cmds
]))

# 6. Clase para fondo membretada
class BackgroundDocTemplate(SimpleDocTemplate):
    def __init__(self, *args, fondo_path=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.fondo_path = fondo_path

    def handle_pageBegin(self):
        super().handle_pageBegin()
        self.canv.saveState()
        if self.fondo_path:
            self.canv.drawImage(
                self.fondo_path,
                0, 0,
                width=ancho_hoja,
                height=alto_hoja
            )
        self.canv.restoreState()

# 7. Generar el PDF
nombre_archivo = "proforma_final_completa.pdf"
fondo_imagen = "fondomembretada.jpg"  # Asegúrate de tener este archivo en el directorio

doc = BackgroundDocTemplate(
    nombre_archivo,
    pagesize=A4,
    rightMargin=margen_der,
    leftMargin=margen_izq,
    topMargin=margen_sup,
    bottomMargin=margen_inf,
    fondo_path=fondo_imagen
)

elements = []
elements.append(tabla)
doc.build(elements)

print("¡PDF generado correctamente, con todos los campos y márgenes ajustados!")
    