from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle
)
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image
)
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase.pdfmetrics import stringWidth

class PDFPresupuestoService:

    def __init__(self):

        self.carpeta_presupuestos = (
            Path.home()
            / "Documents"
            / "GestionStock"
            / "Presupuestos"
        )

        self.carpeta_presupuestos.mkdir(
            parents=True,
            exist_ok=True
        )


    # ==========================================
    # GENERAR PDF
    # ==========================================

    def generar_pdf(

        self,
        id_presupuesto,
        fecha,
        cliente,
        
        productos,
        total,
        logo_path=None

    ):

        ruta_pdf = (
            self.carpeta_presupuestos
            / f"Presupuesto_{id_presupuesto}.pdf"
        )


        documento = SimpleDocTemplate(

            str(ruta_pdf),

            pagesize=A4,

            rightMargin=2 * cm,

            leftMargin=2 * cm,

            topMargin=2 * cm,

            bottomMargin=2 * cm

        )


        elementos = []

        estilos = getSampleStyleSheet()


        # ======================================
        # LOGO
        # ======================================

        if logo_path:

            logo = Image(

                str(logo_path),

                width=4 * cm,

                height=4 * cm

            )

            elementos.append(
                logo
            )


        # ==========================================
        # CABECERA DEL NEGOCIO
        # ==========================================

        estilo_nombre_empresa = ParagraphStyle(
            "NombreEmpresa",
            fontName="Helvetica-Bold",
            fontSize=28,
            leading=32,
            alignment=TA_CENTER,
            spaceAfter=8
        )

        estilo_datos_empresa = ParagraphStyle(
            "DatosEmpresa",
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            alignment=TA_CENTER,
            spaceAfter=2
        )

        nombre_empresa = Paragraph(
            "SAN PABLO",
            estilo_nombre_empresa
        )

        telefono_empresa = Paragraph(
            "Teléfono: 3756-418521",
            estilo_datos_empresa
        )

        direccion_empresa = Paragraph(
            "Dirección: Lisandro de la Torre & Manuel Ledesma,<br/>"
            "Gobernador Ingeniero Valentín Virasoro",
            estilo_datos_empresa
        )

        elementos.append(nombre_empresa)
        elementos.append(telefono_empresa)
        elementos.append(direccion_empresa)

        elementos.append(Spacer(1, 25))


        # ======================================
        # TITULO
        # ======================================

        titulo = Paragraph(

            "<b>PRESUPUESTO</b>",

            estilos["Title"]

        )

        elementos.append(
            titulo
        )


        elementos.append(
            Spacer(
                1,
                0.5 * cm
            )
        )


        # ======================================
        # INFORMACION DEL CLIENTE
        # ======================================

        datos_cliente = [

            [
                "Presupuesto Nº:",
                str(id_presupuesto)
            ],

            [
                "Fecha:",
                fecha
            ],

            [
                "Cliente:",
                cliente
            ]

          

        ]


        tabla_cliente = Table(

            datos_cliente,

            colWidths=[
                4 * cm,
                11 * cm
            ]

        )


        tabla_cliente.setStyle(

            TableStyle([

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (0, -1),
                    "Helvetica-Bold"
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "PADDING",
                    (0, 0),
                    (-1, -1),
                    8
                )

            ])

        )


        elementos.append(
            tabla_cliente
        )


        elementos.append(
            Spacer(
                1,
                0.8 * cm
            )
        )


        # ======================================
        # TABLA DE PRODUCTOS
        # ======================================

        datos_tabla = [

            [
                "Producto",
                "Descripción",
                "Cant.",
                "Precio",
                "Subtotal"
            ]

        ]


        for producto in productos:

            subtotal = (
                producto["cantidad"]
                * producto["precio"]
            )


            datos_tabla.append([

                producto["nombre"],

                producto.get(
                    "descripcion",
                    ""
                ),

                str(
                    producto["cantidad"]
                ),

                f"${producto['precio']:,.2f}",

                f"${subtotal:,.2f}"

            ])


        # IMPORTANTE:
        # ESTA PARTE VA FUERA DEL FOR

        tabla_productos = Table(

            datos_tabla,

            colWidths=[

                4 * cm,

                5 * cm,

                1.5 * cm,

                2.5 * cm,

                2.5 * cm

            ]

        )


        tabla_productos.setStyle(

            TableStyle([

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#2E6E9E")
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold"
                ),

                (
                    "ALIGN",
                    (2, 1),
                    (-1, -1),
                    "RIGHT"
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "PADDING",
                    (0, 0),
                    (-1, -1),
                    7
                )

            ])

        )


        elementos.append(
            tabla_productos
        )


        elementos.append(
            Spacer(
                1,
                0.7 * cm
            )
        )


        # ======================================
        # TOTAL
        # ======================================

        estilo_total = ParagraphStyle(

            "Total",

            parent=estilos["Heading2"],

            alignment=2

        )


        texto_total = Paragraph(

            f"<b>TOTAL: ${total:,.2f}</b>",

            estilo_total

        )


        elementos.append(
            texto_total
        )


        elementos.append(
            Spacer(
                1,
                1 * cm
            )
        )


        # ======================================
        # OBSERVACION
        # ======================================

        observacion = Paragraph(

            "Este presupuesto no representa una factura ni una venta confirmada.",

            estilos["Normal"]

        )


        elementos.append(
            observacion
        )


        # ======================================
        # GENERAR ARCHIVO
        # ======================================

        documento.build(
            elementos
        )


        return ruta_pdf