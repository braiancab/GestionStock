from pathlib import Path
from datetime import datetime
from openpyxl import Workbook, load_workbook


class PresupuestoService:

    def __init__(self):

        self.carpeta_datos = (
            Path.home()
            / "AppData"
            / "Local"
            / "GestionStock"
            / "datos"
        )

        self.carpeta_datos.mkdir(
            parents=True,
            exist_ok=True
        )

        self.ruta_excel = (
            self.carpeta_datos
            / "presupuestos.xlsx"
        )

        self.encabezados = [
            "ID Presupuesto",
            "Fecha",
            "Cliente",
            "ID Producto",
            "Producto",
            "Descripcion",
            "Rubro",
            "Cantidad",
            "Precio Unitario",
            "Subtotal",
            "Total Presupuesto",
            "Estado"
        ]

        self.crear_excel_si_no_existe()


    # ==========================================
    # CREAR EXCEL
    # ==========================================

    def crear_excel_si_no_existe(self):

        if not self.ruta_excel.exists():

            wb = Workbook()

            ws = wb.active

            ws.title = "Presupuestos"

            ws.append(self.encabezados)

            wb.save(self.ruta_excel)

            wb.close()


    # ==========================================
    # OBTENER SIGUIENTE ID
    # ==========================================

    def obtener_siguiente_id(self):

        wb = load_workbook(self.ruta_excel)

        ws = wb["Presupuestos"]

        ids = []

        for fila in ws.iter_rows(
            min_row=2,
            values_only=True
        ):

            if isinstance(fila[0], int):

                ids.append(fila[0])

        wb.close()

        if not ids:
            return 1

        return max(ids) + 1


    # ==========================================
    # REGISTRAR PRESUPUESTO
    # ==========================================

    def registrar_presupuesto(
        self,
        cliente,
        productos,
        total_presupuesto
    ):

        wb = load_workbook(self.ruta_excel)

        ws = wb["Presupuestos"]

        id_presupuesto = (
            self.obtener_siguiente_id()
        )

        fecha = datetime.now().strftime(
            "%d/%m/%Y %H:%M:%S"
        )

        for producto in productos:

            subtotal = (
                producto["cantidad"]
                * producto["precio"]
            )

            ws.append([
                id_presupuesto,
                fecha,
                cliente,
                producto["id"],
                producto["nombre"],
                producto["descripcion"],
                producto["rubro"],
                producto["cantidad"],
                producto["precio"],
                subtotal,
                total_presupuesto,
                "Pendiente"
            ])

        wb.save(self.ruta_excel)

        wb.close()

        return id_presupuesto