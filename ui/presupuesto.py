import customtkinter as ctk
from tkinter import ttk, messagebox


class VentanaPresupuesto:

    def __init__(
        self,
        ventana_padre,
        excel_service
        ):

        self.ventana_padre = (
            ventana_padre
        )

        self.excel_service = (
            excel_service
        )

        self.productos_presupuesto = []

        self.ventana = ctk.CTkToplevel(
            ventana_padre
        )

        self.ventana.title(
            "Nuevo Presupuesto"
        )

        self.ventana.geometry(
            "1100x700"
        )

        self.crear_interfaz()

    def crear_interfaz(self):

        frame_cliente = ctk.CTkFrame(
            self.ventana
        )

        frame_cliente.pack(
            fill="x",
            padx=20,
            pady=10
        )


        ctk.CTkLabel(
            frame_cliente,
            text="Cliente:"
        ).grid(
            row=0,
            column=0,
            padx=10,
            pady=10
        )


        self.entry_cliente = ctk.CTkEntry(
            frame_cliente,
            width=300
        )

        self.entry_cliente.grid(
            row=0,
            column=1,
            padx=10,
            pady=10
        )


        ctk.CTkLabel(
            frame_cliente,
            text="Teléfono:"
        ).grid(
            row=0,
            column=2,
            padx=10,
            pady=10
        )


    def calcular_total(self):

        total = 0

        for producto in self.productos_presupuesto:

            subtotal = (

                producto["cantidad"]

                * producto["precio"]

            )

            total += subtotal

        return total

        self.label_total = ctk.CTkLabel(

            self.ventana,

            text="TOTAL: $0.00",

            font=(
                "Arial",
                22,
                "bold"
            )

        )

        self.label_total.pack(
            pady=10
        )

    def actualizar_total(self):

        total = self.calcular_total()

        self.label_total.configure(

            text=(
                f"TOTAL: ${total:,.2f}"
            )

        )