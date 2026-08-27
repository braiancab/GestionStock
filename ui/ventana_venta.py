import customtkinter as ctk
from tkinter import ttk, messagebox

from services.venta_service import VentaService


class VentanaVenta:

    def __init__(
        self,
        ventana_padre,
        excel_service,
        callback_venta
    ):

        self.ventana_padre = ventana_padre

        self.excel_service = excel_service

        self.callback_venta = callback_venta

        self.venta_service = VentaService()

        self.productos = (
            self.excel_service
            .obtener_productos_todos_los_rubros()
        )

        self.total_venta = 0
        self.productos_venta = []
        self.producto_seleccionado = None

        # ==========================================
        # CREAR VENTANA
        # ==========================================

        self.ventana = ctk.CTkToplevel(
            ventana_padre
        )

        self.ventana.title(
            "Registrar Venta"
        )

        self.ventana.state("zoomed")

        self.ventana.minsize(
            850,
            600
        )

        self.ventana.transient(
            ventana_padre
        )

        self.ventana.grab_set()

        self.crear_interfaz()


    # ==========================================
    # INTERFAZ
    # ==========================================

    def crear_interfaz(self):

        self.ventana.grid_columnconfigure(
            0,
            weight=1
        )

        self.ventana.grid_rowconfigure(
            2,
            weight=1
        )

        # ==========================================
        # TÍTULO
        # ==========================================

        titulo = ctk.CTkLabel(
            self.ventana,
            text="REGISTRAR VENTA",
            font=ctk.CTkFont(
                size=26,
                weight="bold"
            )
        )

        titulo.grid(
            row=0,
            column=0,
            padx=25,
            pady=(20, 10),
            sticky="w"
        )


        # ==========================================
        # CLIENTE
        # ==========================================

        frame_cliente = ctk.CTkFrame(
            self.ventana
        )

        frame_cliente.grid(
            row=1,
            column=0,
            padx=25,
            pady=10,
            sticky="ew"
        )

        frame_cliente.grid_columnconfigure(
            1,
            weight=1
        )


        label_cliente = ctk.CTkLabel(
            frame_cliente,
            text="Cliente:"
        )

        label_cliente.grid(
            row=0,
            column=0,
            padx=15,
            pady=15
        )


        self.entry_cliente = ctk.CTkEntry(
            frame_cliente,
            placeholder_text="Nombre de la persona"
        )

        self.entry_cliente.grid(
            row=0,
            column=1,
            padx=15,
            pady=15,
            sticky="ew"
        )


        # ==========================================
        # FRAME PRODUCTOS
        # ==========================================

        frame_agregar = ctk.CTkFrame(
            self.ventana
        )

        frame_agregar.grid(
            row=2,
            column=0,
            padx=25,
            pady=10,
            sticky="nsew"
        )

        frame_agregar.grid_columnconfigure(
            0,
            weight=1
        )

        frame_agregar.grid_rowconfigure(
            1,
            weight=1
        )


        # ==========================================
        # CONTROLES
        # ==========================================

        frame_controles = ctk.CTkFrame(
            frame_agregar,
            fg_color="transparent"
        )

        frame_controles.grid(
            row=0,
            column=0,
            padx=10,
            pady=10,
            sticky="ew"
        )


        # ==========================================
        # BUSCADOR DE PRODUCTO
        # ==========================================

        label_producto = ctk.CTkLabel(
            frame_controles,
            text="Buscar producto:"
        )

        label_producto.grid(
            row=0,
            column=0,
            padx=5,
            pady=5
        )


        self.entry_busqueda = ctk.CTkEntry(
            frame_controles,
            width=400,
            placeholder_text="Nombre, descripción o variante..."
        )

        self.entry_busqueda.grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        self.entry_busqueda.bind(
            "<KeyRelease>",
            self.buscar_productos
        )

        self.entry_busqueda.bind(
            "<Down>",
            self.seleccionar_siguiente
        )

        self.entry_busqueda.bind(
            "<Up>",
            self.seleccionar_anterior
        )

        self.entry_busqueda.bind(
            "<Return>",
            self.seleccionar_producto_teclado
        )


        # ==========================================
        # RESULTADOS DE BÚSQUEDA
        # ==========================================

        self.lista_resultados = ttk.Treeview(
            frame_controles,
            columns=(
                "producto",
                "variante",
                "stock",
                "precio"
            ),
            show="headings",
            height=5
        )

        self.lista_resultados.heading(
            "producto",
            text="Producto"
        )

        self.lista_resultados.heading(
            "variante",
            text="Variante"
        )

        self.lista_resultados.heading(
            "stock",
            text="Stock"
        )

        self.lista_resultados.heading(
            "precio",
            text="Precio"
        )


        self.lista_resultados.column(
            "producto",
            width=250
        )

        self.lista_resultados.column(
            "variante",
            width=250
        )

        self.lista_resultados.column(
            "stock",
            width=80,
            anchor="center"
        )

        self.lista_resultados.column(
            "precio",
            width=120,
            anchor="e"
        )


        self.lista_resultados.grid(
            row=1,
            column=0,
            columnspan=4,
            padx=5,
            pady=(0, 10),
            sticky="ew"
        )

        self.lista_resultados.bind(
            "<<TreeviewSelect>>",
            self.seleccionar_producto
        )

        self.lista_resultados.bind(
            "<Down>",
            self.seleccionar_siguiente
        )

        self.lista_resultados.bind(
            "<Up>",
            self.seleccionar_anterior
        )

        self.lista_resultados.bind(
            "<Return>",
            self.seleccionar_producto_teclado
        )

        self.lista_resultados.bind(
            "<Double-1>",
            self.seleccionar_producto
        )

        # ==========================================
        # STOCK
        # ==========================================

        self.label_stock = ctk.CTkLabel(
            frame_controles,
            text="Stock: -"
        )

        self.label_stock.grid(
            row=2,
            column=0,
            padx=5,
            pady=5,
            sticky="w"
        )


        # ==========================================
        # CANTIDAD
        # ==========================================

        label_cantidad = ctk.CTkLabel(
            frame_controles,
            text="Cantidad:"
        )

        label_cantidad.grid(
            row=3,
            column=0,
           # padx=5,
          #  pady=10
        )


        self.entry_cantidad = ctk.CTkEntry(
            frame_controles,
            width=80,
            placeholder_text="1"
        )


        self.entry_cantidad.grid(
            row=3,
            column=1,
            padx=5,
            pady=10,
            sticky="w"
        )

        self.entry_cantidad.bind(
            "<Return>",
            self.agregar_producto_teclado
        )

        # ==========================================
        # PRECIO
        # ==========================================

        self.label_precio = ctk.CTkLabel(
            frame_controles,
            text="Precio: $0.00"
        )

        self.label_precio.grid(
            row=3,
            column=2,
            columnspan=2,
           # padx=15,
           # pady=5,
           # sticky="w"
        )

        # ==========================================
        # BOTÓN AGREGAR
        # ==========================================

        boton_agregar = ctk.CTkButton(
            frame_controles,
            text="Agregar",
            command=self.agregar_producto
        )

        boton_agregar.grid(
            row=3,
            column=4,
            padx=10,
            pady=10
        )


        # ==========================================
        # TABLA
        # ==========================================

        frame_tabla = ctk.CTkFrame(
            frame_agregar
        )

        frame_tabla.grid(
            row=1,
            column=0,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        frame_tabla.grid_rowconfigure(
            0,
            weight=1
        )

        frame_tabla.grid_columnconfigure(
            0,
            weight=1
        )


        columnas = (
            "producto",
            "variante",
            "cantidad",
            "precio",
            "subtotal"
        )


        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=columnas,
            show="headings"
        )


        self.tabla.heading(
            "producto",
            text="Producto"
        )

        self.tabla.heading(
            "variante",
            text="Variante"
        )

        self.tabla.heading(
            "cantidad",
            text="Cantidad"
        )

        self.tabla.heading(
            "precio",
            text="Precio"
        )

        self.tabla.heading(
            "subtotal",
            text="Subtotal"
        )


        self.tabla.column(
            "producto",
            width=250
        )

        self.tabla.column(
            "variante",
            width=150
        )

        self.tabla.column(
            "cantidad",
            width=100,
            anchor="center"
        )

        self.tabla.column(
            "precio",
            width=130,
            anchor="e"
        )

        self.tabla.column(
            "subtotal",
            width=130,
            anchor="e"
        )


        self.tabla.grid(
            row=0,
            column=0,
            sticky="nsew"
        )


        scrollbar = ttk.Scrollbar(
            frame_tabla,
            orient="vertical",
            command=self.tabla.yview
        )

        scrollbar.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        self.tabla.configure(
            yscrollcommand=scrollbar.set
        )


        # ==========================================
        # PARTE INFERIOR
        # ==========================================
#####
        frame_inferior = ctk.CTkFrame(
            self.ventana,
            fg_color="transparent"
        )

        frame_inferior.grid(
            row=3,
            column=0,
            padx=25,
            pady=15,
            sticky="ew"
        )


        boton_eliminar = ctk.CTkButton(
            frame_inferior,
            text="Eliminar producto",
            command=self.eliminar_producto
        )

        boton_eliminar.grid(
            row=0,
            column=0,
            padx=5
        )


        self.label_total = ctk.CTkLabel(
            frame_inferior,
            text="TOTAL: $0.00",
            font=ctk.CTkFont(
                size=22,
                weight="bold"
            )
        )

        self.label_total.grid(
            row=0,
            column=1,
            padx=30
        )

        # ==========================================
        # PAGO EN EFECTIVO (OPCIONAL)
        # ==========================================

        frame_pago = ctk.CTkFrame(
            frame_inferior,
            fg_color="transparent"
        )

        frame_pago.grid(
            row=1,
            column=0,
            columnspan=4,
            pady=(15, 0),
            sticky="w"
        )

        ctk.CTkLabel(
            frame_pago,
            text="Paga con:"
        ).grid(row=0, column=0, padx=(0, 5))

        self.entry_pago = ctk.CTkEntry(
            frame_pago,
            width=120,
            placeholder_text="Opcional"
        )

        self.entry_pago.grid(
            row=0,
            column=1,
            padx=5
        )

        self.entry_pago.bind(
            "<KeyRelease>",
            self.calcular_vuelto
        )

        self.label_vuelto = ctk.CTkLabel(
            frame_pago,
            text="Vuelto: $0.00",
            font=ctk.CTkFont(size=16, weight="bold")
        )

        self.label_vuelto.grid(
            row=0,
            column=2,
            padx=(20, 0)
        )


        boton_registrar = ctk.CTkButton(
            frame_inferior,
            text="REGISTRAR VENTA",
            command=self.registrar_venta,
            width=180,
            height=40
        )

        boton_registrar.grid(
            row=0,
            column=2,
            padx=5
        )


        boton_cancelar = ctk.CTkButton(
            frame_inferior,
            text="Cancelar",
            command=self.ventana.destroy
        )

        boton_cancelar.grid(
            row=0,
            column=3,
            padx=5
        )


    # ==========================================
    # OBTENER NOMBRES ÚNICOS
    # ==========================================

    def obtener_nombres_productos(self):

        nombres = []

        for producto in self.productos:

            nombre = str(
                producto["nombre"] or ""
            ).strip()

            if nombre and nombre not in nombres:

                nombres.append(nombre)

        return nombres


    # ==========================================
    # OBTENER PRODUCTO SELECCIONADO
    # ==========================================

    def obtener_producto_seleccionado(self):

        return self.producto_seleccionado
    # ==========================================
    # AGREGAR PRODUCTO
    # ==========================================

    def agregar_producto(self):

        producto = self.obtener_producto_seleccionado()

        # ==================================================
        # VALIDAR PRODUCTO
        # ==================================================

        if producto is None:

            messagebox.showwarning(
                "Producto",
                "Seleccione un producto.",
                parent=self.ventana
            )

            return


        # ==================================================
        # CANTIDAD
        # ==================================================

        cantidad_texto = (
            self.entry_cantidad.get().strip()
        )


        try:

            cantidad = int(
                cantidad_texto
            )

        except ValueError:

            messagebox.showerror(
                "Cantidad inválida",
                "La cantidad debe ser un número entero.",
                parent=self.ventana
            )

            return


        if cantidad <= 0:

            messagebox.showerror(
                "Cantidad inválida",
                "La cantidad debe ser mayor a cero.",
                parent=self.ventana
            )

            return


        # ==================================================
        # STOCK
        # ==================================================

        stock = producto["stock"] or 0


        # ==================================================
        # COMPROBAR LO YA AGREGADO
        # ==================================================

        cantidad_ya_agregada = 0


        for item in self.productos_venta:

           if (
                item["id"] == producto["id"]
                and
                item["rubro"] == producto["rubro"]
            ):

                cantidad_ya_agregada += (
                    item["cantidad"]
                )


        if cantidad_ya_agregada + cantidad > stock:

            variante = producto.get("descripcion", "")

            if variante:

                texto_producto = (
                    f"Producto: {producto['nombre']}\n"
                    f"Variante: {variante}\n\n"
                )

            else:

                texto_producto = (
                    f"Producto: {producto['nombre']}\n\n"
                )


            messagebox.showwarning(

                "Stock insuficiente",

                texto_producto +
                f"Stock disponible: {stock}\n"
                f"Ya agregado: {cantidad_ya_agregada}\n"
                f"Solicitado: {cantidad}",

                parent=self.ventana

            )

            return


        # ==================================================
        # AGREGAR PRODUCTO A LA VENTA
        # ==================================================

        producto_venta = {

            "id": producto["id"],

            "nombre": producto["nombre"],

            # Si tiene variante utiliza la descripción.
            # Si no tiene variante queda vacío.
            "descripcion": (
                producto.get("descripcion") or ""
            ),

            "cantidad": cantidad,

            "precio": float(
                producto["precio"] or 0
            ),

            "rubro": producto["rubro"]

        }


        self.productos_venta.append(
            producto_venta
        )


        # ==================================================
        # ACTUALIZAR TABLA
        # ==================================================

        self.actualizar_tabla()


        # ==================================================
        # LIMPIAR CAMPOS
        # ==================================================

        self.entry_cantidad.delete(
            0,
            "end"
        )


        for item in self.lista_resultados.get_children():

            self.lista_resultados.delete(item)


        self.producto_seleccionado = None


        self.label_stock.configure(
            text="Stock: -"
        )


        self.label_precio.configure(
            text="Precio: $0.00"
        )


        self.entry_busqueda.focus()
    # ==========================================
    # ACTUALIZAR TABLA
    # ==========================================

    def actualizar_tabla(self):

        for elemento in self.tabla.get_children():

            self.tabla.delete(elemento)


        total = 0


        for producto in self.productos_venta:

            subtotal = (
                producto["cantidad"]
                * producto["precio"]
            )


            total += subtotal


            self.tabla.insert(

                "",

                "end",

                values=(

                    producto["nombre"],

                    producto["descripcion"],

                    producto["cantidad"],

                    f"${producto['precio']:,.2f}",

                    f"${subtotal:,.2f}"

                )

            )
#####3
        self.total_venta = total

        self.label_total.configure(
            text=f"TOTAL: ${total:,.2f}"
        )

        self.calcular_vuelto()

        self.label_total.configure(

            text=f"TOTAL: ${total:,.2f}"

        )


    # ==========================================
    # ELIMINAR PRODUCTO
    # ==========================================

    def eliminar_producto(self):

        seleccion = (
            self.tabla.selection()
        )


        if not seleccion:

            messagebox.showwarning(
                "Selección",
                "Seleccione un producto de la venta.",
                parent=self.ventana
            )

            return


        indice = self.tabla.index(
            seleccion[0]
        )


        del self.productos_venta[
            indice
        ]


        self.actualizar_tabla()


    # ==========================================
    # REGISTRAR VENTA
    # ==========================================

    def registrar_venta(self):

        cliente = (
            self.entry_cliente.get().strip()
        )

        if not cliente:

            messagebox.showwarning(
                "Cliente",
                "Ingrese el nombre de la persona.",
                parent=self.ventana
            )

            self.entry_cliente.focus()

            return


        if not self.productos_venta:

            messagebox.showwarning(
                "Venta vacía",
                "Debe agregar al menos un producto.",
                parent=self.ventana
            )

            return


        # ==========================================
        # CALCULAR TOTAL
        # ==========================================

        total = 0

        for producto in self.productos_venta:

            total += (
                producto["cantidad"]
                * producto["precio"]
            )


        # ==========================================
        # PAGO EN EFECTIVO OPCIONAL
        # ==========================================

        texto_pago = (
            self.entry_pago.get().strip()
        )


        if texto_pago:

            try:

                paga = float(
                    texto_pago
                )

            except ValueError:

                messagebox.showerror(
                    "Pago",
                    "El importe recibido no es válido.",
                    parent=self.ventana
                )

                return


            if paga < total:

                messagebox.showwarning(
                    "Pago insuficiente",
                    "El dinero recibido es menor al total de la venta.",
                    parent=self.ventana
                )

                return


            vuelto = paga - total

        else:

            paga = None
            vuelto = None


        # ==========================================
        # CONFIRMAR
        # ==========================================

        confirmar = messagebox.askyesno(

            "Confirmar venta",

            f"Cliente: {cliente}\n\n"
            f"Total: ${total:,.2f}\n\n"
            f"¿Desea registrar esta venta?",

            parent=self.ventana

        )


        if not confirmar:

            return


        try:

          

            # ==========================================
            # VALIDAR STOCK
            # ==========================================

            for producto in self.productos_venta:

                # --------------------------------------
                # BUSCAR DIRECTAMENTE POR ID Y RUBRO
                # --------------------------------------

                producto_stock = None

                for p in self.productos:

                    if (
                        p["id"] == producto["id"]
                        and
                        p["rubro"] == producto["rubro"]
                    ):
                        producto_stock = p
                        break

                # --------------------------------------
                # PRODUCTO NO ENCONTRADO
                # --------------------------------------

                if producto_stock is None:

                    raise ValueError(
                        f"No se encontró el producto:\n"
                        f"{producto['nombre']}"
                    )

                # --------------------------------------
                # VALIDAR STOCK
                # --------------------------------------

                stock = producto_stock.get("stock") or 0

                if producto["cantidad"] > stock:

                    descripcion = (
                        producto.get("descripcion") or ""
                    )

                    if descripcion:

                        texto_producto = (
                            f"Producto: {producto['nombre']}\n"
                            f"Variante: {descripcion}\n\n"
                        )

                    else:

                        texto_producto = (
                            f"Producto: {producto['nombre']}\n\n"
                        )

                    raise ValueError(
                        texto_producto +
                        f"Stock disponible: {stock}\n"
                        f"Cantidad solicitada: {producto['cantidad']}"
                    )

            # ==========================================
            # DESCONTAR STOCK
            # ==========================================

            for producto in self.productos_venta:

                self.excel_service.descontar_stock(
                    id_producto=producto["id"],
                    cantidad=producto["cantidad"],
                    rubro=producto["rubro"]
                )


            # ==========================================
            # GUARDAR VENTA
            # ==========================================

            id_venta = (

                self.venta_service.registrar_venta(

                    cliente=cliente,

                    productos=self.productos_venta,

                    total_venta=total

                )

            )


            # ==========================================
            # CONFIRMACIÓN
            # ==========================================

            messagebox.showinfo(

                "Venta registrada",

                f"La venta se registró correctamente.\n\n"

                f"Venta N°: {id_venta}\n"

                f"Cliente: {cliente}\n"

                f"Total: ${total:,.2f}",

                parent=self.ventana

            )


            self.callback_venta()

            self.ventana.destroy()


        except Exception as error:

            messagebox.showerror(

                "Error",

                f"No se pudo registrar la venta.\n\n"
                f"Detalle: {error}",

                parent=self.ventana

            )
    def calcular_vuelto(self, event=None):

        texto = self.entry_pago.get().strip()

        if texto == "":
            self.label_vuelto.configure(
                text="Vuelto: $0.00"
            )
            return

        try:
            paga = float(texto)

        except ValueError:

            self.label_vuelto.configure(
                text="Vuelto: -"
            )
            return

        vuelto = paga - self.total_venta

        if vuelto < 0:

            self.label_vuelto.configure(
                text=f"Faltan: ${abs(vuelto):,.2f}",
                text_color="red"
            )

        else:

            self.label_vuelto.configure(
                text=f"Vuelto: ${vuelto:,.2f}",
                text_color="green"
            )

    def filtrar_productos(self, event=None):

        texto = self.combo_producto.get().strip().lower()

        if not texto:
            self.combo_producto.configure(
                values=self.obtener_nombres_productos()
            )
            return

        resultados = []

        for producto in self.productos:

            nombre = str(
                producto["nombre"] or ""
            ).strip()

            descripcion = str(
                producto["descripcion"] or ""
            ).strip()

            texto_completo = (
                f"{nombre} {descripcion}"
            ).lower()

            if texto in texto_completo:

                if nombre not in resultados:

                    resultados.append(nombre)

        self.combo_producto.configure(
            values=resultados
        )

    # ==========================================
    # BUSCAR PRODUCTOS
    # ==========================================

    def buscar_productos(self, event=None):

        texto = (
            self.entry_busqueda
            .get()
            .strip()
            .lower()
        )

        # Limpiar resultados anteriores
        for item in self.lista_resultados.get_children():
            self.lista_resultados.delete(item)

        self.producto_seleccionado = None

        self.label_stock.configure(
            text="Stock: -"
        )

        self.label_precio.configure(
            text="Precio: $0.00"
        )

        if not texto:
            return

        palabras = texto.split()

        resultados = []

        for producto in self.productos:

            nombre = str(
                producto.get("nombre") or ""
            ).strip()

            descripcion = str(
                producto.get("descripcion") or ""
            ).strip()

            texto_producto = (
                f"{nombre} {descripcion}"
            ).lower()

            # Todas las palabras escritas
            # deben aparecer en producto o descripción
            coincide = all(
                palabra in texto_producto
                for palabra in palabras
            )

            if coincide:

                resultados.append(producto)

        # Mostrar resultados
        for producto in resultados[:50]:

            nombre = str(
                producto.get("nombre") or ""
            ).strip()

            descripcion = str(
                producto.get("descripcion") or ""
            ).strip()

            stock = producto.get("stock") or 0
            precio = producto.get("precio") or 0

            self.lista_resultados.insert(
                "",
                "end",
                values=(
                    nombre,
                    descripcion,
                    stock,
                    f"${float(precio):,.2f}"
                ),
                tags=(str(producto["id"]),)
            )

        # ==========================================
        # SELECCIONAR PRODUCTO
        # ==========================================

    def seleccionar_producto(self, event=None):

            seleccion = self.lista_resultados.selection()

            if not seleccion:
                return

            item = seleccion[0]

            valores = self.lista_resultados.item(
                item,
                "values"
            )

            if not valores:
                return

            nombre = str(
                valores[0] or ""
            ).strip()

            descripcion = str(
                valores[1] or ""
            ).strip()


            producto_encontrado = None


            # ==========================================
            # BUSCAR EL PRODUCTO EXACTO
            # ==========================================

            for producto in self.productos:

                producto_nombre = str(
                    producto.get("nombre") or ""
                ).strip()

                producto_descripcion = str(
                    producto.get("descripcion") or ""
                ).strip()


                if (
                    producto_nombre == nombre
                    and
                    producto_descripcion == descripcion
                ):

                    producto_encontrado = producto

                    break


            if producto_encontrado is None:

                return


            # ==========================================
            # GUARDAR PRODUCTO SELECCIONADO
            # ==========================================

            self.producto_seleccionado = (
                producto_encontrado
            )


            # ==========================================
            # MOSTRAR STOCK
            # ==========================================

            stock = (
                producto_encontrado.get("stock") or 0
            )


            self.label_stock.configure(
                text=f"Stock: {stock}"
            )


            # ==========================================
            # MOSTRAR PRECIO
            # ==========================================

            precio = (
                producto_encontrado.get("precio") or 0
            )


            self.label_precio.configure(
                text=f"Precio: ${float(precio):,.2f}"
            )
    # ==========================================
    # NAVEGACIÓN CON TECLADO
    # ==========================================

    def seleccionar_siguiente(self, event=None):

        items = self.lista_resultados.get_children()

        if not items:
            return "break"

        seleccion = self.lista_resultados.selection()

        # Si todavía no hay nada seleccionado
        if not seleccion:

            item = items[0]

            self.lista_resultados.selection_set(item)
            self.lista_resultados.focus(item)
            self.lista_resultados.see(item)

            self.seleccionar_producto()

            # IMPORTANTE:
            # pasar el foco al Treeview
            self.lista_resultados.focus_set()

            return "break"

        # Ya hay uno seleccionado
        indice = items.index(seleccion[0])

        siguiente = indice + 1

        if siguiente >= len(items):
            siguiente = len(items) - 1

        item = items[siguiente]

        self.lista_resultados.selection_set(item)
        self.lista_resultados.focus(item)
        self.lista_resultados.see(item)

        self.seleccionar_producto()

        return "break"

    def seleccionar_anterior(self, event=None):

        items = self.lista_resultados.get_children()

        if not items:
            return "break"

        seleccion = self.lista_resultados.selection()

        if not seleccion:

            item = items[0]

            self.lista_resultados.selection_set(item)
            self.lista_resultados.focus(item)
            self.lista_resultados.see(item)

            self.seleccionar_producto()

            self.lista_resultados.focus_set()

            return "break"

        indice = items.index(seleccion[0])

        anterior = indice - 1

        if anterior < 0:
            anterior = 0

        item = items[anterior]

        self.lista_resultados.selection_set(item)
        self.lista_resultados.focus(item)
        self.lista_resultados.see(item)

        self.seleccionar_producto()

        return "break"

    def seleccionar_producto_teclado(self, event=None):

        seleccion = self.lista_resultados.selection()

        if not seleccion:

            items = self.lista_resultados.get_children()

            if items:

                self.lista_resultados.selection_set(
                    items[0]
                )

                self.lista_resultados.focus(
                    items[0]
                )

                self.seleccionar_producto()

        self.entry_cantidad.focus()

        return "break"

    def agregar_producto_teclado(self, event=None):

        self.agregar_producto()

        return "break"