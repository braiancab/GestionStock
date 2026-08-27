import customtkinter as ctk
from tkinter import messagebox


class FormularioProducto:

  

    def __init__(self,ventana_padre,excel_service,callback_guardado,producto=None,modo_variante=False):

        self.ventana_padre = ventana_padre

        self.excel_service = excel_service

        self.callback_guardado = callback_guardado

        self.producto = producto

        self.modo_variante = modo_variante

        self.modo_edicion = producto is not None

        self.modo_edicion = (
        producto is not None
        and not modo_variante
        )
        # ==========================================
        # CREAR VENTANA
        # ==========================================

        self.ventana = ctk.CTkToplevel(
            ventana_padre
        )

        if modo_variante:

            titulo_ventana = "Agregar Variante"

        elif producto is not None:

            titulo_ventana = "Modificar Producto"

        else:

            titulo_ventana = "Nuevo Producto"


        self.ventana.title(
            titulo_ventana
        )

        self.ventana.geometry(
            "500x600"
        )

        self.ventana.resizable(
            False,
            False
        )


        # Mantener ventana al frente

        self.ventana.transient(
            ventana_padre
        )

        self.ventana.grab_set()


        # Crear interfaz

        self.crear_interfaz()


    def crear_interfaz(self):

        
        # ==========================================
        # CONFIGURACIÓN
        # ==========================================

        self.ventana.grid_columnconfigure(
            0,
            weight=1
        )

    
        # ==========================================
        # TÍTULO
        # ==========================================

        if self.modo_variante:

            texto_titulo = "Agregar Variante"

        elif self.modo_edicion:

            texto_titulo = "Modificar Producto"

        else:

            texto_titulo = "Nuevo Producto"


        titulo = ctk.CTkLabel(
            self.ventana,
            text=texto_titulo,
            font=ctk.CTkFont(
            size=24,
            weight="bold"
            )
        )

        titulo.grid(
            row=0,
            column=0,
            padx=30,
            pady=(25, 20)
        )


        # ==========================================
        # FRAME DEL FORMULARIO
        # ==========================================

        frame_formulario = ctk.CTkFrame(
            self.ventana
        )

        frame_formulario.grid(
            row=1,
            column=0,
            padx=30,
            pady=10,
            sticky="ew"
        )

        frame_formulario.grid_columnconfigure(
            1,
            weight=1
        )


        # ==========================================
        # NOMBRE
        # ==========================================

        label_nombre = ctk.CTkLabel(
            frame_formulario,
            text="Nombre:"
        )

        label_nombre.grid(
            row=0,
            column=0,
            padx=15,
            pady=10,
            sticky="w"
        )


        self.entry_nombre = ctk.CTkEntry(
            frame_formulario,
            placeholder_text="Nombre del producto"
        )

        self.entry_nombre.grid(
            row=0,
            column=1,
            padx=15,
            pady=10,
            sticky="ew"
        )


        # ==========================================
        # DESCRIPCIÓN
        # ==========================================

        label_descripcion = ctk.CTkLabel(
            frame_formulario,
            text="Descripción:"
        )

        label_descripcion.grid(
            row=1,
            column=0,
            padx=15,
            pady=10,
            sticky="w"
        )


        self.entry_descripcion = ctk.CTkEntry(
            frame_formulario,
            placeholder_text="Descripción del producto"
        )

        self.entry_descripcion.grid(
            row=1,
            column=1,
            padx=15,
            pady=10,
            sticky="ew"
        )


        # ==========================================
        # CATEGORÍA
        # ==========================================

        label_categoria = ctk.CTkLabel(
            frame_formulario,
            text="Categoría:"
        )

        label_categoria.grid(
            row=2,
            column=0,
            padx=15,
            pady=10,
            sticky="w"
        )


        self.entry_categoria = ctk.CTkEntry(
            frame_formulario,
            placeholder_text="Ej: Remeras"
        )

        self.entry_categoria.grid(
            row=2,
            column=1,
            padx=15,
            pady=10,
            sticky="ew"
        )


        # ==========================================
        # STOCK
        # ==========================================

        label_stock = ctk.CTkLabel(
            frame_formulario,
            text="Stock:"
        )

        label_stock.grid(
            row=3,
            column=0,
            padx=15,
            pady=10,
            sticky="w"
        )


        self.entry_stock = ctk.CTkEntry(
            frame_formulario,
            placeholder_text="Cantidad disponible"
        )

        self.entry_stock.grid(
            row=3,
            column=1,
            padx=15,
            pady=10,
            sticky="ew"
        )


        # ==========================================
        # PRECIO
        # ==========================================

        label_precio = ctk.CTkLabel(
            frame_formulario,
            text="Precio:"
        )

        label_precio.grid(
            row=4,
            column=0,
            padx=15,
            pady=10,
            sticky="w"
        )


        self.entry_precio = ctk.CTkEntry(
            frame_formulario,
            placeholder_text="Ej: 15000"
        )

        self.entry_precio.grid(
            row=4,
            column=1,
            padx=15,
            pady=10,
            sticky="ew"
        )


       


        # ==========================================
        # FRAME BOTONES
        # ==========================================

        frame_botones = ctk.CTkFrame(
            self.ventana,
            fg_color="transparent"
        )

        #frame_botones.grid(
           # row=2,
           # column=0,
          #  padx=30,
         #   pady=20
        #)

        frame_botones.grid(
            row=2,
            column=0,
            padx=30,
            pady=20
        )
        
        # ==========================================
        # BOTÓN GUARDAR
        # ==========================================

        if self.modo_variante:

            texto_boton = "Guardar variante"

        elif self.modo_edicion:

            texto_boton = "Guardar cambios"

        else:

            texto_boton = "Guardar producto"

        boton_guardar = ctk.CTkButton(
            frame_botones,
            text=texto_boton,
            command=self.guardar_producto
        )

        boton_guardar.grid(
            row=0,
            column=0,
            padx=10
        )




        # ==========================================
        # BOTÓN CANCELAR
        # ==========================================

        boton_cancelar = ctk.CTkButton(
            frame_botones,
            text="Cancelar",
            command=self.cancelar
        )

        boton_cancelar.grid(
            row=0,
            column=1,
            padx=10
        )


        # ==========================================
        # CONFIGURAR ENTER
        # ==========================================

        self.configurar_enter()

         # ==========================================
        # CARGAR DATOS SI ESTAMOS EDITANDO
        # ==========================================

        #if self.modo_edicion:

         #   self.cargar_datos_producto()

        if self.producto is not None:

            self.cargar_datos_producto()
   
    # ==========================================
    # CARGAR DATOS DEL PRODUCTO
    # ==========================================

    def cargar_datos_producto(self):

        self.entry_nombre.insert(
            0,
            str(self.producto["nombre"] or "")
        )

        self.entry_descripcion.insert(
            0,
            str(self.producto["descripcion"] or "")
        )

        self.entry_categoria.insert(
            0,
            str(self.producto["categoria"] or "")
        )

        self.entry_stock.insert(
            0,
            str(self.producto["stock"] or "")
        )

        self.entry_precio.insert(
            0,
            str(self.producto["precio"] or "")
        )

      

    # ==========================================
    # GUARDAR PRODUCTO
    # ==========================================

    def guardar_producto(self):

        nombre = self.entry_nombre.get().strip()

        descripcion = self.entry_descripcion.get().strip()

        categoria = self.entry_categoria.get().strip()

        stock = self.entry_stock.get().strip()

        precio = self.entry_precio.get().strip()

      



        # ==========================================
        # VALIDAR CAMPOS OBLIGATORIOS
        # ==========================================

        if not nombre:

            messagebox.showwarning(
                "Campo obligatorio",
                "Debe ingresar el nombre del producto.",
                parent=self.ventana
            )

            self.entry_nombre.focus()

            return



        if not stock:

            messagebox.showwarning(
                "Campo obligatorio",
                "Debe ingresar el stock.",
                parent=self.ventana
            )

            self.entry_stock.focus()

            return


        if not precio:

            messagebox.showwarning(
                "Campo obligatorio",
                "Debe ingresar el precio.",
                parent=self.ventana
            )

            self.entry_precio.focus()

            return


        # ==========================================
        # CONVERTIR NÚMEROS
        # ==========================================

        try:

            stock = int(stock)

            precio = float(
                precio.replace(",", ".")
            )

            


        except ValueError:

            messagebox.showerror(
                "Datos inválidos",
                "Stock y stock mínimo deben ser números enteros.\n"
                "El precio debe ser un número válido.",
                parent=self.ventana
            )

            return


        # ==========================================
        # VALIDAR VALORES NEGATIVOS
        # ==========================================

        if stock < 0:

            messagebox.showerror(
                "Stock inválido",
                "El stock no puede ser negativo.",
                parent=self.ventana
            )

            return


        if precio < 0:

            messagebox.showerror(
                "Precio inválido",
                "El precio no puede ser negativo.",
                parent=self.ventana
            )

            return


       


        # ==========================================
        # GUARDAR EN EXCEL
        # ==========================================

        try:

            # ==========================================
            # MODO CREAR
            # ==========================================

            if not self.modo_edicion:

                nuevo_id = self.excel_service.agregar_producto(

                    nombre=nombre,

                    descripcion=descripcion,

                    categoria=categoria,

                    stock=stock,

                    precio=precio

                    

                )


                messagebox.showinfo(

                    "Producto guardado",

                    f"El producto se guardó correctamente.\n\n"
                    f"ID asignado: {nuevo_id}",

                    parent=self.ventana

                )


            # ==========================================
            # MODO EDITAR
            # ==========================================

            else:

                id_producto = self.producto["id"]


                self.excel_service.actualizar_producto(

                    id_producto=id_producto,

                    nombre=nombre,

                    descripcion=descripcion,

                    categoria=categoria,

                    stock=stock,

                    precio=precio,

                    

                )


                messagebox.showinfo(

                    "Producto actualizado",

                    "El producto se actualizó correctamente.",

                    parent=self.ventana

                )


        # ==========================================
        # ACTUALIZAR TABLA PRINCIPAL
        # ==========================================

            self.callback_guardado()


        # ==========================================
        # CERRAR FORMULARIO
        # ==========================================

            self.ventana.destroy()


        except Exception as error:

            messagebox.showerror(

                "Error",

                f"No se pudo guardar el producto.\n\n"
                f"Detalle: {error}",

                parent=self.ventana

            )

    # ==========================================
    # CANCELAR
    # ==========================================

    def cancelar(self):

        self.ventana.destroy()


    # ==========================================
    # RECORRER CON ENTER
    # ==========================================

    def configurar_enter(self):

        self.entry_nombre.bind(
            "<Return>",
            lambda event: self.entry_descripcion.focus()
        )

        self.entry_descripcion.bind(
            "<Return>",
            lambda event: self.entry_categoria.focus()
        )

        self.entry_categoria.bind(
            "<Return>",
            lambda event: self.entry_stock.focus()
        )

        self.entry_stock.bind(
            "<Return>",
            lambda event: self.entry_precio.focus()
        )

        self.entry_precio.bind(
            "<Return>",
            lambda event: self.guardar_producto()
        )

        