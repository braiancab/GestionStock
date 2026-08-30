import customtkinter as ctk
import tkinter as tk
from tkinter import ttk,messagebox

from services.excel_service import ExcelService
from ui.formulario_producto import FormularioProducto
from ui.ventana_venta import VentanaVenta
from ui.historial_ventas import HistorialVentas
from ui.ventana_presupuesto import VentanaPresupuesto
from ui.historial_presupuesto import HistorialPresupuestos
from ui.actualizar_productos import VentanaActualizarPrecios
from ui.reportes import VentanaReportes
import sys

class VentanaPrincipal:

    def __init__(self):

        self.rubro_actual = "ferreteria"

        self.excel_service = ExcelService(
            self.rubro_actual
        )


        self.ventana = ctk.CTk()
        self.ventana.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)

        self.ventana.title("Pablo Gestor Stock")


        self.ventana.state("zoomed")      # Maximizada

        self.ventana.minsize(900, 500)

        # Modifica la fuente global de todos los componentes de tipo Menu
        self.ventana.option_add("*Menu.font", ("Arial", 18))

        self.crear_menu()


        # Configuración del diseño
        self.crear_interfaz()

        # Cargar productos del Excel
        self.cargar_productos()


        


    def crear_interfaz(self):

        # ==========================================
        # CONFIGURACIÓN DE LA VENTANA
        # ==========================================

        self.ventana.grid_rowconfigure(1, weight=1)

        self.ventana.grid_columnconfigure(0, weight=1)


        # ==========================================
        # TÍTULO
        # ==========================================

        titulo = ctk.CTkLabel(
            self.ventana,
            text="PABLO STOCK",
            font=ctk.CTkFont(
                size=26,
                weight="bold"
            )
        )

        titulo.grid(
            row=0,
            column=0,
            padx=20,
            pady=(20, 10),
            sticky="w"
        )

        self.label_rubro = ctk.CTkLabel(
            self.ventana,
            text=f"RUBRO: {self.rubro_actual.upper()}",
            font=ctk.CTkFont(
                size=22,
                weight="bold"
            )
        )

        self.label_rubro.place(
            relx=0.5,
            y=45,
            anchor="center"
        )


        # ==========================================
        # FRAME PRINCIPAL
        # ==========================================

        frame_principal = ctk.CTkFrame(
            self.ventana
        )

        frame_principal.grid(
            row=1,
            column=0,
            padx=20,
            pady=10,
            sticky="nsew"
        )

        frame_principal.grid_rowconfigure(
            3,
            weight=1
        )

        frame_principal.grid_columnconfigure(
            0,
            weight=1
        )

        # ==========================================
        # FRAME DE BÚSQUEDA
        # ==========================================

        frame_busqueda = ctk.CTkFrame(
            frame_principal,
            fg_color="transparent"
        )

        frame_busqueda.grid(
            row=0,
            column=0,
            padx=15,
            pady=15,
            sticky="ew"
        )

        frame_busqueda.grid_columnconfigure(
            1,
            weight=1
        )
     

        # Texto Buscar

        label_buscar = ctk.CTkLabel(
            frame_busqueda,
            text="Buscar producto:"
        )

        label_buscar.grid(
            row=0,
            column=0,
            padx=(0, 10)
        )


        # Campo de búsqueda

        self.entry_buscar = ctk.CTkEntry(
            frame_busqueda,
            placeholder_text="Ingrese el nombre del producto..."
        )

        self.entry_buscar.grid(
            row=0,
            column=1,
            padx=10,
            sticky="ew"
        )
        self.entry_buscar.bind(
            "<Return>",
            lambda evento: self.buscar_productos()
        )

        # Botón buscar

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar",
            command=self.buscar_productos
        )

        boton_buscar.grid(
            row=0,
            column=2,
            padx=5
        )


        # Botón mostrar todos

        boton_mostrar_todos = ctk.CTkButton(
            frame_busqueda,
            text="Mostrar todos",
            command=self.cargar_productos
        )

        boton_mostrar_todos.grid(
            row=0,
            column=3,
            padx=5
        )


       

        # ==========================================
        # BOTONES DE RUBRO
        # ==========================================

        frame_rubros = ctk.CTkFrame(
            frame_principal,
            fg_color="transparent"
        )

        frame_rubros.grid(
            row=2,
            column=0,
            padx=15,
            pady=(0, 10)
        )

        # ==========================================
        # BOTONES DE RUBRO
        # ==========================================

        boton_ferreteria = ctk.CTkButton(
            frame_rubros,
            text="Ferretería",
            fg_color="#338637",
            hover_color="#0A500F",
            command=lambda: self.cambiar_rubro(
                "ferreteria"
            )
        )

        boton_ferreteria.grid(
            row=0,
            column=0,
            padx=(0, 10)
        )


        boton_refrigeracion = ctk.CTkButton(
            frame_rubros,
            text="Refrigeración",
            fg_color="#338637",
            hover_color="#0A500F",
            command=lambda: self.cambiar_rubro(
                "refrigeracion"
            )
        )

        boton_refrigeracion.grid(
            row=0,
            column=1,
            padx=10
        )


        boton_electricidad = ctk.CTkButton(
            frame_rubros,
            text="Electricidad",
            fg_color="#338637",
            hover_color="#0A500F",
            command=lambda: self.cambiar_rubro(
                "electricidad"
            )
        )

        boton_electricidad.grid(
            row=0,
            column=2,
            padx=10
        )


        # ==========================================
        # TABLA
        # ==========================================

        frame_tabla = ctk.CTkFrame(
            frame_principal
        )

        frame_tabla.grid(
            row=3,
            column=0,
            padx=15,
            pady=(0, 15),
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


        # ==========================================
        # CONFIGURAR ESTILO DE TABLA
        # ==========================================

        estilo = ttk.Style()

        estilo.configure(
            "Treeview",
            rowheight=30,
            font=("Arial", 10)
        )

        
        estilo.configure(
            "Treeview.Heading",
            font=("Arial", 10, "bold")
        )
       


        # ==========================================
        # CREAR TABLA
        # ==========================================

        columnas = (
            "id",
            "nombre",
            "descripcion",
            "categoria",
            "stock",
            "precio"
        )


        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=columnas,
            show="headings"
        )

        self.tabla.tag_configure(
            "bajo_stock",
            background="#DA7B7B"   # rojo suave
        )
        # Encabezados

        self.tabla.heading(
            "id",
            text="ID"
        )

        self.tabla.heading(
            "nombre",
            text="Producto"
        )

        self.tabla.heading(
            "descripcion",
            text="Descripción"
        )

        self.tabla.heading(
            "categoria",
            text="Categoría"
        )

        self.tabla.heading(
            "stock",
            text="Stock"
        )

        self.tabla.heading(
            "precio",
            text="Precio"
        )

       

        # Ancho de columnas

        self.tabla.column(
            "id",
            width=50,
            anchor="center"
        )

        self.tabla.column(
            "nombre",
            width=180
        )

        self.tabla.column(
            "descripcion",
            width=220
        )

        self.tabla.column(
            "categoria",
            width=120
        )

        self.tabla.column(
            "stock",
            width=80,
            anchor="center"
        )

        self.tabla.column(
            "precio",
            width=100,
            anchor="e"
        )

      


        # ==========================================
        # SCROLLBAR
        # ==========================================

        scrollbar = ttk.Scrollbar(
            frame_tabla,
            orient="vertical",
            command=self.tabla.yview
        )

        self.tabla.configure(
            yscrollcommand=scrollbar.set
        )


        # Colocar tabla

        self.tabla.grid(
            row=0,
            column=0,
            sticky="nsew"
        )


        # Colocar scrollbar

        scrollbar.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        # Evento de doble clic sobre una fila de la tabla
        self.tabla.bind("<Double-1>", self.abrir_opciones_producto)

    # ==========================================
    # CARGAR PRODUCTOS
    # ==========================================

    def cargar_productos(self):

        # Eliminar productos actuales de la tabla

        for elemento in self.tabla.get_children():

            self.tabla.delete(elemento)


        # Obtener productos del Excel

        productos = self.excel_service.obtener_productos()


        # Agregar productos a la tabla
     
        for producto in productos:

            tag = ()
        
            if producto["stock"] <= 3:
                tag = ("bajo_stock",)
            self.tabla.insert(
                "",
                "end",
                values=(
                    producto["id"],
                    producto["nombre"],
                    producto["descripcion"],
                    producto["categoria"],
                    producto["stock"],
                    f"${float(producto['precio']):,.2f}"
                ),
                tags=tag
            )

    # ==========================================
    # ABRIR HISTORIAL DE VENTAS
    # ==========================================

    def abrir_historial_ventas(self):

        HistorialVentas(
            self.ventana
        )



    def abrir_historial_presupuestos(self):
        HistorialPresupuestos(
            self.ventana
        )

    def abrir_actualizar_precios(self):
        VentanaActualizarPrecios(
            ventana_padre=self.ventana, callback_actualizado=self.cargar_productos
        )
    # ==========================================
    # BUSCAR PRODUCTOS
    # ==========================================

    def buscar_productos(self):

        texto_busqueda = self.entry_buscar.get().strip().lower()


        # Limpiar tabla

        for elemento in self.tabla.get_children():

            self.tabla.delete(elemento)


        # Obtener todos los productos

        productos = self.excel_service.obtener_productos()


        # Filtrar productos

        for producto in productos:

            nombre = str(
                producto["nombre"] or ""
            ).lower()

            tag = ()

            if producto["stock"] <= 3:
                tag = ("bajo_stock",)

            if texto_busqueda in nombre:

                self.tabla.insert(
                    "",
                    "end",
                    values=(
                        producto["id"],
                        producto["nombre"],
                        producto["descripcion"],
                        producto["categoria"],
                        producto["stock"],
                        f"${float(producto['precio']):,.2f}"
                    ),
                     tags=tag
                )


    # ==========================================
    # NUEVO PRODUCTO
    # ==========================================

    def nuevo_producto(self):

        FormularioProducto(
            ventana_padre=self.ventana,
            excel_service=self.excel_service,
            callback_guardado=self.cargar_productos
        )



    def abrir_reportes(self):
        VentanaReportes(self.ventana)
    # ==========================================
    # MODIFICAR PRODUCTO
    # ==========================================

    def modificar_producto(self):

    # ==========================================
    # OBTENER SELECCIÓN
    # ==========================================

        seleccion = self.tabla.selection()


    # ==========================================
    # VALIDAR SELECCIÓN
    # ==========================================

        if not seleccion:

            from tkinter import messagebox

            messagebox.showwarning(

                "Producto no seleccionado",

                "Debe seleccionar un producto de la tabla.",

                parent=self.ventana

            )

            return


    # ==========================================
    # OBTENER DATOS DE LA FILA
    # ==========================================

        elemento = self.tabla.item(
            seleccion[0]
        )


        datos = elemento["values"]


    # ==========================================
    # CONVERTIR DATOS A DICCIONARIO
    # ==========================================

        producto = {

            "id": datos[0],

            "nombre": datos[1],

            "descripcion": datos[2],

            "categoria": datos[3],

            "stock": datos[4],

           "precio": str(datos[5]).replace("$", "").replace(",", "")

            

        }


    # ==========================================
    # ABRIR FORMULARIO DE EDICIÓN
    # ==========================================

        FormularioProducto(

            ventana_padre=self.ventana,

            excel_service=self.excel_service,

            callback_guardado=self.cargar_productos,

            producto=producto

        )

        # ==========================================
    # ELIMINAR PRODUCTO
    # ==========================================

    def eliminar_producto(self):

        # ==========================================
        # OBTENER SELECCIÓN
        # ==========================================

        seleccion = self.tabla.selection()


        # ==========================================
        # VALIDAR SELECCIÓN
        # ==========================================

        if not seleccion:

            messagebox.showwarning(

                "Producto no seleccionado",

                "Debe seleccionar un producto de la tabla.",

                parent=self.ventana

            )

            return


        # ==========================================
        # OBTENER DATOS DE LA FILA
        # ==========================================

        elemento = self.tabla.item(
            seleccion[0]
        )


        datos = elemento["values"]


        # ==========================================
        # OBTENER DATOS DEL PRODUCTO
        # ==========================================

        id_producto = datos[0]

        nombre_producto = datos[1]

        stock_producto = datos[4]


        # ==========================================
        # CONFIRMAR ELIMINACIÓN
        # ==========================================

        confirmar = messagebox.askyesno(

            "Confirmar eliminación",

            f"¿Está seguro de eliminar este producto?\n\n"

            f"ID: {id_producto}\n"

            f"Producto: {nombre_producto}\n"

            f"Stock actual: {stock_producto}\n\n"

            f"Esta acción no se puede deshacer.",

            parent=self.ventana

        )


        # ==========================================
        # CANCELAR
        # ==========================================

        if not confirmar:

            return


        # ==========================================
        # ELIMINAR PRODUCTO
        # ==========================================

        try:

            self.excel_service.eliminar_producto(

                id_producto=id_producto

            )


            # ==========================================
            # MOSTRAR CONFIRMACIÓN
            # ==========================================

            messagebox.showinfo(

                "Producto eliminado",

                f"El producto '{nombre_producto}' "
                f"se eliminó correctamente.",

                parent=self.ventana

            )


            # ==========================================
            # ACTUALIZAR TABLA
            # ==========================================

            self.cargar_productos()


        except Exception as error:

            messagebox.showerror(

                "Error",

                f"No se pudo eliminar el producto.\n\n"

                f"Detalle: {error}",

                parent=self.ventana

            )


    # ==========================================
    # ABRIR VENTANA DE VENTA
    # ==========================================

    def abrir_venta(self):

    # Servicio especial para ventas.
    # No pertenece a un único rubro.
        excel_service_ventas = ExcelService()

        VentanaVenta(
            ventana_padre=self.ventana,
            excel_service=excel_service_ventas,
            callback_venta=self.cargar_productos
        )


    # ==========================================
    # EJECUTAR APLICACIÓN
    # ==========================================

    def ejecutar(self):

        self.ventana.mainloop()

    # ==========================================
    # agregar variante
    # ==========================================
    def agregar_variante(self):

        seleccion = self.tabla.selection()

        if not seleccion:

            messagebox.showwarning(
                "Producto no seleccionado",
                "Seleccione un producto para agregar una variante.",
                parent=self.ventana
            )

            return

        elemento = self.tabla.item(
            seleccion[0]
        )

        datos = elemento["values"]

    # ==========================================
    # CONVERTIR DATOS A DICCIONARIO
    # ==========================================

        producto = {

            "id": datos[0],

            "nombre": datos[1],

            "descripcion": datos[2],

            "categoria": datos[3],

            "stock": datos[4],

            "precio": str(datos[5]).replace("$", "").replace(",", "")

            

        }

    # ==========================================
    # ABRIR FORMULARIO DE VARIANTE
    # ==========================================

        FormularioProducto(

            ventana_padre=self.ventana,

            excel_service=self.excel_service,

            callback_guardado=self.cargar_productos,

            producto=producto,

            modo_variante=True

        )

    # ==========================================
    # CAMBIAR RUBRO
    # ==========================================

    def cambiar_rubro(self, rubro):

        self.rubro_actual = rubro

          # Actualizar título
        self.label_rubro.configure(
            text=f"RUBRO: {rubro.upper()}"
        )

        # Crear servicio correspondiente
        self.excel_service = ExcelService(
            rubro
        )

        # Limpiar búsqueda
        self.entry_buscar.delete(
            0,
            "end"
        )

        # Cargar productos del nuevo Excel
        self.cargar_productos()

    def crear_menu(self):

        barra_menu = tk.Menu(
            self.ventana
            )

        # ==========================
        # MENU PRODUCTOS
        # ==========================

        menu_productos = tk.Menu(
            barra_menu,
            tearoff=0,
            font=("Arial", 14)
        )

        menu_productos.add_command(
            label="Nuevo producto",
            command=self.nuevo_producto
        )

        menu_productos.add_command(
            label="Modificar producto",
            command=self.modificar_producto
        )

        menu_productos.add_command(
            label="Eliminar producto",
            command=self.eliminar_producto
        )

        menu_productos.add_separator()

        menu_productos.add_command(
            label="Agregar variante",
            command=self.agregar_variante
        )

        barra_menu.add_cascade(
            label="Productos",
            menu=menu_productos
        )

        # ==========================
        # MENU VENTAS
        # ==========================

        menu_ventas = tk.Menu(
            barra_menu,
            tearoff=0,
            font=("Arial", 14)
        )

        menu_ventas.add_command(
            label="Registrar venta",
            command=self.abrir_venta
        )

        barra_menu.add_cascade(
            label="Ventas",
            menu=menu_ventas
        )

        menu_ventas.add_command(
            label="Ver historial ventas",
            command=self.abrir_historial_ventas
        )
       

        menu_presupuestos = tk.Menu(
            barra_menu,
            tearoff=0,
            font=("Arial", 14)
        )

        menu_presupuestos.add_command(
            label="Nuevo presupuesto",
            command=self.nuevo_presupuesto
        )

        menu_presupuestos.add_command(
            label="Historial de presupuestos",
            command=self.abrir_historial_presupuestos
        )

        barra_menu.add_cascade(
            label="Presupuestos",
            menu=menu_presupuestos
        )


        ###
        # MENU REPORTES
        ###

        menu_reportes = tk.Menu(
            barra_menu,
            tearoff=0,
            font=("Arial", 14)
        )

        menu_reportes.add_command(
            label="Reporte de ventas", 
            command=self.abrir_reportes
        )

        barra_menu.add_cascade(
            label="Reportes",
            menu=menu_reportes
        )

        # ==========================
        # ASIGNAR MENU
        # ==========================
        
        menu_actualizar = tk.Menu(
            barra_menu,
            tearoff=0,
            font=("Arial", 14)
        )
        barra_menu.add_cascade(
            label="Actualizar",
            menu=menu_actualizar
        )
        menu_actualizar.add_command(
            label="Actualizar precios",
            command=self.abrir_actualizar_precios
        )

             

        # ==========================
        # ASIGNAR MENU
        # ==========================

        self.ventana.config(menu=barra_menu)

    def nuevo_presupuesto(self):

        VentanaPresupuesto(
        self.ventana,
        self.excel_service
    )

    # ==========================================
    # VENTANA DE OPCIONES AL HACER DOBLE CLIC
    # ==========================================
    def abrir_opciones_producto(self, evento=None):
      seleccion = self.tabla.selection()

      if not seleccion:
        return

      elemento = self.tabla.item(seleccion[0])
      datos = elemento["values"]

      id_prod = datos[0]
      nombre_prod = datos[1]

      # Crear ventana modal emergente
      ventana_opciones = ctk.CTkToplevel(self.ventana)
      ventana_opciones.title("Opciones de Producto")
      ventana_opciones.geometry("400x320")
      ventana_opciones.resizable(False, False)
      ventana_opciones.transient(self.ventana)
      ventana_opciones.grab_set()

      # Título con nombre del producto
      lbl_titulo = ctk.CTkLabel(
          ventana_opciones,
          text=f"PRODUCTO #{id_prod}",
          font=ctk.CTkFont(size=20, weight="bold"),
      )
      lbl_titulo.pack(pady=(20, 5))

      lbl_subtitulo = ctk.CTkLabel(
          ventana_opciones,
          text=f"{nombre_prod}",
          font=ctk.CTkFont(size=14),
          text_color="gray",
      )
      lbl_subtitulo.pack(pady=(0, 20))

      # Frame para los botones de acción
      frame_acciones = ctk.CTkFrame(ventana_opciones, fg_color="transparent")
      frame_acciones.pack(fill="both", expand=True, padx=30, pady=10)

      # Auxiliares para cerrar la ventana modal y llamar a la función seleccionada
      def accion_modificar():
        ventana_opciones.destroy()
        self.modificar_producto()

      def accion_variante():
        ventana_opciones.destroy()
        self.agregar_variante()

      def accion_eliminar():
        ventana_opciones.destroy()
        self.eliminar_producto()

      # Botón Modificar
      btn_modificar = ctk.CTkButton(
          frame_acciones,
          text="Modificar producto",
          command=accion_modificar,
          height=38,
          font=ctk.CTkFont(size=13, weight="bold"),
      )
      btn_modificar.pack(fill="x", pady=6)

      # Botón Agregar Variante
      btn_variante = ctk.CTkButton(
          frame_acciones,
          text="Agregar variante",
          command=accion_variante,
          fg_color="#2b8a3e",
          hover_color="#1f632d",
          height=38,
          font=ctk.CTkFont(size=13, weight="bold"),
      )
      btn_variante.pack(fill="x", pady=6)

      # Botón Eliminar
      btn_eliminar = ctk.CTkButton(
          frame_acciones,
          text="Eliminar producto",
          command=accion_eliminar,
          fg_color="#c92a2a",
          hover_color="#961f1f",
          height=38,
          font=ctk.CTkFont(size=13, weight="bold"),
      )
      btn_eliminar.pack(fill="x", pady=6)

    def cerrar_aplicacion(self):
    # Oculta la ventana primero para dar sensación de cierre instantáneo
        self.ventana.withdraw()

    # Destruye la ventana capturando cualquier alerta en segundo plano de Tkinter
        try:
            self.ventana.quit()
            self.ventana.destroy()
        except Exception:
            pass

       