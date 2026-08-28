import tkinter as tk
from tkinter import ttk, messagebox


# Ventana secundaria: Muestra el historial de presupuestos
class VentanaHistorialPresupuestos(tk.Toplevel):

  def __init__(self, parent):
    super().__init__(parent)
    self.title("Historial de Presupuestos")
    self.geometry("600x400")

    # Hacer que la ventana sea modal (opcional: impide tocar la ventana principal)
    self.transient(parent)
    self.grab_set()

    self.crear_interfaz()

  def crear_interfaz(self):
    lbl_titulo = tk.Label(
        self, text="Presupuestos Generados", font=("Arial", 14, "bold")
    )
    lbl_titulo.pack(pady=10)

    # Tabla (Treeview) para listar presupuestos
    columnas = ("id", "cliente", "fecha", "total")
    self.tabla = ttk.Treeview(self, columns=columnas, show="headings")

    self.tabla.heading("id", text="Nº Presupuesto")
    self.tabla.heading("cliente", text="Cliente")
    self.tabla.heading("fecha", text="Fecha")
    self.tabla.heading("total", text="Total ($)")

    self.tabla.column("id", width=100, anchor="center")
    self.tabla.column("cliente", width=200)
    self.tabla.column("fecha", width=100, anchor="center")
    self.tabla.column("total", width=100, anchor="e")

    self.tabla.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Cargar datos de ejemplo o desde base de datos
    self.cargar_datos()

  def cargar_datos(self):
    # Aquí puedes llamar a tu base de datos (SQLite, JSON, etc.)
    presupuestos_ejemplo = [
        ("001", "Juan Pérez", "2026-08-20", "15000.00"),
        ("002", "María Gómez", "2026-08-25", "8500.50"),
    ]
    for p in presupuestos_ejemplo:
      self.tabla.insert("", tk.END, values=p)


# Ventana Principal
class VentanaPrincipal(tk.Tk):

  def __init__(self):
    super().__init__()
    self.title("Sistema de Gestión de Stock")
    self.geometry("800x600")

    self.crear_menu()

  def crear_menu(self):
    barra_menu = tk.Menu(self)

    # Menú desplegable para 'Presupuestos'
    menu_presupuestos = tk.Menu(barra_menu, tearoff=0)

    # CONEXIÓN DEL MENÚ A LA VENTANA:
    # Usamos la opción command llamando al método correspondiente
    menu_presupuestos.add_command(
        label="Nuevo Presupuesto", command=self.abrir_nuevo_presupuesto
    )
    menu_presupuestos.add_command(
       
        label="Historial de Presupuestos",
        command=self.abrir_historial_presupuestos,
        
    )

    # Añadir el submenú a la barra principal
    barra_menu.add_cascade(label="Presupuestos", menu=menu_presupuestos)

    # Configurar la ventana principal para usar este menú
    self.config(menu=barra_menu)

  def abrir_historial_presupuestos(self):
    # Instancia y abre la ventana secundaria
    VentanaHistorialPresupuestos(self)

  def abrir_nuevo_presupuesto(self):
    pass


if __name__ == "__main__":
  app = VentanaPrincipal()
  app.mainloop()