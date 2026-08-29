import customtkinter as ctk
from tkinter import messagebox
from services.excel_service import ExcelService


class VentanaActualizarPrecios:

  def __init__(self, ventana_padre, callback_actualizado=None):
    self.ventana_padre = ventana_padre
    self.callback_actualizado = callback_actualizado

    # Servicio general de Excel sin rubro fijo
    self.excel_service = ExcelService()

    # Mapeo de nombres para la interfaz
    self.rubros_map = {
        "Ferretería": "ferreteria",
        "Refrigeración": "refrigeracion",
        "Electricidad": "electricidad",
    }

    # ==========================================
    # VENTANA
    # ==========================================
    self.ventana = ctk.CTkToplevel(ventana_padre)
    self.ventana.title("Actualizar Precios por Rubro")
    self.ventana.geometry("500x420")
    self.ventana.resizable(False, False)

    self.ventana.transient(ventana_padre)
    self.ventana.grab_set()

    self.crear_interfaz()

  def crear_interfaz(self):
    titulo = ctk.CTkLabel(
        self.ventana,
        text="ACTUALIZAR PRECIOS",
        font=ctk.CTkFont(size=22, weight="bold"),
    )
    titulo.pack(pady=(25, 15))

    frame_form = ctk.CTkFrame(self.ventana)
    frame_form.pack(fill="both", expand=True, padx=30, pady=10)

    lbl_rubro = ctk.CTkLabel(
        frame_form,
        text="Seleccionar Rubro:",
        font=ctk.CTkFont(size=14, weight="bold"),
    )
    lbl_rubro.pack(anchor="w", padx=20, pady=(15, 5))

    opciones_rubro = ["Todos los rubros"] + list(self.rubros_map.keys())
    self.combo_rubro = ctk.CTkOptionMenu(
        frame_form, values=opciones_rubro, width=380, height=35
    )
    self.combo_rubro.pack(padx=20, pady=(0, 15))

    lbl_porcentaje = ctk.CTkLabel(
        frame_form,
        text="Porcentaje de modificación (%):",
        font=ctk.CTkFont(size=14, weight="bold"),
    )
    lbl_porcentaje.pack(anchor="w", padx=20, pady=(5, 5))

    self.entry_porcentaje = ctk.CTkEntry(
        frame_form,
        width=380,
        height=35,
        placeholder_text="Ej: 15 (aumenta 15%) o -10 (descuento 10%)",
    )
    self.entry_porcentaje.pack(padx=20, pady=(0, 10))

    lbl_info = ctk.CTkLabel(
        frame_form,
        text="* Use números positivos para aumentos y negativos para descuentos.",
        font=ctk.CTkFont(size=11),
        text_color="gray",
    )
    lbl_info.pack(anchor="w", padx=20, pady=(0, 15))

    frame_botones = ctk.CTkFrame(self.ventana, fg_color="transparent")
    frame_botones.pack(fill="x", padx=30, pady=(10, 20))

    btn_aplicar = ctk.CTkButton(
        frame_botones,
        text="Aplicar Cambio",
        command=self.aplicar_actualizacion,
        fg_color="#1f538d",
        hover_color="#14375e",
        width=180,
        height=38,
        font=ctk.CTkFont(size=14, weight="bold"),
    )
    btn_aplicar.pack(side="left", padx=(20, 10))

    btn_cancelar = ctk.CTkButton(
        frame_botones,
        text="Cancelar",
        command=self.ventana.destroy,
        fg_color="gray",
        hover_color="#555555",
        width=140,
        height=38,
        font=ctk.CTkFont(size=14),
    )
    btn_cancelar.pack(side="right", padx=(10, 20))

  def aplicar_actualizacion(self):
    texto_porcentaje = self.entry_porcentaje.get().strip().replace(",", ".")

    try:
      porcentaje = float(texto_porcentaje)
      if porcentaje == 0:
        messagebox.showwarning(
            "Atención",
            "El porcentaje ingresado debe ser distinto de 0.",
            parent=self.ventana,
        )
        return
    except ValueError:
      messagebox.showerror(
          "Error",
          "Por favor, ingrese un número válido para el porcentaje.",
          parent=self.ventana,
      )
      return

    seleccion_combo = self.combo_rubro.get()
    rubro_clave = self.rubros_map.get(
        seleccion_combo, None
    )  # None si elige "Todos los rubros"

    tipo_operacion = "Aumento" if porcentaje > 0 else "Descuento"
    mensaje_conf = (
        f"¿Está seguro de aplicar un {tipo_operacion} del {abs(porcentaje)}% "
        f"a: '{seleccion_combo}'?"
    )

    confirmar = messagebox.askyesno(
        "Confirmar Actualización", mensaje_conf, parent=self.ventana
    )
    if not confirmar:
      return

    try:
      # Llama al método de ExcelService
      modificados = self.excel_service.actualizar_precios_por_porcentaje(
          porcentaje=porcentaje, rubro=rubro_clave
      )

      messagebox.showinfo(
          "Éxito",
          f"Se actualizaron correctamente los precios de {modificados} productos.",
          parent=self.ventana,
      )

      if self.callback_actualizado:
        self.callback_actualizado()

      self.ventana.destroy()

    except Exception as error:
      messagebox.showerror(
          "Error",
          f"Ocurrió un error al intentar actualizar los precios:\n\n{error}",
          parent=self.ventana,
      )