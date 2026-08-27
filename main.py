import customtkinter as ctk

from services.excel_service import ExcelService
from ui.ventana_principal import VentanaPrincipal

def iniciar_aplicacion():

    excel_service = ExcelService()
    aplicacion = VentanaPrincipal()

    aplicacion.ejecutar()
   

   


if __name__ == "__main__":
    iniciar_aplicacion()