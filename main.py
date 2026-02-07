import flet as ft
import os
import sys
import json
import logging
import AutoExcel as AE
import seguridad
from ui_components import *
from utils import *
import ctypes

try:
    myappid = "Isaias.AutoExcel.0.1.0"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

# --- Mantenemos tu lógica de recursos ---
if getattr(sys, 'frozen', False):
    base = sys._MEIPASS
else:
    base = os.path.dirname(__file__)

class TextboxHandler(logging.Handler):
    def __init__(self, textbox, page):
        super().__init__()
        self.textbox = textbox
        self.page = page

    def emit(self, record):
        msg = self.format(record)
        self.textbox.value += msg + "\n"
        self.page.update()

class AutoExcelApp:
    def __init__(self, page: ft.Page, debug=False):

        self.page = page
        self.debug = debug
        self.page.title = "AutoExcel - Generador de Nóminas"
        self.page.window.app_id = "Isaias.AutoExcel"
        
        
        # Logger configurado PRIMERO (antes de cualquier log_debug)
        self.logger = logging.getLogger("AutoExcel")
        self.logger.setLevel(logging.DEBUG)
        
        # Formato para los logs
        formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s', 
                                     datefmt='%Y-%m-%d %H:%M:%S')
        
        # Configuración de ventana para 0.28.3
        self.page.window.width = 800
        self.page.window.height = 500 if not self.debug else 750
        self.page.window.resizable = False
        page.window.maximizable = False
        self.page.window.center()

        
        ruta_icono = AE.obtener_ruta_recurso("static/logo.ico")
        if os.path.exists(ruta_icono):
            self.page.window.icon = ruta_icono
            log_debug(self, "INIT", f"✓ Icono cargado desde: {ruta_icono}")
        else:
            log_debug(self, "INIT", f"⚠️ Icono no encontrado en: {ruta_icono}")
        
        self.usuarios_data = [] # Lista de diccionarios [usuario1, usuario2, ...]
        self.lista_usuarios_ui = ft.Column(spacing=10, scroll=ft.ScrollMode.ADAPTIVE) 
        self.usuario_logueado = None 
        self.es_oscuro = False    

        build_ui(self)
        log_debug(self, "INIT", f"Tema del sistema detectado: {'OSCURO' if self.es_oscuro else 'CLARO'}")
        log_debug(self, "INIT", "✓ Interfaz de usuario construida")

        # FilePicker con on_result en el constructor
        self.file_picker_tipo = None
        self.file_picker = ft.FilePicker(on_result=self._file_picker_result)
        self.page.overlay.append(self.file_picker)
        log_debug(self, "INIT", "✓ FilePicker inicializado")



        
        # Configurar handlers según el modo
        if self.debug:
            # Modo debug: mostrar en pantalla
            text_handler = TextboxHandler(self.log_textbox, self.page)
            text_handler.setFormatter(formatter)
            self.logger.addHandler(text_handler)
            log_debug(self, "INIT", "✓ Handler de debug activado (pantalla)")
        else:
            # Modo normal: guardar en archivo
            try:
                file_handler = logging.FileHandler('autoexcel.log', encoding='utf-8', mode='w') 
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
                log_debug(self, "INIT", "✓ Handler de archivo activado (autoexcel.log)")
            except Exception as e:
                # Si no se puede crear el archivo, al menos registramos el error en consola
                console_handler = logging.StreamHandler()
                console_handler.setFormatter(formatter)
                self.logger.addHandler(console_handler)
                self.logger.error(f"No se pudo crear archivo de log: {e}")
            
        log_debug(self, "INIT", "✓✓✓ Aplicación inicializada correctamente")
    
    def _file_picker_result(self, e: ft.FilePickerResultEvent):
        if self.file_picker_tipo == "db" and e.files:
            ruta = e.files[0].path
            self.db_path_field.value = ruta
            log_debug(self, "FILEPICKER", f"✓ Base de datos seleccionada: {ruta}")
        elif self.file_picker_tipo == "folder" and e.path:
            self.output_path_field.value = e.path
            log_debug(self, "FILEPICKER", f"✓ Carpeta de salida seleccionada: {e.path}")
        else:
            log_debug(self, "FILEPICKER", "✗ Selección cancelada por el usuario")
        
        self.page.update()




def main(page: ft.Page):
    AutoExcelApp(page, debug=False)

if __name__ == "__main__":
    ft.app(
        target=main,
        assets_dir="static")