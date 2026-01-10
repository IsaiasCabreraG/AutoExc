import flet as ft
import os
import sys
import json
import logging
import AutoExcel as AE

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
        
        # Configuración de ventana para 0.28.3
        self.page.window.width = 700
        self.page.window.height = 400 if not self.debug else 650
        self.page.window.resizable = False
        
        ruta_icono = AE.obtener_ruta_recurso("static/logo.ico")
        if os.path.exists(ruta_icono):
            self.page.window.icon = ruta_icono
        
        # Logger configurado
        self.logger = logging.getLogger("AutoExcel")
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(levelname)s: %(message)s')
        
        # FilePicker con on_result en el constructor
        self.file_picker_tipo = None
        self.file_picker = ft.FilePicker(on_result=self._file_picker_result)
        self.page.overlay.append(self.file_picker)
        
        self.build_ui()
        
        if self.debug:
            text_handler = TextboxHandler(self.log_textbox, self.page)
            text_handler.setFormatter(formatter)
            self.logger.addHandler(text_handler)
        
        self.refrescar_años_db()
        self.cargar_configuracion()
    
    def build_ui(self):
        self.main_view = self.build_main_view()
        self.config_view = self.build_config_view()
        
        # Contenedor de vistas
        self.views_container = ft.Column(
            controls=[self.main_view, self.config_view],
            spacing=0
        )
        
        self.page.add(self.views_container)
        
        if self.debug:
            self.debug_view = self.build_debug_view()
            self.page.add(self.debug_view)
    
    def build_main_view(self):
        self.nomina_field = ft.TextField(
            label="Nómina",
            hint_text="Ej: 12177",
            width=140,
            text_align=ft.TextAlign.CENTER
        )
        
        self.año_dropdown = ft.Dropdown(
            label="Año",
            width=120,
            value="----",
            options=[ft.dropdown.Option("----")],
            disabled=True
        )
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Row(
                        controls=[
                            ft.Container(expand=1),
                            ft.Text("AutoExcel", size=32, weight=ft.FontWeight.BOLD),
                            ft.Container(
                                content=ft.IconButton(
                                    icon=ft.Icons.SETTINGS,
                                    on_click=lambda _: self.cambiar_ventana("config")
                                ),
                                expand=1,
                                alignment=ft.alignment.center_right,
                                padding=ft.padding.only(right=10)
                            )
                        ],
                    ),
                    ft.Container(height=50),
                    ft.Text("Ingrese Nómina y Año:", size=14),
                    ft.Row(
                        controls=[self.nomina_field, self.año_dropdown],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=10
                    ),
                    ft.Container(height=20),
                    ft.ElevatedButton(
                        text="Generar Nómina",
                        width=200,
                        height=45,
                        on_click=self.generar_nomina
                    )
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            visible=True
        )

    def build_config_view(self):
        self.rut_field = ft.TextField(label="RUT Empresa", width=330)
        self.db_path_field = ft.TextField(label="Base de Datos", width=330, read_only=True)
        self.output_path_field = ft.TextField(label="Carpeta de Salida", width=330, read_only=True)
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Row([
                        ft.IconButton(
                            icon=ft.Icons.ARROW_BACK,
                            on_click=lambda _: self.confirmar_salida()
                        ),
                        ft.Text(
                            "Configuración",
                            size=32,
                            weight=ft.FontWeight.BOLD,
                            expand=True,
                            text_align=ft.TextAlign.CENTER
                        ),
                        ft.Container(width=40)
                    ]),
                    ft.Container(height=20),
                    self.rut_field,
                    ft.Row([
                        self.db_path_field,
                        ft.IconButton(
                            icon=ft.Icons.FILE_OPEN,
                            on_click=lambda _: self.buscar_archivo("db")
                        )
                    ], alignment=ft.MainAxisAlignment.CENTER),
                    ft.Row([
                        self.output_path_field,
                        ft.IconButton(
                            icon=ft.Icons.FOLDER_OPEN,
                            on_click=lambda _: self.buscar_archivo("folder")
                        )
                    ], alignment=ft.MainAxisAlignment.CENTER),
                    ft.Row([
                        ft.ElevatedButton(
                            "Guardar",
                            width=150,
                            on_click=lambda _: self.guardar_configuracion()
                        ),
                        ft.OutlinedButton(
                            "Cancelar",
                            width=150,
                            on_click=lambda _: self.confirmar_salida()
                        )
                    ], alignment=ft.MainAxisAlignment.CENTER)
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            visible=False
        )

    def build_debug_view(self):
        self.log_textbox = ft.TextField(
            multiline=True,
            read_only=True,
            min_lines=5,
            text_size=11
        )
        return ft.Container(
            content=ft.Column([
                ft.Text("Consola Debug"),
                self.log_textbox
            ]),
            padding=10
        )

    # --- LÓGICA DE NEGOCIO ---
    def cambiar_ventana(self, ventana):
        self.main_view.visible = (ventana == "main")
        self.config_view.visible = (ventana == "config")
        self.page.update()

    def buscar_archivo(self, tipo):
        self.file_picker_tipo = tipo
        if tipo == "db":
            self.file_picker.pick_files(
                dialog_title="Seleccionar Base de Datos",
                allowed_extensions=["accdb"]
            )
        else:
            self.file_picker.get_directory_path(
                dialog_title="Seleccionar Carpeta de Salida"
            )

    def _file_picker_result(self, e: ft.FilePickerResultEvent):
        if self.file_picker_tipo == "db" and e.files:
            self.db_path_field.value = e.files[0].path
        elif self.file_picker_tipo == "folder" and e.path:
            self.output_path_field.value = e.path
        self.page.update()

    def guardar_configuracion(self):
        """Guarda la configuración y actualiza los años disponibles"""
        config_data = {
            "ruta_db": self.db_path_field.value,
            "ruta_output": self.output_path_field.value,
            "rut_empresa": self.rut_field.value
        }
        
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
            
            self.logger.info("Configuración guardada exitosamente")
            self.mostrar_dialogo("Éxito", "Configuración guardada correctamente.")
            self.cambiar_ventana("main")
            self.refrescar_años_db()
        except Exception as e:
            self.logger.error(f"Error al guardar configuración: {e}")
            self.mostrar_dialogo("Error", f"No se pudo guardar la configuración: {e}", True)

    def cargar_configuracion(self):
        """Carga la configuración desde config.json"""
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.db_path_field.value = data.get("ruta_db", "")
                    self.output_path_field.value = data.get("ruta_output", "")
                    self.rut_field.value = data.get("rut_empresa", "")
                self.logger.debug("Configuración cargada correctamente")
                self.page.update()
            except Exception as e:
                self.logger.error(f"Error al cargar configuración: {e}")

    def refrescar_años_db(self):
        """Obtiene los años disponibles de la base de datos usando AutoExcel"""
        try:
            self.logger.debug("Obteniendo años disponibles de la base de datos...")
            lista_años = AE.obtener_años_disponibles()
            
            if lista_años:
                self.año_dropdown.disabled = False
                self.año_dropdown.options = [ft.dropdown.Option(str(a)) for a in lista_años]
                self.año_dropdown.value = str(lista_años[0])
                self.logger.info(f"Años disponibles cargados: {len(lista_años)} años encontrados")
            else:
                self.año_dropdown.disabled = True
                self.año_dropdown.options = [ft.dropdown.Option("----")]
                self.año_dropdown.value = "----"
                self.logger.warning("No se encontraron años disponibles en la base de datos")
                
            self.page.update()
        except Exception as e:
            self.logger.error(f"Error al obtener años: {e}")
            self.año_dropdown.disabled = True

    def generar_nomina(self, e):
        """Genera la nómina usando AutoExcel"""
        nomina = self.nomina_field.value
        año = self.año_dropdown.value
        
        # Validaciones
        if not nomina or not nomina.isdigit():
            self.mostrar_dialogo("Error", "Ingrese una nómina válida (solo números)", True)
            return
        
        if not año or año == "----":
            self.mostrar_dialogo("Error", "Seleccione un año válido", True)
            return
        
        # Verificar configuración
        if not os.path.exists("config.json"):
            self.mostrar_dialogo("Error", "Configure la aplicación antes de generar nóminas", True)
            self.cambiar_ventana("config")
            return
        
        try:
            self.logger.info(f"Iniciando generación de nómina {nomina} para el año {año}")
            
            # Abrir conexión a la base de datos
            conn = AE.open_database()
            if not conn:
                self.mostrar_dialogo("Error", "No se pudo conectar a la base de datos. Verifique la configuración.", True)
                return
            
            try:
                cursor = conn.cursor()
                
                # Obtener datos del egreso
                self.logger.debug(f"Buscando datos de egreso para nómina {nomina}...")
                datos = AE.datos_egreso(cursor, int(nomina), int(año))
                
                if not datos:
                    self.mostrar_dialogo(
                        "Advertencia", 
                        f"No se encontraron registros para la nómina {nomina} en el año {año}", 
                        True
                    )
                    return
                
                self.logger.info(f"Se encontraron {len(datos)} registros")
                
                # Obtener RUT de empresa desde configuración
                with open("config.json", "r", encoding="utf-8") as f:
                    config = json.load(f)
                    rut_empresa = config.get("rut_empresa", "")
                
                # Generar el archivo Excel
                self.logger.debug("Generando archivo Excel...")
                total_pago = AE.llenar_nomina_xlsx(datos, nomina, año, rut_empresa)
                
                if total_pago is not None:
                    self.mostrar_dialogo(
                        "Éxito", 
                        f"Nómina generada correctamente\n\n"
                        f"Registros: {len(datos)}\n"
                        f"Total: ${total_pago:,.0f}".replace(",", ".")
                    )
                    self.logger.info(f"Nómina {nomina}-{año} generada exitosamente. Total: ${total_pago:,.0f}")
                else:
                    self.mostrar_dialogo("Error", "Ocurrió un error al generar el archivo Excel", True)
                    
            finally:
                conn.close()
                self.logger.debug("Conexión a base de datos cerrada")
                
        except Exception as ex:
            self.logger.error(f"Error al generar nómina: {ex}")
            self.mostrar_dialogo("Error", f"Error al generar nómina:\n{str(ex)}", True)

    def mostrar_dialogo(self, titulo, mensaje, es_error=False):
        """Muestra un diálogo modal con un mensaje"""
        def cerrar(e):
            dialogo.open = False
            self.page.update()
        
        dialogo = ft.AlertDialog(
            title=ft.Text(titulo),
            content=ft.Text(mensaje),
            actions=[ft.TextButton("OK", on_click=cerrar)]
        )
        self.page.overlay.append(dialogo)
        dialogo.open = True
        self.page.update()

    def confirmar_salida(self):
        """Vuelve a la ventana principal sin guardar cambios"""
        self.cargar_configuracion()  # Restaura los valores originales
        self.cambiar_ventana("main")

def main(page: ft.Page):
    AutoExcelApp(page, debug=False)

if __name__ == "__main__":
    ft.app(target=main)