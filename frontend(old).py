import flet as ft
import os
import sys
import json
import logging
import AutoExcel as AE

if getattr(sys, 'frozen', False):
    base = sys._MEIPASS
else:
    base = os.path.dirname(__file__)

class TextboxHandler(logging.Handler):
    """Handler personalizado para mostrar logs en el TextBox de Flet"""
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
        self.page.window.width = 500
        self.page.window.height = 300 if not self.debug else 550
        self.page.window.resizable = False
        self.page.window.center()
        self.page.theme_mode = ft.ThemeMode.SYSTEM
        
        # Configurar icono
        ruta_icono = AE.obtener_ruta_recurso("static/logo.ico")
        if os.path.exists(ruta_icono):
            self.page.window.icon = ruta_icono
        
        # Configurar logger
        self.logger = logging.getLogger("AutoExcel")
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(levelname)s: %(message)s')
        
        file_handler = logging.FileHandler("autoexcel_debug.log", mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        # Variables de estado
        self.año_seleccionado = "----"
        
        # Construir la UI
        self.build_ui()
        
        # Si está en modo debug, agregar handler al textbox
        if self.debug:
            text_handler = TextboxHandler(self.log_textbox, self.page)
            text_handler.setFormatter(formatter)
            self.logger.addHandler(text_handler)
        
        self.logger.info("Aplicación iniciada correctamente")
        self.refrescar_años_db()
        self.cargar_configuracion()
    
    def build_ui(self):
        """Construye toda la interfaz de usuario"""
        # Vista Principal
        self.main_view = self.build_main_view()
        
        # Vista de Configuración
        self.config_view = self.build_config_view()
        
        # Vista de Debug (si está habilitado)
        if self.debug:
            self.debug_view = self.build_debug_view()
        
        # Agregar vista principal inicialmente
        self.page.add(self.main_view)
        if self.debug:
            self.page.add(self.debug_view)
    
    def build_main_view(self):
        """Construye la vista principal"""
        # Campo de entrada de nómina
        self.nomina_field = ft.TextField(
            label="Nómina",
            hint_text="Ej: 12177",
            width=140,
            on_change=self.validar_numeros,
            text_align=ft.TextAlign.CENTER
        )
        
        # Dropdown de años
        self.año_dropdown = ft.Dropdown(
            label="Año",
            width=100,
            value="----",
            options=[ft.dropdown.Option("----")],
            disabled=True
        )
        
        # Botón de configuración
        config_btn = ft.IconButton(
            icon=ft.icons.SETTINGS,
            tooltip="Configuración",
            on_click=lambda _: self.cambiar_ventana("config")
        )
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    # Header
                    ft.Row(
                        controls=[
                            ft.Container(expand=1),
                            ft.Text("AutoExcel", size=32, weight=ft.FontWeight.BOLD),
                            ft.Container(
                                content=config_btn,
                                expand=1,
                                alignment=ft.alignment.center_right,
                                padding=ft.padding.only(right=10)
                            )
                        ],
                        alignment=ft.MainAxisAlignment.CENTER
                    ),
                    ft.Container(height=40),
                    
                    # Inputs
                    ft.Text("Ingrese Nómina y Año:", size=14),
                    ft.Container(height=5),
                    ft.Row(
                        controls=[
                            self.nomina_field,
                            self.año_dropdown
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=10
                    ),
                    ft.Container(height=20),
                    
                    # Botón generar
                    ft.ElevatedButton(
                        text="Generar Nómina",
                        width=150,
                        height=40,
                        on_click=self.generar_nomina
                    )
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=0
            ),
            padding=ft.padding.only(top=15),
            visible=True
        )
    
    def build_config_view(self):
        """Construye la vista de configuración"""
        # Botón volver
        back_btn = ft.IconButton(
            icon=ft.icons.ARROW_BACK,
            tooltip="Volver",
            on_click=lambda _: self.confirmar_salida()
        )
        
        # Campos de configuración
        self.rut_field = ft.TextField(
            label="RUT Empresa",
            hint_text="12.345.678-9",
            width=400
        )
        
        self.db_path_field = ft.TextField(
            label="Base de Datos",
            hint_text="Seleccione un archivo...",
            width=330,
            read_only=True
        )
        
        self.output_path_field = ft.TextField(
            label="Carpeta de Salida",
            hint_text="Seleccione una carpeta...",
            width=330,
            read_only=True
        )
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    # Header
                    ft.Row(
                        controls=[
                            ft.Container(
                                content=back_btn,
                                padding=ft.padding.only(left=10)
                            ),
                            ft.Container(expand=1),
                            ft.Text("Configuración", size=32, weight=ft.FontWeight.BOLD),
                            ft.Container(expand=1),
                            ft.Container(width=50)
                        ]
                    ),
                    ft.Container(height=25),
                    
                    # Campos
                    ft.Container(
                        content=ft.Column(
                            controls=[
                                self.rut_field,
                                ft.Container(height=10),
                                
                                ft.Row(
                                    controls=[
                                        self.db_path_field,
                                        ft.IconButton(
                                            icon=ft.icons.FOLDER_OPEN,
                                            tooltip="Buscar archivo",
                                            on_click=lambda _: self.buscar_archivo("db")
                                        )
                                    ],
                                    spacing=5
                                ),
                                ft.Container(height=10),
                                
                                ft.Row(
                                    controls=[
                                        self.output_path_field,
                                        ft.IconButton(
                                            icon=ft.icons.FOLDER_OPEN,
                                            tooltip="Buscar carpeta",
                                            on_click=lambda _: self.buscar_archivo("folder")
                                        )
                                    ],
                                    spacing=5
                                )
                            ],
                            horizontal_alignment=ft.CrossAxisAlignment.CENTER
                        ),
                        padding=20
                    ),
                    ft.Container(height=20),
                    
                    # Botones
                    ft.Row(
                        controls=[
                            ft.ElevatedButton(
                                text="Guardar",
                                width=150,
                                on_click=lambda _: self.guardar_configuracion()
                            ),
                            ft.OutlinedButton(
                                text="Cancelar",
                                width=150,
                                on_click=lambda _: self.confirmar_salida()
                            )
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=20
                    )
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            padding=ft.padding.only(top=15),
            visible=False
        )
    
    def build_debug_view(self):
        """Construye la vista de debug"""
        self.log_textbox = ft.TextField(
            value="",
            multiline=True,
            read_only=True,
            min_lines=8,
            max_lines=8,
            width=460,
            text_style=ft.TextStyle(font_family="Consolas", size=11)
        )
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text("Consola de Depuración", weight=ft.FontWeight.BOLD),
                    self.log_textbox
                ],
                spacing=10
            ),
            padding=20,
            visible=True
        )
    
    def validar_numeros(self, e):
        """Valida que solo se ingresen números"""
        valor = e.control.value
        if valor and not valor.isdigit():
            e.control.value = valor[:-1]
            self.page.update()
    
    def cambiar_ventana(self, ventana):
        """Cambia entre vistas"""
        if ventana == "config":
            self.main_view.visible = False
            self.config_view.visible = True
            self.cargar_configuracion()
        else:  # main
            self.main_view.visible = True
            self.config_view.visible = False
        self.page.update()
    
    def buscar_archivo(self, tipo):
        """Abre el diálogo de selección de archivo/carpeta"""
        def on_result(e: ft.FilePickerResultEvent):
            if tipo == "db" and e.files:
                self.db_path_field.value = e.files[0].path
                self.logger.info(f"Base de datos seleccionada: {e.files[0].path}")
            elif tipo == "folder" and e.path:
                self.output_path_field.value = e.path
                self.logger.info(f"Carpeta de salida seleccionada: {e.path}")
            self.page.update()
        
        if tipo == "db":
            picker = ft.FilePicker(on_result=on_result)
            self.page.overlay.append(picker)
            self.page.update()
            picker.pick_files(
                dialog_title="Seleccionar Base de Datos",
                allowed_extensions=["accdb"],
                allow_multiple=False
            )
        else:  # folder
            picker = ft.FilePicker(on_result=on_result)
            self.page.overlay.append(picker)
            self.page.update()
            picker.get_directory_path(dialog_title="Seleccionar Carpeta de Salida")
    
    def guardar_configuracion(self):
        """Guarda la configuración en JSON"""
        config_data = {
            "ruta_db": self.db_path_field.value or "",
            "ruta_output": self.output_path_field.value or "",
            "rut_empresa": self.rut_field.value or ""
        }
        
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
            
            self.mostrar_dialogo("Éxito", "Configuración guardada correctamente.", es_error=False)
            self.logger.info("Configuración guardada exitosamente.")
            self.cambiar_ventana("main")
            self.refrescar_años_db()
        except Exception as e:
            self.mostrar_dialogo("Error", f"No se pudo guardar: {e}", es_error=True)
            self.logger.error(f"Error al guardar configuración: {e}")
    
    def cargar_configuracion(self):
        """Carga la configuración desde JSON"""
        self.logger.info("Cargando configuración desde JSON...")
        
        if not os.path.exists("config.json"):
            self.logger.warning("No se encontró config.json. Creando uno nuevo.")
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump({"ruta_db": "", "ruta_output": "", "rut_empresa": ""}, f, indent=4)
        
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                config_data = json.load(f)
            
            self.db_path_field.value = config_data.get("ruta_db", "")
            self.output_path_field.value = config_data.get("ruta_output", "")
            self.rut_field.value = config_data.get("rut_empresa", "")
            
            self.logger.debug("Configuración cargada correctamente.")
            self.page.update()
        except Exception as e:
            self.logger.error(f"Error al cargar configuración: {e}")
    
    def confirmar_salida(self):
        """Confirma la salida sin guardar cambios"""
        self.logger.info("Confirmando salida de la configuración...")
        
        config_guardada = {"ruta_db": "", "ruta_output": "", "rut_empresa": ""}
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                config_guardada = json.load(f)
        
        if (self.db_path_field.value != config_guardada.get("ruta_db", "") or
            self.output_path_field.value != config_guardada.get("ruta_output", "") or
            self.rut_field.value != config_guardada.get("rut_empresa", "")):
            
            self.logger.warning("Cambios detectados sin guardar.")
            
            def cerrar_dialogo(e):
                dialogo.open = False
                if e.control.text == "Sí":
                    self.cambiar_ventana("main")
                self.page.update()
            
            dialogo = ft.AlertDialog(
                modal=True,
                title=ft.Text("Cambios sin guardar"),
                content=ft.Text("Has realizado cambios en la configuración. ¿Deseas salir sin guardar?"),
                actions=[
                    ft.TextButton("No", on_click=cerrar_dialogo),
                    ft.TextButton("Sí", on_click=cerrar_dialogo)
                ]
            )
            self.page.overlay.append(dialogo)
            dialogo.open = True
            self.page.update()
        else:
            self.cambiar_ventana("main")
    
    def generar_nomina(self, e):
        """Genera la nómina"""
        nomina = self.nomina_field.value
        
        if not nomina or not nomina.strip().isdigit():
            self.mostrar_dialogo("Entrada Inválida", "Por favor, ingrese un número de nómina válido.", es_error=True)
            self.logger.warning("Intento de generar nómina sin número válido.")
            return
        
        self.logger.info(f"Iniciando generación de nómina N°{nomina}...")
        
        conn = AE.open_database()
        if conn is None:
            self.mostrar_dialogo("Error de Conexión", 
                               "No se pudo conectar a la base de datos.\nVerifica la configuración de la ruta.", 
                               es_error=True)
            self.logger.error("No se pudo conectar a la base de datos.")
            return
        
        self.logger.debug("Conexión a la base de datos establecida.")
        
        try:
            cur = conn.cursor()
            egreso_datos = AE.datos_egreso(cur, nomina, self.año_dropdown.value)
            
            if egreso_datos:
                self.logger.info(f"Datos obtenidos para nómina N°{nomina}-{self.año_dropdown.value}: {len(egreso_datos)} registros.")
                total = AE.llenar_nomina_xlsx(egreso_datos, nomina, self.año_dropdown.value, RutEmpresa=self.rut_field.value)
                total_formateado = f"${total:,.0f}".replace(",", ".")
                
                mensaje = (
                    f"Proceso Completado con Éxito\n\n"
                    f"Nómina N°: {nomina}\n"
                    f"Año: {self.año_dropdown.value}\n"
                    f"Total a Pagar: {total_formateado}\n"
                    f"Registros: {len(egreso_datos)}\n\n"
                    f"Ubicación: {self.output_path_field.value}"
                )
                self.mostrar_dialogo("Nómina Generada", mensaje, es_error=False)
            else:
                self.mostrar_dialogo("Sin Resultados", 
                                   f"No se encontró ningún egreso con el número: {nomina}", 
                                   es_error=True)
                self.logger.warning(f"No se encontraron datos para nómina N°{nomina}-{self.año_dropdown.value}.")
        
        except Exception as e:
            self.mostrar_dialogo("Error Crítico", f"Ocurrió un error inesperado:\n{e}", es_error=True)
            self.logger.error(f"Error durante la generación: {e}")
        
        finally:
            if conn:
                conn.close()
    
    def refrescar_años_db(self):
        """Actualiza el dropdown de años desde la base de datos"""
        self.logger.info("Intentando obtener años de la base de datos...")
        lista_años = AE.obtener_años_disponibles()
        
        if lista_años:
            self.año_dropdown.disabled = False
            self.año_dropdown.options = [ft.dropdown.Option(año) for año in lista_años]
            self.año_dropdown.value = lista_años[0]
            self.logger.info(f"Dropdown actualizado con {len(lista_años)} años.")
        else:
            self.año_dropdown.disabled = True
            self.año_dropdown.options = [ft.dropdown.Option("----")]
            self.año_dropdown.value = "----"
            self.logger.warning("No se pudo acceder a los años. Dropdown deshabilitado.")
        
        self.page.update()
    
    def mostrar_dialogo(self, titulo, mensaje, es_error=False):
        """Muestra un diálogo de información o error"""
        def cerrar(e):
            dialogo.open = False
            self.page.update()
        
        dialogo = ft.AlertDialog(
            modal=True,
            title=ft.Text(titulo),
            content=ft.Text(mensaje),
            actions=[ft.TextButton("OK", on_click=cerrar)]
        )
        self.page.overlay.append(dialogo)
        dialogo.open = True
        self.page.update()


def main(page: ft.Page):
    AutoExcelApp(page, debug=False)


if __name__ == "__main__":
    ft.app(target=main)