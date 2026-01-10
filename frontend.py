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
        self.page.window.center()

        
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
        
        brillo_sistema = self.page.platform_brightness
        self.es_oscuro = brillo_sistema == ft.Brightness.DARK or brillo_sistema is None
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
            text_align=ft.TextAlign.CENTER,
                input_filter=ft.InputFilter(
                allow=True, 
                regex_string=r"^[0-9]*$", 
                replacement_string=""
            )
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
        self.rut_field = ft.TextField(label="RUT Empresa",width=330,hint_text="12345678-K")                
        self.db_path_field = ft.TextField(label="Base de Datos", width=330, read_only=True)
        self.output_path_field = ft.TextField(label="Carpeta de Salida", width=330, read_only=True)
        self.tema_btn = ft.IconButton(
            icon=ft.Icons.LIGHT_MODE_OUTLINED if self.es_oscuro else ft.Icons.DARK_MODE_OUTLINED,
            tooltip="Cambiar tema",
            on_click=self.cambiar_tema_handler 
        )
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
                        self.tema_btn,
                    ]),
                    ft.Container(height=20),
                    ft.Row([
                        self.rut_field,
                        ft.Container( width=40
                        )
                    ], alignment=ft.MainAxisAlignment.CENTER),
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
        """Cambia entre vistas y asegura que los datos estén sincronizados"""
        self.cargar_configuracion() 
        if ventana == "config":
            self.main_view.visible = False
            self.config_view.visible = True
        else:
            self.main_view.visible = True
            self.config_view.visible = False
        
        self.page.update()

    def buscar_archivo(self, tipo):
        self.file_picker_tipo = tipo
        if tipo == "db":
            self.file_picker.pick_files(
                dialog_title="Seleccionar Base de Datos",
                allowed_extensions=["accdb", "mdb"]
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
        """Valida y guarda la configuración asegurando campos obligatorios"""
        
        # 1. Verificar que ningún campo esté vacío
        if not self.rut_field.value or not self.db_path_field.value or not self.output_path_field.value:
            self.logger.warning("Intento de guardado con campos vacíos")
            self.mostrar_dialogo(
                "Campos Incompletos", 
                "Todos los campos son obligatorios (RUT, Base de Datos y Carpeta de Salida).", 
                es_error=True
            )
            return

        # 2. Capitalizar y limpiar espacios en blanco del RUT
        rut_input = self.rut_field.value.strip().upper()
        
        # 3. Definir el patrón del RUT
        import re
        patron_rut = r"^\d+-[0-9K]$"

        # 4. Verificar si el RUT cumple con el formato
        if not re.match(patron_rut, rut_input):
            self.logger.warning(f"RUT con formato inválido: {rut_input}")
            self.mostrar_dialogo(
                "Error de Formato", 
                "El RUT debe ser: números, guion y dígito verificador.\nEjemplo: 12345678-K", 
                es_error=True
            )
            return

        # 5. Si todo es válido, preparamos los datos
        config_data = {
            "ruta_db": self.db_path_field.value,
            "ruta_output": self.output_path_field.value,
            "rut_empresa": rut_input,
            "dark_mode": self.es_oscuro  # Guardamos también la preferencia de tema
        }
        
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
            
            self.logger.info("Configuración guardada exitosamente.")
            self.mostrar_dialogo("Éxito", "Configuración guardada correctamente.")
            self.cambiar_ventana("main")
            self.refrescar_años_db()
        except Exception as e:
            self.logger.error(f"Error al guardar: {e}")
            self.mostrar_dialogo("Error", f"No se pudo guardar el archivo: {e}", True)

    def cargar_configuracion(self):
        """Carga la configuración desde config.json"""
        brillo_sistema = self.page.platform_brightness
        tema_sistema = brillo_sistema == ft.Brightness.DARK or brillo_sistema is None
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.db_path_field.value = data.get("ruta_db", "")
                    self.output_path_field.value = data.get("ruta_output", "")
                    self.rut_field.value = data.get("rut_empresa", "")
                    self.es_oscuro = data.get("dark_mode",tema_sistema)
                    
                    # Aplicar a la página
                    self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT
                    
                    # Cambiamos el icono del botón según la configuración
                    if hasattr(self, "tema_btn"): # Verificamos que el botón ya exista
                        self.tema_btn.icon = (
                            ft.Icons.LIGHT_MODE_OUTLINED if self.es_oscuro 
                            else ft.Icons.DARK_MODE_OUTLINED
                        )
                        self.tema_btn.tooltip = "Cambiar a modo claro" if self.es_oscuro else "Cambiar a modo oscuro"

                self.actualizar_estilos_ui() 
                self.logger.debug("Configuración cargada correctamente")
                self.page.update()
            except Exception as e:
                self.logger.error(f"Error al cargar configuración: {e}")
        else:
            self.db_path_field.value = ""
            self.output_path_field.value = ""
            self.rut_field.value = ""
            self.es_oscuro = tema_sistema
            self.logger.debug("No hay config.json, campos limpiados.")
        self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT

    def refrescar_años_db(self):
        """Obtiene años y muestra diálogo si hay error o está vacío"""
        try:
            self.logger.info("Refrescando años desde la base de datos...")
            lista_años = AE.obtener_años_disponibles()
            
            # Si la lista es None (error) o está vacía [] (no se encontraron datos o error manejado)
            if not lista_años:
                # Forzamos el salto al bloque except para mostrar el diálogo
                raise Exception("No se obtuvieron datos de la base de datos.")

            # Si llegamos aquí, es que hay datos
            self.año_dropdown.disabled = False
            self.año_dropdown.options = [ft.dropdown.Option(str(a)) for a in lista_años]
            self.año_dropdown.value = str(lista_años[0])
            self.logger.info(f"Se cargaron {len(lista_años)} años.")

        except Exception as e:
            self.logger.error(f"Fallo al cargar años: {e}")
            
            # Limpiar y deshabilitar el dropdown
            self.año_dropdown.disabled = True
            self.año_dropdown.options = [ft.dropdown.Option("----")]
            self.año_dropdown.value = "----"
            
            # DISPARAR DIÁLOGO
            self.mostrar_dialogo(
                "Error de Base de Datos",
                "No se pudo obtener información de la base de datos.\n\n"
                "Asegúrese de que:\n"
                "1. La ruta en Configuración sea correcta.\n"
                "2. El archivo .accdb no esté abierto o bloqueado.\n"
                "3. El driver de Microsoft Access esté instalado.",
                es_error=True
            )
        
        self.page.update()
        
        self.page.update()
    def generar_nomina(self, e):
        """Genera la nómina usando AutoExcel"""
        nomina = self.nomina_field.value
        año = self.año_dropdown.value
        
        # 1. Validaciones iniciales
        if not nomina or not nomina.isdigit():
            self.mostrar_dialogo("Error", "Ingrese una nómina válida (solo números)", True)
            return
        
        if not año or año == "----":
            self.mostrar_dialogo("Error", "Seleccione un año válido", True)
            return
        
        # 2. Cargar configuración para obtener rutas y RUT
        if not os.path.exists("config.json"):
            self.mostrar_dialogo("Error", "Configure la aplicación antes de generar nóminas", True)
            self.cambiar_ventana("config")
            return

        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
            rut_empresa = config.get("rut_empresa", "")
            ruta_salida = config.get("ruta_output", "") # Obtenemos la carpeta de destino
        
        try:
            self.logger.info(f"Iniciando generación de nómina {nomina} para el año {año}")
            conn = AE.open_database()
            if not conn:
                self.mostrar_dialogo("Error", "No se pudo conectar a la base de datos.", True)
                return
            
            try:
                cursor = conn.cursor()
                datos = AE.datos_egreso(cursor, int(nomina), int(año))
                
                if not datos:
                    self.mostrar_dialogo(
                        "Advertencia", 
                        f"No se encontraron registros para la nómina {nomina} en el año {año}", 
                        True
                    )
                    return
                
                # 3. Generar el archivo Excel
                total_pago = AE.llenar_nomina_xlsx(datos, nomina, año, rut_empresa)
                
                if total_pago is not None:
                    # Formateamos el mensaje de éxito con toda la información solicitada
                    mensaje_exito = (
                        f"Nómina: {nomina}\n"
                        f"Año: {año}\n"
                        f"Registros: {len(datos)}\n"
                        f"Total a Pagar: ${total_pago:,.0f}\n\n"
                        f"Ubicación: {ruta_salida}"
                    ).replace(",", ".") # Reemplazo para formato moneda chileno

                    self.mostrar_dialogo("Nómina Generada con Éxito", mensaje_exito)
                    self.logger.info(f"Nómina {nomina}-{año} generada en {ruta_salida}")
                else:
                    self.mostrar_dialogo("Error", "Ocurrió un error al generar el archivo Excel", True)
                    
            finally:
                conn.close()
                
        except Exception as ex:
            self.logger.error(f"Error al generar nómina: {ex}")
            self.mostrar_dialogo("Error", f"Error crítico:\n{str(ex)}", True)

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
            """Confirma la salida comparando el estado actual con el archivo JSON"""
            self.logger.info("Verificando cambios antes de salir...")
            
            # 1. Cargar la configuración actual desde el disco para comparar
            config_en_disco = {"ruta_db": "", "ruta_output": "", "rut_empresa": "", "dark_mode": ""}
            if os.path.exists("config.json"):
                try:
                    with open("config.json", "r", encoding="utf-8") as f:
                        config_en_disco = json.load(f)
                except Exception as e:
                    self.logger.error(f"Error al leer para comparar: {e}")

            # 2. Verificar si hay diferencias
            hay_cambios = (
                self.es_oscuro != config_en_disco.get("dark_mode", "") or
                self.db_path_field.value != config_en_disco.get("ruta_db", "") or
                self.output_path_field.value != config_en_disco.get("ruta_output", "") or
                self.rut_field.value != config_en_disco.get("rut_empresa", "")
            )

            if hay_cambios:
                self.logger.warning("Cambios detectados. Pidiendo confirmación.")
                
                def cerrar_y_salir(e):
                    dialogo.open = False
                    self.cargar_configuracion()
                    self.cambiar_ventana("main") # Salir al menú principal
                    self.page.update()

                def solo_cerrar(e):
                    dialogo.open = False
                    self.page.update()

                dialogo = ft.AlertDialog(
                    modal=True,
                    title=ft.Text("Cambios sin guardar"),
                    content=ft.Text("Has realizado cambios en la configuración. ¿Deseas salir sin guardar?"),
                    actions=[
                        ft.TextButton("Sí, salir", on_click=cerrar_y_salir),
                        ft.TextButton("No, volver", on_click=solo_cerrar),
                    ],
                    actions_alignment=ft.MainAxisAlignment.END,
                )
                
                self.page.overlay.append(dialogo)
                dialogo.open = True
                self.page.update()
            else:
                # Si no hay cambios, ir directo al main
                self.logger.info("Sin cambios detectados, regresando al menú.")
                self.cambiar_ventana("main")
                self.refrescar_años_db()

    def cambiar_tema_handler(self, e):
        """Alterna entre modo claro y oscuro y actualiza los estilos visuales"""
        # 1. Invertir el estado lógico
        self.es_oscuro = not self.es_oscuro
        
        # 2. Aplicar el modo a la página de Flet
        self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT
        
        # 3. Determinar el nuevo color de los bordes
        color_borde = ft.Colors.WHITE if self.es_oscuro else ft.Colors.BLACK
        
        # 4. Actualizar el icono del botón
        self.tema_btn.icon = (
            ft.Icons.LIGHT_MODE_OUTLINED if self.es_oscuro 
            else ft.Icons.DARK_MODE_OUTLINED
        )
        self.tema_btn.tooltip = "Modo Claro" if self.es_oscuro else "Modo Oscuro"

        # 5. Actualizar los bordes de todos los campos de texto activos
        lista_campos = [
            self.nomina_field, 
            self.rut_field, 
            self.db_path_field, 
            self.output_path_field
        ]
        
        for campo in lista_campos:
            campo.border_color = color_borde
            campo.focused_border_color = color_borde
        
        # 6. Refrescar la página para aplicar los cambios
        self.actualizar_estilos_ui()
        self.page.update()

    def actualizar_estilos_ui(self):
        """Actualiza bordes e iconos de todos los controles según self.es_oscuro"""
        color = ft.Colors.WHITE if self.es_oscuro else ft.Colors.BLACK
        
        # 1. Lista de todos los controles que tienen bordes
        controles = [
            self.nomina_field, 
            self.año_dropdown,   # <-- Agregamos el dropdown
            self.rut_field, 
            self.db_path_field, 
            self.output_path_field
        ]
        
        # 2. Aplicar color a cada uno
        for control in controles:
            control.border_color = color
            control.focused_border_color = color

        # 3. Actualizar el botón de tema (opcional pero recomendado aquí)
        if hasattr(self, "tema_btn"):
            self.tema_btn.icon = (
                ft.Icons.LIGHT_MODE_OUTLINED if self.es_oscuro 
                else ft.Icons.DARK_MODE_OUTLINED
            )


def main(page: ft.Page):
    AutoExcelApp(page, debug=False)

if __name__ == "__main__":
    ft.app(target=main)