import flet as ft
import os
import sys
import json
import logging
import AutoExcel as AE
import seguridad

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
        
        # Logger configurado PRIMERO (antes de cualquier log_debug)
        self.logger = logging.getLogger("AutoExcel")
        self.logger.setLevel(logging.DEBUG)
        
        # Formato para los logs
        formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s', 
                                     datefmt='%Y-%m-%d %H:%M:%S')
        
        # Configuración de ventana para 0.28.3
        self.page.window.width = 700
        self.page.window.height = 400 if not self.debug else 650
        self.page.window.resizable = False
        self.page.window.center()

        
        ruta_icono = AE.obtener_ruta_recurso("static/logo.ico")
        if os.path.exists(ruta_icono):
            self.page.window.icon = ruta_icono
            self.log_debug("INIT", f"✓ Icono cargado desde: {ruta_icono}")
        else:
            self.log_debug("INIT", f"⚠️ Icono no encontrado en: {ruta_icono}")
        
        # FilePicker con on_result en el constructor
        self.file_picker_tipo = None
        self.file_picker = ft.FilePicker(on_result=self._file_picker_result)
        self.page.overlay.append(self.file_picker)
        self.log_debug("INIT", "✓ FilePicker inicializado")

        self.lista_usuarios = ft.ListView(expand=True, spacing=5)
        self.usuario_seleccionado = None

        
        brillo_sistema = self.page.platform_brightness
        self.es_oscuro = brillo_sistema == ft.Brightness.DARK or brillo_sistema is None
        self.log_debug("INIT", f"Tema del sistema detectado: {'OSCURO' if self.es_oscuro else 'CLARO'}")
        
        self.build_ui()
        # La aplicación ahora inicia mostrando SOLO el login
        self.mostrar_login() 
        self.page.update()
        self.log_debug("INIT", "✓ Interfaz de usuario construida")
        
        # Configurar handlers según el modo
        if self.debug:
            # Modo debug: mostrar en pantalla
            text_handler = TextboxHandler(self.log_textbox, self.page)
            text_handler.setFormatter(formatter)
            self.logger.addHandler(text_handler)
            self.log_debug("INIT", "✓ Handler de debug activado (pantalla)")
        else:
            # Modo normal: guardar en archivo
            try:
                file_handler = logging.FileHandler('autoexcel.log', encoding='utf-8')
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
                self.log_debug("INIT", "✓ Handler de archivo activado (autoexcel.log)")
            except Exception as e:
                # Si no se puede crear el archivo, al menos registramos el error en consola
                console_handler = logging.StreamHandler()
                console_handler.setFormatter(formatter)
                self.logger.addHandler(console_handler)
                self.logger.error(f"No se pudo crear archivo de log: {e}")
            
        

        self.log_debug("INIT", "✓✓✓ Aplicación inicializada correctamente")
    
    def log_debug(self, categoria, mensaje):
        """Helper para logs más estructurados"""
        self.logger.debug(f"[{categoria}] {mensaje}")
    
    def build_ui(self):
        self.log_debug("UI", "→ Construyendo vista principal...")
        self.main_view = self.build_main_view()
        
        self.log_debug("UI", "→ Construyendo vista de configuración...")
        self.config_view = self.build_config_view()

        self.log_debug("UI", "→ Construyendo vista de login...")
        self.login_view = self.build_login_view()
        
        # Contenedor de vistas
        self.views_container = ft.Column(
            controls=[self.main_view, self.config_view, self.login_view],
            spacing=0
        )
        
        self.page.add(self.views_container)
        
        if self.debug:
            self.log_debug("UI", "→ Construyendo vista de debug...")
            self.debug_view = self.build_debug_view()
            self.page.add(self.debug_view)
            self.log_debug("UI", "✓ Vista de debug agregada")
    
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
            visible=False
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
            max_lines=10,
            text_size=11,
            expand=False
        )
        return ft.Container(
            content=ft.Column([
                ft.Text("Consola Debug", weight=ft.FontWeight.BOLD),
                ft.Container(
                    content=self.log_textbox,
                    height=200,  # Altura fija del contenedor
                )
            ]),
            padding=10
        )
    
    def build_login_view(self):
        self.user_field = ft.TextField(label="Usuario", width=280, prefix_icon=ft.Icons.PERSON)
        self.pass_field = ft.TextField(
            label="Contraseña", 
            width=280, 
            password=True, 
            can_reveal_password=True,
            prefix_icon=ft.Icons.LOCK,
            on_submit=lambda _: self.intentar_login() # Enter para entrar
        )
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Icon(ft.Icons.LOCK_PERSON_OUTLINED, size=80, color=ft.Colors.BLUE),
                    ft.Text("Acceso al Sistema", size=24, weight=ft.FontWeight.BOLD),
                    self.user_field,
                    self.pass_field,
                    ft.Container(height=10),
                    ft.ElevatedButton(
                        "Iniciar Sesión", 
                        width=280, 
                        height=45,
                        on_click=lambda _: self.intentar_login()
                    ),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER,
            ),
            expand=True,
            padding=ft.padding.only(top=50),
            visible=True 
        )

    def intentar_login(self):
        user = self.user_field.value
        pw = self.pass_field.value
        
        self.log_debug("AUTH", f"Intento de login para usuario: {user}")
        
        if seguridad.validar_credenciales(user, pw):
            self.log_debug("AUTH", "✓ Acceso concedido")
            self.refrescar_años_db()
            self.cargar_configuracion()
            self.login_view.visible = False
            self.main_view.visible = True
            self.page.update()
        else:
            self.log_debug("AUTH", "✗ Acceso denegado: credenciales incorrectas")
            self.mostrar_dialogo("Error de Acceso", "Usuario o contraseña incorrectos.", es_error=True)

    def mostrar_login(self):
        self.login_view.visible = True
        self.main_view.visible = False
        self.config_view.visible = False

    # --- LÓGICA DE NEGOCIO ---
    def cambiar_ventana(self, ventana):
        """Cambia entre vistas y asegura que los datos estén sincronizados"""
        self.log_debug("NAV", f"→ Cambiando a ventana: {ventana.upper()}")
        self.cargar_configuracion() 
        
        if ventana == "config":
            self.main_view.visible = False
            self.config_view.visible = True
            self.log_debug("NAV", "✓ Vista de CONFIGURACIÓN activada")
        else:
            self.main_view.visible = True
            self.config_view.visible = False
            self.log_debug("NAV", "✓ Vista PRINCIPAL activada")
        
        self.page.update()

    def buscar_archivo(self, tipo):
        self.file_picker_tipo = tipo
        self.log_debug("FILEPICKER", f"→ Abriendo selector de tipo: {tipo.upper()}")
        
        if tipo == "db":
            self.file_picker.pick_files(
                dialog_title="Seleccionar Base de Datos",
                allowed_extensions=["accdb", "mdb"]
            )
            self.log_debug("FILEPICKER", "Esperando selección de archivo .accdb/.mdb")
        else:
            self.file_picker.get_directory_path(
                dialog_title="Seleccionar Carpeta de Salida"
            )
            self.log_debug("FILEPICKER", "Esperando selección de carpeta")

    def _file_picker_result(self, e: ft.FilePickerResultEvent):
        if self.file_picker_tipo == "db" and e.files:
            ruta = e.files[0].path
            self.db_path_field.value = ruta
            self.log_debug("FILEPICKER", f"✓ Base de datos seleccionada: {ruta}")
        elif self.file_picker_tipo == "folder" and e.path:
            self.output_path_field.value = e.path
            self.log_debug("FILEPICKER", f"✓ Carpeta de salida seleccionada: {e.path}")
        else:
            self.log_debug("FILEPICKER", "✗ Selección cancelada por el usuario")
        
        self.page.update()

    def guardar_configuracion(self):
        """Valida y guarda la configuración asegurando campos obligatorios"""
        self.log_debug("CONFIG", "═══ INICIANDO GUARDADO DE CONFIGURACIÓN ═══")
        
        # 1. Verificar que ningún campo esté vacío
        if not self.rut_field.value or not self.db_path_field.value or not self.output_path_field.value:
            self.log_debug("CONFIG", "✗ Validación fallida: Campos vacíos detectados")
            self.log_debug("CONFIG", f"  • RUT: {'✓' if self.rut_field.value else '✗ VACÍO'}")
            self.log_debug("CONFIG", f"  • DB: {'✓' if self.db_path_field.value else '✗ VACÍO'}")
            self.log_debug("CONFIG", f"  • Salida: {'✓' if self.output_path_field.value else '✗ VACÍO'}")
            
            self.mostrar_dialogo(
                "Campos Incompletos", 
                "Todos los campos son obligatorios (RUT, Base de Datos y Carpeta de Salida).", 
                es_error=True
            )
            return

        # 2. Capitalizar y limpiar espacios en blanco del RUT
        rut_input = self.rut_field.value.strip().upper()
        self.log_debug("CONFIG", f"RUT procesado: '{self.rut_field.value}' → '{rut_input}'")
        
        # 3. Definir el patrón del RUT
        import re
        patron_rut = r"^\d+-[0-9K]$"

        # 4. Verificar si el RUT cumple con el formato
        if not re.match(patron_rut, rut_input):
            self.log_debug("CONFIG", f"✗ RUT inválido: '{rut_input}' no cumple patrón {patron_rut}")
            self.mostrar_dialogo(
                "Error de Formato", 
                "El RUT debe ser: números, guion y dígito verificador.\nEjemplo: 12345678-K", 
                es_error=True
            )
            return

        self.log_debug("CONFIG", f"✓ RUT válido: {rut_input}")

        # 5. Si todo es válido, preparamos los datos
        config_data = {
            "ruta_db": self.db_path_field.value,
            "ruta_output": self.output_path_field.value,
            "rut_empresa": rut_input,
            "dark_mode": self.es_oscuro
        }
        
        self.log_debug("CONFIG", "Datos a guardar:")
        self.log_debug("CONFIG", f"  • DB: {config_data['ruta_db']}")
        self.log_debug("CONFIG", f"  • Output: {config_data['ruta_output']}")
        self.log_debug("CONFIG", f"  • RUT: {config_data['rut_empresa']}")
        self.log_debug("CONFIG", f"  • Tema oscuro: {config_data['dark_mode']}")
        
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
            
            self.log_debug("CONFIG", "✓ Archivo config.json guardado exitosamente")
            self.mostrar_dialogo("Éxito", "Configuración guardada correctamente.")
            self.cambiar_ventana("main")
            self.refrescar_años_db()
            self.log_debug("CONFIG", "═══ GUARDADO COMPLETADO ═══")
        except Exception as e:
            self.log_debug("CONFIG", f"✗✗✗ ERROR CRÍTICO al guardar: {type(e).__name__}: {e}")
            self.mostrar_dialogo("Error", f"No se pudo guardar el archivo: {e}", True)

    def cargar_configuracion(self):
        """Carga la configuración desde config.json"""
        self.log_debug("CONFIG", "─── Cargando configuración desde disco ───")
        
        brillo_sistema = self.page.platform_brightness
        tema_sistema = brillo_sistema == ft.Brightness.DARK or brillo_sistema is None
        
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.db_path_field.value = data.get("ruta_db", "")
                    self.output_path_field.value = data.get("ruta_output", "")
                    self.rut_field.value = data.get("rut_empresa", "")
                    self.es_oscuro = data.get("dark_mode", tema_sistema)
                    
                    self.log_debug("CONFIG", "✓ config.json leído correctamente:")
                    self.log_debug("CONFIG", f"  • DB: {self.db_path_field.value}")
                    self.log_debug("CONFIG", f"  • Output: {self.output_path_field.value}")
                    self.log_debug("CONFIG", f"  • RUT: {self.rut_field.value}")
                    self.log_debug("CONFIG", f"  • Tema: {'OSCURO' if self.es_oscuro else 'CLARO'}")
                    
                    # Aplicar a la página
                    self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT
                    
                    # Cambiamos el icono del botón según la configuración
                    if hasattr(self, "tema_btn"):
                        self.tema_btn.icon = (
                            ft.Icons.LIGHT_MODE_OUTLINED if self.es_oscuro 
                            else ft.Icons.DARK_MODE_OUTLINED
                        )
                        self.tema_btn.tooltip = "Cambiar a modo claro" if self.es_oscuro else "Cambiar a modo oscuro"

                self.actualizar_estilos_ui() 
                self.page.update()
            except json.JSONDecodeError as e:
                self.log_debug("CONFIG", f"✗ Error JSON: {e}")
            except Exception as e:
                self.log_debug("CONFIG", f"✗ Error al cargar: {type(e).__name__}: {e}")
        else:
            self.log_debug("CONFIG", "⚠️ config.json no existe, usando valores por defecto")
            self.db_path_field.value = ""
            self.output_path_field.value = ""
            self.rut_field.value = ""
            self.es_oscuro = tema_sistema
        
        self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT

    def refrescar_años_db(self):
        """Obtiene años y muestra diálogo si hay error o está vacío"""
        self.log_debug("DB", "═══ REFRESCANDO AÑOS DISPONIBLES ═══")
        
        try:
            lista_años = AE.obtener_años_disponibles()
            self.log_debug("DB", f"Años obtenidos: {lista_años}")
            
            # Si la lista es None (error) o está vacía [] (no se encontraron datos o error manejado)
            if not lista_años:
                self.log_debug("DB", "✗ Lista vacía o None recibida")
                raise Exception("No se obtuvieron datos de la base de datos.")

            # Si llegamos aquí, es que hay datos
            self.año_dropdown.disabled = False
            self.año_dropdown.options = [ft.dropdown.Option(str(a)) for a in lista_años]
            self.año_dropdown.value = str(lista_años[0])
            self.log_debug("DB", f"✓ {len(lista_años)} años cargados exitosamente")
            self.log_debug("DB", f"  Años disponibles: {', '.join(map(str, lista_años))}")
            self.log_debug("DB", f"  Año por defecto: {lista_años[0]}")

        except Exception as e:
            self.log_debug("DB", f"✗✗✗ ERROR al cargar años: {type(e).__name__}: {e}")
            
            # Limpiar y deshabilitar el dropdown
            self.año_dropdown.disabled = True
            self.año_dropdown.options = [ft.dropdown.Option("----")]
            self.año_dropdown.value = "----"
            self.log_debug("DB", "Dropdown deshabilitado y reseteado a '----'")
            
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
        self.log_debug("DB", "═══ REFRESCO FINALIZADO ═══")
        
    def generar_nomina(self, e):
        """Genera la nómina usando AutoExcel"""
        nomina = self.nomina_field.value
        año = self.año_dropdown.value
        
        self.log_debug("NOMINA", "╔═══════════════════════════════════════╗")
        self.log_debug("NOMINA", "║   INICIANDO GENERACIÓN DE NÓMINA    ║")
        self.log_debug("NOMINA", "╚═══════════════════════════════════════╝")
        self.log_debug("NOMINA", f"Input recibido → Nómina: '{nomina}', Año: '{año}'")
        
        # 1. Validaciones iniciales
        if not nomina or not nomina.isdigit():
            self.log_debug("NOMINA", f"✗ Validación fallida: Nómina inválida '{nomina}'")
            self.mostrar_dialogo("Error", "Ingrese una nómina válida (solo números)", True)
            return
        
        if not año or año == "----":
            self.log_debug("NOMINA", f"✗ Validación fallida: Año no seleccionado '{año}'")
            self.mostrar_dialogo("Error", "Seleccione un año válido", True)
            return
        
        self.log_debug("NOMINA", "✓ Validaciones de entrada pasadas")
        
        # 2. Cargar configuración para obtener rutas y RUT
        if not os.path.exists("config.json"):
            self.log_debug("NOMINA", "✗ config.json no existe, redirigiendo a configuración")
            self.mostrar_dialogo("Error", "Configure la aplicación antes de generar nóminas", True)
            self.cambiar_ventana("config")
            return

        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
            rut_empresa = config.get("rut_empresa", "")
            ruta_salida = config.get("ruta_output", "")
        
        self.log_debug("NOMINA", f"Configuración cargada:")
        self.log_debug("NOMINA", f"  • RUT Empresa: {rut_empresa}")
        self.log_debug("NOMINA", f"  • Ruta Salida: {ruta_salida}")
        
        try:
            self.log_debug("NOMINA", "→ Abriendo conexión a base de datos...")
            conn = AE.open_database()
            
            if not conn:
                self.log_debug("NOMINA", "✗✗✗ ERROR: No se pudo establecer conexión con la BD")
                self.mostrar_dialogo("Error", "No se pudo conectar a la base de datos.", True)
                return
            
            self.log_debug("NOMINA", "✓ Conexión establecida")
            
            try:
                cursor = conn.cursor()
                self.log_debug("NOMINA", f"→ Consultando datos para nómina {nomina}, año {año}...")
                
                datos = AE.datos_egreso(cursor, int(nomina), int(año))
                
                if not datos:
                    self.log_debug("NOMINA", f"⚠️ No se encontraron registros para nómina {nomina} en año {año}")
                    self.mostrar_dialogo(
                        "Advertencia", 
                        f"No se encontraron registros para la nómina {nomina} en el año {año}", 
                        True
                    )
                    return
                
                self.log_debug("NOMINA", f"✓ {len(datos)} registros obtenidos de la BD")
                
                # 3. Generar el archivo Excel
                self.log_debug("NOMINA", "→ Generando archivo Excel...")
                total_pago = AE.llenar_nomina_xlsx(datos, nomina, año, rut_empresa)
                
                if total_pago is not None:
                    self.log_debug("NOMINA", f"✓✓✓ NÓMINA GENERADA EXITOSAMENTE")
                    self.log_debug("NOMINA", f"  • Registros procesados: {len(datos)}")
                    self.log_debug("NOMINA", f"  • Total a pagar: ${total_pago:,.0f}")
                    self.log_debug("NOMINA", f"  • Ubicación: {ruta_salida}")
                    
                    mensaje_exito = (
                        f"Nómina: {nomina}\n"
                        f"Año: {año}\n"
                        f"Registros: {len(datos)}\n"
                        f"Total a Pagar: ${total_pago:,.0f}\n\n"
                        f"Ubicación: {ruta_salida}"
                    ).replace(",", ".")

                    self.mostrar_dialogo("Nómina Generada con Éxito", mensaje_exito)
                else:
                    self.log_debug("NOMINA", "✗✗✗ ERROR: llenar_nomina_xlsx retornó None")
                    self.mostrar_dialogo("Error", "Ocurrió un error al generar el archivo Excel", True)
                    
            finally:
                conn.close()
                self.log_debug("NOMINA", "Conexión a BD cerrada")
                
        except Exception as ex:
            self.log_debug("NOMINA", f"✗✗✗ EXCEPCIÓN CRÍTICA: {type(ex).__name__}: {ex}")
            import traceback
            self.log_debug("NOMINA", f"Traceback:\n{traceback.format_exc()}")
            self.mostrar_dialogo("Error", f"Error crítico:\n{str(ex)}", True)
        
        self.log_debug("NOMINA", "╚═══════════════════════════════════════╝")

    def mostrar_dialogo(self, titulo, mensaje, es_error=False):
        """Muestra un diálogo modal con un mensaje"""
        tipo = "ERROR" if es_error else "INFO"
        self.log_debug("DIALOG", f"[{tipo}] {titulo}: {mensaje[:50]}...")
        
        def cerrar(e):
            dialogo.open = False
            self.page.update()
            self.log_debug("DIALOG", f"Diálogo '{titulo}' cerrado")
        
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
        self.log_debug("NAV", "─── Verificando cambios antes de salir ───")
        
        # 1. Cargar la configuración actual desde el disco para comparar
        config_en_disco = {"ruta_db": "", "ruta_output": "", "rut_empresa": "", "dark_mode": ""}
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    config_en_disco = json.load(f)
            except Exception as e:
                self.log_debug("NAV", f"✗ Error al leer config para comparar: {e}")

        # 2. Verificar si hay diferencias
        cambios = {
            "tema": self.es_oscuro != config_en_disco.get("dark_mode", ""),
            "db": self.db_path_field.value != config_en_disco.get("ruta_db", ""),
            "output": self.output_path_field.value != config_en_disco.get("ruta_output", ""),
            "rut": self.rut_field.value != config_en_disco.get("rut_empresa", "")
        }
        
        hay_cambios = any(cambios.values())
        
        if hay_cambios:
            self.log_debug("NAV", "⚠️ CAMBIOS DETECTADOS:")
            for campo, cambio in cambios.items():
                if cambio:
                    self.log_debug("NAV", f"  • {campo.upper()} modificado")
            
            def cerrar_y_salir(e):
                dialogo.open = False
                self.log_debug("NAV", "Usuario confirmó salir SIN GUARDAR")
                self.cargar_configuracion()
                self.cambiar_ventana("main")
                self.page.update()

            def solo_cerrar(e):
                dialogo.open = False
                self.log_debug("NAV", "Usuario canceló la salida, permaneciendo en configuración")
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
            self.log_debug("NAV", "✓ Sin cambios detectados, regresando al menú principal")
            self.cambiar_ventana("main")
            self.refrescar_años_db()

    def cambiar_tema_handler(self, e):
        """Alterna entre modo claro y oscuro y actualiza los estilos visuales"""
        tema_anterior = "OSCURO" if self.es_oscuro else "CLARO"
        
        # 1. Invertir el estado lógico
        self.es_oscuro = not self.es_oscuro
        
        tema_nuevo = "OSCURO" if self.es_oscuro else "CLARO"
        self.log_debug("TEMA", f"Cambio de tema: {tema_anterior} → {tema_nuevo}")
        
        # 2. Aplicar el modo a la página de Flet
        self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT
        
        # 3. Determinar el nuevo color de los bordes
        color_borde = ft.Colors.WHITE if self.es_oscuro else ft.Colors.BLACK
        self.log_debug("TEMA", f"Color de borde aplicado: {color_borde}")
        
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
        self.log_debug("TEMA", "✓ Tema aplicado y UI actualizada")

    def actualizar_estilos_ui(self):
        """Actualiza bordes e iconos de todos los controles según self.es_oscuro"""
        self.log_debug("UI", "→ Actualizando estilos de la interfaz...")
        
        color = ft.Colors.WHITE if self.es_oscuro else ft.Colors.BLACK
        
        # 1. Lista de todos los controles que tienen bordes
        controles = [
            self.nomina_field, 
            self.año_dropdown,
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
        
        self.log_debug("UI", f"✓ Estilos actualizados para tema {'OSCURO' if self.es_oscuro else 'CLARO'}")



def main(page: ft.Page):
    AutoExcelApp(page, debug=False)

if __name__ == "__main__":
    ft.app(target=main)