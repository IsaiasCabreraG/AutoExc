import flet as ft
from utils import *
import datetime

def build_ui(self):
    self.page.locale_configuration = ft.LocaleConfiguration(
        supported_locales=[ft.Locale("es", "CL")],  # o "ES" para España
        current_locale=ft.Locale("es", "CL")
    )
    self.page.update()
    verificar_primer_usuario(self)
    log_debug(self, "UI", "→ Construyendo vista principal...")
    self.main_view = build_main_view(self)
    
    log_debug(self, "UI", "→ Construyendo vista de configuración...")
    self.config_view = build_config_view(self)

    log_debug(self, "UI", "→ Construyendo vista de login...")
    self.login_view = build_login_view(self)
    
    # Contenedor de vistas
    self.views_container = ft.Column(
        controls=[self.main_view, self.config_view, self.login_view],
        spacing=0,
        expand=True
    )
    
    self.page.add(self.views_container)
    
    if self.debug:
        log_debug(self, "UI", "→ Construyendo vista de debug...")
        self.debug_view = build_debug_view(self)
        self.page.add(self.debug_view)
        log_debug(self, "UI", "✓ Vista de debug agregada")

    cambiar_ventana(self, "login")
    cargar_configuracion(self)
    self.page.update()


def build_main_view(self):


    
    self._actualizando_fecha = False
    self.correos_seleccionados = []
    # --- Configuración del DatePicker ---
    self.date_picker = ft.DatePicker(
        on_change=lambda e: actualizar_fecha(e),
        on_dismiss=lambda _: bloquear_foco()
    )
    self.page.overlay.append(self.date_picker)


    def actualizar_fecha(e):
        if e.control.value:
            self._actualizando_fecha = True  # ← bloqueamos el on_focus
            self.fecha_field.value = e.control.value.strftime("%d/%m/%Y")
            self.fecha_field.error_text = None
            self.date_picker.open = False
            self.page.update()

    # --- Campos de Entrada ---
    self.nomina_field = ft.TextField(
        label="Nómina",
        hint_text="Ej: 12177",
        width=140,
        text_align=ft.TextAlign.CENTER,
        input_filter=ft.InputFilter(allow=True, regex_string=r"^[0-9]*$", replacement_string=""),
        on_submit=lambda _: generar_nomina(self)
    )
    
    self.año_dropdown = ft.Dropdown(
        label="Año", width=120, value="----",
        options=[ft.dropdown.Option("----")],
        disabled=True
    )

    estilo_texto_check = ft.TextStyle(size=13)

    # --- Campo de Fecha (Compacto) ---
    self.fecha_field = ft.TextField(
        label="Fecha de abono",
        read_only=True,
        width=200,
        height=35,
        text_size=12,
        prefix_icon=ft.Icons.CALENDAR_TODAY,
        on_focus=lambda _: abrir_datepicker(),
        on_click=lambda _: abrir_datepicker(), 
    )

    # --- MODAL Correos --- 
    self.lista_modal_check = ft.Column(spacing=0) # Aquí se cargarán los checks luego

    def cargar_checks_correos():
        self.lista_modal_check.controls.clear()
        
        if hasattr(self, 'correos_banco') and self.correos_banco:
            # Si la lista lógica está vacía y el reporte está activo, 
            # asumimos que es la primera carga y seleccionamos todos.
            # Pero si se vació por el reset de generar_nomina, se quedará vacía.
            if not self.correos_seleccionados:
                 # Solo autorellenar si es la primera vez absoluta que se activa el check
                 # Si prefieres que siempre aparezcan vacíos tras resetear, quita esta línea.
                 pass 

            for correo in self.correos_banco:
                self.lista_modal_check.controls.append(
                    ft.Checkbox(
                        label=correo,
                        # CRÍTICO: El valor visual DEPENDE de si está en la lista lógica
                        value=correo in self.correos_seleccionados, 
                        label_style=ft.TextStyle(size=14),
                        on_change=actualizar_seleccion_local
                    )
                )
        else:
            self.lista_modal_check.controls.append(
                ft.Text("No hay correos configurados.", color=ft.Colors.RED_400)
            )
        self.page.update()

    def actualizar_seleccion_local(e):
        correo = e.control.label
        if e.control.value:
            if correo not in self.correos_seleccionados:
                self.correos_seleccionados.append(correo)
        else:
            if correo in self.correos_seleccionados:
                self.correos_seleccionados.remove(correo)

    
    
    self.dialogo_reporte = ft.AlertDialog(
        title=ft.Text("Opciones de Reporte Banco"),
        content=ft.Container(
            content=self.lista_modal_check,
            width=300,
            height=200, # Ajustable según necesites
        ),
        actions=[
            ft.TextButton("Aceptar", on_click=lambda _: self.cerrar_modal_reporte())
        ],
    )
    self.page.overlay.append(self.dialogo_reporte)

    def cerrar_modal_reporte():
        self.dialogo_reporte.open = False
        self.page.update()
    
    self.cerrar_modal_reporte = cerrar_modal_reporte

    def bloquear_foco():
        self._actualizando_fecha = True
        self.page.update()

    def abrir_datepicker():
        if self._actualizando_fecha:
            self._actualizando_fecha = False  # ← reseteamos y salimos
            return
        self.date_picker.open = True
        self.page.update()

    # --- Checkboxes ---
    self.plantilla_nomina_check = ft.Checkbox(label="Plantilla IConstruye", value=True, label_style=estilo_texto_check, visual_density=ft.VisualDensity.COMPACT)
    self.plantilla_r2_check = ft.Checkbox(label="Plantilla Banco", value=True, label_style=estilo_texto_check, visual_density=ft.VisualDensity.COMPACT)

    def toggle_fecha_field(e):
        valor = e.control.value
        self.fecha_field.disabled = not valor
        self.fecha_field.opacity = 1.0 if valor else 0.5
        
        if not valor:
            self.fecha_field.value = ""
            self.fecha_field.error_text = None
            self.correos_seleccionados = [] 
            log_debug(self, "UI", "Selección de correos reseteada (Checkbox desactivado)")
        else:
            cargar_checks_correos()
            self.dialogo_reporte.open = True
            
        self.page.update()

    self.reporte_banco_check = ft.Checkbox(
        value=False,
        visual_density=ft.VisualDensity.COMPACT,
        on_change=toggle_fecha_field, 
    )

    etiqueta_reporte = ft.Row(
        controls=[
            ft.Text("Reporte Plantilla Banco", style=estilo_texto_check),
            ft.IconButton(
                icon=ft.Icons.EMAIL_OUTLINED,
                icon_size=18,
                icon_color=ft.Colors.BLUE_400,
                tooltip="Configurar destinatarios",
                on_click=lambda _: abrir_modal_desde_icono()
            )
        ],
        spacing=5,
        tight=True
    )
    def abrir_modal_desde_icono():
        if self.reporte_banco_check.value:
            cargar_checks_correos()
            self.dialogo_reporte.open = True
            self.page.update()
    
    # --- Contenedor de la Lista (Con indentación) ---
    self.lista_container = ft.Container(
        content=ft.Column(
            controls=[
                self.plantilla_nomina_check,
                self.plantilla_r2_check,
                # Aquí juntamos el checkbox con su etiqueta e icono
                ft.Row([self.reporte_banco_check, etiqueta_reporte], spacing=0),
                
                ft.Container(
                    content=self.fecha_field,
                    padding=ft.padding.only(left=30, top=0, bottom=0)
                ),
            ],
            spacing=2,
            horizontal_alignment=ft.CrossAxisAlignment.START,
        ),
        padding=ft.padding.symmetric(vertical=10, horizontal=20),
        border=ft.border.all(1, ft.Colors.OUTLINE_VARIANT),
        bgcolor=ft.Colors.BLUE_GREY_50 if not self.es_oscuro else ft.Colors.GREY_900, 
        width=280,
    )

    self.btn_generar = ft.ElevatedButton(
        text="Generar Nómina",
        width=200, height=45,
        on_click=lambda _: generar_nomina(self)
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
                                on_click=lambda _: cambiar_ventana(self, "config")
                            ),
                            expand=1,
                            alignment=ft.alignment.center_right,
                            padding=ft.padding.only(right=10)
                        )
                    ],
                ),
                
                ft.Container(height=40),
                ft.Text("Ingrese Nómina y Año:", size=14, weight=ft.FontWeight.W_500),
                ft.Row(
                    controls=[self.nomina_field, self.año_dropdown],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=10
                ),
                
                ft.Text("Plantillas a rellenar", size=14,),

                # LISTA CON FONDO DIFERENCIADO Y MÁS ANCHA
                ft.Row(
                    alignment=ft.MainAxisAlignment.CENTER,
                    controls=[
                        self.lista_container
                    ]
                ),
                
                self.btn_generar
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        visible=False
    )


def build_config_view(self):
    # 1. Componentes del Header
    self.tema_switch = ft.Switch(
        label="Modo Oscuro",
        value=self.es_oscuro, 
        active_color=ft.Colors.BLUE_400,
        on_change=lambda _: cambiar_tema_handler(self)
    )

    self.config_title = ft.Text(
        "Parámetros del Sistema", 
        size=28, 
        weight=ft.FontWeight.BOLD, 
        expand=True, 
        text_align="center"
    )

    # --- VISTA 1: PARÁMETROS ---
    self.rut_field = ft.TextField(label="RUT Empresa", width=330, hint_text="12345678-K")                
    self.db_path_field = ft.TextField(label="Base de Datos", width=330, read_only=True, on_click=lambda _: buscar_archivo(self, "db"))
    self.output_path_field = ft.TextField(label="Carpeta de Salida", width=330, read_only=True, on_click=lambda _: buscar_archivo(self, "folder"))

    self.vista_parametros = ft.Column([
        ft.Container(height=20),
        ft.Row([self.rut_field, ft.Container(width=40)], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([
            self.db_path_field,
            ft.IconButton(icon=ft.Icons.FILE_OPEN, on_click=lambda _: buscar_archivo(self, "db"))
        ], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([
            self.output_path_field,
            ft.IconButton(icon=ft.Icons.FOLDER_OPEN, on_click=lambda _: buscar_archivo(self, "folder"))
        ], alignment=ft.MainAxisAlignment.CENTER),
        ft.Container(height=20),
        ft.Row([
            ft.ElevatedButton("Guardar", width=150, on_click=lambda _: guardar_configuracion(self)),
            ft.OutlinedButton("Cancelar", width=150, on_click=lambda _: confirmar_salida(self))
        ], alignment=ft.MainAxisAlignment.CENTER) 
    ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER, visible=True, expand=True)

    # --- VISTA 2: USUARIOS ---
    self.nuevo_usuario_field = ft.TextField(label="Nuevo Usuario", width=180, prefix_icon=ft.Icons.PERSON_ADD_OUTLINED)
    self.nueva_pass_field = ft.TextField(label="Contraseña", width=200, password=True, can_reveal_password=True, prefix_icon=ft.Icons.KEY_OUTLINED)
    self.check_admin = ft.Checkbox(label="Admin", value=False)
    
    self.barra_creacion_usuarios = ft.Row([
        self.nuevo_usuario_field,
        self.nueva_pass_field,
        self.check_admin,
        ft.ElevatedButton(
            "Crear", 
            icon=ft.Icons.ADD, 
            # Fondo azul y texto blanco para resaltar
            style=ft.ButtonStyle(
                color=ft.Colors.WHITE,
                bgcolor=ft.Colors.BLUE_700,
            ),
            on_click=lambda _: agregar_usuario_handler(self) 
        ),
        ft.Divider()
    ], alignment=ft.MainAxisAlignment.CENTER)

    self.lista_usuarios_ui = ft.Column(expand=True, scroll=ft.ScrollMode.ADAPTIVE, spacing=0)

    self.vista_usuarios = ft.Column([
        self.barra_creacion_usuarios, 
        ft.Text("Usuarios Registrados:", size=16, weight=ft.FontWeight.W_500),
        ft.Container(content=self.lista_usuarios_ui, expand=True, padding=ft.padding.only(top=10))
    ], visible=False, expand=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER)


# --- VISTA 3: CORREOS ---
    # Campos para cuando NO hay información
    self.remitente_input = ft.TextField(label="Correo Remitente (Gmail)", width=250, hint_text="ejemplo@gmail.com")
    self.pass_remitente_input = ft.TextField(label="Contraseña de Aplicación", width=250, password=True, can_reveal_password=True)
    
    # Contenedor dinámico para la parte superior
    self.seccion_remitente_ui = ft.Column(horizontal_alignment=ft.CrossAxisAlignment.CENTER)

    self.nuevo_correo_field = ft.TextField(label="Nuevo Destinatario", width=250, prefix_icon=ft.Icons.EMAIL_OUTLINED,height=40,content_padding=ft.padding.only(left=5, right=5, top=5, bottom=5), # Espacio interno
    text_size=14, on_submit=lambda _: agregar_correo_handler(self))
    
    self.barra_creacion_correos = ft.Row([
        self.nuevo_correo_field,
        ft.ElevatedButton(
            "Agregar", 
            icon=ft.Icons.ADD, 
            style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_700),
            on_click=lambda _: agregar_correo_handler(self) 
        ),
    ], alignment=ft.MainAxisAlignment.CENTER)

    self.lista_correos_ui = ft.Column(expand=True, scroll=ft.ScrollMode.ADAPTIVE, spacing=0)

    self.vista_correos = ft.Column([
        ft.Text("Configuración del Remitente", size=18, weight=ft.FontWeight.BOLD),
        self.seccion_remitente_ui,
        ft.Divider(height=5),
        ft.Text("Destinatarios de Reporte Banco", size=18, weight=ft.FontWeight.BOLD),
        self.barra_creacion_correos, 
        ft.Container(content=self.lista_correos_ui, expand=True)
    ], visible=False, expand=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER)

    # --- ESTRUCTURA DE CONTENIDO ---
    self.content_area = ft.Container(content=self.vista_parametros, expand=True, padding=ft.padding.all(5))

    # --- NAVIGATION RAIL (SIN TRAILING) ---
    self.nav_rail = ft.NavigationRail(
        selected_index=0,
        label_type=ft.NavigationRailLabelType.ALL,
        min_width=100,
        group_alignment=-0.9,
        destinations=[
            ft.NavigationRailDestination(icon=ft.Icons.SETTINGS_OUTLINED, selected_icon=ft.Icons.SETTINGS, label="Parámetros"),
            ft.NavigationRailDestination(icon=ft.Icons.PERSON_OUTLINED, selected_icon=ft.Icons.PERSON, label="Usuarios"),
            ft.NavigationRailDestination(icon=ft.Icons.EMAIL_OUTLINED, selected_icon=ft.Icons.EMAIL, label="Correos")
        ],
        on_change=lambda e: cambiar_sub_config(self, e.control.selected_index),
    )

    # --- BOTÓN DE LOGOUT (Superpuesto con Stack) ---
    self.logout_button = ft.IconButton(
        icon=ft.Icons.LOGOUT_ROUNDED,
        icon_color=ft.Colors.RED_400,
        icon_size=28,
        tooltip="Cerrar Sesión",
        on_click=lambda _: cerrar_sesion_handler(self)
    )

    # Stack para superponer el botón sobre el NavigationRail
    nav_with_logout = ft.Stack([
        self.nav_rail,
        ft.Container(
            content=self.logout_button,
            left=0,
            bottom=0,
            width=100,  # Mismo ancho que min_width del NavigationRail
            alignment=ft.alignment.center,
            padding=ft.padding.only(bottom=10)
        )
    ])

    # --- RENDERIZADO FINAL ---
    return ft.Container(
        content=ft.Column([
            # Header
            ft.Row([
                ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda _: confirmar_salida(self)),
                self.config_title,
                self.tema_switch,
            ]),
            ft.Divider(height=1),
            ft.Row([
                nav_with_logout,
                ft.VerticalDivider(width=1),
                self.content_area
            ], expand=True) 
        ], expand=True), 
        expand=True, 
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
        on_submit=lambda _: intentar_login(self) # Enter para entrar
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
                    on_click=lambda _: intentar_login(self)
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            alignment=ft.MainAxisAlignment.CENTER,
        ),
        expand=True,
        padding=ft.padding.only(top=50),
        visible=False 
    )

    
def verificar_primer_usuario(self):
    """Comprueba si existen usuarios. Si no, fuerza la creación del primero."""
    usuarios = seguridad.listar_usuarios()
    
    if not usuarios:
        log_debug(self, "AUTH", "No se detectaron usuarios. Solicitando creación de administrador inicial.")
        
        # 1. Definimos la función de creación primero para poder referenciarla
        def crear_primer_admin(e):
            u = user_input.value.strip()
            p = pass_input.value.strip()
            
            if u and p:
                exito, msg = seguridad.registrar_usuario(u, p, rol="admin")
                if exito:
                    dialogo_inicial.open = False
                    self.page.update()
                    mostrar_dialogo(self, "Configuración Inicial", f"Administrador '{u}' creado con éxito.")
                else:
                    user_input.error_text = msg
                    self.page.update()
            else:
                user_input.error_text = "Complete todos los campos"
                self.page.update()

        # 2. Agregamos on_submit a los inputs
        user_input = ft.TextField(
            label="Usuario Administrador", 
            width=280, 
            autofocus=True,
            on_submit=crear_primer_admin # <--- ENTER AQUÍ
        )
        pass_input = ft.TextField(
            label="Contraseña", 
            width=280, 
            password=True, 
            can_reveal_password=True,
            on_submit=crear_primer_admin # <--- ENTER AQUÍ
        )

        dialogo_inicial = ft.AlertDialog(
            modal=True,
            title=ft.Text("Primer Inicio: Crear Administrador"),
            content=ft.Column([
                ft.Text("No se encontraron usuarios en el sistema.\nPor favor, cree la cuenta de administrador inicial."),
                user_input,
                pass_input
            ], tight=True, spacing=20),
            actions=[
                ft.ElevatedButton("Crear Administrador", on_click=crear_primer_admin)
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER
        )

        self.page.overlay.append(dialogo_inicial)
        dialogo_inicial.open = True
        self.page.update()
    
    log_debug(self, "AUTH", "Continuando con el flujo normal.")