import seguridad 
import flet as ft
import os
import json
import AutoExcel as AE
import re
from dotenv import load_dotenv



def log_debug(self, categoria, mensaje):
        """Helper para logs más estructurados"""
        self.logger.debug(f"[{categoria}] {mensaje}")
    

def intentar_login(self):
    user = self.user_field.value
    pw = self.pass_field.value
    
    log_debug(self, "AUTH", f"Intento de login para usuario: {user}")
    
    if seguridad.validar_credenciales(user, pw):
        log_debug(self, "AUTH", "✓ Acceso concedido")
        self.usuario_logueado = {
        "user": user,
        "rol": seguridad.obtener_rol(user) # Asegúrate de tener esta función en seguridad.py
    }
        cambiar_ventana(self, "main")
    else:
        log_debug(self, "AUTH", "✗ Acceso denegado: credenciales incorrectas")
        self.pass_field.value = "" 
        self.user_field.value = ""  

        mostrar_dialogo(self, "Error de Acceso", "Usuario o contraseña incorrectos.", es_error=True)


# --- LÓGICA DE NEGOCIO ---
def cambiar_ventana(self, ventana):
    """Cambia entre vistas y asegura que los datos estén sincronizados"""
    log_debug(self, "NAV", f"→ Cambiando a ventana: {ventana.upper()}")
    cargar_configuracion(self)
    
    # Resetear visibilidad de todas
    self.config_view.visible = False
    self.main_view.visible = False
    self.login_view.visible = False

    if ventana == "config":
        self.config_view.visible = True
        self.page.update() 
        if hasattr(self, "nav_rail"):
            cambiar_sub_config(self, 0)
        log_debug(self, "NAV", "✓ Vista de CONFIGURACIÓN activada")
        
    elif ventana == "main":
        refrescar_años_db(self)
        self.main_view.visible = True
        log_debug(self, "NAV", "✓ Vista PRINCIPAL activada")
        
    elif ventana == "login":
        self.login_view.visible = True
        log_debug(self, "NAV", "✓ Vista de LOGIN activada")
    

    self.page.update()

def buscar_archivo(self, tipo):
    self.file_picker_tipo = tipo
    log_debug(self, "FILEPICKER", f"→ Abriendo selector de tipo: {tipo.upper()}")
    
    if tipo == "db":
        self.file_picker.pick_files(
            dialog_title="Seleccionar Base de Datos",
            allowed_extensions=["accdb", "mdb"])
        log_debug(self, "FILEPICKER", "Esperando selección de archivo .accdb/.mdb")
    elif tipo == "folder":
        self.file_picker.get_directory_path(
            dialog_title="Seleccionar Carpeta de Salida")
        log_debug(self, "FILEPICKER", "Esperando selección de carpeta")

def guardar_configuracion(self):
    """Valida y guarda la configuración asegurando campos obligatorios"""
    log_debug(self, "CONFIG", "═══ INICIANDO GUARDADO DE CONFIGURACIÓN ═══")
    
    if not self.rut_field.value or not self.db_path_field.value or not self.output_path_field.value:
        log_debug(self, "CONFIG", "✗ Validación fallida: Campos vacíos detectados")
        log_debug(self, "CONFIG", f"  • RUT: {'✓' if self.rut_field.value else '✗ VACÍO'}")
        log_debug(self, "CONFIG", f"  • DB: {'✓' if self.db_path_field.value else '✗ VACÍO'}")
        log_debug(self, "CONFIG", f"  • Salida: {'✓' if self.output_path_field.value else '✗ VACÍO'}")
        
        mostrar_dialogo(self, 
            "Campos Incompletos", 
            "Todos los campos son obligatorios (RUT, Base de Datos y Carpeta de Salida).", 
            es_error=True
        )
        return

    # Validar formato del RUT
    rut_input = self.rut_field.value.strip().upper()
    log_debug(self, "CONFIG", f"RUT procesado: '{self.rut_field.value}' → '{rut_input}'")
    patron_rut = r"^\d+-[0-9K]$"
    if not re.match(patron_rut, rut_input):
        log_debug(self, "CONFIG", f"✗ RUT inválido: '{rut_input}' no cumple patrón {patron_rut}")
        mostrar_dialogo(self, 
            "Error de Formato", 
            "El RUT debe ser: números, guion y dígito verificador.\nEjemplo: 12345678-K", 
            es_error=True
        )
        return
    log_debug(self, "CONFIG", f"✓ RUT válido: {rut_input}")

    # Todo Valido
    config_data = {
        "ruta_db": self.db_path_field.value,
        "ruta_output": self.output_path_field.value,
        "rut_empresa": rut_input,
        "dark_mode": self.es_oscuro
    }
    
    log_debug(self, "CONFIG", "Datos a guardar:")
    log_debug(self, "CONFIG", f"  • DB: {config_data['ruta_db']}")
    log_debug(self, "CONFIG", f"  • Output: {config_data['ruta_output']}")
    log_debug(self, "CONFIG", f"  • RUT: {config_data['rut_empresa']}")
    log_debug(self, "CONFIG", f"  • Tema oscuro: {config_data['dark_mode']}")
    
    try:
        with open("internal/config.json", "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)
        
        log_debug(self, "CONFIG", "✓ Archivo internal/config.json guardado exitosamente")
        mostrar_dialogo(self, "Éxito", "Configuración guardada correctamente.")
        cambiar_ventana(self, "main")
        log_debug(self, "CONFIG", "═══ GUARDADO COMPLETADO ═══")
    except Exception as e:
        log_debug(self, "CONFIG", f"✗✗✗ ERROR CRÍTICO al guardar: {type(e).__name__}: {e}")
        mostrar_dialogo(self, "Error", f"No se pudo guardar el archivo: {e}", True)

def cargar_configuracion(self):
    """Carga la configuración o crea una inicial basada en el sistema si no existe."""
    log_debug(self, "CONFIG", "─── Cargando configuración desde disco ───")
    
    # 1. Detectar el estado del sistema operativo inmediatamente
    brillo_sistema = self.page.platform_brightness
    tema_sistema = brillo_sistema == ft.Brightness.DARK or brillo_sistema is None
    self.correo_remitente, self.correo_password = seguridad.obtener_credenciales_email()
    actualizar_vista_correos(self)
    
    if os.path.exists("internal/config.json"):
        try:
            with open("internal/config.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                self.db_path_field.value = data.get("ruta_db", "")
                self.output_path_field.value = data.get("ruta_output", "")
                self.rut_field.value = data.get("rut_empresa", "")
                # Priorizar lo que diga el JSON
                self.es_oscuro = data.get("dark_mode", tema_sistema) 
                self.correos_banco = data.get("correos_banco", [])
                
                log_debug(self, "CONFIG", f"✓ Cargado desde JSON. Modo Oscuro: {self.es_oscuro}")
        except Exception as e:
            log_debug(self, "CONFIG", f"✗ Error JSON: {e}")
            self.es_oscuro = tema_sistema
    else:
        # 2. SI NO HAY JSON: Creamos uno nuevo con la preferencia del sistema
        log_debug(self, "CONFIG", f"⚠️ Sin internal/config.json. Creando inicial con tema del sistema: {tema_sistema}")
        self.es_oscuro = tema_sistema
        
        # Valores vacíos para los campos técnicos
        self.db_path_field.value = ""
        self.output_path_field.value = ""
        self.rut_field.value = ""
        
        # Guardar automáticamente para que el bug del switch desaparezca
        config_inicial = {
            "ruta_db": "",
            "ruta_output": "",
            "rut_empresa": "",
            "dark_mode": self.es_oscuro
        }
        try:
            with open("internal/config.json", "w", encoding="utf-8") as f:
                json.dump(config_inicial, f, indent=4)
            log_debug(self, "CONFIG", "✓ Archivo internal/config.json inicial creado con éxito.")
        except Exception as e:
            log_debug(self, "CONFIG", f"✗ No se pudo crear el config inicial: {e}")

    # 3. SINCRONIZACIÓN FINAL: Forzar que la UI coincida con el estado lógico
    self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT
    
    if hasattr(self, "tema_switch") and self.tema_switch is not None:
        self.tema_switch.value = self.es_oscuro # Forzamos la posición del Switch
    
    actualizar_estilos_campos(self)
    self.page.update() # Refresco vital para Flet

def refrescar_años_db(self):
    """Obtiene años y muestra diálogo si hay error o está vacío"""
    log_debug(self, "DB", "═══ REFRESCANDO AÑOS DISPONIBLES ═══")
    
    try:
        lista_años = AE.obtener_años_disponibles()
        log_debug(self, "DB", f"Años obtenidos: {lista_años}")
        
        # Si la lista es None (error) o está vacía [] (no se encontraron datos o error manejado)
        if not lista_años:
            log_debug(self, "DB", "✗ Lista vacía o None recibida")
            raise Exception("No se obtuvieron datos de la base de datos.")

        # Si llegamos aquí, es que hay datos
        self.año_dropdown.disabled = False
        self.año_dropdown.options = [ft.dropdown.Option(str(a)) for a in lista_años]
        self.año_dropdown.value = str(lista_años[0])
        log_debug(self, "DB", f"✓ {len(lista_años)} años cargados exitosamente")
        log_debug(self, "DB", f"  Años disponibles: {', '.join(map(str, lista_años))}")
        log_debug(self, "DB", f"  Año por defecto: {lista_años[0]}")

    except Exception as e:
        log_debug(self, "DB", f"✗✗✗ ERROR al cargar años: {type(e).__name__}: {e}")
        
        # Limpiar y deshabilitar el dropdown
        self.año_dropdown.disabled = True
        self.año_dropdown.options = [ft.dropdown.Option("----")]
        self.año_dropdown.value = "----"
        log_debug(self, "DB", "Dropdown deshabilitado y reseteado a '----'")
        
        # DISPARAR DIÁLOGO
        mostrar_dialogo(self, 
            "Error de Base de Datos",
            "No se pudo obtener información de la base de datos.\n\n"
            "Asegúrese de que:\n"
            "1. La ruta en Configuración sea correcta.\n"
            "2. El archivo .accdb no esté abierto o bloqueado.\n"
            "3. El driver de Microsoft Access esté instalado.",
            es_error=True
        )
    
    self.page.update()
    log_debug(self, "DB", "═══ REFRESCO FINALIZADO ═══")
    
def generar_nomina(self):
    """Genera la nómina usando AutoExcel"""
    nomina = self.nomina_field.value
    año = self.año_dropdown.value
    
    log_debug(self, "NOMINA", "╔═══════════════════════════════════════╗")
    log_debug(self, "NOMINA", "║   INICIANDO GENERACIÓN DE NÓMINA    ║")
    log_debug(self, "NOMINA", "╚═══════════════════════════════════════╝")
    log_debug(self, "NOMINA", f"Input recibido → Nómina: '{nomina}', Año: '{año}'")
    log_debug(self, "NOMINA", f"Usar plantilla banco: {'Sí' if self.plantilla_r2_check.value else 'No'}")

    self.btn_generar.disabled = True
    self.btn_generar.text = "Procesando..."
    self.page.update()

    log_debug(self, "NOMINA", "═══ INICIANDO PROCESO BLOQUEADO ═══")

    
    try:
        # 1. Validaciones iniciales
        if not nomina or not nomina.isdigit():
            log_debug(self, "NOMINA", f"✗ Validación fallida: Nómina inválida '{nomina}'")
            mostrar_dialogo(self, "Error", "Ingrese una nómina válida (solo números)", True)
            return
        
        if not año or año == "----":
            log_debug(self, "NOMINA", f"✗ Validación fallida: Año no seleccionado '{año}'")
            mostrar_dialogo(self, "Error", "Seleccione un año válido", True)
            return
        
        if ((self.reporte_banco_check.value == True) and (self.plantilla_r2_check.value == False)):
            log_debug(self, "NOMINA", f"✗ Validación fallida: reporte banco  {self.reporte_banco_check.value} pero plantilla banco {self.plantilla_r2_check.value} ")
            mostrar_dialogo(self, "Error", "No puede generar Reporte PLantilla Banco sin generar Plantilla Banco", True)
            return


        log_debug(self, "NOMINA", "✓ Validaciones de entrada pasadas")
        
        # 2. Cargar configuración para obtener rutas y RUT
        if not os.path.exists("internal/config.json"):
            log_debug(self, "NOMINA", "✗ internal/config.json no existe, redirigiendo a configuración")
            mostrar_dialogo(self, "Error", "Configure la aplicación antes de generar nóminas", True)
            cambiar_ventana(self, "config")
            return

        with open("internal/config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
            rut_empresa = config.get("rut_empresa", "")
            ruta_salida = config.get("ruta_output", "")
        
        log_debug(self, "NOMINA", f"Configuración cargada:")
        log_debug(self, "NOMINA", f"  • RUT Empresa: {rut_empresa}")
        log_debug(self, "NOMINA", f"  • Ruta Salida: {ruta_salida}")
    
        log_debug(self, "NOMINA", "→ Abriendo conexión a base de datos...")
        conn = AE.open_database()
        
        if not conn:
            log_debug(self, "NOMINA", "✗✗✗ ERROR: No se pudo establecer conexión con la BD")
            mostrar_dialogo(self, "Error", "No se pudo conectar a la base de datos.", True)
            return
        
        log_debug(self, "NOMINA", "✓ Conexión establecida")
        
        try:
            cursor = conn.cursor()
            datos = AE.datos_egreso(cursor, int(nomina), int(año))
            
            if not datos:
                mostrar_dialogo(self, "Advertencia", "No se encontraron registros.", True)
                return

            # Lista para acumular los mensajes de éxito de cada módulo
            mensajes_exito = []

            # 1. Generar Nómina Estándar
            if self.plantilla_nomina_check.value:
                log_debug(self, "NOMINA", "checkbox de plantilla banco está activo, intentando generar Nómina Estándar...")
                resultado_nomina = AE.llenar_nomina_xlsx(datos, nomina, año, rut_empresa)
                if resultado_nomina:
                    mensajes_exito.append(resultado_nomina)
                else:
                    mostrar_dialogo(self, "Error", "Fallo al generar Nómina Estándar", True)
                    return

            # 2. Generar R2 si aplica
            if self.plantilla_r2_check.value:
                log_debug(self, "NOMINA", "checkbox de plantilla banco está activo, intentando generar R2...")
                mensaje_r2 = AE.llenarR2(datos, nomina, año, self.correo_remitente, self.correo_password, self.correos_banco, self.reporte_banco_check.value)
                if mensaje_r2:
                    mensajes_exito.append(mensaje_r2)
                else:
                    log_debug(self, "NOMINA", "⚠️ Error al generar R2 (continuando...)")

            # 3. Mostrar Diálogo Consolidado
            if mensajes_exito:
                cuerpo_mensaje = "\n".join(mensajes_exito)
                mostrar_dialogo(self, "Proceso Finalizado", cuerpo_mensaje)
                log_debug(self, "NOMINA", "✓✓✓ Mensaje de éxito mostrado")

        finally:
            conn.close()
            log_debug(self, "NOMINA", "Conexión a BD cerrada")
            
            
    except Exception as ex:
        log_debug(self, "NOMINA", f"✗✗✗ EXCEPCIÓN CRÍTICA: {type(ex).__name__}: {ex}")
        import traceback
        log_debug(self, "NOMINA", f"Traceback:\n{traceback.format_exc()}")
        mostrar_dialogo(self, "Error", f"Error crítico:\n{str(ex)}", True)
    
    finally:
        # --- 3. DESBLOQUEAR UI ---
        # Pase lo que pase (éxito o error), devolvemos el botón a su estado normal
        self.btn_generar.disabled = False
        self.btn_generar.text = "Generar Nómina"
        
        # self.loader.visible = False
        
        self.page.update()
        log_debug(self, "NOMINA", "═══ PROCESO FINALIZADO Y UI LIBERADA ═══")
    


def mostrar_dialogo(self, titulo, mensaje, es_error=False, acciones=None):
    """Muestra un diálogo modal. Si acciones es None, pone un botón 'OK'."""
    tipo = "ERROR" if es_error else "INFO"
    log_debug(self, "DIALOG", f"[{tipo}] {titulo}: {mensaje[:50]}...")
    
    def cerrar(e):
        dialogo.open = False
        self.page.update()

    # Si no hay acciones, usamos el botón por defecto
    if acciones is None:
        acciones = [ft.TextButton("OK", on_click=cerrar)]

    dialogo = ft.AlertDialog(
        modal=True if acciones else False, # Modal si requiere decisión
        title=ft.Text(titulo),
        content=ft.Text(mensaje),
        actions=acciones
    )
    
    self.page.overlay.append(dialogo)
    dialogo.open = True
    self.page.update()

def confirmar_salida(self):
    """Confirma la salida comparando el estado actual con el archivo JSON"""
    log_debug(self, "NAV", "─── Verificando cambios antes de salir ───")
    
    # 1. Cargar la configuración actual desde el disco para comparar
    config_en_disco = {"ruta_db": "", "ruta_output": "", "rut_empresa": "", "dark_mode": ""}
    if os.path.exists("internal/config.json"):
        try:
            with open("internal/config.json", "r", encoding="utf-8") as f:
                config_en_disco = json.load(f)
        except Exception as e:
            log_debug(self, "NAV", f"✗ Error al leer config para comparar: {e}")

    # 2. Verificar si hay diferencias
    cambios = {
        "tema": self.es_oscuro != config_en_disco.get("dark_mode", ""),
        "db": self.db_path_field.value != config_en_disco.get("ruta_db", ""),
        "output": self.output_path_field.value != config_en_disco.get("ruta_output", ""),
        "rut": self.rut_field.value != config_en_disco.get("rut_empresa", "")
    }
    
    hay_cambios = any(cambios.values())
    
    if hay_cambios:
        def cerrar_y_salir(e):
            # Aquí tendrías que cerrar el diálogo manualmente antes de cambiar
            self.page.overlay[-1].open = False # Accede al último diálogo
            log_debug(self, "NAV", "Usuario confirmó salir SIN GUARDAR")
            cambiar_ventana(self, "main")

        def solo_cerrar(e):
            self.page.overlay[-1].open = False
            self.page.update()

        # USANDO EL HELPER
        acciones_confirmacion = [
            ft.TextButton("Sí, salir", on_click=cerrar_y_salir),
            ft.TextButton("No, volver", on_click=solo_cerrar),
        ]
        
        mostrar_dialogo(
            self, 
            "Cambios sin guardar", 
            "Has realizado cambios en la configuración. ¿Deseas salir sin guardar?",
            acciones=acciones_confirmacion
        )
    else:
        cambiar_ventana(self, "main")
        refrescar_años_db(self)

def cambiar_tema_handler(self):
    """Alterna entre modo claro y oscuro usando el Switch"""
    # 1. Sincronizamos el estado con el valor del switch
    self.es_oscuro = self.tema_switch.value
    
    # 2. Aplicar modo a la página
    self.page.theme_mode = ft.ThemeMode.DARK if self.es_oscuro else ft.ThemeMode.LIGHT
    

    # 4. Sincronizar bordes de campos y actualizar
    actualizar_estilos_campos(self)
    self.page.update()
    log_debug(self, "TEMA", f"✓ UI cambiada a modo {'OSCURO' if self.es_oscuro else 'CLARO'}")


def actualizar_estilos_campos(self):
    """Actualiza solo los colores de los bordes y el fondo de la lista según el tema actual"""
    color_borde = ft.Colors.WHITE if self.es_oscuro else ft.Colors.BLACK
    
    # Definimos colores para el contenedor de la lista
    fondo_lista = ft.Colors.GREY_900 if self.es_oscuro else ft.Colors.BLUE_GREY_50
    borde_lista = ft.Colors.GREY_700 if self.es_oscuro else ft.Colors.OUTLINE_VARIANT
    
    lista_campos = [
        self.nomina_field, 
        self.rut_field, 
        self.db_path_field, 
        self.output_path_field,
        self.año_dropdown,
        self.user_field,
        self.pass_field,
        self.nuevo_usuario_field,
        self.nueva_pass_field,
        self.nuevo_correo_field,
        self.pass_remitente_input,
        self.remitente_input
    ]
    
    for campo in lista_campos:
        if campo:
            campo.border_color = color_borde
            campo.focused_border_color = color_borde
            
    # Actualizamos el contenedor de la lista si existe
    if hasattr(self, "lista_container") and self.lista_container:
        self.lista_container.bgcolor = fondo_lista
        self.lista_container.border = ft.border.all(1, borde_lista)
    
    log_debug(self, "TEMA", f"Bordes y lista actualizados. Modo oscuro: {self.es_oscuro}")



def cambiar_sub_config(self, index):
    """Cambia la sub-vista dentro de la configuración y actualiza el título"""
    # 1. Actualizar el índice del rail
    self.nav_rail.selected_index = index
    self.vista_parametros.visible=False
    self.vista_usuarios.visible=False
    self.vista_correos.visible=False
    
    # 2. Gestionar visibilidad y Título
    if index == 0:
        self.config_title.value = "Parámetros del Sistema"
        self.vista_parametros.visible = True
        self.content_area.content = self.vista_parametros
        log_debug(self, "CONFIG", "Sub-vista: PARÁMETROS")
    elif index == 1:
        self.config_title.value = "Gestión de Usuarios"
        self.vista_usuarios.visible = True
        self.content_area.content = self.vista_usuarios
        log_debug(self, "CONFIG", "Sub-vista: USUARIOS")
        # Aprovechamos de refrescar la lista al entrar
        actualizar_lista_usuarios(self)
    elif index == 2:
        self.config_title.value = "Gestión de Correos"
        self.vista_correos.visible = True
        self.content_area.content = self.vista_correos
        log_debug(self, "CONFIG", "Sub-vista: CORREOS")
        actualizar_vista_correos(self)
        
    self.page.update()


def actualizar_lista_usuarios(self):
    """Carga usuarios y aplica restricciones visuales según el rol del usuario logueado"""
    log_debug(self, "USER", "Sincronizando lista de usuarios con restricciones de rol")
    
    # 1. Obtener datos y determinar permisos
    usuarios_actuales = seguridad.listar_usuarios()
    self.usuarios_data = usuarios_actuales
    
    # Verificamos si el usuario activo es admin
    # (Asumimos que self.usuario_logueado es un dict: {'user': 'nombre', 'rol': 'admin'})
    rol_actual = self.usuario_logueado.get("rol", "user") if self.usuario_logueado else "user"
    es_admin = (rol_actual.lower() == "admin")

    # 2. Controlar visibilidad de la barra de creación superior
    if hasattr(self, "barra_creacion_usuarios"):
        self.barra_creacion_usuarios.visible = es_admin

    # 3. Limpiar y dibujar lista
    self.lista_usuarios_ui.controls.clear()
    
    if not usuarios_actuales:
        self.lista_usuarios_ui.controls.append(
            ft.Container(
                content=ft.Text("No hay usuarios registrados", color=ft.Colors.GREY_500, italic=True),
                padding=20, alignment=ft.alignment.center
            )
        )
    else:
        for indice, datos in enumerate(usuarios_actuales):
            nombre = datos['user']
            rol = datos.get('rol', 'user')
            
            # Estilo del Badge de rol
            badge_color = ft.Colors.BLUE_700 if rol.lower() == "admin" else ft.Colors.BLUE_GREY_400
            
            nueva_fila = ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.PERSON_ROUNDED, color=ft.Colors.BLUE_400),
                    ft.Text(nombre, size=16, weight=ft.FontWeight.W_500),
                    
                    # Badge de ROL
                    ft.Container(
                        content=ft.Text(rol.upper(), size=10, color="white", weight="bold"),
                        bgcolor=badge_color,
                        padding=ft.padding.symmetric(horizontal=8, vertical=2),
                        border_radius=5,
                        margin=ft.margin.only(left=10)
                    ),
                    ft.Container(expand=True),
                    
                    # BOTONES: Solo visibles si el logueado es ADMIN
                    ft.IconButton(
                        icon=ft.Icons.LOCK_RESET,
                        icon_color=ft.Colors.BLUE_GREY_400,
                        tooltip="Cambiar contraseña",
                        visible=es_admin, # <--- Restricción de rol
                        on_click=lambda _: editar_usuario_handler(self,indice)
                    ),
                    ft.IconButton(
                        icon=ft.Icons.DELETE_OUTLINE,
                        icon_color=ft.Colors.RED_400,
                        tooltip="Eliminar usuario",
                        visible=es_admin, # <--- Restricción de rol
                        on_click=lambda e, u=nombre: eliminar_usuario_handler(self, u)
                    ),
                ], alignment=ft.MainAxisAlignment.CENTER),
                padding=ft.padding.symmetric(horizontal=15, vertical=5),
                border=ft.border.only(bottom=ft.border.BorderSide(1, ft.Colors.OUTLINE_VARIANT)),
            )
            self.lista_usuarios_ui.controls.append(nueva_fila)
            
    self.page.update()


def agregar_usuario_handler(self):
    user = self.nuevo_usuario_field.value.strip()
    pw = self.nueva_pass_field.value.strip()
    
    # Determinamos el string del rol basado en el checkbox
    rol_seleccionado = "admin" if self.check_admin.value else "user"
    
    if not user or not pw:
        mostrar_dialogo(self, "Error", "Debe completar usuario y contraseña", True)
        return

    # Pasamos el rol a la función de registro
    exito, mensaje = seguridad.registrar_usuario(user, pw, rol=rol_seleccionado)
    
    if exito:
        log_debug(self, "USER", f"✓ {mensaje}")
        # Limpiamos los campos
        self.nuevo_usuario_field.value = ""
        self.nueva_pass_field.value = ""
        self.check_admin.value = False
        actualizar_lista_usuarios(self)
    else:
        mostrar_dialogo(self, "Error", mensaje, True)
    
    self.page.update()

def eliminar_usuario_handler(self, nombre_usuario):
    """Llama a seguridad para borrar y refresca la lista"""
    # Evitar que el usuario se borre a sí mismo si está logueado (opcional)
    
    def confirmar_borrado(e):
        self.page.overlay[-1].open = False # Cerrar diálogo
        if seguridad.eliminar_usuario(nombre_usuario):
            log_debug(self, "USER", f"Usuario '{nombre_usuario}' eliminado correctamente")
            actualizar_lista_usuarios(self)
        else:
            mostrar_dialogo(self, "Error", "No se pudo eliminar el usuario", True)

    # Diálogo de confirmación
    mostrar_dialogo(
        self,
        "Confirmar eliminación",
        f"¿Estás seguro de que deseas eliminar al usuario '{nombre_usuario}'?",
        acciones=[
            ft.TextButton("Eliminar", on_click=confirmar_borrado, icon=ft.Icons.DELETE, icon_color="red"),
            ft.TextButton("Cancelar", on_click=lambda e: self.page.overlay.__getitem__(-1).__setattr__('open', False) or self.page.update())
        ]
    )

def editar_usuario_handler(self, nombre_usuario):
    """Maneja la lógica para editar la contraseña de un usuario existente"""
    
    log_debug(self, "USER", f"Preparando cambio de contraseña para: {nombre_usuario}")

    # 2. Campo de texto para la nueva contraseña
    nueva_pass_field = ft.TextField(
        label="Nueva Contraseña",
        password=True,
        can_reveal_password=True,
        autofocus=True,
        prefix_icon=ft.Icons.LOCK_OUTLINED
    )

    def guardar_click(e):
        nueva_pw = nueva_pass_field.value.strip()
        if not nueva_pw:
            nueva_pass_field.error_text = "La contraseña no puede estar vacía"
            self.page.update()
            return

        # 3. Llamar a la función de seguridad.py
        if seguridad.cambiar_password(nombre_usuario, nueva_pw):
            log_debug(self, "USER", f"✓ Contraseña actualizada para '{nombre_usuario}'")
            dialogo_editar.open = False
            self.page.update()
            mostrar_dialogo(self, "Éxito", f"Contraseña de '{nombre_usuario}' actualizada correctamente.")
            # Refrescamos la lista por si acaso, aunque el nombre no cambie
            actualizar_lista_usuarios(self)
        else:
            log_debug(self, "USER", f"✗ Error al cambiar contraseña de '{nombre_usuario}'")
            mostrar_dialogo(self, "Error", "No se pudo actualizar la contraseña.", es_error=True)

    def cancelar_click(e):
        dialogo_editar.open = False
        self.page.update()

    # 4. Crear y mostrar el diálogo
    dialogo_editar = ft.AlertDialog(
        modal=True,
        title=ft.Text(f"Editar Usuario: {nombre_usuario}"),
        content=ft.Column([
            ft.Text("Ingrese la nueva contraseña para este usuario:"),
            nueva_pass_field
        ], tight=True),
        actions=[
            ft.TextButton("Guardar", on_click=guardar_click),
            ft.TextButton("Cancelar", on_click=cancelar_click),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    self.page.overlay.append(dialogo_editar)
    dialogo_editar.open = True
    self.page.update()


def actualizar_vista_correos(self):
    """Refresca toda la Vista 3: Remitente (Binario) y Destinatarios (JSON)"""
    log_debug(self, "EMAIL", "═══ SINCRONIZANDO VISTA DE CORREOS ═══")
    
    # ---------------------------------------------------------
    # PARTE 1: REFRESCAR REMITENTE (Datos Sensibles)
    # ---------------------------------------------------------
    email_remitente, pass_app = seguridad.obtener_credenciales_email()
    self.seccion_remitente_ui.controls.clear()

    if email_remitente and pass_app:
        self.seccion_remitente_ui.controls.append(
            ft.Row([
                ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400),
                ft.Text(f"Configurado: {email_remitente}", size=16, weight=ft.FontWeight.W_500),
                ft.IconButton(
                    icon=ft.Icons.EDIT,
                    icon_color=ft.Colors.BLUE_400,
                    tooltip="Editar credenciales",
                    on_click=lambda _: mostrar_edicion_remitente(self, email_remitente)
                )
            ], alignment=ft.MainAxisAlignment.CENTER)
        )
    else:
        self.seccion_remitente_ui.controls.extend([
            ft.Row([self.remitente_input, self.pass_remitente_input], alignment=ft.MainAxisAlignment.CENTER),
            ft.ElevatedButton(
                "Guardar Credenciales", 
                icon=ft.Icons.SAVE,
                on_click=lambda _: guardar_remitente_handler(self)
            )
        ])

    # ---------------------------------------------------------
    # PARTE 2: REFRESCAR LISTA DE DESTINATARIOS (JSON)
    # ---------------------------------------------------------
    try:
        if os.path.exists("internal/config.json"):
            with open("internal/config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                correos_actuales = config.get("correos_banco", [])
        else:
            correos_actuales = []
    except Exception as e:
        log_debug(self, "USER", f"✗ Error al leer internal/config.json: {e}")
        correos_actuales = []

    # Permisos de Admin
    rol_actual = self.usuario_logueado.get("rol", "user") if self.usuario_logueado else "user"
    es_admin = (rol_actual.lower() == "admin")
    
    if hasattr(self, "barra_creacion_correos"):
        self.barra_creacion_correos.visible = es_admin

    self.lista_correos_ui.controls.clear()
    
    if not correos_actuales:
        self.lista_correos_ui.controls.append(
            ft.Container(
                content=ft.Text("No hay destinatarios registrados", color=ft.Colors.GREY_500, italic=True),
                padding=20, alignment=ft.alignment.center
            )
        )
    else:
        for correo in correos_actuales:
            fila = ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.MARK_EMAIL_READ_OUTLINED, color=ft.Colors.BLUE_400),
                    ft.Text(correo, size=15),
                    ft.Container(expand=True),
                    ft.IconButton(
                        icon=ft.Icons.DELETE_OUTLINE,
                        icon_color=ft.Colors.RED_400,
                        visible=es_admin,
                        on_click=lambda e, u=correo: eliminar_correo_handler(self, u)
                    ),
                ]),
                padding=ft.padding.symmetric(horizontal=15, vertical=2),
                border=ft.border.only(bottom=ft.border.BorderSide(1, ft.Colors.OUTLINE_VARIANT)),
            )
            self.lista_correos_ui.controls.append(fila)

    # ---------------------------------------------------------
    # FINALIZACIÓN
    # ---------------------------------------------------------
    log_debug(self, "EMAIL", "✓ Vista de correos sincronizada")
    self.page.update()



def agregar_correo_handler(self):
    correo = self.nuevo_correo_field.value.strip()
    
    
    if not correo:
        mostrar_dialogo(self, "Campo Vacío", "Por favor, ingrese un correo electrónico.", True)
        return

    if validar_formato_correo(correo) is False:
        log_debug(self, "USER", f"✗ Validación fallida: Formato de correo inválido '{correo}'")
        mostrar_dialogo(self, "Formato Inválido", 
                        "El formato del correo electrónico no es válido.\nEjemplo: usuario@empresa.cl", True)
        self.nuevo_correo_field.value = ""
        return

    # 2. Lógica de guardado en internal/config.json
    try:
        # Aseguramos que el archivo existe o manejamos su lectura
        if not os.path.exists("internal/config.json"):
             config = {}
        else:
            with open("internal/config.json", "r", encoding="utf-8") as f:
                config = json.load(f)

        correos_banco = config.get("correos_banco", [])
        
        if correo in correos_banco:
            mostrar_dialogo(self, "Error", "El correo ya está registrado en la lista.", True)
            self.nuevo_correo_field.value = ""
            return
        correos_banco.append(correo)
        config["correos_banco"] = correos_banco
        
        with open("internal/config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
        
        log_debug(self, "USER", f"✓ Correo '{correo}' validado y agregado a configuración")
        
        # 3. Limpieza y actualización de UI
        actualizar_vista_correos(self)
        self.nuevo_correo_field.value = ""
        self.nuevo_correo_field.error_text = None # Limpia errores visuales si los hubiera
        
    except Exception as e:
        log_debug(self, "USER", f"✗ Error al agregar correo: {e}")
        mostrar_dialogo(self, "Error", f"No se pudo agregar el correo: {e}", True)

    self.page.update()


def eliminar_correo_handler(self, correo):
    """Elimina un correo de internal/config.json y actualiza la UI"""
    def confirmar_borrado(e):
        self.page.overlay[-1].open = False # Cerrar diálogo
        try:
            with open("internal/config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                correos_banco = config.get("correos_banco", [])
                if correo in correos_banco:
                    correos_banco.remove(correo)
                    config["correos_banco"] = correos_banco
            with open("internal/config.json", "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4)
            log_debug(self, "USER", f"✓ Correo '{correo}' eliminado de configuración")
            actualizar_vista_correos(self)
        except Exception as e:
            log_debug(self, "USER", f"✗ Error al eliminar correo: {e}")
            mostrar_dialogo(self, "Error", f"No se pudo eliminar el correo: {e}", True)

    # Diálogo de confirmación
    mostrar_dialogo(
        self,
        "Confirmar eliminación",
        f"¿Estás seguro de que deseas eliminar el correo '{correo}'?",
        acciones=[
            ft.TextButton("Eliminar", on_click=confirmar_borrado, icon=ft.Icons.DELETE, icon_color="red"),
            ft.TextButton("Cancelar", on_click=lambda e: self.page.overlay.__getitem__(-1).__setattr__('open', False) or self.page.update())
        ]
    )



def cerrar_sesion_handler(self):
    """Limpia los datos del usuario activo y regresa al login"""
    log_debug(self, "AUTH", f"Cerrando sesión del usuario: {self.usuario_logueado.get('user') if self.usuario_logueado else 'Desconocido'}")
    
    # 1. Borrar datos de sesión
    self.usuario_logueado = None
    
    # 2. Limpiar campos de texto sensibles
    if hasattr(self, "user_field"): self.user_field.value = ""
    if hasattr(self, "pass_field"): self.pass_field.value = ""
    if hasattr(self, "nuevo_usuario_field"): self.nuevo_usuario_field.value = ""
    if hasattr(self, "nueva_pass_field"): self.nueva_pass_field.value = ""
    
    # 3. Mandar al login
    cambiar_ventana(self, "login")
    
    log_debug(self, "NAV", "✓ Sesión cerrada, redirigido a Login")



def guardar_remitente_handler(self):
    email = self.remitente_input.value
    password = self.pass_remitente_input.value
    
    if not email or not password:
        mostrar_dialogo(self, "Error", "Debe completar ambos campos", True)
        return
    
    if validar_formato_correo(email) is False:
        log_debug(self, "USER", f"✗ Validación fallida: Formato de correo inválido '{email}'")
        mostrar_dialogo(self, "Formato Inválido", 
                        "El formato del correo electrónico no es válido.\nEjemplo: ejemplo@gmail.com", True)
        return
        
    seguridad.guardar_credenciales_email(email, password)
    self.remitente_input.value = ""
    self.pass_remitente_input.value = ""
    cargar_configuracion(self)
    actualizar_vista_correos(self)
    mostrar_dialogo(self, "Éxito", "Credenciales de remitente guardadas de forma segura.")

def mostrar_edicion_remitente(self, email_actual):
    # Simplemente limpiamos los datos en seguridad y refrescamos la vista 
    # para que vuelvan a salir los campos de entrada
    self.remitente_input.value = email_actual
    self.seccion_remitente_ui.controls.clear()
    self.seccion_remitente_ui.controls.extend([
        ft.Row([self.remitente_input, self.pass_remitente_input], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([
            ft.ElevatedButton("Actualizar", on_click=lambda _: guardar_remitente_handler(self)),
            ft.TextButton("Cancelar", on_click=lambda _:actualizar_vista_correos(self))
        ], alignment=ft.MainAxisAlignment.CENTER)
    ])
    self.page.update()

def validar_formato_correo(correo: str) -> bool:
    """
    Verifica si un string tiene una estructura de correo electrónico válida.
    Retorna True si es válido, False si no.
    """
    # Explicación del patrón:
    # ^[a-zA-Z0-9._%+-]+ : Empieza con letras, números o caracteres permitidos.
    # @                  : Debe contener un arroba.
    # [a-zA-Z0-9.-]+     : Dominio con letras o números.
    # \.[a-zA-Z]{2,}$    : Punto seguido de al menos 2 letras (com, cl, org, etc).
    
    patron = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    if correo and re.match(patron, correo.strip()):
        return True
    return False
        
