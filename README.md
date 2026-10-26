# AutoExcel

Aplicación de escritorio para Windows que automatiza la generación de **nóminas de pago** a partir de los egresos de una base de datos contable, reemplazando un proceso que antes se hacía a mano.

Desarrollada para la unidad de tesorería de una empresa constructora, donde sigue en uso.

## El problema

Para pagar a proveedores, tesorería debía copiar manualmente los datos de cada egreso (proveedor, documento, monto, banco, cuenta) desde el sistema contable a la plantilla Excel de nómina. Era un trabajo repetitivo y propenso a errores de transcripción.

## Qué hace

1. El usuario indica el **número de nómina (egreso)** y el **año**.
2. La app consulta la base de datos (Microsoft Access, vía ODBC) y obtiene las facturas asociadas a ese egreso, junto con los datos bancarios de cada proveedor.
3. Llena la **plantilla Excel de nómina de pagos** y la guarda en la carpeta de salida configurada.
4. Opcionalmente genera la **plantilla del banco** y envía por correo un reporte con el detalle y el monto total a los destinatarios elegidos, indicando la fecha de pago.

## Características

- Interfaz gráfica con tema claro/oscuro, hecha con [Flet](https://flet.dev).
- Inicio de sesión con **usuarios y roles**; las credenciales se guardan cifradas, nunca en texto plano.
- Pantalla de configuración: ruta de la base de datos, carpeta de salida y RUT de la empresa.
- Detección automática de los años disponibles en la base de datos.
- Gestión de destinatarios y del remitente del correo desde la propia app.
- Validaciones antes de generar (nómina numérica, año válido, fecha y destinatarios cuando se envía el reporte).
- Vista de depuración con el log en pantalla.
- Se distribuye como un único `.exe` (PyInstaller): el usuario final no instala Python.

## Tecnologías

Python · Flet · pyodbc (Access) · openpyxl · smtplib · cryptography (Fernet) · PyInstaller

## Estructura

| Archivo | Rol |
|---|---|
| `main.py` | Punto de entrada y aplicación Flet |
| `ui_components.py` | Vistas: login, principal, configuración, depuración |
| `utils.py` | Lógica de la interfaz: generación, configuración, usuarios y correos |
| `AutoExcel.py` | Consulta a la base de datos, llenado de las plantillas Excel y envío del reporte |
| `seguridad.py` | Usuarios, contraseñas y credenciales de correo |
| `static/` | Plantillas Excel e iconos |

---

## Requisitos

- Windows (64 bits)
- Python 3.14 (con `pip`)
- Driver ODBC de Access: *Microsoft Access Database Engine* (64 bits). Necesario para ejecutar la app y para probarla después de compilar.

## Compilar la aplicación (Windows)

Desde una terminal (PowerShell o CMD) en la carpeta del proyecto:

1. Crear y activar un entorno virtual:

   ```powershell
   python -m venv venv
   venv\Scripts\activate
   ```

2. Instalar las dependencias (incluye PyInstaller):

   ```powershell
   pip install -r requirements.txt
   ```

3. Subir la versión en `version_info.txt` (si corresponde). Hay que cambiar las cuatro apariciones:
   `filevers`, `prodvers`, `FileVersion` y `ProductVersion`.

4. Cerrar `AutoExcel.exe` si está abierto (si no, PyInstaller falla con `Acceso denegado`) y compilar:

   ```powershell
   python -m PyInstaller AutoExcel.spec --noconfirm --clean
   ```

   Usar `--clean` evita que PyInstaller reutilice caché y deje fuera librerías nuevas.

5. El ejecutable queda en `dist\AutoExcel.exe`. El `.spec` ya incluye la carpeta `static` (plantillas e icono) dentro del exe.

## Distribuir

- Entregar `AutoExcel.exe`. Junto a él debe estar la carpeta `internal`, donde la app guarda su configuración (`config.json`), credenciales (`sys_data.bin`) y log.
- Para actualizar una instalación existente sin perder la configuración, basta con reemplazar el `.exe` y mantener la misma carpeta `internal`.
- Los equipos donde se ejecute deben tener el driver de Access indicado arriba, y `config.json` debe apuntar a la ruta correcta de la base de datos (`.accdb`).

## Notas

- Si el exe se cierra al abrir con `ModuleNotFoundError`, falta una dependencia en `requirements.txt`; agregarla y compilar de nuevo con `--clean`.
