import json
import os
import hashlib
import sys
from cryptography.fernet import Fernet

# ============================================================
# CONFIGURACIÓN DE SEGURIDAD
# ============================================================
LLAVE_MAESTRA = b'uX6_L7p8vY2R3nE5tW1qA9sD4fG6hJ8kL0mN2bP4rQ8='
cipher_suite = Fernet(LLAVE_MAESTRA)

# ============================================================
# RUTA JUNTO AL EJECUTABLE (CORREGIDO)
# ============================================================
if getattr(sys, 'frozen', False):
    # Si es el .exe, usamos la carpeta donde vive el ejecutable físico
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Si es el script .py, usamos la carpeta del proyecto
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARCHIVO_USUARIOS = os.path.join(BASE_DIR, "internal/sys_data.bin")

# ============================================================
# FUNCIONES INTERNAS
# ============================================================

def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def _guardar_archivo(datos: dict) -> None:
    json_data = json.dumps(datos).encode("utf-8")
    encriptado = cipher_suite.encrypt(json_data)
    with open(ARCHIVO_USUARIOS, "wb") as f:
        f.write(encriptado)

def _cargar_archivo() -> dict:
    if not os.path.exists(ARCHIVO_USUARIOS):
        return {}
    try:
        with open(ARCHIVO_USUARIOS, "rb") as f:
            contenido = f.read()
        desencriptado = cipher_suite.decrypt(contenido)
        return json.loads(desencriptado.decode("utf-8"))
    except Exception:
        return {}

# ============================================================
# FUNCIONES PÚBLICAS
# ============================================================

def registrar_usuario(usuario: str, password: str, rol: str = "user") -> tuple:
    usuarios = _cargar_archivo()
    if usuario in usuarios:
        return False, "El usuario ya existe."

    usuarios[usuario] = {
        "pass": _hash_password(password),
        "rol": rol
    }
    _guardar_archivo(usuarios)
    return True, f"Usuario {usuario} registrado como {rol}."

def validar_credenciales(usuario: str, password: str) -> bool:
    if usuario.startswith("__sys_"): return False  # Bloqueo de seguridad
    usuarios = _cargar_archivo()
    if usuario not in usuarios or "pass" not in usuarios[usuario]:
        return False
    return usuarios[usuario]["pass"] == _hash_password(password)

def obtener_rol(usuario: str) -> str:
    usuarios = _cargar_archivo()
    if usuario in usuarios:
        return usuarios[usuario].get("rol", "admin")
    return None

def eliminar_usuario(usuario: str) -> bool:
    usuarios = _cargar_archivo()
    if usuario not in usuarios:
        return False
    del usuarios[usuario]
    _guardar_archivo(usuarios)
    return True

def listar_usuarios():
    usuarios = _cargar_archivo()
    return [
        {"user": k, "rol": v.get("rol", "admin")} 
        for k, v in usuarios.items() 
        if not k.startswith("__sys_")
    ]

def cambiar_password(usuario: str, nueva_password: str) -> bool:
    usuarios = _cargar_archivo()
    if usuario not in usuarios:
        return False
    # Actualizamos solo el hash, manteniendo el diccionario interno
    usuarios[usuario]["pass"] = _hash_password(nueva_password)
    _guardar_archivo(usuarios)
    return True

def guardar_credenciales_email(email: str, password_app: str) -> None:
    """Guarda credenciales de correo en el binario cifrado."""
    datos = _cargar_archivo()
    # Usamos una clave especial que no sea un nombre de usuario
    datos["__sys_config_correo__"] = {
        "email": email,
        "pass_app": password_app
    }
    _guardar_archivo(datos)

def obtener_credenciales_email() -> tuple:
    """Recupera (email, pass_app) del binario cifrado."""
    datos = _cargar_archivo()
    config = datos.get("__sys_config_correo__", {})
    return config.get("email", ""), config.get("pass_app", "")