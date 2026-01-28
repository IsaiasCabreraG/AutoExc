import json
import os
import hashlib
from cryptography.fernet import Fernet

# ============================================================
# CONFIGURACIÓN DE SEGURIDAD
# ============================================================
# ⚠️ Genera esta llave SOLO UNA VEZ con Fernet.generate_key()
# ⚠️ No la cambies o perderás acceso a los usuarios guardados
LLAVE_MAESTRA = b'uX6_L7p8vY2R3nE5tW1qA9sD4fG6hJ8kL0mN2bP4rQ8='
cipher_suite = Fernet(LLAVE_MAESTRA)

# ============================================================
# RUTA LOCAL (MISMA CARPETA DEL PROYECTO)
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_USUARIOS = os.path.join(BASE_DIR, "sys_data.bin")

# ============================================================
# FUNCIONES INTERNAS
# ============================================================

def _hash_password(password: str) -> str:
    """Genera un hash SHA-256 de la contraseña."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _guardar_archivo(datos: dict) -> None:
    """Encripta y guarda el archivo de usuarios."""
    json_data = json.dumps(datos).encode("utf-8")
    encriptado = cipher_suite.encrypt(json_data)

    with open(ARCHIVO_USUARIOS, "wb") as f:
        f.write(encriptado)


def _cargar_archivo() -> dict:
    """Carga y desencripta el archivo de usuarios."""
    if not os.path.exists(ARCHIVO_USUARIOS):
        return {}

    try:
        with open(ARCHIVO_USUARIOS, "rb") as f:
            contenido = f.read()

        desencriptado = cipher_suite.decrypt(contenido)
        return json.loads(desencriptado.decode("utf-8"))

    except Exception:
        # Archivo corrupto o llave incorrecta
        return {}

# ============================================================
# FUNCIONES PÚBLICAS
# ============================================================

def registrar_usuario(usuario: str, password: str, rol: str = "user") -> tuple:
    """Registra un nuevo usuario con un rol específico."""
    usuarios = _cargar_archivo()

    if usuario in usuarios:
        return False, "El usuario ya existe."

    # Guardamos como un diccionario interno
    usuarios[usuario] = {
        "pass": _hash_password(password),
        "rol": rol
    }
    _guardar_archivo(usuarios)
    return True, f"Usuario {usuario} registrado como {rol}."

def validar_credenciales(usuario: str, password: str) -> bool:
    """Valida usuario y contraseña buscando en la nueva estructura."""
    usuarios = _cargar_archivo()

    if usuario not in usuarios:
        return False

    # Accedemos a la clave 'pass' dentro del usuario
    return usuarios[usuario]["pass"] == _hash_password(password)

def obtener_rol(usuario: str) -> str:
    """Devuelve el rol del usuario o None si no existe."""
    usuarios = _cargar_archivo()
    if usuario in usuarios:
        return usuarios[usuario].get("rol", "admin")
    return None


def eliminar_usuario(usuario: str) -> bool:
    """Elimina un usuario."""
    usuarios = _cargar_archivo()

    if usuario not in usuarios:
        return False

    del usuarios[usuario]
    _guardar_archivo(usuarios)
    return True


def listar_usuarios():
    """Devuelve una lista de diccionarios con usuarios y sus roles."""
    usuarios = _cargar_archivo()

    return [{"user": k, "rol": v.get("rol", "admin")} for k, v in usuarios.items()]




def cambiar_password(usuario: str, nueva_password: str) -> bool:
    """Cambia la contraseña de un usuario."""
    usuarios = _cargar_archivo()

    if usuario not in usuarios:
        return False

    usuarios[usuario] = _hash_password(nueva_password)
    _guardar_archivo(usuarios)
    return True
