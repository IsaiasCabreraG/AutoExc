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

def registrar_usuario(usuario: str, password: str):
    """Registra un nuevo usuario."""
    usuarios = _cargar_archivo()

    if usuario in usuarios:
        return False, "El usuario ya existe."

    usuarios[usuario] = _hash_password(password)
    _guardar_archivo(usuarios)

    return True, "Usuario registrado exitosamente."


def validar_credenciales(usuario: str, password: str) -> bool:
    """Valida usuario y contraseña."""
    usuarios = _cargar_archivo()

    if usuario not in usuarios:
        return False

    return usuarios[usuario] == _hash_password(password)


def eliminar_usuario(usuario: str) -> bool:
    """Elimina un usuario."""
    usuarios = _cargar_archivo()

    if usuario not in usuarios:
        return False

    del usuarios[usuario]
    _guardar_archivo(usuarios)
    return True


def listar_usuarios():
    """Devuelve una lista de usuarios existentes (sin contraseñas)."""
    return list(_cargar_archivo().keys())


def cambiar_password(usuario: str, nueva_password: str) -> bool:
    """Cambia la contraseña de un usuario."""
    usuarios = _cargar_archivo()

    if usuario not in usuarios:
        return False

    usuarios[usuario] = _hash_password(nueva_password)
    _guardar_archivo(usuarios)
    return True
