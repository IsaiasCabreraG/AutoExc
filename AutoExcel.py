import pyodbc # type: ignore
import openpyxl # type: ignore
import os
import json 
import sys



def open_database():
    """Busca la configuración en el JSON y abre la conexión automáticamente."""
    
    # 1. Verificar si existe el config.json
    if not os.path.exists("config.json"):
        print("Error: No se encontró el archivo de configuración.")
        return None

    # 2. Leer la ruta desde el JSON
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)
        db_path = config.get("ruta_db")

    if not db_path:
        print("Error: La ruta de la base de datos no está configurada.")
        return None

    # 3. El Driver suele ser fijo para Access (.accdb)
    driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    
    try:
        print(f"Conectando a: {db_path}")
        conn = pyodbc.connect(f'DRIVER={driver};DBQ={db_path};')
        return conn
    except Exception as e:
        print(f"Error de conexión: {e}")
        return None

def datos_egreso(cursor: pyodbc.Cursor, NEgreso: int):
    """Funcion encargada de obtener los datos de un egreso especifico y los nombres de columna.
    Args:
        cursor (pyodbc.Cursor): Cursor de la base de datos.
        NEgreso (int): Numero de egreso a buscar.
        
    Returns:
        tuple: (fila_de_datos o None, lista_de_nombres_de_columnas)
    """
    CpbNum = str(NEgreso)
    if len(CpbNum) < 8:
        CpbNum = CpbNum.zfill(8)
    print(f"Buscando egreso con CpbNum: {CpbNum}")

    cursor.execute("SELECT CodAux, MovTipDocRef FROM cwmovim WHERE CpbNum = ?", (CpbNum,))
    fila = cursor.fetchall()
    columnas = [desc[0] for desc in cursor.description] if cursor.description else []
    return fila, columnas

def mostrar_estructura_tabla(connection, tabla):
    """Funcion encargada de mostrar la estructura de una tabla en la base de datos.
    Args:
        connection (pyodbc.Connection): Conexion a la base de datos.
        tabla (str): Nombre de la tabla a inspeccionar.
    """
    cursor = connection.cursor()
    
    print(f"{'Columna':<25} | {'Tipo de Dato':<15} | {'Tamaño'}")
    print("-" * 55)
    
    # El método .columns devuelve metadatos de la tabla
    for row in cursor.columns(table=tabla):
        nombre_col = row.column_name
        tipo_dato = row.type_name
        tamano = row.column_size
        
        print(f"{nombre_col:<25} | {tipo_dato:<15} | {tamano}")




def obtener_ruta_recurso(relativa):
    """ Obtiene la ruta absoluta para recursos (iconos, plantillas, etc) """
    try:
        # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relativa)

def llenar_nomina_xlsx(datos_db, numero_nomina):
    """Llena una plantilla de Excel con datos obtenidos de la base de datos.
    Args:
        datos_db (list): Lista de tuplas con los datos obtenidos de la base de datos.
        numero_nomina (str): Numero de nomina para nombrar el archivo de salida.
    """
    if not os.path.exists("config.json"):
        print("Error: No existe config.json")
        return

    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)
        carpeta_salida = config.get("ruta_output", "")

    nombre_archivo = f"Nomina_{numero_nomina}.xlsx"
    ruta_final_salida = os.path.join(carpeta_salida, nombre_archivo)
    ruta_plantilla = obtener_ruta_recurso("static/Plantilla_NominaPagoV2.xlsx")

    try:
        wb = openpyxl.load_workbook(ruta_plantilla)
        ws = wb.active

        if not isinstance(datos_db, list):
            datos_db = [datos_db]

        fila_inicio = 11
        for i, registro in enumerate(datos_db):
            fila_actual = fila_inicio + i
            ws.cell(row=fila_actual, column=2).value = registro[0]
            ws.cell(row=fila_actual, column=3).value = registro[1]
        wb.save(ruta_final_salida)
        print(f"Éxito: Archivo guardado en {ruta_final_salida}")
        
    except Exception as e:
        print(f"Error al procesar Excel: {e}")

        

if __name__ == "__main__":

    conn = open_database()
    if conn:
        cursor = conn.cursor()
        datos, columnas = datos_egreso(cursor, 123456)
        print("Datos obtenidos:", datos)
        print("Columnas:", columnas)
        mostrar_estructura_tabla(conn, "cwmovim")
        llenar_nomina_xlsx(datos, "123456")
        conn.close()

