import pyodbc # type: ignore
import openpyxl # type: ignore
import os
import json 



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

def tabular_datos(cursor: pyodbc.Cursor, datos):
    """
    Toma el resultado de un fetch (tupla o lista de tuplas) 
    y lo imprime como una tabla usando los metadatos del cursor.
    """
    if not datos:
        print("Sin datos para mostrar.")
        return

    # 1. Normalizar: si es una sola tupla (fetchone), meterla en una lista
    filas = [datos] if isinstance(datos, pyodbc.Row) or isinstance(datos, tuple) else datos
    
    # 2. Obtener encabezados desde el cursor
    headers = [desc[0] for desc in cursor.description]
    
    # 3. Calcular anchos de columna (comparando encabezado vs datos)
    # Convertimos todo a string para medir largo
    anchos = []
    for i, col in enumerate(headers):
        largo_max_datos = max([len(str(fila[i])) for fila in filas])
        largo_cabecera = len(str(col))
        anchos.append(max(largo_max_datos, largo_cabecera, 5)) # Mínimo 5 espacios

    # 4. Construir separadores y filas
    linea = "+" + "+".join("-" * (a + 2) for a in anchos) + "+"
    
    # Imprimir Cabecera
    print(linea)
    encabezado_str = "| " + " | ".join(f"{str(h):<{anchos[i]}}" for i, h in enumerate(headers)) + " |"
    print(encabezado_str)
    print(linea.replace("-", "=")) # Separador visual fuerte para el header

    # Imprimir Filas
    for fila in filas:
        fila_str = "| " + " | ".join(f"{str(fila[i]):<{anchos[i]}}" for i, d in enumerate(fila)) + " |"
        print(fila_str)
    
    print(linea)



def llenar_nomina_xlsx(datos_db, ruta_plantilla, ruta_salida):
    """
    datos_db: Puede ser una lista de tuplas (fetchall) o una tupla sola (fetchone).
    """
    # 1. Cargar el libro y la hoja activa
    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active # Esto toma la primera pestaña

    # 2. Asegurarnos de que los datos sean una lista (por si viene un fetchone solo)
    if not isinstance(datos_db, list):
        datos_db = [datos_db]

    # 3. Empezar a escribir desde la fila 10
    fila_inicio = 11
    
    print(f"Insertando {len(datos_db)} filas en el Excel...")

    for i, registro in enumerate(datos_db):
        fila_actual = fila_inicio + i
        
        # registro[0], registro[1], etc., dependen del orden de tu SELECT
        # Ejemplo basado en una nómina típica (RUT, Nombre, Monto...):
        ws.cell(row=fila_actual, column=2).value = registro[0] # Columna A
        ws.cell(row=fila_actual, column=3).value = registro[1] # Columna B
        
    # 4. Guardar el archivo
    wb.save(ruta_salida)
    print(f"Archivo guardado exitosamente como: {ruta_salida}")

        

if __name__ == "__main__":

    db_dir = "."
    db_name = 'Database1'
    db_ext = '.accdb'
    driver = 'microsoft access driver (*.mdb, *.accdb)'

    plantilla_path = "./static/Plantilla_NominaPagoV2.xlsx"
    ruta_salida = "./output/NominaPago_Generado.xlsx"


    connection = open_database(db_dir, db_name, db_ext, driver)
    print("Conexion exitosa a la base de datos.")
    cursor = connection.cursor()

    # 3. Ejemplo de SELECT (Leer datos)
    mostrar_estructura_tabla(connection, "cwmovim")
    egreso_numero = 12177 
    egreso_datos, columnas = datos_egreso(cursor, egreso_numero)
    if egreso_datos:
        print("Datos del egreso encontrado:")
        print(egreso_datos)
        llenar_nomina_xlsx(egreso_datos, plantilla_path, ruta_salida)
    else:
        print("No se encontro ningun egreso con ese numero.")




