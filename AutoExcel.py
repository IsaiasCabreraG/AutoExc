import pyodbc # type: ignore
import openpyxl # type: ignore
import os
import json 
import sys
import logging
import datetime

MAPEO_DOCS = {
    'FT': '33: Factura Electrónica'
}

MAPEO_PAGO = {
    'CC': 'CC: Cuenta Corriente',
    'CA': 'CA: Cuenta de Ahorro',
    'VV': 'VV: Vale Vista',
    'CH': 'CH: Cheque',
    'OT': 'OT: Otro'
}





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

def datos_egreso(cursor, NEgreso: int, año: int):
    """
    Busca datos de egreso y los retorna como una lista de diccionarios con nombres claros.
    Args:
        cursor (pyodbc.Cursor): Cursor de la base de datos.
        NEgreso (int): Número de egreso a buscar.
        año (int): Año del egreso.
    Returns:
        list: Lista de diccionarios con los datos del egreso.
    """
    # 1. Preparar el número de egreso (rellenar con ceros hasta 8 caracteres)
    CpbNum = str(NEgreso).zfill(8)
    CpbAno = str(año)
    codigos_permitidos = tuple(MAPEO_DOCS.keys())
    filtro_sql = str(codigos_permitidos).replace(",)", ")")
    
    print(f"Buscando egreso N°: {CpbNum} para el año: {CpbAno}")

    # 2. Consulta SQL con las columnas solicitadas
    # Nota: Asegúrate de que los nombres de las columnas coincidan exactamente con tu DB
    query = f"""
        SELECT 
            T1.CodAux, 
            T2.NomAux, 
            T1.MovNumDocRef, 
            T1.MovTipDocRef, 
            T1.MovDebe, 
            T1.FecPag
        FROM cwmovim AS T1
        LEFT JOIN cwtauxi AS T2 ON T1.CodAux = T2.CodAux
        WHERE T1.CpbNum = ? 
          AND T1.CpbAno = ?
          AND T1.CodAux <> '0000000000'
          AND T1.MovTipDocRef IN {filtro_sql}
    """
    
    try:
        cursor.execute(query, (CpbNum, CpbAno))
        filas = cursor.fetchall()
        
        # 3. Transformar a estructura de diccionarios (Clave: Valor)
        registros_mapeados = []
        for fila in filas:
            registro = {
                'Rut Proovedor':   fila[0],
                'Razon Social':    fila[1],
                'N documento':     fila[2],
                'Tipo Documento':  fila[3],
                'Monto Documento': fila[4],
                'Fecha Emision':   fila[5],
                'Total Pagar':     fila[4]  # MovDebe se asigna a ambos según tu solicitud
            }
            registros_mapeados.append(registro)
            
        return registros_mapeados
        
    except Exception as e:
        print(f"Error en la consulta SQL: {e}")
        return []

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

def llenar_nomina_xlsx(datos_db, numero_nomina, año):
    """
    Llena la plantilla de Excel usando diccionarios mapeados.
    """
    logger = logging.getLogger("AutoExcel")
    

    if not os.path.exists("config.json"):
        logger.error("No se encontró config.json para obtener la ruta de salida.")
        return

    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)
        carpeta_salida = config.get("ruta_output", "")

    if not carpeta_salida or not os.path.exists(carpeta_salida):
        logger.error(f"La ruta de salida no es válida: {carpeta_salida}")
        return

    nombre_archivo = f"Nomina_{numero_nomina}-{año}.xlsx"
    ruta_final_salida = os.path.join(carpeta_salida, nombre_archivo)
    ruta_plantilla = obtener_ruta_recurso("static/Plantilla_NominaPagoV2.xlsx")

    try:
        logger.debug(f"Cargando plantilla desde: {ruta_plantilla}")
        wb = openpyxl.load_workbook(ruta_plantilla)
        ws = wb.active
        fila_inicio = 11
        
        for i, registro in enumerate(datos_db):
            fila_actual = fila_inicio + i
            # Mapeo explícito a celdas
            ws.cell(row=fila_actual, column=2).value = registro['Rut Proovedor']   # B
            ws.cell(row=fila_actual, column=3).value = registro['Razon Social']    # C
            ws.cell(row=fila_actual, column=4).value = registro['N documento']     # D
            ws.cell(row=fila_actual, column=5).value = MAPEO_DOCS.get(registro['Tipo Documento'], '99: Otro Documento')  # E (Mapeado)
            ws.cell(row=fila_actual, column=6).value = registro['Monto Documento'] # F
            ws.cell(row=fila_actual, column=7).value = registro['Fecha Emision']   # G
            ws.cell(row=fila_actual, column=12).value = registro['Total Pagar']    # L
            ws.cell(row=fila_actual, column=13).value = MAPEO_PAGO.get('CC')       # M (Mapeado)

        wb.save(ruta_final_salida)
        logger.info(f"Nómina generada con éxito: {nombre_archivo}")

    except Exception as e:
        logger.error(f"Error al llenar Excel: {e}")
        
    except PermissionError:
        logger.error(f"Error: El archivo {nombre_archivo} está abierto. Ciérralo e intenta de nuevo.")
        raise Exception("El archivo de destino está abierto. Por favor, ciérrelo.")
    except Exception as e:
        logger.error(f"Error inesperado al llenar el Excel: {e}")
        raise e
    
def obtener_años_disponibles():
    """
    Intenta conectar a la DB usando la ruta de config.json 
    y retorna la lista de años.
    Returns:
        list: Lista de años disponibles en la tabla cwmovim.
    """
    años = []
    conn = open_database() # Tu función que ya abre la base de datos
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT CpbAno FROM cwmovim ORDER BY CpbAno DESC")
            años = [str(fila[0]) for fila in cursor.fetchall()]
        except Exception as e:
            print(f"Error al obtener años: {e}")
        finally:
            conn.close()
    return años

def inspeccionar_tipos_documento():
    """Consulta los valores únicos y la frecuencia de la columna MovTipDocRef."""
    conn = open_database()
    resultados = []
    if conn:
        try:
            cursor = conn.cursor()
            # Esta consulta trae el valor y cuántas veces se repite
            query = """
                SELECT MovTipDocRef, COUNT(*) as Total
                FROM cwmovim
                GROUP BY MovTipDocRef
                ORDER BY COUNT(*) DESC
            """
            cursor.execute(query)
            resultados = cursor.fetchall()
        except Exception as e:
            print(f"Error al inspeccionar documentos: {e}")
        finally:
            conn.close()
    return resultados

        

if __name__ == "__main__":
    # 1. Intentar conectar
    conn = open_database()
    
    if conn:
        try:
            cursor = conn.cursor()
            
            # --- NUEVA FUNCIÓN: Inspección de tipos ---
            print("\n--- REVISANDO CÓDIGOS DE DOCUMENTOS EN LA DB ---")
            tipos = inspeccionar_tipos_documento() # Asegúrate de que esta función esté definida arriba
            for t in tipos:
                print(f"Código en DB: '{t[0]}' | Cantidad de registros: {t[1]}")
            print("-----------------------------------------------\n")

            # 2. Probar obtención de datos (Ajustado con AÑO)
            # Cambia 123456 y 2025 por datos que sepas que existen para probar
            nomina_test = 12177 
            año_test = 2025
            
            datos = datos_egreso(cursor, nomina_test, año_test)
            
            if datos:
                print(f"Éxito: Se obtuvieron {len(datos)} registros.")
                # 3. Probar generación de Excel
                llenar_nomina_xlsx(datos, str(nomina_test))
            else:
                print(f"Aviso: No se encontraron datos para la nómina {nomina_test} en {año_test}")

        except Exception as e:
            print(f"Error durante la prueba: {e}")
        finally:
            conn.close()
            print("\nConexión cerrada.")
    else:
        print("No se pudo establecer la conexión inicial. Revisa la ruta en config.json.")

