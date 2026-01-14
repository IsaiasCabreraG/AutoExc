import pyodbc # type: ignore
import openpyxl # type: ignore
import os
import json 
import sys
import logging

# Configurar logger para este módulo
logger = logging.getLogger("AutoExcel")

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
MAPEO_BANCOS = {
    '001': '001 BANCO DE CHILE'
}


def log_debug(categoria, mensaje):
    """Helper para logs estructurados en AutoExcel"""
    logger.debug(f"[{categoria}] {mensaje}")


def open_database():
    """Busca la configuración en el JSON y abre la conexión automáticamente."""
    log_debug("DB_CONN", "→ Iniciando conexión a base de datos...")
    
    # 1. Verificar si existe el config.json
    if not os.path.exists("config.json"):
        log_debug("DB_CONN", "✗ No se encontró config.json")
        return None

    # 2. Leer la ruta desde el JSON
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
            db_path = config.get("ruta_db")
        
        log_debug("DB_CONN", f"Ruta DB desde config: {db_path}")
    except Exception as e:
        log_debug("DB_CONN", f"✗ Error al leer config.json: {e}")
        return None

    if not db_path:
        log_debug("DB_CONN", "✗ La ruta de la base de datos no está configurada")
        return None

    # 3. El Driver suele ser fijo para Access (.accdb)
    driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    
    try:
        log_debug("DB_CONN", f"→ Intentando conectar con driver: {driver}")
        conn = pyodbc.connect(f'DRIVER={driver};DBQ={db_path};')
        log_debug("DB_CONN", "✓ Conexión establecida exitosamente")
        return conn
    except Exception as e:
        log_debug("DB_CONN", f"✗✗✗ Error de conexión: {type(e).__name__}: {e}")
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
    log_debug("DB_QUERY", "═══ CONSULTANDO DATOS DE EGRESO ═══")
    
    # 1. Preparar el número de egreso (rellenar con ceros hasta 8 caracteres)
    CpbNum = str(NEgreso).zfill(8)
    CpbAno = str(año)
    codigos_permitidos = tuple(MAPEO_DOCS.keys())
    filtro_sql = str(codigos_permitidos).replace(",)", ")")
    
    log_debug("DB_QUERY", f"Parámetros de búsqueda:")
    log_debug("DB_QUERY", f"  • Egreso original: {NEgreso} → Formateado: {CpbNum}")
    log_debug("DB_QUERY", f"  • Año: {CpbAno}")
    log_debug("DB_QUERY", f"  • Filtro tipos doc: {filtro_sql}")

    # 2. Consulta SQL con las columnas solicitadas
    query = f"""
        SELECT 
            T2.RutAux,
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
        log_debug("DB_QUERY", "→ Ejecutando consulta SQL...")
        cursor.execute(query, (CpbNum, CpbAno))
        filas = cursor.fetchall()
        
        log_debug("DB_QUERY", f"✓ Consulta ejecutada: {len(filas)} filas obtenidas")
        
        # 3. Transformar a estructura de diccionarios (Clave: Valor)
        registros_mapeados = []
        for i, fila in enumerate(filas):
            registro = {
                'Rut Proovedor':   fila[0].strip().replace(".", ""),
                'Razon Social':    fila[1],
                'N documento':     fila[2],
                'Tipo Documento':  fila[3],
                'Monto Documento': fila[4],
                'Fecha Emision':   fila[5],
                'Total Pagar':     fila[4]
            }
            registros_mapeados.append(registro)
            
            if i < 3:  # Mostrar solo los primeros 3 registros en debug
                log_debug("DB_QUERY", f"  Registro {i+1}: RUT={registro['Rut Proovedor']}, "
                         f"Razón Social={registro['Razon Social'][:30]}..., "
                         f"Monto=${registro['Monto Documento']:,.0f}")
        
        if len(filas) > 3:
            log_debug("DB_QUERY", f"  ... y {len(filas) - 3} registros más")
        
        log_debug("DB_QUERY", "═══ CONSULTA FINALIZADA ═══")
        return registros_mapeados
        
    except Exception as e:
        log_debug("DB_QUERY", f"✗✗✗ Error en consulta SQL: {type(e).__name__}: {e}")
        return []


def mostrar_estructura_tabla(connection, tabla):
    """Funcion encargada de mostrar la estructura de una tabla en la base de datos.
    Args:
        connection (pyodbc.Connection): Conexion a la base de datos.
        tabla (str): Nombre de la tabla a inspeccionar.
    """
    log_debug("DB_INSPECT", f"→ Inspeccionando estructura de tabla: {tabla}")
    cursor = connection.cursor()
    
    log_debug("DB_INSPECT", f"{'Columna':<25} | {'Tipo de Dato':<15} | {'Tamaño'}")
    log_debug("DB_INSPECT", "-" * 55)
    
    # El método .columns devuelve metadatos de la tabla
    for row in cursor.columns(table=tabla):
        nombre_col = row.column_name
        tipo_dato = row.type_name
        tamano = row.column_size
        
        log_debug("DB_INSPECT", f"{nombre_col:<25} | {tipo_dato:<15} | {tamano}")


def obtener_ruta_recurso(relativa):
    """ Obtiene la ruta absoluta para recursos (iconos, plantillas, etc) """
    try:
        # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
        base_path = sys._MEIPASS
        log_debug("RESOURCE", f"Ejecutando desde PyInstaller: {base_path}")
    except Exception:
        base_path = os.path.abspath(".")
        log_debug("RESOURCE", f"Ejecutando desde script: {base_path}")

    ruta_completa = os.path.join(base_path, relativa)
    log_debug("RESOURCE", f"Recurso solicitado: {relativa} → {ruta_completa}")
    
    return ruta_completa


def llenar_nomina_xlsx(datos_db, numero_nomina, año, RutEmpresa=""):
    """
    Llena la plantilla de Excel usando diccionarios mapeados.
    """
    log_debug("EXCEL", "╔═══════════════════════════════════════╗")
    log_debug("EXCEL", "║     GENERANDO ARCHIVO EXCEL         ║")
    log_debug("EXCEL", "╚═══════════════════════════════════════╝")
    
    log_debug("EXCEL", f"Parámetros:")
    log_debug("EXCEL", f"  • Nómina: {numero_nomina}")
    log_debug("EXCEL", f"  • Año: {año}")
    log_debug("EXCEL", f"  • RUT Empresa: {RutEmpresa}")
    log_debug("EXCEL", f"  • Registros a procesar: {len(datos_db)}")

    if not os.path.exists("config.json"):
        log_debug("EXCEL", "✗ No se encontró config.json para obtener la ruta de salida")
        return None

    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
            carpeta_salida = config.get("ruta_output", "")
        
        log_debug("EXCEL", f"Carpeta de salida: {carpeta_salida}")
    except Exception as e:
        log_debug("EXCEL", f"✗ Error al leer config.json: {e}")
        return None

    if not carpeta_salida or not os.path.exists(carpeta_salida):
        log_debug("EXCEL", f"✗ La ruta de salida no es válida: {carpeta_salida}")
        return None

    nombre_archivo = f"Nomina_{numero_nomina}-{año}.xlsx"
    ruta_final_salida = os.path.join(carpeta_salida, nombre_archivo)
    ruta_plantilla = obtener_ruta_recurso("static/Plantilla_NominaPagoV2.xlsx")
    RutEmpresa = RutEmpresa.strip(".-")

    log_debug("EXCEL", f"Archivo destino: {nombre_archivo}")
    log_debug("EXCEL", f"Ruta completa: {ruta_final_salida}")

    try:
        log_debug("EXCEL", "→ Cargando plantilla Excel...")
        
        if not os.path.exists(ruta_plantilla):
            log_debug("EXCEL", f"✗ Plantilla no encontrada: {ruta_plantilla}")
            return None
        
        wb = openpyxl.load_workbook(ruta_plantilla)
        ws = wb.active
        log_debug("EXCEL", f"✓ Plantilla cargada: {ruta_plantilla}")
        
        fila_inicio = 11
        
        # Encabezado
        log_debug("EXCEL", "→ Escribiendo encabezado...")
        ws.cell(row=5, column=3).value = int(str(numero_nomina) + str(año))  # C5
        ws.cell(row=6, column=3).value = RutEmpresa     # C6
        log_debug("EXCEL", f"  • C5 (ID Nómina): {numero_nomina}{año}")
        log_debug("EXCEL", f"  • C6 (RUT): {RutEmpresa}")
        
        # Datos
        log_debug("EXCEL", f"→ Escribiendo {len(datos_db)} registros...")
        totalPago = 0
        
        for i, registro in enumerate(datos_db):
            totalPago += int(registro['Total Pagar'])
            fila_actual = fila_inicio + i
            
            # Mapeo explícito a celdas
            ws.cell(row=fila_actual, column=2).value = registro['Rut Proovedor']   # B
            ws.cell(row=fila_actual, column=3).value = registro['Razon Social']    # C
            ws.cell(row=fila_actual, column=4).value = registro['N documento']     # D
            ws.cell(row=fila_actual, column=5).value = MAPEO_DOCS.get(registro['Tipo Documento'], '99: Otro Documento')  # E
            ws.cell(row=fila_actual, column=6).value = registro['Monto Documento'] # F
            ws.cell(row=fila_actual, column=7).value = registro['Fecha Emision']   # G
            ws.cell(row=fila_actual, column=12).value = registro['Total Pagar']    # L
            ws.cell(row=fila_actual, column=13).value = MAPEO_PAGO.get('CC')       # M
            ws.cell(row=fila_actual, column=14).value = MAPEO_BANCOS.get('001')    # N
            
            if i < 3:
                log_debug("EXCEL", f"  Fila {fila_actual}: {registro['Razon Social'][:30]}... - ${registro['Total Pagar']:,.0f}")
        
        if len(datos_db) > 3:
            log_debug("EXCEL", f"  ... {len(datos_db) - 3} filas más escritas")
        
        log_debug("EXCEL", f"✓ Total acumulado: ${totalPago:,.0f}")
        
        # Guardar
        log_debug("EXCEL", "→ Guardando archivo...")
        wb.save(ruta_final_salida)
        log_debug("EXCEL", f"✓✓✓ Archivo guardado exitosamente: {nombre_archivo}")
        log_debug("EXCEL", "╚═══════════════════════════════════════╝")
        
        return totalPago

    except PermissionError:
        log_debug("EXCEL", f"✗✗✗ ERROR: Archivo bloqueado - {nombre_archivo} está abierto")
        logger.error(f"El archivo {nombre_archivo} está abierto. Ciérralo e intenta de nuevo.")
        raise Exception("El archivo de destino está abierto. Por favor, ciérrelo.")
    
    except FileNotFoundError as e:
        log_debug("EXCEL", f"✗✗✗ ERROR: Archivo no encontrado: {e}")
        raise e
    
    except Exception as e:
        log_debug("EXCEL", f"✗✗✗ ERROR INESPERADO: {type(e).__name__}: {e}")
        import traceback
        log_debug("EXCEL", f"Traceback:\n{traceback.format_exc()}")
        raise e

    
def obtener_años_disponibles():
    """
    Intenta conectar a la DB usando la ruta de config.json 
    y retorna la lista de años.
    Returns:
        list: Lista de años disponibles en la tabla cwmovim.
    """
    log_debug("DB_AÑOS", "→ Obteniendo años disponibles...")
    años = []
    conn = open_database()
    
    if conn:
        try:
            cursor = conn.cursor()
            log_debug("DB_AÑOS", "→ Ejecutando consulta DISTINCT de años...")
            cursor.execute("SELECT DISTINCT CpbAno FROM cwmovim ORDER BY CpbAno DESC")
            años = [str(fila[0]) for fila in cursor.fetchall()]
            log_debug("DB_AÑOS", f"✓ {len(años)} años encontrados: {', '.join(años)}")
        except Exception as e:
            log_debug("DB_AÑOS", f"✗ Error al obtener años: {type(e).__name__}: {e}")
        finally:
            conn.close()
            log_debug("DB_AÑOS", "Conexión cerrada")
    else:
        log_debug("DB_AÑOS", "✗ No se pudo conectar a la base de datos")
    
    return años


def inspeccionar_tipos_documento():
    """Consulta los valores únicos y la frecuencia de la columna MovTipDocRef."""
    log_debug("DB_INSPECT", "→ Inspeccionando tipos de documento...")
    conn = open_database()
    resultados = []
    
    if conn:
        try:
            cursor = conn.cursor()
            query = """
                SELECT MovTipDocRef, COUNT(*) as Total
                FROM cwmovim
                GROUP BY MovTipDocRef
                ORDER BY COUNT(*) DESC
            """
            cursor.execute(query)
            resultados = cursor.fetchall()
            
            log_debug("DB_INSPECT", f"✓ {len(resultados)} tipos de documento encontrados:")
            for t in resultados:
                log_debug("DB_INSPECT", f"  • Código: '{t[0]}' | Registros: {t[1]}")
                
        except Exception as e:
            log_debug("DB_INSPECT", f"✗ Error al inspeccionar documentos: {e}")
        finally:
            conn.close()
    
    return resultados

        

if __name__ == "__main__":
    # Configurar logging para pruebas standalone
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(levelname)s: %(message)s'
    )
    
    log_debug("TEST", "╔═══════════════════════════════════════╗")
    log_debug("TEST", "║      INICIANDO PRUEBAS DE MÓDULO    ║")
    log_debug("TEST", "╚═══════════════════════════════════════╝")
    
    # 1. Intentar conectar
    conn = open_database()
    
    if conn:
        try:
            cursor = conn.cursor()
            
            # --- Inspección de tipos ---
            log_debug("TEST", "\n--- REVISANDO CÓDIGOS DE DOCUMENTOS EN LA DB ---")
            tipos = inspeccionar_tipos_documento()
            log_debug("TEST", "-----------------------------------------------\n")

            # 2. Probar obtención de datos
            nomina_test = 12177 
            año_test = 2025
            
            log_debug("TEST", f"→ Probando con nómina {nomina_test}, año {año_test}")
            datos = datos_egreso(cursor, nomina_test, año_test)
            
            if datos:
                log_debug("TEST", f"✓ Prueba exitosa: {len(datos)} registros obtenidos")
                # 3. Probar generación de Excel
                total = llenar_nomina_xlsx(datos, str(nomina_test), str(año_test))
                if total:
                    log_debug("TEST", f"✓✓✓ Excel generado con total: ${total:,.0f}")
            else:
                log_debug("TEST", f"⚠️ No se encontraron datos para la nómina {nomina_test} en {año_test}")

        except Exception as e:
            log_debug("TEST", f"✗✗✗ Error durante la prueba: {type(e).__name__}: {e}")
        finally:
            conn.close()
            log_debug("TEST", "Conexión cerrada")
            log_debug("TEST", "╚═══════════════════════════════════════╝")
    else:
        log_debug("TEST", "✗ No se pudo establecer la conexión inicial")
        log_debug("TEST", "Revisa la ruta en config.json")