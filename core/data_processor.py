import pandas as pd
import numpy as np

# ==========================================
# SIMULACIÓN DE DATOS DE ENTRADA (Tus archivos)
# ==========================================

# 1. Simulamos la Base de Datos de Colaboradores (Pestaña 'Colaboradores')
# Nota cómo incluimos demografías planas y multinivel (Área de Desempeño 1, 2, 3)
data_colaboradores = {
    'Identificación': [101, 102, 103, 104, 105],
    'Nombre': ['Carlos Pérez', 'Ana Gómez', 'Luis Martínez', 'María Rodriguez', 'Jorge Ruiz'],
    'E-Mail': ['carlos@empresa.com', 'ana@empresa.com', 'luis@empresa.com', 'maria@empresa.com', 'jorge@empresa.com'],
    'Lugar de Trabajo': ['Piedecuesta', 'Bucaramanga', 'Piedecuesta', 'Floridablanca', 'Bucaramanga'],
    'Nivel de Cargo': ['Coordinador', 'Operativo', 'Gerente', 'Operativo', 'Coordinador'],
    'Área de Desempeño 1': ['Gerencia Hospital', 'Gerencia Hospital', 'Presidencia', 'Gerencia Hospital', 'Presidencia'],
    'Área de Desempeño 2': ['Urgencias', 'Urgencias', 'Finanzas', 'Pediatría', 'Finanzas'],
    'Área de Desempeño 3': ['Enfermería TAM', 'Médicos', 'Contabilidad', 'Enfermería Hospital', 'Tesorería']
}
df_colaboradores_mock = pd.DataFrame(data_colaboradores)

# 2. Simulamos el reporte de participación diario (Pestaña 'Reporte de participación')
# Solo contiene las identificaciones de los que YA contestaron
data_participacion = {
    'Identificación': [101, 103, 105],
    'Fecha Respuesta': ['2026-06-12 08:00', '2026-06-12 08:30', '2026-06-12 09:15']
}
df_participacion_mock = pd.DataFrame(data_participacion)


# ==========================================
# MOTOR DE PROCESAMIENTO EN PYTHON
# ==========================================

def procesar_datos_participacion(df_colaboradores, df_participacion):
    """
    Función principal que identifica demografías dinámicamente,
    realiza el XLOOKUP (merge) y calcula los conteos para los reportes.
    """
    
    # PASO 1: Identificar demografías dinámicamente
    # Definimos cuáles son las columnas fijas obligatorias
    columnas_fijas = ['Identificación', 'Nombre', 'E-Mail']
    
    # Cualquier columna que NO sea fija, automáticamente se vuelve una demografía
    demografias_detectadas = [col for col in df_colaboradores.columns if col not in columnas_fijas]
    
    print(f"🔍 Demografías detectadas dinámicamente: {demografias_detectadas}\n")
    
    # PASO 2: El XLOOKUP Automático usando el E-Mail como llave
    # Traemos todos los colaboradores y les cruzamos el reporte de participación por correo
    df_cruce = pd.merge(df_colaboradores, df_participacion, on='E-Mail', how='left')
    
    # Creamos la columna booleana para saber quién contestó
    # (Suponiendo que el reporte de participación trae una columna llamada 'Fecha Respuesta' o similar)
    df_cruce['Contestó'] = df_cruce['Fecha Respuesta'].notna()
    
    # PASO 3: Generar los conteos estilo "Rango H66:I70" para CADA demografía
    resultados_reporte = {}
    
    for demografia in demografias_detectadas:
        # Agrupamos por la demografía y contamos:
        # 1. Cuántos colaboradores totales hay en esa opción (Total Colaboradores)
        # 2. Cuántos de ellos tienen 'Contestó' en True (Han contestado)
        conteo = df_cruce.groupby(demografia).agg(
            Total_Colaboradores=('Identificación', 'count'),
            Han_Contestado=('Contestó', 'sum') # Sumar True cuenta como 1
        ).reset_index()
        
        # Guardamos el resultado de esta demografía en nuestro diccionario de reportes
        resultados_reporte[demografia] = conteo
        
    return resultados_reporte, df_cruce


# ==========================================
# EJECUCIÓN DEL PROCESO
# ==========================================

# Ejecutamos la función enviándole los dataframes simuldados
reportes_finales, df_detallado = procesar_datos_participacion(df_colaboradores_mock, df_participacion_mock)

# ==========================================
# DEMOSTRACIÓN DE RESULTADOS (Lo que iría a la Web App)
# ==========================================

# Ejemplo 1: Ver el reporte de "Lugar de Trabajo"
print("📊 REPORTE: Lugar de Trabajo")
print(reportes_finales['Lugar de Trabajo'].to_string(index=False))
print("-" * 40)

# Ejemplo 2: Ver el reporte multinivel de "Área de Desempeño 1"
print("📊 REPORTE JERÁRQUICO: Área de Desempeño 1")
print(reportes_finales['Área de Desempeño 1'].to_string(index=False))
print("-" * 40)

# Ejemplo 3: Ver el reporte multinivel profundo de "Área de Desempeño 3"
print("📊 REPORTE JERÁRQUICO COMPLETO: Área de Desempeño 3")
print(reportes_finales['Área de Desempeño 3'].to_string(index=False))
print("-" * 40)