import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import io
import warnings
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime

warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(
    page_title="Análisis Gestión Inmobiliaria",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("📊 Sistema de Análisis Inmobiliario")
st.markdown("---")

# Hojas objetivo (nombres exactos del Excel)
HOJAS_OBJETIVO = ['Operaciones', 'Comisiones-equipo', 'Cuotas-comisiones', 'Clientes']

# ==================== FUNCIONES DE PROCESAMIENTO ====================

def limpiar_monto(valor):
    """Convierte valores de monto a float, manejando strings con símbolos y espacios"""
    if pd.isna(valor):
        return np.nan
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor_limpio = valor.replace('$', '').replace(',', '').replace('.', '').replace(' ', '').strip()
        try:
            return float(valor_limpio)
        except:
            return np.nan
    return np.nan

@st.cache_data
def cargar_y_procesar_excel(uploaded_file):
    """Carga y procesa el archivo Excel completo"""
    try:
        dataframes = {}
        errores_carga = {}
        
        # Cargar todas las hojas
        for hoja in HOJAS_OBJETIVO:
            try:
                df_temp = pd.read_excel(uploaded_file, sheet_name=hoja)
                
                # Para Comisiones-equipo: detectar y quedarse solo con la primera tabla
                if hoja == 'Comisiones-equipo':
                    filas_vacias = df_temp.isnull().all(axis=1)
                    indices_vacias = df_temp[filas_vacias].index.tolist()
                    
                    if indices_vacias:
                        primera_fila_vacia = indices_vacias[0]
                        df_temp = df_temp.iloc[:primera_fila_vacia]
                    
                    columnas_vacias = df_temp.isnull().all(axis=0)
                    if columnas_vacias.any():
                        df_temp = df_temp.loc[:, ~columnas_vacias]
                
                dataframes[hoja] = df_temp
            except Exception as e:
                errores_carga[hoja] = str(e)
        
        # Limpieza de datos
        columnas_a_eliminar = {
            'Operaciones': ['Observaciones'],
            'Comisiones-equipo': ['Rol', 'Agentes', '$', 'Rol.1', 'Tipo operación', '% comisión'],
            'Cuotas-comisiones': ['Metodo de pago'],
            'Clientes': ['Teléfono', 'DNI o CUIT']
        }
        
        for nombre_hoja, df in dataframes.items():
            # Eliminar filas completamente vacías
            df = df.dropna(how='all')
            
            # Eliminar filas sin Nº Operación
            columna_operacion = None
            if 'Nº Operación' in df.columns:
                columna_operacion = 'Nº Operación'
            elif 'Nº de operación' in df.columns:
                columna_operacion = 'Nº de operación'
            
            if columna_operacion:
                df = df[df[columna_operacion].notna()].copy()
                if df[columna_operacion].dtype == 'object':
                    df = df[df[columna_operacion].astype(str).str.strip() != ''].copy()
            
            # Eliminar columnas específicas
            if nombre_hoja in columnas_a_eliminar:
                columnas_eliminar = columnas_a_eliminar[nombre_hoja]
                columnas_encontradas = [col for col in columnas_eliminar if col in df.columns]
                if columnas_encontradas:
                    df = df.drop(columns=columnas_encontradas)
            
            dataframes[nombre_hoja] = df
        
        # Normalización de datos
        columnas_montos = {
            'Operaciones': ['Comisión total', 'Cobrado', 'Saldo'],
            'Comisiones-equipo': ['Monto comisión', 'Saldo a pagar'],
            'Cuotas-comisiones': ['Importe'],
            'Clientes': []
        }
        
        columnas_enteros = {
            'Operaciones': ['Nº Operación', 'Cuotas', 'Días estimados'],
            'Comisiones-equipo': ['Nº Operación'],
            'Cuotas-comisiones': ['Nº Operación', 'Cuota'],
            'Clientes': ['Edad', 'Nº de operación']
        }
        
        for nombre_hoja, df in dataframes.items():
            # Normalizar montos
            if nombre_hoja in columnas_montos:
                for col_monto in columnas_montos[nombre_hoja]:
                    if col_monto in df.columns:
                        df[col_monto] = df[col_monto].apply(limpiar_monto)
                        df[col_monto] = df[col_monto].round(2)
            
            # Normalizar enteros
            if nombre_hoja in columnas_enteros:
                for col_entero in columnas_enteros[nombre_hoja]:
                    if col_entero in df.columns:
                        df[col_entero] = pd.to_numeric(df[col_entero], errors='coerce').astype('Int64')
            
            dataframes[nombre_hoja] = df
        
        return dataframes, errores_carga
    except Exception as e:
        return None, {'general': str(e)}

def generar_reporte_operaciones(df_ops):
    """Genera el reporte de operaciones"""
    reporte = {}
    
    # Resumen financiero
    total_cobrado = df_ops['Cobrado'].sum() if 'Cobrado' in df_ops.columns else 0
    total_comision = df_ops['Comisión total'].sum() if 'Comisión total' in df_ops.columns else 0
    total_saldo = df_ops['Saldo'].sum() if 'Saldo' in df_ops.columns else 0
    
    reporte['resumen_financiero'] = {
        'Total Comisión': total_comision,
        'Total Cobrado': total_cobrado,
        'Total Saldo Pendiente': total_saldo,
        'Porcentaje Cobrado': (total_cobrado/total_comision*100) if total_comision > 0 else 0
    }
    
    # Indicadores por tipo
    if 'Tipo' in df_ops.columns:
        tipo_counts = df_ops['Tipo'].value_counts()
        reporte['indicadores'] = {
            'Alquileres': tipo_counts.get('alquiler', 0),
            'Ventas': tipo_counts.get('venta', 0),
            'Total Operaciones': len(df_ops)
        }
    
    # Operaciones con saldo pendiente
    if 'Saldo' in df_ops.columns:
        operaciones_pendientes = df_ops[df_ops['Saldo'] > 0].copy()
        reporte['operaciones_pendientes'] = operaciones_pendientes
    
    # KPI: Comisión total por Agente
    if 'Agente' in df_ops.columns and 'Comisión total' in df_ops.columns:
        # Filtrar registros válidos (con agente y comisión)
        df_ops_agentes = df_ops[
            df_ops['Agente'].notna() & 
            (df_ops['Agente'].astype(str).str.strip() != '') &
            df_ops['Comisión total'].notna()
        ].copy()
        
        if len(df_ops_agentes) > 0:
            comision_por_agente = df_ops_agentes.groupby('Agente')['Comisión total'].agg(['sum', 'count']).reset_index()
            comision_por_agente.columns = ['Agente', 'Comisión Total', 'Cantidad Operaciones']
            comision_por_agente['Comisión Total'] = comision_por_agente['Comisión Total'].round(2)
            comision_por_agente = comision_por_agente.sort_values('Comisión Total', ascending=False)
            reporte['comision_por_agente'] = comision_por_agente
    
    return reporte

def generar_reporte_comisiones(df_com):
    """Genera el reporte de comisiones por agente"""
    # Calcular pendiente y pagado
    def calcular_pendiente(row):
        monto = row['Monto comisión']
        pagado = str(row['Pagado']).strip().upper() if pd.notna(row['Pagado']) else ''
        porcentaje = row['Porcentaje pagado'] if pd.notna(row['Porcentaje pagado']) else 0
        
        if pagado == 'NO':
            return monto
        elif pagado == 'SI' and porcentaje < 1.0:
            return monto * (1 - porcentaje)
        else:
            return 0.0
    
    def calcular_pagado(row):
        monto = row['Monto comisión']
        pagado = str(row['Pagado']).strip().upper() if pd.notna(row['Pagado']) else ''
        porcentaje = row['Porcentaje pagado'] if pd.notna(row['Porcentaje pagado']) else 0
        
        if pagado == 'SI':
            return monto * porcentaje
        else:
            return 0.0
    
    df_com['Pendiente'] = df_com.apply(calcular_pendiente, axis=1)
    df_com['Pagado_calculado'] = df_com.apply(calcular_pagado, axis=1)
    
    # Filtrar registros válidos
    if 'Nº Operación' in df_com.columns:
        df_com_limpio = df_com[df_com['Nº Operación'].notna()].copy()
    else:
        df_com_limpio = df_com.copy()
    
    df_com_limpio = df_com_limpio[df_com_limpio['Agente'].notna() & 
                                   (df_com_limpio['Agente'].astype(str).str.strip() != '')].copy()
    
    # Agrupar por agente
    resumen_agentes = df_com_limpio.groupby('Agente').agg({
        'Monto comisión': 'sum',
        'Pagado_calculado': 'sum',
        'Pendiente': 'sum'
    }).reset_index()
    
    resumen_agentes.columns = ['Agente', 'Total Comisión', 'Total Pagado', 'Total Pendiente']
    resumen_agentes['Total Comisión'] = resumen_agentes['Total Comisión'].round(2)
    resumen_agentes['Total Pagado'] = resumen_agentes['Total Pagado'].round(2)
    resumen_agentes['Total Pendiente'] = resumen_agentes['Total Pendiente'].round(2)
    
    return resumen_agentes

def generar_reporte_clientes(df_clientes):
    """Genera el reporte de clientes"""
    reporte = {}
    
    total_clientes = len(df_clientes)
    reporte['total_clientes'] = total_clientes
    
    if 'Edad' in df_clientes.columns:
        edades_validas = df_clientes[df_clientes['Edad'] <= 70]['Edad']
        edades_invalidas = df_clientes[df_clientes['Edad'] > 70]
        
        if len(edades_validas) > 0:
            reporte['edad'] = {
                'promedio': edades_validas.mean(),
                'minima': edades_validas.min(),
                'maxima': edades_validas.max(),
                'clientes_mayores_70': len(edades_invalidas)
            }
        
        # Rango etario
        def clasificar_edad(edad):
            if pd.isna(edad) or edad > 70:
                return 'Sin clasificar (>70 o sin dato)'
            elif edad < 25:
                return 'Menos de 25 años'
            elif edad < 35:
                return '25-34 años'
            elif edad < 45:
                return '35-44 años'
            elif edad < 55:
                return '45-54 años'
            elif edad <= 70:
                return '55-70 años'
            else:
                return 'Sin clasificar (>70 o sin dato)'
        
        df_clientes['Rango Etario'] = df_clientes['Edad'].apply(clasificar_edad)
        distribucion_etaria = df_clientes['Rango Etario'].value_counts().sort_index()
        
        reporte['distribucion_etaria'] = pd.DataFrame({
            'Rango Etario': distribucion_etaria.index,
            'Cantidad': distribucion_etaria.values,
            'Porcentaje': (distribucion_etaria.values / total_clientes * 100).round(2)
        })
    
    return reporte

def generar_flujo_caja(dataframes):
    """Genera el informe de flujo de caja"""
    # Lo cobrado
    cobrado_total = 0
    if 'Operaciones' in dataframes:
        df_ops = dataframes['Operaciones']
        if 'Cobrado' in df_ops.columns:
            cobrado_total = df_ops['Cobrado'].sum()
    
    # Lo por cobrar
    por_cobrar_total = 0
    if 'Operaciones' in dataframes:
        df_ops = dataframes['Operaciones']
        if 'Saldo' in df_ops.columns:
            por_cobrar_total = df_ops['Saldo'].sum()
    
    # Lo por pagar y lo pagado
    por_pagar_total = 0
    pagado_total = 0
    
    if 'Comisiones-equipo' in dataframes:
        df_com = dataframes['Comisiones-equipo']
        
        def calcular_pendiente(row):
            monto = row['Monto comisión'] if 'Monto comisión' in row else 0
            pagado = str(row['Pagado']).strip().upper() if pd.notna(row.get('Pagado')) else ''
            porcentaje = row['Porcentaje pagado'] if pd.notna(row.get('Porcentaje pagado')) else 0
            
            if pagado == 'NO':
                return monto
            elif pagado == 'SI' and porcentaje < 1.0:
                return monto * (1 - porcentaje)
            else:
                return 0.0
        
        def calcular_pagado(row):
            monto = row['Monto comisión'] if 'Monto comisión' in row else 0
            pagado = str(row['Pagado']).strip().upper() if pd.notna(row.get('Pagado')) else ''
            porcentaje = row['Porcentaje pagado'] if pd.notna(row.get('Porcentaje pagado')) else 0
            
            if pagado == 'SI':
                return monto * porcentaje
            else:
                return 0.0
        
        if 'Pendiente' not in df_com.columns:
            df_com['Pendiente'] = df_com.apply(calcular_pendiente, axis=1)
        if 'Pagado_calculado' not in df_com.columns:
            df_com['Pagado_calculado'] = df_com.apply(calcular_pagado, axis=1)
        
        if 'Nº Operación' in df_com.columns:
            df_com_limpio = df_com[df_com['Nº Operación'].notna()].copy()
        else:
            df_com_limpio = df_com.copy()
        
        if 'Pendiente' in df_com_limpio.columns:
            por_pagar_total = df_com_limpio['Pendiente'].sum()
        if 'Pagado_calculado' in df_com_limpio.columns:
            pagado_total = df_com_limpio['Pagado_calculado'].sum()
    
    balance_neto = cobrado_total - pagado_total
    flujo_futuro_neto = por_cobrar_total - por_pagar_total
    
    return {
        'cobrado_total': cobrado_total,
        'pagado_total': pagado_total,
        'por_cobrar_total': por_cobrar_total,
        'por_pagar_total': por_pagar_total,
        'balance_neto': balance_neto,
        'flujo_futuro_neto': flujo_futuro_neto
    }

def generar_pdf(dataframes, reportes):
    """Genera un PDF con todo el reporte"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()
    
    # Título
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=30
    )
    story.append(Paragraph("📊 Reporte Ejecutivo - Gestión Inmobiliaria", title_style))
    story.append(Paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    # Sección Operaciones
    story.append(Paragraph("💰 OPERACIONES", styles['Heading2']))
    if 'Operaciones' in reportes:
        ops = reportes['Operaciones']
        if 'resumen_financiero' in ops:
            rf = ops['resumen_financiero']
            data = [
                ['Concepto', 'Monto'],
                ['Total Comisión', f"${rf['Total Comisión']:,.2f}"],
                ['Total Cobrado', f"${rf['Total Cobrado']:,.2f}"],
                ['Total Saldo Pendiente', f"${rf['Total Saldo Pendiente']:,.2f}"],
                ['Porcentaje Cobrado', f"{rf['Porcentaje Cobrado']:.2f}%"]
            ]
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(table)
            story.append(Spacer(1, 0.2*inch))
    
    story.append(PageBreak())
    
    # Sección Comisiones
    story.append(Paragraph("👥 COMISIONES POR AGENTE", styles['Heading2']))
    if 'Comisiones-equipo' in reportes:
        com_df = reportes['Comisiones-equipo']
        if isinstance(com_df, pd.DataFrame) and len(com_df) > 0:
            # Convertir DataFrame a lista para la tabla
            data = [['Agente', 'Total Comisión', 'Total Pagado', 'Total Pendiente']]
            for _, row in com_df.iterrows():
                data.append([
                    str(row['Agente']),
                    f"${row['Total Comisión']:,.2f}",
                    f"${row['Total Pagado']:,.2f}",
                    f"${row['Total Pendiente']:,.2f}"
                ])
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(table)
            story.append(Spacer(1, 0.2*inch))
    
    story.append(PageBreak())
    
    # Sección Clientes
    story.append(Paragraph("👥 CLIENTES", styles['Heading2']))
    if 'Clientes' in reportes:
        cli = reportes['Clientes']
        if 'total_clientes' in cli:
            story.append(Paragraph(f"Total de Clientes: {cli['total_clientes']}", styles['Normal']))
        if 'edad' in cli:
            edad = cli['edad']
            story.append(Paragraph(f"Promedio de Edad: {edad['promedio']:.1f} años", styles['Normal']))
            story.append(Paragraph(f"Edad Mínima: {edad['minima']} años", styles['Normal']))
            story.append(Paragraph(f"Edad Máxima: {edad['maxima']} años", styles['Normal']))
        story.append(Spacer(1, 0.2*inch))
    
    story.append(PageBreak())
    
    # Sección Flujo de Caja
    story.append(Paragraph("💰 FLUJO DE CAJA", styles['Heading2']))
    if 'Flujo de Caja' in reportes:
        fc = reportes['Flujo de Caja']
        data = [
            ['Concepto', 'Monto'],
            ['Total Cobrado', f"${fc['cobrado_total']:,.2f}"],
            ['Total Pagado', f"${fc['pagado_total']:,.2f}"],
            ['Por Cobrar', f"${fc['por_cobrar_total']:,.2f}"],
            ['Por Pagar', f"${fc['por_pagar_total']:,.2f}"],
            ['Balance Neto', f"${fc['balance_neto']:,.2f}"],
            ['Flujo Futuro Neto', f"${fc['flujo_futuro_neto']:,.2f}"]
        ]
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# ==================== FUNCIONES PORTFOLIO ====================

import re

def procesar_precio_y_tipo_portfolio(precio):
    """Procesa la columna Precio para detectar tipo de operación"""
    if pd.isna(precio):
        return None, None
    
    precio_str = str(precio).strip()
    precio_str_upper = precio_str.upper()
    
    # CASO 1: "consultar precio"
    if 'CONSULTAR PRECIO' in precio_str_upper:
        return None, 'Venta'
    
    # CASO 2: U$D o USD
    if 'U$D' in precio_str_upper or 'USD' in precio_str_upper:
        numeros = re.sub(r'[^\d.,]', '', precio_str)
        numeros = numeros.replace('.', '').replace(',', '.')
        try:
            precio_numerico = float(numeros) if numeros else None
            return precio_numerico, 'Venta'
        except:
            return None, 'Venta'
    
    # CASO 3: Precio 0
    numeros_temp = re.sub(r'[^\d.,]', '', precio_str)
    numeros_temp = numeros_temp.replace('.', '').replace(',', '.')
    try:
        precio_temp = float(numeros_temp) if numeros_temp else None
        if precio_temp == 0:
            return 0.0, 'Venta'
    except:
        pass
    
    # CASO 4: Alquiler (solo número)
    numeros = re.sub(r'[^\d.,]', '', precio_str)
    numeros = numeros.replace(',', '.')
    try:
        precio_numerico = float(numeros) if numeros else None
        return precio_numerico, 'Alquiler'
    except:
        return None, 'Alquiler'

def extraer_superficie_numerica(superficie):
    """Extrae el valor numérico de Superficie total (formato latino)"""
    if pd.isna(superficie):
        return None
    
    superficie_str = str(superficie).strip()
    
    # Buscar posición de "m" (ignorar todo desde ahí)
    pos_m = -1
    for i, char in enumerate(superficie_str):
        if char.lower() == 'm':
            pos_m = i
            break
    
    if pos_m > 0:
        superficie_str = superficie_str[:pos_m].strip()
    
    # Formato latino: punto (.) es separador de miles, coma (,) es separador decimal
    numeros = re.sub(r'[^\d.,]', '', superficie_str)
    
    if ',' in numeros:
        numeros = numeros.replace('.', '').replace(',', '.')
    else:
        numeros = numeros.replace('.', '')
    
    try:
        return float(numeros) if numeros else None
    except:
        return None

@st.cache_data
def cargar_y_procesar_portfolio(uploaded_file):
    """Carga y procesa el archivo Excel de Portfolio"""
    try:
        # Cargar todas las hojas
        excel_file = pd.ExcelFile(uploaded_file)
        dataframes = {}
        
        for hoja in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=hoja)
            dataframes[hoja] = df
        
        # Procesar cada hoja
        for nombre_hoja, df in dataframes.items():
            # Limpieza: eliminar filas vacías
            df = df.dropna(how='all')
            
            # Eliminar columnas especificadas
            columnas_a_eliminar = [
                'Id Aviso', 'Superficie cubierta', 'Fondo libre', 'Antiguedad',
                'Cocheras descubiertas', 'Cocheras cubiertas', 'Cocheras semicubiertas',
                'Plantas', 'Nombre de propietario', 'Email de propietario', 'Celular de propietario'
            ]
            columnas_existentes = [col for col in columnas_a_eliminar if col in df.columns]
            if columnas_existentes:
                df = df.drop(columns=columnas_existentes)
            
            # Procesar Precio y Tipo de Operación
            if 'Precio' in df.columns:
                resultados = df['Precio'].apply(procesar_precio_y_tipo_portfolio)
                df['Precio_numerico'] = resultados.apply(lambda x: x[0])
                df['Tipo de operación'] = resultados.apply(lambda x: x[1])
            
            # Procesar Superficie Total
            if 'Superficie total' in df.columns:
                df['Superficie_total_num'] = df['Superficie total'].apply(extraer_superficie_numerica)
            
            dataframes[nombre_hoja] = df
        
        return dataframes, {}
    except Exception as e:
        return None, {'general': str(e)}

def generar_reporte_portfolio(df):
    """Genera todos los reportes del Portfolio"""
    reportes = {}
    
    # KPI 1: Total de propiedades
    reportes['total_propiedades'] = len(df)
    
    # KPI 2: Distribución por Tipo de Operación
    if 'Tipo de operación' in df.columns:
        reportes['distribucion_operacion'] = df['Tipo de operación'].value_counts().to_dict()
    
    # KPI 3: Distribución por Tipo de Propiedad
    if 'Tipo de propiedad' in df.columns:
        reportes['distribucion_tipo_propiedad'] = df['Tipo de propiedad'].value_counts().to_dict()
    
    # KPI 4: Distribución por Estado
    if 'Estado' in df.columns:
        reportes['distribucion_estado'] = df['Estado'].value_counts().to_dict()
    
    # KPI 5: Precios Promedio por Tipo de Operación
    if 'Precio_numerico' in df.columns and 'Tipo de operación' in df.columns:
        precios_por_tipo = df.groupby('Tipo de operación')['Precio_numerico'].agg(['mean', 'median', 'min', 'max', 'count'])
        reportes['precios_por_operacion'] = precios_por_tipo.to_dict('index')
    
    # KPI 6: Precios Promedio por Tipo de Propiedad
    if 'Precio_numerico' in df.columns and 'Tipo de propiedad' in df.columns:
        precios_por_tipo_prop = df.groupby('Tipo de propiedad')['Precio_numerico'].agg(['mean', 'median', 'min', 'max', 'count'])
        reportes['precios_por_tipo_propiedad'] = precios_por_tipo_prop.to_dict('index')
    
    # KPI 7: Top 10 Ubicaciones
    if 'Ubicación' in df.columns:
        reportes['top_ubicaciones'] = df['Ubicación'].value_counts().head(10).to_dict()
    
    # KPI 8: Distribución por Usuario Asignado
    if 'Usuario asignado' in df.columns:
        reportes['distribucion_usuario'] = df['Usuario asignado'].value_counts().to_dict()
    
    # KPI 9-10: Dormitorios y Baños (excluyendo TERRENO, CAMPO, GALPON, LOCAL)
    tipos_a_excluir = ['TERRENO', 'CAMPO', 'GALPON', 'LOCAL']
    if 'Dormitorios' in df.columns and 'Tipo de propiedad' in df.columns:
        mask_dorm = df['Tipo de propiedad'].notna()
        mask_dorm = mask_dorm & (~df['Tipo de propiedad'].str.upper().isin(tipos_a_excluir))
        df_dorm = df[mask_dorm]
        if len(df_dorm) > 0:
            dorm_clean = pd.to_numeric(df_dorm['Dormitorios'], errors='coerce')
            reportes['distribucion_dormitorios'] = dorm_clean.value_counts().to_dict()
    
    if 'Baños' in df.columns and 'Tipo de propiedad' in df.columns:
        mask_banos = df['Tipo de propiedad'].notna()
        mask_banos = mask_banos & (~df['Tipo de propiedad'].str.upper().isin(tipos_a_excluir))
        df_banos = df[mask_banos]
        if len(df_banos) > 0:
            banos_clean = pd.to_numeric(df_banos['Baños'], errors='coerce')
            reportes['distribucion_banos'] = banos_clean.value_counts().to_dict()
    
    # KPI 11: Precio por m² de Ventas por Tipo de Propiedad
    if all(col in df.columns for col in ['Precio_numerico', 'Tipo de operación', 'Superficie_total_num', 'Tipo de propiedad', 'Estado']):
        df_ventas = df[(df['Tipo de operación'] == 'Venta') & (df['Estado'] == 'Vigente')].copy()
        df_ventas = df_ventas[df_ventas['Precio_numerico'].notna() & (df_ventas['Precio_numerico'] > 0)]
        df_ventas = df_ventas[df_ventas['Superficie_total_num'].notna() & (df_ventas['Superficie_total_num'] > 0)]
        
        if len(df_ventas) > 0:
            df_ventas['Precio_m2'] = df_ventas['Precio_numerico'] / df_ventas['Superficie_total_num']
            df_ventas = df_ventas[df_ventas['Precio_m2'].notna() & np.isfinite(df_ventas['Precio_m2'])]
            
            precio_m2_por_tipo = []
            for tipo_prop in df_ventas['Tipo de propiedad'].unique():
                if pd.notna(tipo_prop):
                    df_tipo = df_ventas[df_ventas['Tipo de propiedad'] == tipo_prop].copy()
                    if len(df_tipo) > 0:
                        Q1 = df_tipo['Precio_m2'].quantile(0.25)
                        Q3 = df_tipo['Precio_m2'].quantile(0.75)
                        IQR = Q3 - Q1
                        lower_bound = Q1 - 1.5 * IQR
                        upper_bound = Q3 + 1.5 * IQR
                        df_tipo_clean = df_tipo[(df_tipo['Precio_m2'] >= lower_bound) & (df_tipo['Precio_m2'] <= upper_bound)]
                        if len(df_tipo_clean) > 0:
                            precio_m2_por_tipo.append({
                                'Tipo de Propiedad': tipo_prop,
                                'Promedio (USD/m²)': df_tipo_clean['Precio_m2'].mean(),
                                'Mediana (USD/m²)': df_tipo_clean['Precio_m2'].median(),
                                'Mínimo (USD/m²)': df_tipo_clean['Precio_m2'].min(),
                                'Máximo (USD/m²)': df_tipo_clean['Precio_m2'].max(),
                                'Cantidad': len(df_tipo_clean)
                            })
            if precio_m2_por_tipo:
                reportes['precio_m2_por_tipo'] = pd.DataFrame(precio_m2_por_tipo)
    
    # KPI 12: Conteo por Estado
    if 'Estado' in df.columns:
        conteo_estado = df['Estado'].value_counts().sort_values(ascending=False)
        reportes['conteo_estado'] = pd.DataFrame({
            'Estado': conteo_estado.index,
            'Cantidad': conteo_estado.values,
            'Porcentaje': (conteo_estado.values / len(df) * 100).round(2)
        })
    
    # KPI 13: Propiedades con Argenprop
    if all(col in df.columns for col in ['Argenprop', 'Slug', 'Precio']):
        def es_valido_argenprop(valor):
            if pd.isna(valor):
                return False
            valor_str = str(valor).strip()
            if valor_str == '-' or valor_str == '':
                return False
            try:
                num = float(valor_str)
                return num > 0
            except:
                return valor_str != '-'
        
        mask_argenprop = df['Argenprop'].apply(es_valido_argenprop)
        df_argenprop = df[mask_argenprop][['Slug', 'Precio', 'Argenprop']].copy()
        if len(df_argenprop) > 0:
            reportes['propiedades_argenprop'] = df_argenprop
    
    return reportes

# ==================== FUNCIONES ADMINISTRATIVO ====================

@st.cache_data
def cargar_y_procesar_administrativo(uploaded_file):
    """Carga y procesa el archivo Excel del módulo Administrativo"""
    try:
        # Cargar todas las hojas
        excel_file = pd.ExcelFile(uploaded_file)
        dataframes = {}
        
        for hoja in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=hoja)
            dataframes[hoja] = df
        
        # Procesar cada hoja
        for nombre_hoja, df in dataframes.items():
            # Limpieza: eliminar filas vacías
            df = df.dropna(how='all')
            
            # Eliminar columnas especificadas
            columnas_a_eliminar = [
                'Nombre del lead',
                'Compañía',
                'Compañía del lead',
                'Modificado por',
                'Etiquetas del lead',
                'Tareas próximas',
                'Cerrado el',
                'utm_content',
                'utm_medium',
                'utm_campaign',
                'utm_source',
                'utm_term',
                'utm_referrer',
                'referrer',
                'gclientid',
                'gclid',
                'fbclid',
                'Cargo (contacto)',
                'Correo (contacto)',
                'E-mail priv. (contacto)',
                'Otro e-mail (contacto)',
                'Teléfono oficina (contacto)',
                'Teléfono oficina directo (contacto)',
                'Teléfono celular (contacto)',
                'Fax (contacto)',
                'Teléfono de casa (contacto)',
                'Otro teléfono (contacto)',
                'Términos y condiciones (contacto)',
                'Nota 1',
                'Nota 2',
                'Nota 3',
                'Nota 4',
                'Nota 5'
            ]
            columnas_existentes = [col for col in columnas_a_eliminar if col in df.columns]
            if columnas_existentes:
                df = df.drop(columns=columnas_existentes)
            
            # Eliminar duplicados basados en "Contacto principal"
            if 'Contacto principal' in df.columns:
                df = df.drop_duplicates(subset=['Contacto principal'], keep='first')
            
            dataframes[nombre_hoja] = df
        
        return dataframes, {}
    except Exception as e:
        return None, {'general': str(e)}

def generar_reporte_administrativo(df):
    """Genera todos los reportes del módulo Administrativo"""
    reportes = {}
    
    # KPI 1: Total de Contactos Principales
    reportes['total_contactos'] = len(df)
    
    # KPI 2: Distribución por Responsable
    if 'Responsable' in df.columns:
        distribucion_resp = df['Responsable'].value_counts().to_dict()
        reportes['distribucion_responsable'] = distribucion_resp
        
        # Crear DataFrame para mostrar
        df_responsable = pd.DataFrame({
            'Responsable': list(distribucion_resp.keys()),
            'Cantidad': list(distribucion_resp.values())
        })
        df_responsable['Porcentaje'] = (df_responsable['Cantidad'] / reportes['total_contactos'] * 100).round(2)
        reportes['tabla_responsable'] = df_responsable
    
    # KPI 3: Análisis de Fecha de Creación
    if 'Fecha de Creación' in df.columns:
        # Convertir a datetime con formato día/mes/año
        def parsear_fecha(fecha):
            """Parsea fechas en formato día/mes/año (formato latino)"""
            if pd.isna(fecha):
                return pd.NaT
            
            # Si ya es datetime, retornar
            if isinstance(fecha, pd.Timestamp):
                return fecha
            
            fecha_str = str(fecha).strip()
            
            # Intentar diferentes formatos
            formatos = [
                '%d/%m/%Y',      # 01/09/2024
                '%d-%m-%Y',      # 01-09-2024
                '%Y-%m-%d',      # 2024-09-01
                '%d/%m/%Y %H:%M:%S',  # 01/09/2024 10:30:00
                '%Y-%m-%d %H:%M:%S',  # 2024-09-01 10:30:00
            ]
            
            for formato in formatos:
                try:
                    return pd.to_datetime(fecha_str, format=formato)
                except:
                    continue
            
            # Si ningún formato funciona, intentar parseo automático pero con dayfirst=True
            try:
                return pd.to_datetime(fecha_str, dayfirst=True, errors='coerce')
            except:
                return pd.NaT
        
        df['Fecha de Creación'] = df['Fecha de Creación'].apply(parsear_fecha)
        df_fechas_validas = df[df['Fecha de Creación'].notna()].copy()
        
        if len(df_fechas_validas) > 0:
            # Crear columnas auxiliares
            df_fechas_validas['Año'] = df_fechas_validas['Fecha de Creación'].dt.year
            df_fechas_validas['Mes'] = df_fechas_validas['Fecha de Creación'].dt.month
            df_fechas_validas['Año-Mes'] = df_fechas_validas['Fecha de Creación'].dt.to_period('M')
            df_fechas_validas['Semana'] = df_fechas_validas['Fecha de Creación'].dt.to_period('W')
            
            # Acumulado total
            df_fechas_validas_sorted = df_fechas_validas.sort_values('Fecha de Creación')
            reportes['fecha_min'] = df_fechas_validas_sorted['Fecha de Creación'].min()
            reportes['fecha_max'] = df_fechas_validas_sorted['Fecha de Creación'].max()
            reportes['total_con_fecha'] = len(df_fechas_validas_sorted)
            
            # Por mes
            conteo_mes = df_fechas_validas.groupby('Año-Mes').size().sort_index()
            acumulado_mes = conteo_mes.cumsum()
            
            df_mes = pd.DataFrame({
                'Año-Mes': [str(p) for p in conteo_mes.index],
                'Cantidad': conteo_mes.values,
                'Acumulado': acumulado_mes.values
            })
            reportes['tabla_mes'] = df_mes
            reportes['conteo_mes'] = conteo_mes.to_dict()
            
            # Por semana
            conteo_semana = df_fechas_validas.groupby('Semana').size().sort_index()
            acumulado_semana = conteo_semana.cumsum()
            
            df_semana = pd.DataFrame({
                'Semana': [str(p) for p in conteo_semana.index],
                'Cantidad': conteo_semana.values,
                'Acumulado': acumulado_semana.values
            })
            reportes['tabla_semana'] = df_semana
            reportes['conteo_semana'] = conteo_semana.to_dict()
    
    return reportes

# ==================== INTERFAZ STREAMLIT ====================

# Sidebar para carga de archivo y selector de módulo
with st.sidebar:
    st.header("📁 Configuración")
    
    # Selector de módulo
    modulo_seleccionado = st.selectbox(
        "🔧 Selecciona el módulo",
        options=["Ejecutivo", "Portfolio", "Administrativo"],
        help="Elige qué tipo de análisis realizar"
    )
    
    st.markdown("---")
    st.header("📁 Carga de Archivo")
    uploaded_file = st.file_uploader(
        "Selecciona un archivo Excel",
        type=['xlsx', 'xls'],
        help="Sube un archivo Excel para analizar"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Archivo cargado: {uploaded_file.name}")

# Contenido principal
if uploaded_file is not None:
    # Cargar y procesar datos según el módulo seleccionado
    if modulo_seleccionado == "Ejecutivo":
        with st.spinner("Cargando y procesando datos..."):
            dataframes, errores = cargar_y_procesar_excel(uploaded_file)
        
        if dataframes is None:
            st.error(f"Error al cargar el archivo: {errores.get('general', 'Error desconocido')}")
        elif len(dataframes) == 0:
            st.error("No se pudieron cargar ninguna de las hojas objetivo.")
        else:
            # Mostrar errores si los hay
            if errores:
                for hoja, error in errores.items():
                    st.warning(f"⚠️ No se pudo cargar la hoja '{hoja}': {error}")
            
            # Generar todos los reportes
            reportes = {}
            
            if 'Operaciones' in dataframes:
                reportes['Operaciones'] = generar_reporte_operaciones(dataframes['Operaciones'])
            
            if 'Comisiones-equipo' in dataframes:
                reportes['Comisiones-equipo'] = generar_reporte_comisiones(dataframes['Comisiones-equipo'])
            
            if 'Clientes' in dataframes:
                reportes['Clientes'] = generar_reporte_clientes(dataframes['Clientes'])
            
            reportes['Flujo de Caja'] = generar_flujo_caja(dataframes)
            
            # Botón para descargar PDF
            st.sidebar.markdown("---")
            st.sidebar.header("📥 Descargar Reporte")
            
            pdf_buffer = generar_pdf(dataframes, reportes)
            st.sidebar.download_button(
                label="📄 Descargar Reporte Completo (PDF)",
                data=pdf_buffer.getvalue(),
                file_name=f"reporte_gestion_inmobiliaria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf"
            )
            
            # Pestañas por rubro
            tab1, tab2, tab3, tab4 = st.tabs([
                "💰 Operaciones",
                "👥 Comisiones",
                "👤 Clientes",
                "💵 Flujo de Caja"
            ])
            
            # TAB 1: OPERACIONES
            with tab1:
                st.header("💰 Análisis de Operaciones")
                
                # Filtro por mes según Fecha cierre
                if 'Operaciones' in dataframes:
                    df_ops_original = dataframes['Operaciones']
                
                    # Verificar si tiene columna Fecha cierre
                    if 'Fecha cierre' in df_ops_original.columns:
                        # Convertir a datetime si no lo es
                        if not pd.api.types.is_datetime64_any_dtype(df_ops_original['Fecha cierre']):
                            df_ops_original['Fecha cierre'] = pd.to_datetime(df_ops_original['Fecha cierre'], errors='coerce')
                        
                        # Obtener meses disponibles
                        df_ops_original['Mes'] = df_ops_original['Fecha cierre'].dt.to_period('M')
                        meses_disponibles = sorted(df_ops_original['Mes'].dropna().unique())
                        
                        # Crear opciones para el selector (formato: "Enero 2025", "Febrero 2025", etc.)
                        meses_nombres = {
                            1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
                            5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
                            9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
                        }
                        opciones_mes = ['Total General'] + [
                            f"{meses_nombres[mes.month]} {mes.year}" for mes in meses_disponibles
                        ]
                        
                        # Selector de mes
                        mes_seleccionado = st.selectbox(
                            "📅 Filtrar por mes (Fecha de cierre)",
                            options=opciones_mes,
                            index=0,
                            help="Selecciona un mes específico o 'Total General' para ver todas las operaciones"
                        )
                        
                        # Filtrar DataFrame según selección
                        if mes_seleccionado == 'Total General':
                            df_ops_filtrado = df_ops_original.copy()
                        else:
                            # Buscar el mes correspondiente
                            mes_buscado = None
                            meses_nombres = {
                                1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
                                5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
                                9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
                            }
                            for mes_period in meses_disponibles:
                                nombre_mes = f"{meses_nombres[mes_period.month]} {mes_period.year}"
                                if nombre_mes == mes_seleccionado:
                                    mes_buscado = mes_period
                                    break
                            
                            if mes_buscado:
                                df_ops_filtrado = df_ops_original[df_ops_original['Mes'] == mes_buscado].copy()
                            else:
                                df_ops_filtrado = df_ops_original.copy()
                        
                        # Eliminar columna temporal Mes
                        if 'Mes' in df_ops_filtrado.columns:
                            df_ops_filtrado = df_ops_filtrado.drop(columns=['Mes'])
                        
                        # Regenerar reporte con datos filtrados
                        reporte_ops_filtrado = generar_reporte_operaciones(df_ops_filtrado)
                    else:
                        # Si no hay Fecha cierre, usar reporte original
                        reporte_ops_filtrado = reportes['Operaciones']
                        mes_seleccionado = "Total General (sin fecha de cierre)"
                else:
                    reporte_ops_filtrado = None
            
            if reporte_ops_filtrado:
                ops = reporte_ops_filtrado
                
                # Mostrar mes seleccionado
                if 'Operaciones' in dataframes and 'Fecha cierre' in dataframes['Operaciones'].columns:
                    st.info(f"📅 Mostrando datos: **{mes_seleccionado}**")
                
                # Resumen financiero
                if 'resumen_financiero' in ops:
                    st.subheader("📊 Resumen Financiero")
                    rf = ops['resumen_financiero']
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Comisión", f"${rf['Total Comisión']:,.2f}")
                    with col2:
                        st.metric("Total Cobrado", f"${rf['Total Cobrado']:,.2f}")
                    with col3:
                        st.metric("Saldo Pendiente", f"${rf['Total Saldo Pendiente']:,.2f}")
                    with col4:
                        st.metric("Porcentaje Cobrado", f"{rf['Porcentaje Cobrado']:.2f}%")
                
                # Indicadores por tipo
                if 'indicadores' in ops:
                    st.subheader("📈 Indicadores por Tipo de Operación")
                    ind = ops['indicadores']
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Alquileres", ind['Alquileres'])
                    with col2:
                        st.metric("Ventas", ind['Ventas'])
                    with col3:
                        st.metric("Total Operaciones", ind['Total Operaciones'])
                
            else:
                st.warning("⚠️ No se pudo generar el reporte de Operaciones")
            
            # TAB 2: COMISIONES
            with tab2:
                st.header("👥 Análisis de Comisiones por Agente")
                
                # KPI: Comisión total por Agente (desde Operaciones)
                if 'Operaciones' in reportes and 'comision_por_agente' in reportes['Operaciones']:
                    st.subheader("💼 Comisión Total Generada por Agente (desde Operaciones)")
                    df_comision_agente = reportes['Operaciones']['comision_por_agente'].copy()
                    
                    # Formatear la columna de comisión para mostrar
                    df_comision_agente_display = df_comision_agente.copy()
                    df_comision_agente_display['Comisión Total'] = df_comision_agente_display['Comisión Total'].apply(lambda x: f"${x:,.2f}")
                    
                    st.dataframe(df_comision_agente_display, use_container_width=True, hide_index=True)
                    
                    # Mostrar métricas individuales
                    st.markdown("---")
                    cols = st.columns(min(len(df_comision_agente), 4))
                    for idx, row in df_comision_agente.iterrows():
                        col_idx = idx % 4
                        with cols[col_idx]:
                            st.metric(
                                row['Agente'],
                                f"${row['Comisión Total']:,.2f}",
                                f"{int(row['Cantidad Operaciones'])} ops"
                            )
                    st.markdown("---")
                
                if 'Comisiones-equipo' in reportes:
                    com_df = reportes['Comisiones-equipo']
                    
                    if isinstance(com_df, pd.DataFrame) and len(com_df) > 0:
                        st.subheader("💰 Resumen por Agente (Comisiones Pagadas/Pendientes)")
                        
                        # Formatear montos para mostrar
                        df_mostrar = com_df.copy()
                        for col in ['Total Comisión', 'Total Pagado', 'Total Pendiente']:
                            if col in df_mostrar.columns:
                                df_mostrar[col] = df_mostrar[col].apply(lambda x: f"${x:,.2f}")
                        
                        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
                        
                        # Resumen general
                        st.subheader("📊 Resumen General")
                        total_comision = com_df['Total Comisión'].sum()
                        total_pagado = com_df['Total Pagado'].sum()
                        total_pendiente = com_df['Total Pendiente'].sum()
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Comisiones", f"${total_comision:,.2f}")
                        with col2:
                            st.metric("Total Pagado", f"${total_pagado:,.2f}")
                        with col3:
                            st.metric("Total Pendiente", f"${total_pendiente:,.2f}")
                        
                        if total_comision > 0:
                            porcentaje_pagado = (total_pagado / total_comision * 100)
                            st.metric("Porcentaje Pagado", f"{porcentaje_pagado:.2f}%")
                    else:
                        st.warning("⚠️ No hay datos de comisiones disponibles")
                else:
                    st.warning("⚠️ No se pudo generar el reporte de Comisiones")
            
            # TAB 3: CLIENTES
            with tab3:
                st.header("👤 Análisis de Clientes")
                
                if 'Clientes' in reportes:
                    cli = reportes['Clientes']
                    
                    # Cantidad de clientes
                    if 'total_clientes' in cli:
                        st.subheader("👥 Cantidad de Clientes")
                        st.metric("Total de Clientes", cli['total_clientes'])
                    
                    # Análisis de edad
                    if 'edad' in cli:
                        st.subheader("📊 Análisis de Edad")
                        edad = cli['edad']
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Promedio de Edad", f"{edad['promedio']:.1f} años")
                        with col2:
                            st.metric("Edad Mínima", f"{edad['minima']} años")
                        with col3:
                            st.metric("Edad Máxima", f"{edad['maxima']} años")
                        
                        if edad.get('clientes_mayores_70', 0) > 0:
                            st.info(f"ℹ️ {edad['clientes_mayores_70']} clientes con edad > 70 (ignorados en promedio, pero contados en total)")
                    
                    # Distribución etaria
                    if 'distribucion_etaria' in cli:
                        st.subheader("📊 Distribución por Rango Etario")
                        st.dataframe(cli['distribucion_etaria'], use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ No se pudo generar el reporte de Clientes")
            
            # TAB 4: FLUJO DE CAJA
            with tab4:
                st.header("💵 Flujo de Caja - Informe Ejecutivo")
                
                if 'Flujo de Caja' in reportes:
                    fc = reportes['Flujo de Caja']
                    
                    st.subheader("📊 Resumen Financiero")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.metric("Total Cobrado", f"${fc['cobrado_total']:,.2f}")
                        st.metric("Total Pagado", f"${fc['pagado_total']:,.2f}")
                        st.metric("Balance Neto", f"${fc['balance_neto']:,.2f}")
                    
                    with col2:
                        st.metric("Por Cobrar", f"${fc['por_cobrar_total']:,.2f}")
                        st.metric("Por Pagar", f"${fc['por_pagar_total']:,.2f}")
                        st.metric("Flujo Futuro Neto", f"${fc['flujo_futuro_neto']:,.2f}")
                    
                    # Tabla resumen
                    resumen_df = pd.DataFrame({
                        'Concepto': [
                            'Total Cobrado (Operaciones)',
                            'Total Pagado (Comisiones)',
                            'Por Cobrar (Saldo Operaciones)',
                            'Por Pagar (Comisiones Pendientes)',
                            'Balance Neto (Cobrado - Pagado)',
                            'Flujo Futuro Neto (Por Cobrar - Por Pagar)'
                        ],
                        'Monto': [
                            fc['cobrado_total'],
                            fc['pagado_total'],
                            fc['por_cobrar_total'],
                            fc['por_pagar_total'],
                            fc['balance_neto'],
                            fc['flujo_futuro_neto']
                        ]
                    })
                    
                    resumen_df['Monto'] = resumen_df['Monto'].apply(lambda x: f"${x:,.2f}")
                    st.dataframe(resumen_df, use_container_width=True, hide_index=True)
                    
                    # Operaciones con saldo pendiente
                    if 'Operaciones' in reportes and 'operaciones_pendientes' in reportes['Operaciones']:
                        ops = reportes['Operaciones']
                        if len(ops['operaciones_pendientes']) > 0:
                            st.markdown("---")
                            st.subheader("📋 Operaciones con Saldo a Cobrar")
                            df_pendientes = ops['operaciones_pendientes'].copy()
                            
                            # Formatear montos para mostrar
                            for col in ['Comisión total', 'Cobrado', 'Saldo']:
                                if col in df_pendientes.columns:
                                    df_pendientes[col] = df_pendientes[col].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")
                            
                            columnas_mostrar = ['Nº Operación', 'Tipo', 'Cliente', 'Comisión total', 'Cobrado', 'Saldo']
                            columnas_disponibles = [col for col in columnas_mostrar if col in df_pendientes.columns]
                            st.dataframe(df_pendientes[columnas_disponibles], use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ No se pudo generar el reporte de Flujo de Caja")
    
    elif modulo_seleccionado == "Portfolio":
        # Cargar y procesar datos de Portfolio
        with st.spinner("Cargando y procesando datos de Portfolio..."):
            dataframes, errores = cargar_y_procesar_portfolio(uploaded_file)
        
        if dataframes is None:
            st.error(f"Error al cargar el archivo: {errores.get('general', 'Error desconocido')}")
        elif len(dataframes) == 0:
            st.error("No se pudieron cargar las hojas del archivo.")
        else:
            # Mostrar errores si los hay
            if errores:
                for hoja, error in errores.items():
                    st.warning(f"⚠️ No se pudo cargar la hoja '{hoja}': {error}")
            
            # Procesar la primera hoja (asumiendo que Portfolio tiene una sola hoja principal)
            nombre_hoja = list(dataframes.keys())[0]
            df_portfolio = dataframes[nombre_hoja]
            
            # Generar reportes
            reportes = generar_reporte_portfolio(df_portfolio)
            
            # Pestañas de Portfolio
            tab1, tab2, tab3, tab4 = st.tabs([
                "📊 Resumen General",
                "💰 Precios",
                "🏢 Propiedades",
                "📍 Ubicaciones"
            ])
            
            # TAB 1: RESUMEN GENERAL
            with tab1:
                st.header("📊 Resumen General del Portfolio")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Propiedades", reportes.get('total_propiedades', 0))
                
                if 'distribucion_operacion' in reportes:
                    st.subheader("📈 Distribución por Tipo de Operación")
                    dist_ops = reportes['distribucion_operacion']
                    for tipo, cantidad in dist_ops.items():
                        porcentaje = (cantidad / reportes['total_propiedades']) * 100
                        st.metric(tipo, f"{cantidad} ({porcentaje:.2f}%)")
                
                if 'conteo_estado' in reportes:
                    st.subheader("📊 Conteo por Estado")
                    
                    # Filtro por Tipo de Operación
                    if 'Tipo de operación' in df_portfolio.columns:
                        tipos_disponibles = ['Todos'] + sorted(df_portfolio['Tipo de operación'].dropna().unique().tolist())
                        tipo_filtro = st.selectbox(
                            "Filtrar por Tipo de Operación:",
                            options=tipos_disponibles,
                            index=0,
                            key="filtro_tipo_operacion_estado"
                        )
                        
                        # Aplicar filtro
                        if tipo_filtro == 'Todos':
                            df_filtrado = df_portfolio.copy()
                        else:
                            df_filtrado = df_portfolio[df_portfolio['Tipo de operación'] == tipo_filtro].copy()
                        
                        # Calcular conteo por estado con filtro
                        if 'Estado' in df_filtrado.columns and len(df_filtrado) > 0:
                            conteo_estado_filtrado = df_filtrado['Estado'].value_counts().sort_values(ascending=False)
                            df_conteo = pd.DataFrame({
                                'Estado': conteo_estado_filtrado.index,
                                'Cantidad': conteo_estado_filtrado.values,
                                'Porcentaje': (conteo_estado_filtrado.values / len(df_filtrado) * 100).round(2)
                            })
                            st.dataframe(df_conteo, use_container_width=True, hide_index=True)
                        else:
                            st.info("No hay datos para el filtro seleccionado.")
                    else:
                        # Si no hay columna Tipo de operación, mostrar el conteo original
                        df_conteo = reportes['conteo_estado'].copy()
                        df_conteo['Porcentaje'] = df_conteo['Porcentaje'].round(2)
                        st.dataframe(df_conteo, use_container_width=True, hide_index=True)
            
            # TAB 2: PRECIOS
            with tab2:
                st.header("💰 Análisis de Precios")
                
                if 'precio_m2_por_tipo' in reportes:
                    st.subheader("Precio por m² de Ventas (USD/m²)")
                    st.info("⚠️ Solo se consideran propiedades con Estado 'Vigente', con precio válido (excluye precio 0 y 'consultar precio'), con superficie total válida, y sin valores atípicos (outliers)")
                    # Formatear a 2 decimales
                    df_precio_m2 = reportes['precio_m2_por_tipo'].copy()
                    for col in ['Promedio (USD/m²)', 'Mediana (USD/m²)', 'Mínimo (USD/m²)', 'Máximo (USD/m²)']:
                        if col in df_precio_m2.columns:
                            df_precio_m2[col] = df_precio_m2[col].round(2)
                    st.dataframe(df_precio_m2, use_container_width=True, hide_index=True)
            
            # TAB 3: PROPIEDADES
            with tab3:
                st.header("🏢 Análisis de Propiedades")
                
                if 'distribucion_tipo_propiedad' in reportes:
                    st.subheader("Distribución por Tipo de Propiedad")
                    dist_tipo = reportes['distribucion_tipo_propiedad']
                    for tipo, cantidad in dist_tipo.items():
                        porcentaje = (cantidad / reportes['total_propiedades']) * 100
                        st.metric(tipo, f"{cantidad} ({porcentaje:.2f}%)")
                
                if 'distribucion_dormitorios' in reportes:
                    st.subheader("Distribución de Dormitorios")
                    dist_dorm = reportes['distribucion_dormitorios']
                    for dorm, cantidad in sorted(dist_dorm.items(), key=lambda x: float(x[0]) if str(x[0]).replace('.','').isdigit() else 0):
                        if pd.notna(dorm):
                            st.write(f"{int(dorm)} dormitorio(s): {cantidad}")
                
                if 'propiedades_argenprop' in reportes:
                    st.subheader("Propiedades con Argenprop")
                    st.dataframe(reportes['propiedades_argenprop'], use_container_width=True, hide_index=True)
            
            # TAB 4: UBICACIONES
            with tab4:
                st.header("📍 Análisis de Ubicaciones")
                
                if 'top_ubicaciones' in reportes:
                    st.subheader("Top 10 Ubicaciones")
                    dist_ubic = reportes['top_ubicaciones']
                    for ubic, cantidad in dist_ubic.items():
                        porcentaje = (cantidad / reportes['total_propiedades']) * 100
                        st.metric(ubic, f"{cantidad} ({porcentaje:.2f}%)")
                
                if 'distribucion_usuario' in reportes:
                    st.subheader("Distribución por Usuario Asignado")
                    dist_user = reportes['distribucion_usuario']
                    for user, cantidad in dist_user.items():
                        porcentaje = (cantidad / reportes['total_propiedades']) * 100
                        st.metric(user, f"{cantidad} ({porcentaje:.2f}%)")
            
    
    elif modulo_seleccionado == "Administrativo":
        # Cargar y procesar datos de Administrativo
        with st.spinner("Cargando y procesando datos administrativos..."):
            dataframes, errores = cargar_y_procesar_administrativo(uploaded_file)
        
        if dataframes is None:
            st.error(f"Error al cargar el archivo: {errores.get('general', 'Error desconocido')}")
        elif len(dataframes) == 0:
            st.error("No se pudieron cargar las hojas del archivo.")
        else:
            # Mostrar errores si los hay
            if errores:
                for hoja, error in errores.items():
                    st.warning(f"⚠️ No se pudo cargar la hoja '{hoja}': {error}")
            
            # Procesar la primera hoja
            nombre_hoja = list(dataframes.keys())[0]
            df_administrativo = dataframes[nombre_hoja]
            
            # Generar reportes
            reportes = generar_reporte_administrativo(df_administrativo)
            
            # Pestañas de Administrativo
            tab1, tab2, tab3 = st.tabs([
                "📊 Resumen General",
                "👥 Responsables",
                "📅 Fechas de Creación"
            ])
            
            # TAB 1: RESUMEN GENERAL
            with tab1:
                st.header("📊 Resumen General del Módulo Administrativo")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total de Contactos", reportes.get('total_contactos', 0))
                
                if 'total_con_fecha' in reportes:
                    with col2:
                        st.metric("Contactos con Fecha", reportes.get('total_con_fecha', 0))
                
                if 'fecha_min' in reportes and 'fecha_max' in reportes:
                    st.subheader("📅 Rango de Fechas")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Fecha Mínima", reportes['fecha_min'].strftime('%d/%m/%Y'))
                    with col2:
                        st.metric("Fecha Máxima", reportes['fecha_max'].strftime('%d/%m/%Y'))
            
            # TAB 2: RESPONSABLES
            with tab2:
                st.header("👥 Distribución por Responsable")
                
                if 'distribucion_responsable' in reportes:
                    dist_resp = reportes['distribucion_responsable']
                    total = reportes['total_contactos']
                    
                    for responsable, cantidad in dist_resp.items():
                        porcentaje = (cantidad / total) * 100
                        st.metric(responsable, f"{cantidad} contactos ({porcentaje:.2f}%)")
                    
                    if 'tabla_responsable' in reportes:
                        st.subheader("Tabla Detallada")
                        st.dataframe(reportes['tabla_responsable'], use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ No se encontró información de Responsables")
            
            # TAB 3: FECHAS DE CREACIÓN
            with tab3:
                st.header("📅 Análisis de Fechas de Creación")
                
                if 'tabla_mes' in reportes:
                    st.subheader("📆 Distribución por Mes")
                    st.dataframe(reportes['tabla_mes'], use_container_width=True, hide_index=True)
                
                if 'tabla_semana' in reportes:
                    st.subheader("📅 Distribución por Semana (últimas 15 semanas)")
                    df_semana_display = reportes['tabla_semana'].tail(15).copy()
                    st.dataframe(df_semana_display, use_container_width=True, hide_index=True)
                
                if 'total_con_fecha' not in reportes:
                    st.warning("⚠️ No se encontró información de Fechas de Creación")

else:
    # Mensaje inicial
    st.info("👈 Por favor, carga un archivo Excel desde la barra lateral para comenzar el análisis.")
    
    st.markdown("""
    ### 📋 Características de la aplicación:
    - ✅ Carga de archivos Excel (.xlsx, .xls) con múltiples hojas
    - 📊 Análisis automático por rubro (Operaciones, Comisiones, Clientes, Flujo de Caja)
    - 🔍 Procesamiento y limpieza automática de datos
    - 📈 Reportes ejecutivos en formato tabular
    - 📄 Descarga de reporte completo en PDF
    
    ### 📑 Hojas que se analizan:
    - **Operaciones**: Datos de operaciones inmobiliarias
    - **Comisiones-equipo**: Información de comisiones por agente
    - **Cuotas-comisiones**: Detalle de cuotas y comisiones
    - **Clientes**: Base de datos de clientes
    
    ### 🚀 Instrucciones:
    1. Usa el panel lateral para cargar tu archivo Excel
    2. Explora las diferentes pestañas para ver cada análisis
    3. Descarga el reporte completo en PDF desde el panel lateral
    """)
