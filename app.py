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
st.title("📊 Análisis de Gestión Inmobiliaria")
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

# ==================== INTERFAZ STREAMLIT ====================

# Sidebar para carga de archivo
with st.sidebar:
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
    # Cargar y procesar datos
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
            
            if 'Operaciones' in reportes:
                ops = reportes['Operaciones']
                
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
                
                # Operaciones con saldo pendiente
                if 'operaciones_pendientes' in ops and len(ops['operaciones_pendientes']) > 0:
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
                    st.success("✅ No hay operaciones con saldo pendiente")
            else:
                st.warning("⚠️ No se pudo generar el reporte de Operaciones")
        
        # TAB 2: COMISIONES
        with tab2:
            st.header("👥 Análisis de Comisiones por Agente")
            
            if 'Comisiones-equipo' in reportes:
                com_df = reportes['Comisiones-equipo']
                
                if isinstance(com_df, pd.DataFrame) and len(com_df) > 0:
                    st.subheader("💰 Resumen por Agente")
                    
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
            else:
                st.warning("⚠️ No se pudo generar el reporte de Flujo de Caja")

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
