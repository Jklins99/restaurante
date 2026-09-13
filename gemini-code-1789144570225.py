import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import pdfplumber
import re
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

# === TUS CONFIGURACIONES ===
SPREADSHEET_ID = '1811VY4-Xa4ZOf7j6MVd5zYFCdhpyxdtuwq1pD5mHlh4'
RUC_RESTAURANTE = '10402504051'
# =========================

st.set_page_config(
    page_title="Gestor de Facturación - Restaurante",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- ESTILOS CSS PERSONALIZADOS ---
st.markdown("""
    <style>
    /* Color scheme del restaurante */
    :root {
        --primary: #2E7D32;      /* Verde profesional */
        --accent: #D32F2F;       /* Rojo cálido */
        --neutral: #424242;      /* Gris oscuro */
        --light-bg: #F5F5F5;     /* Fondo claro */
        --white: #FFFFFF;
    }
    
    /* Título principal */
    .main-title {
        font-size: 2.5em;
        font-weight: 700;
        color: #1B5E20;
        margin-bottom: 0.2em;
        letter-spacing: -0.5px;
    }
    
    /* Secciones */
    .section-title {
        font-size: 1.3em;
        font-weight: 600;
        color: #2E7D32;
        margin-top: 1.5em;
        margin-bottom: 1em;
        border-left: 4px solid #2E7D32;
        padding-left: 12px;
    }
    
    /* Tarjetas de métrica mejoradas */
    .metric-card {
        background: linear-gradient(135deg, #FFFFFF 0%, #F5F5F5 100%);
        border-radius: 8px;
        padding: 1.5em;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid #2E7D32;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    }
    
    .metric-label {
        font-size: 0.85em;
        color: #666;
        font-weight: 500;
        margin-bottom: 0.5em;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .metric-value {
        font-size: 2em;
        font-weight: 700;
        color: #1B5E20;
    }
    
    /* Métrica de alerta */
    .metric-card.alert {
        border-left-color: #D32F2F;
    }
    
    .metric-card.alert .metric-value {
        color: #D32F2F;
    }
    
    /* Botón principal */
    .stButton > button {
        background: linear-gradient(135deg, #2E7D32 0%, #1B5E20 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 6px;
        padding: 0.75em 2em;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(46, 125, 50, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(46, 125, 50, 0.4);
    }
    
    /* Divider mejorado */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #DDD, transparent);
        margin: 2em 0;
    }
    
    /* Info boxes */
    .info-box {
        background: #E8F5E9;
        border-left: 4px solid #2E7D32;
        padding: 1em;
        border-radius: 4px;
        margin: 1em 0;
    }
    
    .warning-box {
        background: #FFEBEE;
        border-left: 4px solid #D32F2F;
        padding: 1em;
        border-radius: 4px;
        margin: 1em 0;
    }
    
    /* Tabla mejorada */
    .dataframe {
        font-size: 0.95em;
    }
    
    /* Subtítulos */
    .subtitle {
        font-size: 1.1em;
        color: #555;
        margin: 1.5em 0 0.5em 0;
        font-weight: 500;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIONES DE GOOGLE SHEETS ---
def obtener_servicio_sheets():
    creds_dict = st.secrets["gcp_service_account"]
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return build('sheets', 'v4', credentials=creds)

def obtener_facturas_registradas(sheets_service):
    registradas = set()
    try:
        resultado = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range='A:K'
        ).execute()
        filas = resultado.get('values', [])
        
        for fila in filas[1:]:
            if len(fila) >= 4:
                llave = f"{fila[3]}-{fila[2]}"  # RUC-Comprobante
                registradas.add(llave)
    except Exception:
        pass
    return registradas

def descargar_historial_sheets(sheets_service):
    try:
        resultado = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range='A:K'
        ).execute()
        filas = resultado.get('values', [])
        if len(filas) > 1:
            # Cabeceras con columnas nuevas de documento y tasa
            cabeceras = ['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Razón Social', 'Base Imponible', 'IGV', 'Total', 'Categoría', 'Documento', 'Tasa %']
            datos = filas[1:]
            
            datos_normalizados = []
            for fila in datos:
                while len(fila) < len(cabeceras):
                    fila.append("0")
                datos_normalizados.append(fila[:len(cabeceras)])
                
            df = pd.DataFrame(datos_normalizados, columns=cabeceras)
            
            # Normalizar columnas numéricas
            for col in ['Base Imponible', 'IGV', 'Total', 'Tasa %']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                
            return df
    except Exception:
        pass
    return pd.DataFrame()

def guardar_en_sheets(sheets_service, datos, categoria):
    # Usar IGV si existe, sino IGV (18%)
    igv_val = datos.get('IGV') if 'IGV' in datos else datos.get('IGV (18%)', 0)
    tipo_doc = datos.get('Documento', 'N/A')
    tasa_igv = datos.get('Tasa %', 18.0)
    
    fila = [
        str(datos['Fecha']), str(datos['Tipo']), str(datos['Comprobante']), str(datos['RUC']), 
        str(datos['Razón Social']), float(datos['Base Imponible']), float(igv_val), 
        float(datos['Total']), str(categoria), str(tipo_doc), str(tasa_igv)
    ]
    cuerpo = {'values': [fila]}
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range='A:K',
        valueInputOption='USER_ENTERED', body=cuerpo
    ).execute()

# --- EXTRACCIÓN Y CLASIFICACIÓN AUTOMÁTICA (XML y PDF) ---
def parse_sunat_xml(file_obj):
    try:
        file_obj.seek(0)
        tree = ET.parse(file_obj)
        root = tree.getroot()
        ns = {'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2', 'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'}
        
        fecha = root.find('.//cbc:IssueDate', ns)
        serie_numero = root.find('.//cbc:ID', ns)
        ruc = root.find('.//cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID', ns)
        razon_social = root.find('.//cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName', ns)
        total_node = root.find('.//cac:LegalMonetaryTotal/cbc:PayableAmount', ns)
        igv_node = root.find('.//cac:TaxTotal/cac:TaxSubtotal/cbc:TaxAmount', ns)
        
        total = float(total_node.text) if total_node is not None else 0.0
        igv = float(igv_node.text) if igv_node is not None else 0.0
        ruc_val = ruc.text if ruc is not None else 'N/A'
        
        # Detectar tipo de documento (Boleta o Factura)
        serie_val = serie_numero.text if serie_numero is not None else 'N/A'
        tipo_doc = "Boleta" if serie_val.startswith('B') else "Factura"
        
        # Determinar tasa de IGV según tipo de documento
        tasa_igv = 10.5 if tipo_doc == "Boleta" else 18.0
        
        categoria = "Venta" if ruc_val == RUC_RESTAURANTE else "Compra"
        
        return {
            "Archivo": file_obj.name, "Tipo": "XML", "Documento": tipo_doc, "Fecha": fecha.text if fecha is not None else 'N/A',
            "Comprobante": serie_val, "RUC": ruc_val, "Razón Social": razon_social.text if razon_social is not None else 'N/A',
            "Base Imponible": round(total - igv, 2), "IGV": round(igv, 2), "Tasa %": tasa_igv, "Total": round(total, 2), 
            "Categoría": categoria, "Estado": "OK"
        }
    except Exception:
        return {"Archivo": file_obj.name, "Tipo": "XML", "Categoría": "Compra", "Estado": "Error de lectura"}

def procesar_factura_pdf(file_obj):
    try:
        file_obj.seek(0)
        texto_completo = ""
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto: texto_completo += texto + "\n"
        
        # --- EXTRAER FECHA DEL DOCUMENTO ---
        fecha_documento = None
        
        # Buscar formato: DD/MM/YYYY (formato peruano típico)
        fecha_match = re.search(r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b', texto_completo)
        
        if fecha_match:
            try:
                dia, mes, anio = fecha_match.groups()
                # Convertir a datetime para validar
                fecha_documento = datetime.datetime(int(anio), int(mes), int(dia))
                fecha_str = fecha_documento.strftime('%Y-%m-%d')
            except:
                fecha_str = datetime.datetime.now().strftime('%Y-%m-%d')
        else:
            # Si no encuentra fecha, usar la actual como fallback
            fecha_str = datetime.datetime.now().strftime('%Y-%m-%d')
        
        ruc_match = re.search(r'\b(10|20)\d{9}\b', texto_completo)
        serie_match = re.search(r'\b[F|E|B][A-Z0-9]{3}-\d{1,8}\b', texto_completo)
        total_match = re.search(r'(?:TOTAL|Total|Importe Total).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo)
        
        total = float(total_match.group(1).replace(',', '')) if total_match else 0.0
        serie_val = serie_match.group(0) if serie_match else "N/A"
        
        # Detectar tipo de documento (Boleta o Factura)
        tipo_doc = "Boleta" if serie_val.startswith('B') else "Factura"
        
        # Determinar tasa de IGV según tipo de documento
        tasa_igv = 10.5 if tipo_doc == "Boleta" else 18.0
        
        # Calcular base e IGV con la tasa correcta
        base_imponible = total / (1 + tasa_igv/100) if total > 0 else 0.0
        igv = total - base_imponible if total > 0 else 0.0
        ruc_val = ruc_match.group(0) if ruc_match else 'N/A'
        
        categoria = "Venta" if ruc_val == RUC_RESTAURANTE else "Compra"
        
        return {
            "Archivo": file_obj.name, "Tipo": "PDF", "Documento": tipo_doc, "Fecha": fecha_str,
            "Comprobante": serie_val, "RUC": ruc_val,
            "Razón Social": "Por verificar (PDF)", "Base Imponible": round(base_imponible, 2),
            "IGV": round(igv, 2), "Tasa %": tasa_igv, "Total": round(total, 2), "Categoría": categoria, 
            "Estado": "OK" if total > 0 else "Revisar Manualmente"
        }
    except Exception:
        return {"Archivo": file_obj.name, "Tipo": "PDF", "Categoría": "Compra", "Estado": "Error de lectura"}

# --- INTERFAZ PRINCIPAL MEJORADA ---
st.markdown('<div class="main-title">🍽️ Gestor de Facturación</div>', unsafe_allow_html=True)
st.markdown('Sistema automático de registro de ventas y compras para tu restaurante', unsafe_allow_html=True)

try:
    sheets_service = obtener_servicio_sheets()
    facturas_ya_registradas = obtener_facturas_registradas(sheets_service)
except Exception:
    sheets_service = None
    facturas_ya_registradas = set()

# --- SECCIÓN 1: DASHBOARD MENSUAL ---
st.markdown('<div class="section-title">📊 Resumen del Periodo</div>', unsafe_allow_html=True)

if sheets_service:
    df_historial = descargar_historial_sheets(sheets_service)
    
    if not df_historial.empty:
        df_historial['Fecha'] = pd.to_datetime(df_historial['Fecha'], errors='coerce')
        df_historial['Mes'] = df_historial['Fecha'].dt.to_period('M').astype(str)
        
        meses_disponibles = sorted(df_historial['Mes'].dropna().unique(), reverse=True)
        
        if meses_disponibles:
            col1, col2 = st.columns([3, 1])
            with col1:
                mes_seleccionado = st.selectbox(
                    "Selecciona el mes a revisar",
                    meses_disponibles,
                    label_visibility="collapsed"
                )
            
            df_mes = df_historial[df_historial['Mes'] == mes_seleccionado]
            ventas_mes = df_mes[df_mes['Categoría'] == 'Venta']
            compras_mes = df_mes[df_mes['Categoría'] == 'Compra']
            
            # Usar columna IGV si existe, sino IGV (18%)
            igv_col = 'IGV' if 'IGV' in df_mes.columns else 'IGV (18%)'
            igv_v = ventas_mes[igv_col].sum() if not ventas_mes.empty else 0.0
            igv_c = compras_mes[igv_col].sum() if not compras_mes.empty else 0.0
            total_v = ventas_mes['Total'].sum() if not ventas_mes.empty else 0.0
            total_c = compras_mes['Total'].sum() if not compras_mes.empty else 0.0
            
            diferencia_igv = igv_v - igv_c
            saldo_a_favor_mes = abs(min(0.0, diferencia_igv))
            
            # Calcular saldo a favor del mes anterior
            fecha_mes_seleccionado = pd.Period(mes_seleccionado, freq='M')
            fecha_mes_anterior = fecha_mes_seleccionado - 1
            mes_anterior_str = fecha_mes_anterior.strftime('%Y-%m')
            
            saldo_anterior = 0.0
            if mes_anterior_str in df_historial['Mes'].values:
                df_mes_anterior = df_historial[df_historial['Mes'] == mes_anterior_str]
                ventas_ant = df_mes_anterior[df_mes_anterior['Categoría'] == 'Venta']
                compras_ant = df_mes_anterior[df_mes_anterior['Categoría'] == 'Compra']
                
                igv_col_ant = 'IGV' if 'IGV' in df_mes_anterior.columns else 'IGV (18%)'
                igv_v_ant = ventas_ant[igv_col_ant].sum() if not ventas_ant.empty else 0.0
                igv_c_ant = compras_ant[igv_col_ant].sum() if not compras_ant.empty else 0.0
                diferencia_anterior = igv_v_ant - igv_c_ant
                saldo_anterior = abs(min(0.0, diferencia_anterior))
            
            # Calcular IGV final considerando saldo anterior
            igv_neto_pagar = max(0.0, diferencia_igv - saldo_anterior)
            saldo_a_favor = max(0.0, saldo_anterior - diferencia_igv)
            
            # --- TARJETAS DE MÉTRICAS MEJORADAS ---
            cols = st.columns(4, gap="medium")
            
            with cols[0]:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">💰 Ingresos</div>
                        <div class="metric-value">S/ {total_v:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            with cols[1]:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">📦 Egresos</div>
                        <div class="metric-value">S/ {total_c:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            with cols[2]:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">📈 IGV Neto del Mes</div>
                        <div class="metric-value">{'+' if diferencia_igv >= 0 else '-'} S/ {abs(diferencia_igv):,.2f}</div>
                        <div style="font-size: 0.75em; color: #999; margin-top: 0.5em;">Cobrado: S/ {igv_v:,.2f} | Crédito: S/ {igv_c:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            with cols[3]:
                if saldo_anterior > 0:
                    st.markdown(f"""
                        <div class="metric-card" style="border-left-color: #FF6F00;">
                            <div class="metric-label">↙️ Saldo Anterior</div>
                            <div class="metric-value" style="color: #FF6F00; font-size: 1.5em;">S/ {saldo_anterior:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">✓ Sin Saldo Anterior</div>
                            <div class="metric-value" style="color: #666; font-size: 1.5em;">-</div>
                        </div>
                        """, unsafe_allow_html=True)
            
            st.divider()
            
            # --- RESUMEN FINAL CON COMPENSACIÓN ---
            st.markdown('<p class="subtitle">📋 Resumen Final (Considerando Arrastre)</p>', unsafe_allow_html=True)
            
            cols_resumen = st.columns(3, gap="medium")
            
            with cols_resumen[0]:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">IGV Neto del Mes</div>
                        <div class="metric-value" style="color: #1565C0;">{'+' if diferencia_igv >= 0 else '-'} S/ {abs(diferencia_igv):,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            with cols_resumen[1]:
                if saldo_anterior > 0:
                    st.markdown(f"""
                        <div class="metric-card" style="border-left-color: #FF6F00;">
                            <div class="metric-label">Menos Saldo Anterior</div>
                            <div class="metric-value" style="color: #FF6F00;">- S/ {saldo_anterior:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Menos Saldo Anterior</div>
                            <div class="metric-value" style="color: #999;">-</div>
                        </div>
                        """, unsafe_allow_html=True)
            
            with cols_resumen[2]:
                if saldo_a_favor > 0:
                    st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">💚 RESULTADO</div>
                            <div class="metric-value" style="color: #2E7D32;">SALDO A FAVOR</div>
                            <div style="font-size: 1.3em; color: #2E7D32; font-weight: 700; margin-top: 0.3em;">S/ {saldo_a_favor:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                        <div class="metric-card alert">
                            <div class="metric-label">🏛️ RESULTADO</div>
                            <div class="metric-value" style="color: #D32F2F;">POR PAGAR</div>
                            <div style="font-size: 1.3em; color: #D32F2F; font-weight: 700; margin-top: 0.3em;">S/ {igv_neto_pagar:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
            
            st.divider()
            
            # --- GRÁFICO MEJORADO ---
            st.markdown('<p class="subtitle">📊 Análisis de IGV del Periodo</p>', unsafe_allow_html=True)
            
            col_grafico1, col_grafico2 = st.columns(2)
            
            with col_grafico1:
                resumen_grafico = pd.DataFrame({
                    "Concepto": ["IGV Ventas", "IGV Compras"],
                    "Monto (S/)": [igv_v, igv_c]
                }).set_index("Concepto")
                st.bar_chart(resumen_grafico)
                st.caption("Diferencia del mes: IGV generado por ventas vs crédito por compras")
            
            with col_grafico2:
                # Resumen de cantidad de documentos
                resumen_docs = pd.DataFrame({
                    "Tipo": ["Ventas", "Compras"],
                    "Cantidad": [len(ventas_mes), len(compras_mes)]
                }).set_index("Tipo")
                st.bar_chart(resumen_docs)
                st.caption(f"Total documentos: {len(ventas_mes) + len(compras_mes)}")
            
            st.divider()
            
            # --- EXPLICACIÓN DEL CÁLCULO ---
            if saldo_anterior > 0:
                st.info(f"""
                    📌 **Cómo se calcula tu obligación tributaria:**
                    
                    1. **IGV Neto del mes**: S/ {diferencia_igv:+,.2f} (Ventas - Compras)
                    2. **Menos Saldo a tu favor de {mes_anterior_str}**: S/ {saldo_anterior:,.2f}
                    3. **Resultado**: {'✅ Saldo a tu favor de S/ ' + f'{saldo_a_favor:,.2f}' if saldo_a_favor > 0 else '🏛️ Por pagar a SUNAT S/ ' + f'{igv_neto_pagar:,.2f}'}
                    
                    El crédito fiscal del mes anterior se compensa automáticamente en este período.
                """)
            
            st.divider()
            
            # --- TABLAS DE DATOS ---
            st.markdown('<p class="subtitle">Detalle de Ventas</p>', unsafe_allow_html=True)
            if not ventas_mes.empty:
                cols_ventas = ['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Total', igv_col]
                cols_ventas = [col for col in cols_ventas if col in ventas_mes.columns]
                
                # Agregar columna Documento y Tasa si existen
                if 'Documento' in ventas_mes.columns:
                    cols_ventas.insert(2, 'Documento')
                if 'Tasa %' in ventas_mes.columns:
                    cols_ventas.insert(3, 'Tasa %')
                
                st.dataframe(
                    ventas_mes[cols_ventas].sort_values('Fecha', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("📭 No hay ventas registradas en este periodo")

            st.markdown('<p class="subtitle">Detalle de Compras</p>', unsafe_allow_html=True)
            if not compras_mes.empty:
                cols_compras = ['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Total', igv_col]
                cols_compras = [col for col in cols_compras if col in compras_mes.columns]
                
                # Agregar columna Documento y Tasa si existen
                if 'Documento' in compras_mes.columns:
                    cols_compras.insert(2, 'Documento')
                if 'Tasa %' in compras_mes.columns:
                    cols_compras.insert(3, 'Tasa %')
                
                st.dataframe(
                    compras_mes[cols_compras].sort_values('Fecha', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("📭 No hay compras registradas en este periodo")
        else:
            st.info("📭 No se encontraron fechas válidas en el historial registrado.")
    else:
        st.info("📭 El Google Sheets aún no tiene registros guardados. Sube tu primer lote abajo.")
else:
    st.error("❌ No se pudo conectar con Google Sheets para cargar el historial.")

st.divider()

# --- INFORMACIÓN SOBRE TASAS DE IGV ---
with st.expander("ℹ️ Tasas de IGV: Boletas vs Facturas"):
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.markdown("""
            **🧾 BOLETAS ELECTRÓNICAS (B)**
            - Serie comienza con: **B**
            - Tasa de IGV: **10.5%**
            - Comprador: Persona natural
            - Ejemplo: B001-000123
        """)
    with col_info2:
        st.markdown("""
            **📋 FACTURAS ELECTRÓNICAS (F)**
            - Serie comienza con: **F**
            - Tasa de IGV: **18%**
            - Comprador: Empresa/RUC
            - Ejemplo: F001-000456
        """)
    st.info("💡 El sistema detecta automáticamente la tasa según el tipo de documento")

st.divider()

# --- SECCIÓN 2: CARGA RÁPIDA DE COMPROBANTES ---
st.markdown('<div class="section-title">📤 Registrar Nuevos Comprobantes</div>', unsafe_allow_html=True)
st.markdown('Carga tus archivos XML o PDF. El sistema detectará automáticamente la fecha de cada documento.', unsafe_allow_html=True)

st.divider()

archivos_subidos = st.file_uploader(
    "Selecciona tus archivos",
    type=['xml', 'pdf'],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

datos_procesados = []

if archivos_subidos:
    for f in archivos_subidos:
        datos = parse_sunat_xml(f) if f.name.lower().endswith('.xml') else procesar_factura_pdf(f)
        llave_actual = f"{datos.get('RUC')}-{datos.get('Comprobante')}"
        
        # Detectar automáticamente el mes del documento
        try:
            fecha_doc = pd.to_datetime(datos.get('Fecha'))
            mes_doc = fecha_doc.strftime('%Y-%m')
            datos['Mes Facturación'] = mes_doc
        except:
            datos['Mes Facturación'] = 'N/A'
        
        if llave_actual in facturas_ya_registradas and datos.get('RUC') != 'N/A':
            datos['Estado'] = '⚠️ Duplicado'
            
        datos_procesados.append(datos)

    df_todos = pd.DataFrame(datos_procesados)

    if not df_todos.empty:
        st.divider()
        
        # Detectar meses únicos en los documentos cargados
        meses_detectados = sorted(df_todos[df_todos['Mes Facturación'] != 'N/A']['Mes Facturación'].unique(), reverse=True)
        
        if meses_detectados:
            st.markdown(f'<p class="subtitle">📅 Meses Detectados: {", ".join(meses_detectados)}</p>', unsafe_allow_html=True)
        
        df_validos = df_todos[df_todos['Estado'] == 'OK']
        
        # Usar columna IGV si existe, sino IGV (18%)
        igv_col_lote = 'IGV' if 'IGV' in df_validos.columns else 'IGV (18%)'
        total_igv_ventas_lote = df_validos[df_validos['Categoría'] == 'Venta'][igv_col_lote].sum() if igv_col_lote in df_validos.columns else 0.0
        total_igv_compras_lote = df_validos[df_validos['Categoría'] == 'Compra'][igv_col_lote].sum() if igv_col_lote in df_validos.columns else 0.0
        igv_neto_lote = max(0, total_igv_ventas_lote - total_igv_compras_lote)
        
        # --- RESUMEN DEL LOTE ---
        st.markdown('<p class="subtitle">Resumen del Lote a Registrar</p>', unsafe_allow_html=True)
        
        cols_lote = st.columns(3, gap="medium")
        
        with cols_lote[0]:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📈 IGV Ventas</div>
                    <div class="metric-value" style="color: #2E7D32;">S/ {total_igv_ventas_lote:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)
        
        with cols_lote[1]:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📦 Crédito Fiscal</div>
                    <div class="metric-value" style="color: #1565C0;">S/ {total_igv_compras_lote:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)
        
        with cols_lote[2]:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">✓ IGV Neto</div>
                    <div class="metric-value" style="color: #F57C00;">S/ {igv_neto_lote:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)
        
        st.divider()
        
        # --- PREVIEW DE DATOS ---
        st.markdown('<p class="subtitle">Vista Previa del Lote</p>', unsafe_allow_html=True)
        
        # Mostrar tabla con mes detectado y tipo de documento
        columnas_display = ['Archivo', 'Mes Facturación', 'Documento', 'Tasa %', 'Categoría', 'Comprobante', 'Total', 'IGV', 'Estado']
        df_display = df_todos[[col for col in columnas_display if col in df_todos.columns]]
        
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True
        )
        
        # --- ADVERTENCIAS ---
        duplicados = df_todos[df_todos['Estado'] == '⚠️ Duplicado']
        if not duplicados.empty:
            st.markdown(f"""
                <div class="warning-box">
                <strong>⚠️ Duplicados Detectados:</strong> {len(duplicados)} comprobante(s) ya están registrados en el sistema. No se incluirán en el registro.
                </div>
                """, unsafe_allow_html=True)
        
        errores = df_todos[df_todos['Estado'].str.contains('Error|Revisar', na=False)]
        if not errores.empty:
            st.markdown(f"""
                <div class="warning-box">
                <strong>⚠️ Revisión Manual:</strong> {len(errores)} archivo(s) necesita(n) revisión. Verifica los datos antes de registrar.
                </div>
                """, unsafe_allow_html=True)
        
        sin_fecha = df_todos[df_todos['Mes Facturación'] == 'N/A']
        if not sin_fecha.empty:
            st.markdown(f"""
                <div class="warning-box">
                <strong>⚠️ Fecha No Detectada:</strong> {len(sin_fecha)} archivo(s) no tiene fecha válida. Verifica que sean documentos válidos.
                </div>
                """, unsafe_allow_html=True)
        
        st.divider()
        
        # --- RESUMEN POR MES ---
        if len(meses_detectados) > 1:
            st.markdown('<p class="subtitle">📊 Resumen por Mes</p>', unsafe_allow_html=True)
            
            for mes in meses_detectados:
                df_mes_actual = df_todos[df_todos['Mes Facturación'] == mes]
                df_mes_validos = df_mes_actual[df_mes_actual['Estado'] == 'OK']
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(f"Documentos ({mes})", len(df_mes_validos))
                with col2:
                    igv_col_mes = 'IGV' if 'IGV' in df_mes_validos.columns else 'IGV (18%)'
                    igv_mes = df_mes_validos[igv_col_mes].sum() if igv_col_mes in df_mes_validos.columns else 0.0
                    st.metric(f"IGV Neto ({mes})", f"S/ {igv_mes:,.2f}")
                with col3:
                    total_mes = df_mes_validos['Total'].sum() if 'Total' in df_mes_validos.columns else 0.0
                    st.metric(f"Total ({mes})", f"S/ {total_mes:,.2f}")
        
        st.divider()
        
        # --- BOTÓN DE REGISTRO ---
        if not df_validos.empty:
            col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
            
            with col_btn1:
                if st.button("🚀 Registrar Lote en Google Sheets", type="primary", use_container_width=True):
                    with st.spinner("💾 Guardando en la base de datos..."):
                        try:
                            exitos = 0
                            for index, fila in df_validos.iterrows():
                                guardar_en_sheets(sheets_service, fila, fila['Categoría'])
                                exitos += 1
                            
                            meses_registrados = ", ".join(sorted(df_validos[df_validos['Estado'] == 'OK']['Mes Facturación'].unique()))
                            st.success(f"✅ ¡Éxito! Se registraron {exitos} comprobante(s) en {meses_registrados}. Recarga arriba para ver los saldos actualizados.")
                        except Exception as e:
                            st.error(f"❌ Error al registrar: {e}")
        else:
            st.warning("⚠️ No hay comprobantes válidos para registrar. Revisa los errores arriba.")
