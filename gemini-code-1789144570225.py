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
            spreadsheetId=SPREADSHEET_ID, range='A:D'
        ).execute()
        filas = resultado.get('values', [])
        
        for fila in filas[1:]:
            if len(fila) >= 4:
                llave = f"{fila[3]}-{fila[2]}"
                registradas.add(llave)
    except Exception:
        pass
    return registradas

def descargar_historial_sheets(sheets_service):
    try:
        resultado = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range='A:I'
        ).execute()
        filas = resultado.get('values', [])
        if len(filas) > 1:
            cabeceras = ['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Razón Social', 'Base Imponible', 'IGV (18%)', 'Total', 'Categoría']
            datos = filas[1:]
            
            datos_normalizados = []
            for fila in datos:
                while len(fila) < len(cabeceras):
                    fila.append("0")
                datos_normalizados.append(fila[:len(cabeceras)])
                
            df = pd.DataFrame(datos_normalizados, columns=cabeceras)
            
            for col in ['Base Imponible', 'IGV (18%)', 'Total']:
                df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                
            return df
    except Exception:
        pass
    return pd.DataFrame()

def guardar_en_sheets(sheets_service, datos, categoria):
    fila = [
        str(datos['Fecha']), str(datos['Tipo']), str(datos['Comprobante']), str(datos['RUC']), 
        str(datos['Razón Social']), float(datos['Base Imponible']), float(datos['IGV (18%)']), 
        float(datos['Total']), str(categoria)
    ]
    cuerpo = {'values': [fila]}
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range='A:I',
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
        
        categoria = "Venta" if ruc_val == RUC_RESTAURANTE else "Compra"
        
        return {
            "Archivo": file_obj.name, "Tipo": "XML", "Fecha": fecha.text if fecha is not None else 'N/A',
            "Comprobante": serie_numero.text if serie_numero is not None else 'N/A',
            "RUC": ruc_val, "Razón Social": razon_social.text if razon_social is not None else 'N/A',
            "Base Imponible": round(total - igv, 2), "IGV (18%)": round(igv, 2), "Total": round(total, 2), 
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
        
        ruc_match = re.search(r'\b(10|20)\d{9}\b', texto_completo)
        serie_match = re.search(r'\b[F|E|B][A-Z0-9]{3}-\d{1,8}\b', texto_completo)
        total_match = re.search(r'(?:TOTAL|Total|Importe Total).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo)
        
        total = float(total_match.group(1).replace(',', '')) if total_match else 0.0
        base_imponible = total / 1.18 if total > 0 else 0.0
        igv = total - base_imponible if total > 0 else 0.0
        ruc_val = ruc_match.group(0) if ruc_match else 'N/A'
        
        categoria = "Venta" if ruc_val == RUC_RESTAURANTE else "Compra"
        
        return {
            "Archivo": file_obj.name, "Tipo": "PDF", "Fecha": datetime.datetime.now().strftime('%Y-%m-%d'),
            "Comprobante": serie_match.group(0) if serie_match else "N/A", "RUC": ruc_val,
            "Razón Social": "Por verificar (PDF)", "Base Imponible": round(base_imponible, 2),
            "IGV (18%)": round(igv, 2), "Total": round(total, 2), "Categoría": categoria, 
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
            
            igv_v = ventas_mes['IGV (18%)'].sum() if not ventas_mes.empty else 0.0
            igv_c = compras_mes['IGV (18%)'].sum() if not compras_mes.empty else 0.0
            total_v = ventas_mes['Total'].sum() if not ventas_mes.empty else 0.0
            total_c = compras_mes['Total'].sum() if not compras_mes.empty else 0.0
            
            diferencia_igv = igv_v - igv_c
            igv_neto_pagar = max(0.0, diferencia_igv)
            saldo_a_favor = abs(min(0.0, diferencia_igv))
            
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
                        <div class="metric-label">📈 IGV Cobrado</div>
                        <div class="metric-value">S/ {igv_v:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            with cols[3]:
                if saldo_a_favor > 0:
                    st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">💚 SALDO A TU FAVOR</div>
                            <div class="metric-value">S/ {saldo_a_favor:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                        <div class="metric-card alert">
                            <div class="metric-label">🏛️ POR PAGAR A SUNAT</div>
                            <div class="metric-value">S/ {igv_neto_pagar:,.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
            
            st.divider()
            
            # --- GRÁFICO MEJORADO ---
            st.markdown('<p class="subtitle">Comparativa IGV del Periodo</p>', unsafe_allow_html=True)
            
            col_grafico1, col_grafico2 = st.columns(2)
            
            with col_grafico1:
                resumen_grafico = pd.DataFrame({
                    "Concepto": ["IGV Ventas", "IGV Compras"],
                    "Monto (S/)": [igv_v, igv_c]
                }).set_index("Concepto")
                st.bar_chart(resumen_grafico)
            
            with col_grafico2:
                # Resumen de cantidad de documentos
                resumen_docs = pd.DataFrame({
                    "Tipo": ["Ventas", "Compras"],
                    "Cantidad": [len(ventas_mes), len(compras_mes)]
                }).set_index("Tipo")
                st.bar_chart(resumen_docs)
            
            st.divider()
            
            # --- TABLAS DE DATOS ---
            st.markdown('<p class="subtitle">Detalle de Ventas</p>', unsafe_allow_html=True)
            if not ventas_mes.empty:
                st.dataframe(
                    ventas_mes[['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Total', 'IGV (18%)']].sort_values('Fecha', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("📭 No hay ventas registradas en este periodo")

            st.markdown('<p class="subtitle">Detalle de Compras</p>', unsafe_allow_html=True)
            if not compras_mes.empty:
                st.dataframe(
                    compras_mes[['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Total', 'IGV (18%)']].sort_values('Fecha', ascending=False),
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

# --- SECCIÓN 2: CARGA RÁPIDA DE COMPROBANTES ---
st.markdown('<div class="section-title">📤 Registrar Nuevos Comprobantes</div>', unsafe_allow_html=True)
st.markdown('Arrastra tus archivos XML o PDF. El sistema detectará automáticamente si es venta o compra.', unsafe_allow_html=True)

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
        
        if llave_actual in facturas_ya_registradas and datos.get('RUC') != 'N/A':
            datos['Estado'] = '⚠️ Duplicado'
            
        datos_procesados.append(datos)

    df_todos = pd.DataFrame(datos_procesados)

    if not df_todos.empty:
        st.divider()
        
        df_validos = df_todos[df_todos['Estado'] == 'OK']
        
        total_igv_ventas_lote = df_validos[df_validos['Categoría'] == 'Venta']['IGV (18%)'].sum() if 'IGV (18%)' in df_validos.columns else 0.0
        total_igv_compras_lote = df_validos[df_validos['Categoría'] == 'Compra']['IGV (18%)'].sum() if 'IGV (18%)' in df_validos.columns else 0.0
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
        st.dataframe(
            df_todos[['Archivo', 'Tipo', 'Categoría', 'RUC', 'Total', 'IGV (18%)', 'Estado']],
            use_container_width=True,
            hide_index=True
        )
        
        # --- ADVERTENCIAS ---
        duplicados = df_todos[df_todos['Estado'] == '⚠️ Duplicado']
        if not duplicados.empty:
            st.markdown(f"""
                <div class="warning-box">
                <strong>⚠️ Atención:</strong> Se detectaron {len(duplicados)} comprobante(s) duplicado(s) que ya están registrados. No se incluirán en el registro.
                </div>
                """, unsafe_allow_html=True)
        
        errores = df_todos[df_todos['Estado'].str.contains('Error|Revisar', na=False)]
        if not errores.empty:
            st.markdown(f"""
                <div class="warning-box">
                <strong>⚠️ Revisión Manual:</strong> {len(errores)} archivo(s) necesita(n) revisión. Verifica los datos antes de registrar.
                </div>
                """, unsafe_allow_html=True)
        
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
                            
                            st.success(f"✅ ¡Éxito! Se registraron {exitos} comprobante(s). Recarga arriba para ver los saldos actualizados.")
                        except Exception as e:
                            st.error(f"❌ Error al registrar: {e}")
        else:
            st.warning("⚠️ No hay comprobantes válidos para registrar. Revisa los errores arriba.")
