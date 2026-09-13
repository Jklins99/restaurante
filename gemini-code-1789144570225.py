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
    :root { --primary: #2E7D32; --accent: #D32F2F; --neutral: #424242; --light-bg: #F5F5F5; --white: #FFFFFF; }
    .main-title { font-size: 2.5em; font-weight: 700; color: #1B5E20; margin-bottom: 0.2em; letter-spacing: -0.5px; }
    .section-title { font-size: 1.3em; font-weight: 600; color: #2E7D32; margin-top: 1.5em; margin-bottom: 1em; border-left: 4px solid #2E7D32; padding-left: 12px; }
    .metric-card { background: linear-gradient(135deg, #FFFFFF 0%, #F5F5F5 100%); border-radius: 8px; padding: 1.5em; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid #2E7D32; transition: transform 0.2s ease, box-shadow 0.2s ease; }
    .metric-card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.12); }
    .metric-label { font-size: 0.85em; color: #666; font-weight: 500; margin-bottom: 0.5em; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-value { font-size: 2em; font-weight: 700; color: #1B5E20; }
    .metric-card.alert { border-left-color: #D32F2F; }
    .metric-card.alert .metric-value { color: #D32F2F; }
    .stButton > button { background: linear-gradient(135deg, #2E7D32 0%, #1B5E20 100%); color: white; font-weight: 600; border: none; border-radius: 6px; padding: 0.75em 2em; transition: all 0.3s ease; box-shadow: 0 2px 8px rgba(46, 125, 50, 0.3); }
    .stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(46, 125, 50, 0.4); }
    hr { border: none; height: 1px; background: linear-gradient(90deg, transparent, #DDD, transparent); margin: 2em 0; }
    .warning-box { background: #FFEBEE; border-left: 4px solid #D32F2F; padding: 1em; border-radius: 4px; margin: 1em 0; }
    .subtitle { font-size: 1.1em; color: #555; margin: 1.5em 0 0.5em 0; font-weight: 500; }
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES DE GOOGLE SHEETS ---
def obtener_servicio_sheets():
    try:
        creds_dict = st.secrets["gcp_service_account"]
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return build('sheets', 'v4', credentials=creds)
    except Exception as e:
        st.error(f"Error de autenticación con Google Sheets: {e}")
        return None

def obtener_facturas_registradas(sheets_service):
    registradas = set()
    if not sheets_service: return registradas
    try:
        resultado = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:K').execute()
        filas = resultado.get('values', [])
        for fila in filas[1:]:
            if len(fila) >= 4:
                llave = f"{fila[3]}-{fila[2]}"  
                registradas.add(llave)
    except Exception:
        pass
    return registradas

def descargar_historial_sheets(sheets_service):
    if not sheets_service: return pd.DataFrame()
    try:
        resultado = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:K').execute()
        filas = resultado.get('values', [])
        if len(filas) > 1:
            cabeceras = ['Fecha', 'Tipo', 'Comprobante', 'RUC', 'Razón Social', 'Base Imponible', 'IGV', 'Total', 'Categoría', 'Documento', 'Tasa %']
            datos = filas[1:]
            datos_normalizados = []
            for fila in datos:
                while len(fila) < len(cabeceras):
                    fila.append("0")
                datos_normalizados.append(fila[:len(cabeceras)])
                
            df = pd.DataFrame(datos_normalizados, columns=cabeceras)
            for col in ['Base Imponible', 'IGV', 'Total']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            return df
    except Exception:
        pass
    return pd.DataFrame()

def guardar_lote_en_sheets(sheets_service, df_validos):
    valores = []
    for _, datos in df_validos.iterrows():
        fila = [
            str(datos.get('Fecha', '')), str(datos.get('Tipo', '')), str(datos.get('Comprobante', '')), 
            str(datos.get('RUC', '')), str(datos.get('Razón Social', '')), float(datos.get('Base Imponible', 0.0)), 
            float(datos.get('IGV', 0.0)), float(datos.get('Total', 0.0)), str(datos.get('Categoría', '')), 
            str(datos.get('Documento', 'N/A')), "N/A"  
        ]
        valores.append(fila)
    if valores:
        cuerpo = {'values': valores}
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='A:K',
            valueInputOption='USER_ENTERED', body=cuerpo
        ).execute()

def obtener_sheet_id(sheets_service, spreadsheet_id):
    metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return metadata.get('sheets', '')[0].get("properties", {}).get("sheetId")

def eliminar_filas_sheets(sheets_service, indices_a_eliminar):
    sheet_id = obtener_sheet_id(sheets_service, SPREADSHEET_ID)
    indices_a_eliminar.sort(reverse=True)
    requests = []
    for idx in indices_a_eliminar:
        requests.append({
            "deleteDimension": {
                "range": { "sheetId": sheet_id, "dimension": "ROWS", "startIndex": idx, "endIndex": idx + 1 }
            }
        })
    if requests:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()

# --- EXTRACCIÓN Y CLASIFICACIÓN (XML y PDF) ---
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
        serie_val = serie_numero.text if serie_numero is not None else 'N/A'
        
        tipo_doc_node = root.find('.//cbc:InvoiceTypeCode', ns)
        if tipo_doc_node is not None:
            codigo = tipo_doc_node.text.strip()
            if codigo == '01': tipo_doc = "Factura"
            elif codigo == '03': tipo_doc = "Boleta"
            elif codigo == '07': tipo_doc = "Nota de Crédito"
            else: tipo_doc = "Otro"
        else:
            serie_upper = serie_val.upper()
            if serie_upper.startswith('B') or serie_upper.startswith('EB'): tipo_doc = "Boleta"
            elif serie_upper.startswith('F') or serie_upper.startswith('E'): tipo_doc = "Factura"
            else: tipo_doc = "Otro"
        
        base_calculada = total - igv
        categoria = "Venta" if ruc_val == RUC_RESTAURANTE else "Compra"
        
        if tipo_doc == "Nota de Crédito":
            total = -abs(total)
            igv = -abs(igv)
            base_calculada = -abs(base_calculada)

        return {
            "Archivo": file_obj.name, "Tipo": "XML", "Documento": tipo_doc, "Fecha": fecha.text if fecha is not None else 'N/A',
            "Comprobante": serie_val, "RUC": ruc_val, "Razón Social": razon_social.text if razon_social is not None else 'N/A',
            "Base Imponible": round(base_calculada, 2), "IGV": round(igv, 2), "Total": round(total, 2), 
            "Categoría": categoria, "Estado": "OK"
        }
    except Exception:
        return {"Archivo": file_obj.name, "Tipo": "XML", "Categoría": "Desconocido", "Estado": "Error de lectura"}

def procesar_factura_pdf(file_obj):
    try:
        file_obj.seek(0)
        texto_completo = ""
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto: texto_completo += texto + "\n"
        
        texto_upper = texto_completo.upper()
        if "ANULADO" in texto_upper or "ANULADA" in texto_upper:
            return {
                "Archivo": file_obj.name, "Tipo": "PDF", "Documento": "Anulado", "Fecha": "N/A",
                "Comprobante": "N/A", "RUC": "N/A", "Razón Social": "Documento Anulado", 
                "Base Imponible": 0.0, "IGV": 0.0, "Total": 0.0, "Categoría": "Descartado", "Estado": "🚫 Factura Anulada"
            }

        fecha_str = None
        fecha_match = re.search(r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b', texto_completo)
        if fecha_match:
            try:
                dia, mes, anio = fecha_match.groups()
                fecha_documento = datetime.datetime(int(anio), int(mes), int(dia))
                fecha_str = fecha_documento.strftime('%Y-%m-%d')
            except:
                pass
        
        ruc_match = re.search(r'\b(10|20)\d{9}\b', texto_completo)
        serie_match = re.search(r'\b([FBE][A-Z0-9]{3}-\d{1,8})\b', texto_completo, re.IGNORECASE)
        serie_val = serie_match.group(1).upper() if serie_match else "N/A"
        
        if "NOTA DE CRÉDITO" in texto_upper or "NOTA DE CREDITO" in texto_upper:
            tipo_doc = "Nota de Crédito"
        elif "BOLETA DE VENTA" in texto_upper or "BOLETA ELECT" in texto_upper:
            tipo_doc = "Boleta"
        elif "FACTURA ELECT" in texto_upper or "FACTURA DE VENTA" in texto_upper:
            tipo_doc = "Factura"
        else:
            if serie_val.startswith('B') or serie_val.startswith('EB'): tipo_doc = "Boleta"
            else: tipo_doc = "Factura"
        
        total_match = re.search(r'(?:TOTAL|Total|Importe Total).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo)
        igv_match = re.search(r'(?:IGV|I\.G\.V\.|Impuesto).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo, re.IGNORECASE)
        base_match = re.search(r'(?:SUBTOTAL|Sub Total|Subtotal|Op\. Gravadas|Base Imponible).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo, re.IGNORECASE)
        
        total = float(total_match.group(1).replace(',', '')) if total_match else 0.0
        igv = float(igv_match.group(1).replace(',', '')) if igv_match else 0.0
        base_imponible = float(base_match.group(1).replace(',', '')) if base_match else (total - igv)
        
        if tipo_doc == "Nota de Crédito":
            total = -abs(total)
            igv = -abs(igv)
            base_imponible = -abs(base_imponible)

        ruc_val = ruc_match.group(0) if ruc_match else 'N/A'
        categoria = "Venta" if ruc_val == RUC_RESTAURANTE else "Compra"
        
        razon_social = "Por verificar (PDF)"
        rs_match = re.search(r'(?:Señor\(es\)|Señor(?:es)?|Cliente|Razón Social|Nombre)\s*:\s*([^\n]+)', texto_completo, re.IGNORECASE)
        
        if rs_match:
            razon_temp = rs_match.group(1).strip()
            if len(razon_temp) > 3 and not re.search(r'\b(RUC|DNI)\b', razon_temp, re.IGNORECASE):
                razon_social = razon_temp
        
        if razon_social == "Por verificar (PDF)":
            lineas = [linea.strip() for linea in texto_completo.split('\n') if linea.strip()]
            for linea in lineas[:5]:
                if len(linea) > 4 and not re.search(r'(FACTURA|BOLETA|RUC|FECHA|\d{11})', linea, re.IGNORECASE):
                    razon_social = linea
                    break

        estado_doc = "OK"
        if total == 0 or not igv_match or not fecha_str:
            estado_doc = "Revisar Manualmente"
            
        return {
            "Archivo": file_obj.name, "Tipo": "PDF", "Documento": tipo_doc, "Fecha": fecha_str if fecha_str else 'N/A',
            "Comprobante": serie_val, "RUC": ruc_val, "Razón Social": razon_social, 
            "Base Imponible": round(base_imponible, 2), "IGV": round(igv, 2), "Total": round(total, 2), 
            "Categoría": categoria, "Estado": estado_doc
        }
    except Exception:
        return {"Archivo": file_obj.name, "Tipo": "PDF", "Categoría": "Desconocido", "Estado": "Error de lectura"}

# --- INTERFAZ PRINCIPAL ---
st.markdown('<div class="main-title">🍽️ Gestor de Facturación</div>', unsafe_allow_html=True)
st.markdown('Sistema automático de registro de ventas y compras para tu restaurante', unsafe_allow_html=True)

sheets_service = obtener_servicio_sheets()
facturas_ya_registradas = obtener_facturas_registradas(sheets_service)

# --- SECCIÓN 1: DASHBOARD MENSUAL Y GESTIÓN ---
st.markdown('<div class="section-title">📊 Resumen y Gestión del Periodo</div>', unsafe_allow_html=True)

if sheets_service:
    df_historial = descargar_historial_sheets(sheets_service)
    
    if not df_historial.empty:
        df_historial['Fecha'] = pd.to_datetime(df_historial['Fecha'], errors='coerce')
        df_historial['Mes'] = df_historial['Fecha'].dt.to_period('M').astype(str)
        meses_disponibles = sorted(df_historial['Mes'].dropna().unique(), reverse=True)
        
        if meses_disponibles:
            col1, col2 = st.columns([3, 1])
            with col1:
                mes_seleccionado = st.selectbox("Selecciona el mes a revisar", meses_disponibles, label_visibility="collapsed")
            
            df_mes = df_historial[df_historial['Mes'] == mes_seleccionado]
            ventas_mes = df_mes[df_mes['Categoría'] == 'Venta']
            compras_mes = df_mes[df_mes['Categoría'] == 'Compra']
            
            igv_col = 'IGV' if 'IGV' in df_mes.columns else 'IGV (18%)'
            igv_v = ventas_mes[igv_col].sum() if not ventas_mes.empty else 0.0
            igv_c = compras_mes[igv_col].sum() if not compras_mes.empty else 0.0
            total_v = ventas_mes['Total'].sum() if not ventas_mes.empty else 0.0
            total_c = compras_mes['Total'].sum() if not compras_mes.empty else 0.0
            
            diferencia_igv = igv_v - igv_c
            saldo_a_favor_mes = abs(min(0.0, diferencia_igv))
            
            fecha_mes_seleccionado = pd.Period(mes_seleccionado, freq='M')
            fecha_mes_anterior = fecha_mes_seleccionado - 1
            mes_anterior_str = fecha_mes_anterior.strftime('%Y-%m')
            
            saldo_anterior = 0.0
            if mes_anterior_str in df_historial['Mes'].values:
                df_mes_anterior = df_historial[df_historial['Mes'] == mes_anterior_str]
                igv_v_ant = df_mes_anterior[df_mes_anterior['Categoría'] == 'Venta'][igv_col].sum() if not df_mes_anterior[df_mes_anterior['Categoría'] == 'Venta'].empty else 0.0
                igv_c_ant = df_mes_anterior[df_mes_anterior['Categoría'] == 'Compra'][igv_col].sum() if not df_mes_anterior[df_mes_anterior['Categoría'] == 'Compra'].empty else 0.0
                saldo_anterior = abs(min(0.0, igv_v_ant - igv_c_ant))
            
            igv_neto_pagar = max(0.0, diferencia_igv - saldo_anterior)
            saldo_a_favor = max(0.0, saldo_anterior - diferencia_igv)
            
            cols = st.columns(4, gap="medium")
            with cols[0]: st.markdown(f'<div class="metric-card"><div class="metric-label">💰 Ingresos</div><div class="metric-value">S/ {total_v:,.2f}</div></div>', unsafe_allow_html=True)
            with cols[1]: st.markdown(f'<div class="metric-card"><div class="metric-label">📦 Egresos</div><div class="metric-value">S/ {total_c:,.2f}</div></div>', unsafe_allow_html=True)
            with cols[2]: st.markdown(f'<div class="metric-card"><div class="metric-label">📈 IGV Neto del Mes</div><div class="metric-value">{"+" if diferencia_igv >= 0 else "-"} S/ {abs(diferencia_igv):,.2f}</div><div style="font-size: 0.75em; color: #999; margin-top: 0.5em;">Cobrado: S/ {igv_v:,.2f} | Crédito: S/ {igv_c:,.2f}</div></div>', unsafe_allow_html=True)
            with cols[3]:
                if saldo_anterior > 0: st.markdown(f'<div class="metric-card" style="border-left-color: #FF6F00;"><div class="metric-label">↙️ Saldo Anterior</div><div class="metric-value" style="color: #FF6F00; font-size: 1.5em;">S/ {saldo_anterior:,.2f}</div></div>', unsafe_allow_html=True)
                else: st.markdown('<div class="metric-card"><div class="metric-label">✓ Sin Saldo Anterior</div><div class="metric-value" style="color: #666; font-size: 1.5em;">-</div></div>', unsafe_allow_html=True)
            
            st.divider()
            st.markdown('<p class="subtitle">📋 Resumen Final (Considerando Arrastre)</p>', unsafe_allow_html=True)
            cols_resumen = st.columns(3, gap="medium")
            with cols_resumen[0]: st.markdown(f'<div class="metric-card"><div class="metric-label">IGV Neto del Mes</div><div class="metric-value" style="color: #1565C0;">{"+" if diferencia_igv >= 0 else "-"} S/ {abs(diferencia_igv):,.2f}</div></div>', unsafe_allow_html=True)
            with cols_resumen[1]:
                if saldo_anterior > 0: st.markdown(f'<div class="metric-card" style="border-left-color: #FF6F00;"><div class="metric-label">Menos Saldo Anterior</div><div class="metric-value" style="color: #FF6F00;">- S/ {saldo_anterior:,.2f}</div></div>', unsafe_allow_html=True)
                else: st.markdown('<div class="metric-card"><div class="metric-label">Menos Saldo Anterior</div><div class="metric-value" style="color: #999;">-</div></div>', unsafe_allow_html=True)
            with cols_resumen[2]:
                if saldo_a_favor > 0: st.markdown(f'<div class="metric-card"><div class="metric-label">💚 RESULTADO</div><div class="metric-value" style="color: #2E7D32;">SALDO A FAVOR</div><div style="font-size: 1.3em; color: #2E7D32; font-weight: 700; margin-top: 0.3em;">S/ {saldo_a_favor:,.2f}</div></div>', unsafe_allow_html=True)
                else: st.markdown(f'<div class="metric-card alert"><div class="metric-label">🏛️ RESULTADO</div><div class="metric-value" style="color: #D32F2F;">POR PAGAR</div><div style="font-size: 1.3em; color: #D32F2F; font-weight: 700; margin-top: 0.3em;">S/ {igv_neto_pagar:,.2f}</div></div>', unsafe_allow_html=True)
            
            st.divider()
            
            # --- TABLAS INTERACTIVAS PARA GESTIÓN DE ELIMINACIÓN ---
            filas_a_eliminar = []
            
            st.markdown('<p class="subtitle">Detalle de Ventas</p>', unsafe_allow_html=True)
            if not ventas_mes.empty:
                ventas_edicion = ventas_mes.copy()
                ventas_edicion.insert(0, '🗑️ Eliminar', False)
                cols_ventas = [col for col in ['Fecha', 'Tipo', 'Documento', 'Comprobante', 'Razón Social', 'Total', igv_col] if col in ventas_mes.columns]
                columnas_display_v = ['🗑️ Eliminar'] + cols_ventas
                
                # Mantenemos el índice original pero ordenamos visualmente por fecha
                ventas_edicion = ventas_edicion.sort_values('Fecha', ascending=False)
                
                editado_v = st.data_editor(
                    ventas_edicion[columnas_display_v],
                    use_container_width=True,
                    disabled=cols_ventas,  # Bloqueamos datos, solo dejamos el checkbox activo
                    key="editor_ventas"
                )
                filas_a_eliminar.extend(editado_v[editado_v['🗑️ Eliminar']].index.tolist())
            else: 
                st.info("📭 No hay ventas registradas en este periodo")

            st.markdown('<p class="subtitle">Detalle de Compras</p>', unsafe_allow_html=True)
            if not compras_mes.empty:
                compras_edicion = compras_mes.copy()
                compras_edicion.insert(0, '🗑️ Eliminar', False)
                cols_compras = [col for col in ['Fecha', 'Tipo', 'Documento', 'Comprobante', 'Razón Social', 'Total', igv_col] if col in compras_mes.columns]
                columnas_display_c = ['🗑️ Eliminar'] + cols_compras
                
                compras_edicion = compras_edicion.sort_values('Fecha', ascending=False)
                
                editado_c = st.data_editor(
                    compras_edicion[columnas_display_c],
                    use_container_width=True,
                    disabled=cols_compras,
                    key="editor_compras"
                )
                filas_a_eliminar.extend(editado_c[editado_c['🗑️ Eliminar']].index.tolist())
            else: 
                st.info("📭 No hay compras registradas en este periodo")
                
            # --- BOTÓN DINÁMICO DE ELIMINACIÓN ---
            if filas_a_eliminar:
                st.warning(f"⚠️ Estás a punto de **ELIMINAR PERMANENTEMENTE** {len(filas_a_eliminar)} registro(s) de este mes.")
                if st.button("🚨 Eliminar Registros Seleccionados", type="primary"):
                    with st.spinner("Eliminando datos de Google Sheets y recalculando saldos..."):
                        try:
                            # Ajuste de índice para la API (+1 porque la API toma la cabecera como 0)
                            indices_api = [idx + 1 for idx in filas_a_eliminar]
                            eliminar_filas_sheets(sheets_service, indices_api)
                            st.success("✅ Registros eliminados exitosamente. Actualizando Dashboard...")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error al eliminar los registros: {e}")

        else:
            st.info("📭 No se encontraron fechas válidas en el historial registrado.")
    else:
        st.info("📭 El Google Sheets aún no tiene registros guardados. Sube tu primer lote abajo.")

st.divider()

# --- SECCIÓN 2: CARGA RÁPIDA DE COMPROBANTES CON CACHÉ ---
st.markdown('<div class="section-title">📤 Registrar Nuevos Comprobantes</div>', unsafe_allow_html=True)

if 'df_procesado' not in st.session_state: st.session_state.df_procesado = pd.DataFrame()
if 'archivos_nombres' not in st.session_state: st.session_state.archivos_nombres = []

archivos_subidos = st.file_uploader("Selecciona tus archivos", type=['xml', 'pdf'], accept_multiple_files=True, label_visibility="collapsed")

if archivos_subidos:
    nombres_actuales = [f.name for f in archivos_subidos]
    if nombres_actuales != st.session_state.archivos_nombres:
        datos_procesados = []
        with st.spinner("Procesando y extrayendo datos (solo una vez por lote)..."):
            for f in archivos_subidos:
                datos = parse_sunat_xml(f) if f.name.lower().endswith('.xml') else procesar_factura_pdf(f)
                llave_actual = f"{datos.get('RUC')}-{datos.get('Comprobante')}"
                
                try: datos['Mes Facturación'] = pd.to_datetime(datos.get('Fecha')).strftime('%Y-%m')
                except: datos['Mes Facturación'] = 'N/A'
                
                if llave_actual in facturas_ya_registradas and datos.get('RUC') != 'N/A':
                    datos['Estado'] = '⚠️ Duplicado'
                datos_procesados.append(datos)

            st.session_state.df_procesado = pd.DataFrame(datos_procesados)
            st.session_state.archivos_nombres = nombres_actuales
    df_todos = st.session_state.df_procesado
else:
    st.session_state.df_procesado = pd.DataFrame()
    st.session_state.archivos_nombres = []
    df_todos = pd.DataFrame()

if not df_todos.empty:
    st.divider()
    df_validos = df_todos[df_todos['Estado'] == 'OK'].copy()
    
    if not df_validos.empty:
        df_validos.insert(0, '🚫 Anulada / Descartar', False)
        st.markdown('<p class="subtitle">Vista Previa del Lote (Marca la casilla si el documento fue anulado)</p>', unsafe_allow_html=True)
        
        columnas_display = ['🚫 Anulada / Descartar', 'Archivo', 'Mes Facturación', 'Documento', 'Razón Social', 'Categoría', 'Comprobante', 'Total', 'IGV']
        df_editado = st.data_editor(
            df_validos[columnas_display],
            use_container_width=True, hide_index=True,
            disabled=columnas_display[1:]
        )
        
        df_final_a_guardar = df_validos[~df_editado['🚫 Anulada / Descartar']].copy()
        st.divider()
        
        igv_col_lote = 'IGV' if 'IGV' in df_final_a_guardar.columns else 'IGV (18%)'
        total_igv_ventas_lote = df_final_a_guardar[df_final_a_guardar['Categoría'] == 'Venta'][igv_col_lote].sum() if not df_final_a_guardar.empty else 0.0
        total_igv_compras_lote = df_final_a_guardar[df_final_a_guardar['Categoría'] == 'Compra'][igv_col_lote].sum() if not df_final_a_guardar.empty else 0.0
        
        cols_lote = st.columns(3, gap="medium")
        with cols_lote[0]: st.metric("📈 IGV Ventas (A Registrar)", f"S/ {total_igv_ventas_lote:,.2f}")
        with cols_lote[1]: st.metric("📦 IGV Compras (A Registrar)", f"S/ {total_igv_compras_lote:,.2f}")
        with cols_lote[2]: st.metric("✅ Total Válidos", len(df_final_a_guardar))
        
        errores = df_todos[df_todos['Estado'].str.contains('Error|Revisar|Duplicado', na=False)]
        if not errores.empty:
            st.markdown(f'<div class="warning-box"><strong>⚠️ Atención:</strong> {len(errores)} archivo(s) requiere(n) revisión manual por errores o duplicados.</div>', unsafe_allow_html=True)
        
        st.divider()
        if not df_final_a_guardar.empty:
            if st.button("🚀 Registrar Lote en Google Sheets", type="primary"):
                if sheets_service:
                    with st.spinner("💾 Guardando todos los registros a la vez..."):
                        try:
                            guardar_lote_en_sheets(sheets_service, df_final_a_guardar)
                            st.success(f"✅ ¡Éxito! Se registraron {len(df_final_a_guardar)} comprobante(s). Recarga la página para ver el dashboard actualizado.")
                            st.session_state.df_procesado = pd.DataFrame()
                            st.session_state.archivos_nombres = []
                        except Exception as e:
                            st.error(f"❌ Error al registrar en lotes: {e}")
                else: st.error("No hay conexión a Google Sheets.")
        else: st.warning("⚠️ Has marcado todos los documentos como anulados o no hay comprobantes válidos para registrar.")
