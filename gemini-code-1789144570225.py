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
RUC_RESTAURANTE = '10402504051' # Tu RUC para detectar ventas automáticas
# =========================

st.set_page_config(page_title="Gestor de Facturación - Restaurante", page_icon="🍽️", layout="wide")

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

# --- INTERFAZ PRINCIPAL (UNICA VISTA) ---
st.title("🍽️ Gestor de Facturación - Restaurante")

try:
    sheets_service = obtener_servicio_sheets()
    facturas_ya_registradas = obtener_facturas_registradas(sheets_service)
except Exception:
    sheets_service = None
    facturas_ya_registradas = set()

# --- SECCIÓN 1: DASHBOARD Y MÉTRICAS DEL MES SELECCIONADO ---
st.subheader("📊 Panel de Control e Historial Mensual")
if sheets_service:
    df_historial = descargar_historial_sheets(sheets_service)
    if not df_historial.empty:
        df_historial['Fecha'] = pd.to_datetime(df_historial['Fecha'], errors='coerce')
        df_historial['Mes'] = df_historial['Fecha'].dt.to_period('M').astype(str)
        
        meses_disponibles = sorted(df_historial['Mes'].dropna().unique(), reverse=True)
        if meses_disponibles:
            mes_seleccionado = st.selectbox("📅 Selecciona el Periodo (Mes)", meses_disponibles)
            
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
            
            # --- TARJETAS DE MÉTRICAS ---
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"Ventas Totales ({mes_seleccionado})", f"S/ {total_v:,.2f}")
            c2.metric(f"Compras Totales ({mes_seleccionado})", f"S/ {total_c:,.2f}")
            c3.metric("IGV Cobrado vs Pagado", f"S/ {igv_v:,.2f} / S/ {igv_c:,.2f}")
            
            if saldo_a_favor > 0:
                c4.metric("💚 SALDO A FAVOR (Crédito)", f"S/ {saldo_a_favor:,.2f}")
            else:
                c4.metric("🏛️ A PAGAR A SUNAT", f"S/ {igv_neto_pagar:,.2f}")
            
            st.divider()
            
            # --- GRÁFICO SIMPLIFICADO Y LIMPIO ---
            st.write(f"### ⚖️ Comparativa de IGV (Ventas vs Compras) - Periodo {mes_seleccionado}")
            resumen_grafico = pd.DataFrame({
                "Concepto": ["IGV Ventas", "IGV Compras"],
                "Monto (S/)": [igv_v, igv_c]
            }).set_index("Concepto")
            
            st.bar_chart(resumen_grafico)
            
            st.divider()
            st.write("### 📤 Registro de Ventas del Mes")
            if not ventas_mes.empty:
                st.dataframe(ventas_mes.drop(columns=['Categoría']), use_container_width=True)
            else:
                st.info("No hay ventas registradas en este periodo.")

            st.write("### 📥 Registro de Compras del Mes")
            if not compras_mes.empty:
                st.dataframe(compras_mes.drop(columns=['Categoría']), use_container_width=True)
            else:
                st.info("No hay compras registradas en este periodo.")
        else:
            st.info("No se encontraron fechas válidas en el historial registrado.")
    else:
        st.info("El Google Sheets aún no tiene registros guardados. Sube tu primer lote abajo.")
else:
    st.error("No se pudo conectar con Google Sheets para cargar el historial.")

st.divider()

# --- SECCIÓN 2: BANDEJA DE CARGA RÁPIDA DE NUEVOS COMPROBANTES ---
st.subheader("📤 Carga Rápida de Nuevos Comprobantes")
st.write("Arrastra tus archivos nuevos (XML o PDF) juntos. El sistema detectará automáticamente si es Venta o Compra.")

archivos_subidos = st.file_uploader("📂 Arrastra o selecciona tus archivos (XML / PDF)", type=['xml', 'pdf'], accept_multiple_files=True)

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
    
    m1, m2, m3 = st.columns(3)
    m1.metric("IGV Ventas (Lote Nuevo)", f"S/ {total_igv_ventas_lote:,.2f}")
    m2.metric("Crédito Fiscal (Lote Nuevo)", f"S/ {total_igv_compras_lote:,.2f}")
    m3.metric("IGV NETO Estimado (Lote Nuevo)", f"S/ {max(0, total_igv_ventas_lote - total_igv_compras_lote):,.2f}")
    
    st.write("### Vista Previa del Lote Actual")
    st.dataframe(df_todos, use_container_width=True)
    
    duplicados = df_todos[df_todos['Estado'] == '⚠️ Duplicado']
    if not duplicados.empty:
        st.warning(f"¡Atención! Se detectaron {len(duplicados)} comprobantes duplicados que ya están registrados en el Excel.")

    st.divider()
    if not df_validos.empty:
        if st.button("🚀 Registrar Lote en Google Sheets", type="primary"):
            with st.spinner("Guardando en la base de datos..."):
                try:
                    exitos = 0
                    for index, fila in df_validos.iterrows():
                        guardar_en_sheets(sheets_service, fila, fila['Categoría'])
                        exitos += 1
                    
                    st.success(f"¡Proceso exitoso! Se registraron {exitos} comprobantes en tu libro mayor. Recarga la página para ver los saldos actualizados arriba.")
                except Exception as e:
                    st.error(f"Ocurrió un error al registrar en Sheets: {e}")
