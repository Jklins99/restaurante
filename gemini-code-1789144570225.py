import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import pdfplumber
import re
import io
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# === TUS IDs DE GOOGLE ===
CARPETA_RAIZ_ID = '1RC6qUV87m5R0PLmQUTLn8cgQUbtitqUs'
SPREADSHEET_ID = '1811VY4-Xa4ZOf7j6MVd5zYFCdhpyxdtuwq1pD5mHlh4'
# =========================

st.set_page_config(page_title="Gestor de Facturación - Restaurante", page_icon="🍽️", layout="wide")

# --- FUNCIONES DE GOOGLE (DRIVE Y SHEETS) ---
def obtener_servicios_google():
    creds_dict = st.secrets["gcp_service_account"]
    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return drive_service, sheets_service

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

def obtener_o_crear_carpeta(drive_service, nombre_carpeta, parent_id):
    query = f"name='{nombre_carpeta}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    resultados = drive_service.files().list(q=query, fields="files(id, name)").execute()
    archivos = resultados.get('files', [])
    if archivos:
        return archivos[0]['id']
    else:
        metadata = {'name': nombre_carpeta, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
        carpeta = drive_service.files().create(body=metadata, fields='id').execute()
        return carpeta.get('id')

def subir_archivo_drive(file_obj, drive_service, id_dia):
    try:
        file_obj.seek(0)
        media = MediaIoBaseUpload(
            io.BytesIO(file_obj.read()), 
            mimetype='application/octet-stream', 
            resumable=True
        )
        metadata = {'name': file_obj.name, 'parents': [id_dia]}
        
        # supportsAllDrives=True evita la restricción de cuota en Service Accounts
        drive_service.files().create(
            body=metadata, 
            media_body=media, 
            fields='id',
            supportsAllDrives=True
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error subiendo archivo {file_obj.name}: {e}")
        return False

# --- FUNCIONES DE EXTRACCIÓN (XML y PDF) ---
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
        
        return {
            "Archivo": file_obj.name, "Tipo": "XML", "Fecha": fecha.text if fecha is not None else 'N/A',
            "Comprobante": serie_numero.text if serie_numero is not None else 'N/A',
            "RUC": ruc.text if ruc is not None else 'N/A', "Razón Social": razon_social.text if razon_social is not None else 'N/A',
            "Base Imponible": round(total - igv, 2), "IGV (18%)": round(igv, 2), "Total": round(total, 2), "Estado": "OK"
        }
    except Exception:
        return {"Archivo": file_obj.name, "Tipo": "XML", "Estado": "Error de lectura"}

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
        
        return {
            "Archivo": file_obj.name, "Tipo": "PDF", "Fecha": datetime.datetime.now().strftime('%Y-%m-%d'),
            "Comprobante": serie_match.group(0) if serie_match else "N/A", "RUC": ruc_match.group(0) if ruc_match else "N/A",
            "Razón Social": "Por verificar (PDF)", "Base Imponible": round(base_imponible, 2),
            "IGV (18%)": round(igv, 2), "Total": round(total, 2), "Estado": "OK" if total > 0 else "Revisar Manualmente"
        }
    except Exception:
        return {"Archivo": file_obj.name, "Tipo": "PDF", "Estado": "Error de lectura"}

# --- INTERFAZ ---
st.title("🍽️ Gestor de Facturación - Sistema Anti-Duplicados")

try:
    drive_service, sheets_service = obtener_servicios_google()
    facturas_ya_registradas = obtener_facturas_registradas(sheets_service)
except Exception:
    facturas_ya_registradas = set()

col1, col2 = st.columns(2)
with col1:
    ventas_files = st.file_uploader("📤 Ventas Emitidas (XML/PDF)", type=['xml', 'pdf'], accept_multiple_files=True, key="ventas")
with col2:
    compras_files = st.file_uploader("📥 Compras y Gastos (XML/PDF)", type=['xml', 'pdf'], accept_multiple_files=True, key="compras")

datos_procesados = []

def analizar_y_filtrar(archivos, categoria):
    for f in archivos:
        datos = parse_sunat_xml(f) if f.name.lower().endswith('.xml') else procesar_factura_pdf(f)
        llave_actual = f"{datos.get('RUC')}-{datos.get('Comprobante')}"
        
        if llave_actual in facturas_ya_registradas and datos.get('RUC') != 'N/A':
            datos['Estado'] = '⚠️ Duplicado'
        
        datos['Categoría'] = categoria
        datos_procesados.append(datos)

if ventas_files: analizar_y_filtrar(ventas_files, "Venta")
if compras_files: analizar_y_filtrar(compras_files, "Compra")

df_todos = pd.DataFrame(datos_procesados)

if not df_todos.empty:
    st.divider()
    df_validos = df_todos[df_todos['Estado'] == 'OK']
    
    total_igv_ventas = df_validos[df_validos['Categoría'] == 'Venta']['IGV (18%)'].sum() if 'IGV (18%)' in df_validos.columns else 0.0
    total_igv_compras = df_validos[df_validos['Categoría'] == 'Compra']['IGV (18%)'].sum() if 'IGV (18%)' in df_validos.columns else 0.0
    
    m1, m2, m3 = st.columns(3)
    m1.metric("IGV Nuevo Cobrado (Ventas)", f"S/ {total_igv_ventas:,.2f}")
    m2.metric("Nuevo Crédito Fiscal (Compras)", f"S/ {total_igv_compras:,.2f}")
    m3.metric("IGV NETO (Este lote)", f"S/ {max(0, total_igv_ventas - total_igv_compras):,.2f}")
    
    st.write("### Vista Previa de Archivos Subidos")
    st.dataframe(df_todos.drop(columns=['Categoría']), use_container_width=True)
    
    duplicados = df_todos[df_todos['Estado'] == '⚠️ Duplicado']
    if not duplicados.empty:
        st.warning(f"¡Atención! Se detectaron {len(duplicados)} facturas que ya habían sido subidas anteriormente.")

    st.divider()
  if not df_validos.empty:
        if st.button("🚀 Registrar Archivos Nuevos y Subir a Drive", type="primary"):
            with st.spinner("Guardando en la base de datos y subiendo archivos..."):
                try:
                    hoy = datetime.datetime.now()
                    id_anio = obtener_o_crear_carpeta(drive_service, hoy.strftime('%Y'), CARPETA_RAIZ_ID)
                    id_mes = obtener_o_crear_carpeta(drive_service, hoy.strftime('%m-%B'), id_anio)
                    id_dia = obtener_o_crear_carpeta(drive_service, hoy.strftime('%d-%m-%Y'), id_mes)
                    
                    todos_los_subidos = (ventas_files or []) + (compras_files or [])
                    
                    exitos_drive = 0
                    exitos_sheets = 0
                    
                    for index, fila in df_validos.iterrows():
                        archivo_original = next((f for f in todos_los_subidos if f.name == fila['Archivo']), None)
                        
                        # 1. Registro obligatorio en Google Sheets (Libro Mayor)
                        try:
                            guardar_en_sheets(sheets_service, fila, fila['Categoría'])
                            exitos_sheets += 1
                        except Exception as err_s:
                            st.error(f"Error registrando {fila['Comprobante']} en Sheets: {err_s}")
                        
                        # 2. Subida física del archivo a Drive
                        if archivo_original:
                            if subir_archivo_drive(archivo_original, drive_service, id_dia):
                                exitos_drive += 1
                    
                    st.success(f"Proceso finalizado: {exitos_sheets} facturas registradas en Sheets y {exitos_drive} archivos guardados en Drive.")
                except Exception as e:
                    st.error(f"Ocurrió un error al procesar el lote: {e}")
