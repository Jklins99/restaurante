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

# ID de la carpeta principal que creaste en Drive
CARPETA_RAIZ_ID = 'PEGA_AQUI_TU_ID_DE_CARPETA'

st.set_page_config(page_title="Gestor de Facturación - Restaurante", page_icon="🍽️", layout="wide")

# --- FUNCIONES DE GOOGLE DRIVE ---
def obtener_servicio_drive():
    creds_dict = st.secrets["gcp_service_account"]
    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

def obtener_o_crear_carpeta(drive_service, nombre_carpeta, parent_id):
    query = f"name='{nombre_carpeta}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    resultados = drive_service.files().list(q=query, fields="files(id, name)").execute()
    archivos = resultados.get('files', [])
    if archivos:
        return archivos[0]['id']
    else:
        metadata_carpeta = {'name': nombre_carpeta, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
        carpeta = drive_service.files().create(body=metadata_carpeta, fields='id').execute()
        return carpeta.get('id')

def subir_archivo_drive(file_obj, drive_service, id_dia):
    try:
        file_obj.seek(0)
        media = MediaIoBaseUpload(io.BytesIO(file_obj.read()), mimetype='application/octet-stream', resumable=True)
        file_metadata = {'name': file_obj.name, 'parents': [id_dia]}
        drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True
    except Exception as e:
        return False

# --- FUNCIONES DE EXTRACCIÓN ---
def parse_sunat_xml(file_obj):
    try:
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
        base_imponible = total - igv
        
        return {
            "Archivo": file_obj.name, "Tipo": "XML", "Comprobante": serie_numero.text if serie_numero is not None else 'N/A',
            "RUC": ruc.text if ruc is not None else 'N/A', "Base Imponible": round(base_imponible, 2),
            "IGV (18%)": round(igv, 2), "Total": round(total, 2), "Estado": "OK"
        }
    except Exception as e:
        return {"Archivo": file_obj.name, "Tipo": "XML", "Estado": f"Error"}

def procesar_factura_pdf(file_obj):
    try:
        texto_completo = ""
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                texto_extraido = page.extract_text()
                if texto_extraido: texto_completo += texto_extraido + "\n"
        
        ruc_match = re.search(r'\b(10|20)\d{9}\b', texto_completo)
        serie_match = re.search(r'\b[F|E|B][A-Z0-9]{3}-\d{1,8}\b', texto_completo)
        total_match = re.search(r'(?:TOTAL|Total|Importe Total).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo)
        
        total = float(total_match.group(1).replace(',', '')) if total_match else 0.0
        base_imponible = total / 1.18 if total > 0 else 0.0
        igv = total - base_imponible if total > 0 else 0.0
        
        return {
            "Archivo": file_obj.name, "Tipo": "PDF", "Comprobante": serie_match.group(0) if serie_match else "N/A",
            "RUC": ruc_match.group(0) if ruc_match else "N/A", "Base Imponible": round(base_imponible, 2),
            "IGV (18%)": round(igv, 2), "Total": round(total, 2), "Estado": "OK" if total > 0 else "Revisar Manualmente"
        }
    except Exception as e:
        return {"Archivo": file_obj.name, "Tipo": "PDF", "Estado": f"Error"}

def enrutador_archivos(file_obj):
    if file_obj.name.lower().endswith('.xml'): return parse_sunat_xml(file_obj)
    elif file_obj.name.lower().endswith('.pdf'): return procesar_factura_pdf(file_obj)
    return {"Archivo": file_obj.name, "Estado": "No soportado"}

# --- INTERFAZ ---
st.title("🍽️ Gestor de Facturación y Cierre de Impuestos")
st.markdown("Sube los archivos diarios. Al finalizar, guárdalos en Google Drive.")

col1, col2 = st.columns(2)
with col1:
    ventas_files = st.file_uploader("📤 Ventas (Facturas Emitidas)", type=['xml', 'pdf'], accept_multiple_files=True)
with col2:
    compras_files = st.file_uploader("📥 Compras (Gastos e Insumos)", type=['xml', 'pdf'], accept_multiple_files=True)

df_ventas = pd.DataFrame([enrutador_archivos(f) for f in ventas_files] if ventas_files else [])
df_compras = pd.DataFrame([enrutador_archivos(f) for f in compras_files] if compras_files else [])

if not df_ventas.empty or not df_compras.empty:
    st.divider()
    
    total_igv_ventas = df_ventas[df_ventas['Estado'] == 'OK']['IGV (18%)'].sum() if 'IGV (18%)' in df_ventas.columns else 0.0
    total_igv_compras = df_compras[df_compras['Estado'] == 'OK']['IGV (18%)'].sum() if 'IGV (18%)' in df_compras.columns else 0.0
    
    m1, m2, m3 = st.columns(3)
    m1.metric("IGV Cobrado (Ventas)", f"S/ {total_igv_ventas:,.2f}")
    m2.metric("Crédito Fiscal (Compras)", f"S/ {total_igv_compras:,.2f}")
    m3.metric("IGV A PAGAR", f"S/ {max(0, total_igv_ventas - total_igv_compras):,.2f}")
    
    st.write("### Detalle de Facturas")
    st.dataframe(pd.concat([df_ventas, df_compras]).fillna(""), use_container_width=True)
    
    st.divider()
    st.subheader("☁️ Guardar en la Nube")
    st.info("Esto creará automáticamente la carpeta de Año/Mes/Día en tu Google Drive y subirá todos los archivos que ves arriba.")
    
    if st.button("🚀 Subir Archivos a Google Drive", type="primary"):
        with st.spinner("Conectando con Google Drive y subiendo archivos..."):
            try:
                drive_service = obtener_servicio_drive()
                hoy = datetime.datetime.now()
                id_anio = obtener_o_crear_carpeta(drive_service, hoy.strftime('%Y'), CARPETA_RAIZ_ID)
                id_mes = obtener_o_crear_carpeta(drive_service, hoy.strftime('%m-%B'), id_anio)
                id_dia = obtener_o_crear_carpeta(drive_service, hoy.strftime('%d-%m-%Y'), id_mes)
                
                todos_los_archivos = (ventas_files or []) + (compras_files or [])
                exitos = 0
                for file in todos_los_archivos:
                    if subir_archivo_drive(file, drive_service, id_dia):
                        exitos += 1
                
                st.success(f"¡Listo! Se guardaron {exitos} de {len(todos_los_archivos)} archivos en Drive.")
            except Exception as e:
                st.error(f"Error de conexión con Drive: Revisa tus credenciales. Detalle: {e}")
