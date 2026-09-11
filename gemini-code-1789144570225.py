import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import pdfplumber
import re

st.set_page_config(page_title="Gestor de Facturación - Restaurante", page_icon="🍽️", layout="wide")

def parse_sunat_xml(file_obj):
    """Extrae los datos de un archivo XML estándar UBL 2.1"""
    try:
        tree = ET.parse(file_obj)
        root = tree.getroot()
        
        ns = {
            'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
            'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'
        }
        
        fecha = root.find('.//cbc:IssueDate', ns)
        fecha = fecha.text if fecha is not None else 'N/A'
        
        serie_numero = root.find('.//cbc:ID', ns)
        serie_numero = serie_numero.text if serie_numero is not None else 'N/A'
        
        ruc = root.find('.//cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID', ns)
        ruc = ruc.text if ruc is not None else 'N/A'
        
        razon_social = root.find('.//cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName', ns)
        razon_social = razon_social.text if razon_social is not None else 'N/A'
        
        total_node = root.find('.//cac:LegalMonetaryTotal/cbc:PayableAmount', ns)
        total = float(total_node.text) if total_node is not None else 0.0
        
        igv_node = root.find('.//cac:TaxTotal/cac:TaxSubtotal/cbc:TaxAmount', ns)
        igv = float(igv_node.text) if igv_node is not None else 0.0
        
        base_imponible = total - igv
        
        return {
            "Archivo": file_obj.name,
            "Tipo": "XML",
            "Fecha": fecha,
            "Comprobante": serie_numero,
            "RUC": ruc,
            "Razón Social": razon_social,
            "Base Imponible": round(base_imponible, 2),
            "IGV (18%)": round(igv, 2),
            "Total": round(total, 2),
            "Estado": "OK"
        }
    except Exception as e:
        return {"Archivo": file_obj.name, "Tipo": "XML", "Estado": f"Error: {e}"}

def procesar_factura_pdf(file_obj):
    """Extrae texto de un PDF mediante expresiones regulares"""
    try:
        texto_completo = ""
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                texto_extraido = page.extract_text()
                if texto_extraido:
                    texto_completo += texto_extraido + "\n"
        
        ruc_match = re.search(r'\b(10|20)\d{9}\b', texto_completo)
        ruc = ruc_match.group(0) if ruc_match else "No detectado"
        
        serie_match = re.search(r'\b[F|E|B][A-Z0-9]{3}-\d{1,8}\b', texto_completo)
        serie = serie_match.group(0) if serie_match else "No detectado"
        
        # Búsqueda de montos (patrón flexible para diferentes formatos de PDF)
        total_match = re.search(r'(?:TOTAL|Total|Importe Total).*?(?:S/|S/\.)?\s*([\d,]+\.\d{2})', texto_completo)
        total = 0.0
        if total_match:
            numero_limpio = total_match.group(1).replace(',', '')
            total = float(numero_limpio)
        
        # Estimación matemática asumiendo que todo el monto está gravado
        base_imponible = total / 1.18 if total > 0 else 0.0
        igv = total - base_imponible if total > 0 else 0.0
        
        return {
            "Archivo": file_obj.name,
            "Tipo": "PDF",
            "Fecha": "En PDF", # Requiere regex adicional según formato
            "Comprobante": serie,
            "RUC": ruc,
            "Razón Social": "En PDF", # Difícil de extraer con regex genérico
            "Base Imponible": round(base_imponible, 2),
            "IGV (18%)": round(igv, 2),
            "Total": round(total, 2),
            "Estado": "OK" if total > 0 else "Revisar Manualmente"
        }
    except Exception as e:
        return {"Archivo": file_obj.name, "Tipo": "PDF", "Estado": f"Error: {e}"}

def enrutador_archivos(file_obj):
    """Deriva el archivo a la función correcta según su extensión"""
    if file_obj.name.lower().endswith('.xml'):
        return parse_sunat_xml(file_obj)
    elif file_obj.name.lower().endswith('.pdf'):
        return procesar_factura_pdf(file_obj)
    else:
        return {"Archivo": file_obj.name, "Estado": "Formato no soportado"}

# --- INTERFAZ DE USUARIO ---
st.title("🍽️ Gestor de Facturación y Cierre de Impuestos")
st.markdown("Sube los archivos (XML o PDF) de las facturas electrónicas diarias.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📤 Ventas (Facturas Emitidas)")
    ventas_files = st.file_uploader("Arrastra los XML/PDF de Ventas aquí", type=['xml', 'pdf'], accept_multiple_files=True, key="ventas")

with col2:
    st.subheader("📥 Compras (Gastos e Insumos)")
    compras_files = st.file_uploader("Arrastra los XML/PDF de Compras aquí", type=['xml', 'pdf'], accept_multiple_files=True, key="compras")

# Procesar los archivos subidos en memoria
ventas_data = [enrutador_archivos(f) for f in ventas_files] if ventas_files else []
compras_data = [enrutador_archivos(f) for f in compras_files] if compras_files else []

df_ventas = pd.DataFrame(ventas_data)
df_compras = pd.DataFrame(compras_data)

st.divider()

# --- DASHBOARD Y RESULTADOS ---
if not df_ventas.empty or not df_compras.empty:
    st.header("📊 Resumen del Día / Periodo")
    
    # Manejar posibles errores donde las columnas de montos no existan aún
    if 'IGV (18%)' in df_ventas.columns:
        total_igv_ventas = df_ventas[df_ventas['Estado'] == 'OK']['IGV (18%)'].sum()
        base_ventas = df_ventas[df_ventas['Estado'] == 'OK']['Base Imponible'].sum()
    else:
        total_igv_ventas = 0.0
        base_ventas = 0.0
        
    if 'IGV (18%)' in df_compras.columns:
        total_igv_compras = df_compras[df_compras['Estado'] == 'OK']['IGV (18%)'].sum()
    else:
        total_igv_compras = 0.0

    igv_a_pagar = max(0, total_igv_ventas - total_igv_compras)
    renta_estimada = base_ventas * 0.015  # Asumiendo 1.5% de Renta MYPE
    
    # Tarjetas de resumen
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Ingresos Totales (Base)", f"S/ {base_ventas:,.2f}")
    m2.metric("IGV Cobrado (Ventas)", f"S/ {total_igv_ventas:,.2f}")
    m3.metric("Crédito Fiscal (Compras)", f"S/ {total_igv_compras:,.2f}")
    m4.metric("💰 TOTAL A PAGAR (IGV + Renta)", f"S/ {(igv_a_pagar + renta_estimada):,.2f}")
    
    tab1, tab2 = st.tabs(["Detalle de Ventas", "Detalle de Compras"])
    with tab1:
        if not df_ventas.empty:
            st.dataframe(df_ventas, use_container_width=True)
        else:
            st.write("No hay ventas registradas.")
    with tab2:
        if not df_compras.empty:
            st.dataframe(df_compras, use_container_width=True)
        else:
            st.write("No hay compras registradas.")
            
    st.caption("Nota: La extracción desde PDF utiliza patrones matemáticos. Si un proveedor incluyó productos exonerados de IGV (ej. verduras frescas), el monto extraído del PDF podría requerir ajuste manual.")
else:
    st.info("Sube algunas facturas XML o PDF para ver el resumen automático.")