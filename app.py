import streamlit as st
import pandas as pd

st.set_page_config(page_title="TaxTech Auditor RD", layout="wide")
st.title("📊 TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal")

with st.sidebar:
    st.header("Configuración del Cliente")
    cliente = st.text_input("Nombre de la Empresa")
    periodo = st.date_input("Período de Análisis")
    st.markdown("---")
    st.header("Parámetros de Materialidad (NIA 320)")
    tipo_empresa = st.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Pérdidas / Sin Fines de Lucro"])
    porcentaje_mat = st.slider("Porcentaje de Materialidad", 0.5, 2.0, 1.0, step=0.1)

st.subheader("1. Carga de Balanza de Comprobación")
archivo_balanza = st.file_uploader("Arrastra tu archivo Excel o CSV generado desde tu software contable (Odoo, QuickBooks, etc.)", type=["xlsx", "csv"])

if archivo_balanza is not None:
    if archivo_balanza.name.endswith('.csv'):
        df = pd.read_csv(archivo_balanza)
    else:
        df = pd.read_excel(archivo_balanza)
    st.success("Balanza cargada exitosamente.")
    st.dataframe(df.head())
    
    st.subheader("2. Análisis de Reglas de Negocio Contable")
    col1, col2 = st.columns(2)
    with col1:
        st.info("### Alertas de Naturaleza y Agrupación DGII")
    with col2:
        st.warning("### Cuentas Materiales y Riesgos Fiscales (Art. 287)")
