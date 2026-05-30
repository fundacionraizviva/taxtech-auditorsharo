import streamlit as st
import pandas as pd
import numpy as np

# Configuración inicial de la página de Streamlit
st.set_page_config(
    page_title="TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal",
    layout="wide"
)

# ==========================================
# 1. PARÁMETROS FISCALES Y DE CONFIGURACIÓN (CTRD)
# ==========================================
# Identificación de naturaleza por primer dígito del catálogo dominicano
NATURALEZAS = {
    '1': 'Debito',  # Activos
    '2': 'Credito', # Pasivos
    '3': 'Credito', # Capital
    '4': 'Credito', # Ingresos
    '5': 'Debito',  # Costos
    '6': 'Debito'   # Gastos
}

# Alertas automáticas según el marco normativo de República Dominicana
PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Riesgo Art. 287 CTRD: Validar deducibilidad, comprobantes con NCF válido y uso de medios de pago para crédito fiscal de ITBIS.',
    'representacion': 'Riesgo Art. 287 CTRD: Gastos de representación. Sujetos a criterios de razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Riesgo Art. 318 CTRD / Reg. 139-98: Retribuciones en especie. Validar que la empresa efectúe el pago del ISR sustitutivo correspondiente.',
    'gasto de personal': 'Riesgo Art. 287 CTRD: Cruce obligatorio con la declaración jurada de TSS (Formulario IR-4) para admitir la deducción.',
    'honorario': 'Riesgo Art. 309 CTRD: Validar aplicación de retenciones fiscales (10% a personas físicas o 2% entre personas jurídicas).'
}

# ==========================================
# 2. FUNCIONES DE LÓGICA DE NEGOCIO Y AUDITORÍA
# ==========================================
def procesar_balanza(file) -> pd.DataFrame:
    """Lee y normaliza la balanza de comprobación previniendo caídas del sistema."""
    try:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
        else:
            df = pd.read_csv(file)
        
        # Homologar nombres de columnas a minúsculas y quitar espacios vacíos
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        # Mapeo de sinónimos comunes para dar flexibilidad al usuario
        mapeo_columnas = {
            'código': 'codigo', 'cuenta': 'cuenta', 'nombre': 'cuenta', 'nombre de cuenta': 'cuenta',
            'débito': 'debito', 'crédito': 'credito', 'saldo final': 'saldo_final', 'saldo': 'saldo_final'
        }
        df = df.rename(columns=mapeo_columnas)
        
        # Validación de campos mínimos obligatorios para auditoría
        columnas_requeridas = {'codigo', 'cuenta', 'debito', 'credito', 'saldo_final'}
        if not columnas_requeridas.issubset(set(df.columns)):
            st.error(f"⚠️ Estructura de archivo incorrecta. Debe incluir las columnas: {list(columnas_requeridas)}")
            return pd.DataFrame()
            
        # Limpieza de valores no numéricos en columnas de importes
        for col in ['debito', 'credito', 'saldo_final']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
        df['codigo'] = df['codigo'].astype(str).str.strip()
        df['cuenta'] = df['cuenta'].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"❌ Error crítico en el procesamiento del archivo: {str(e)}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica algoritmos de auditoría contable y fiscal cuenta por cuenta."""
    alertas_nat = []
    alertas_fisc = []
    
    for _, row in df.iterrows():
        # A) Control de Calidad: Naturaleza del Saldo
        primer_digito = row['codigo'][0]
        nat_esperada = NATURALEZAS.get(primer_digito, None)
        saldo = row['saldo_final']
        
        if nat_esperada == 'Debito' and saldo < 0:
            alertas_nat.append("Saldo Crédito inusual (Naturaleza Débito)")
        elif nat_esperada == 'Credito' and saldo < 0:
            alertas_nat.append("Saldo Débito inusual (Naturaleza Crédito)")
        else:
            alertas_nat.append("Correcto")
            
        # B) Control de Riesgo Fiscal: Art. 287 CTRD
        nombre_cuenta = row['cuenta'].lower()
        alerta_f = "Sin observaciones"
        for palabra, mensaje in PALABRAS_CRITICAS_ART287.items():
            if palabra in nombre_cuenta:
                alerta_f = mensaje
                break
        alertas_fisc.append(alerta_f)
        
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal_rd'] = alertas_fisc
    return df

# ==========================================
# 3. INTERFAZ DE USUARIO (INTEGRACIÓN COMPLETA UI)
# ==========================================

# --- BARRA LATERAL (Preservando los elementos de tu captura de pantalla) ---
st.sidebar.title("Configuración del Cliente")
empresa = st.sidebar.text_input("Nombre de la Empresa", value="Empresa de Prueba SRL")
periodo = st.sidebar.text_input("Período de Análisis", value="2026/05/30")

st.sidebar.markdown("---")
st.sidebar.title("Parámetros de Materialidad (NIA 320)")
tipo_entidad = st.sidebar.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Zonas Francas", "Financieras"])

# Ajuste automático del benchmark sugerido por la NIA 320 según el sector
tasa_referencia = 0.01 if tipo_entidad == "Comercial / Servicios" else 0.005

porcentaje_mp = st.sidebar.slider(
    "Porcentaje de Materialidad", 
    min_value=0.5, 
    max_value=3.0, 
    value=tasa_referencia * 100, 
    step=0.1
) / 100

porcentaje_me = st.sidebar.slider(
    "Porcentaje de Materialidad de Ejecución (ME)", 
    min_value=50, 
    max_value=75, 
    value=75, 
    step=5
) / 100


# --- CUERPO PRINCIPAL ---
st.title("📊 TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal")
st.header("1. Carga de Balanza de Comprobación")
st.markdown("Arrastra tu archivo Excel o CSV generado desde tu software contable (Odoo, QuickBooks, etc.)")

# Cargador de archivos con visibilidad oculta de etiqueta para coincidir exactamente con tu diseño
uploaded_file = st.file_uploader("Upload", type=["xlsx", "csv"], label_visibility="collapsed")

# Lógica condicional de ejecución al recibir los datos
if uploaded_file is not None:
    df_balanza = procesar_balanza(uploaded_file)
    
    if not df_balanza.empty:
        # Ejecución del motor analítico
        df_balanza = analizar_balanza(df_balanza)
        
        # Identificación automática de cuentas control para materialidad NIA 320
        total_activos = df_balanza[df_balanza['codigo'].str.startswith('1')]['saldo_final'].sum()
        total_ingresos = df_balanza[df_balanza['codigo'].str.startswith('4')]['saldo_final'].sum()
        
        # Definición automática del benchmark principal de auditoría
        base_calculo = total_ingresos if total_ingresos > 0 else total_activos
        mp = base_calculo * porcentaje_mp
        me = mp * porcentaje_me
        
        st.markdown("---")
        st.subheader(f"📌 Informe de Auditoría Analítica: {empresa} — Período: {periodo}")
        
        # Módulo de Visualización KPI Dashboards
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Ingresos
