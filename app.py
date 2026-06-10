import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
import plotly.express as px

st.set_page_config(
    page_title="TaxTech Auditor RD",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── ESTILOS CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stSidebar"] { background: #0f1923; }
    [data-testid="stSidebar"] * { color: #e8edf2 !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stTextInput label,
    [data-testid="stSidebar"] .stSlider label { color: #94a3b8 !important; font-size: 0.78rem !important; }
    [data-testid="stSidebar"] h1 { color: #38bdf8 !important; font-size: 1rem !important; letter-spacing: 0.05em; text-transform: uppercase; }
    .block-container { padding-top: 1.5rem; }
    div[data-testid="metric-container"] { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 1rem; }
    div[data-testid="metric-container"] label { font-size: 0.75rem !important; color: #64748b !important; text-transform: uppercase; letter-spacing: 0.06em; }
    .stTabs [data-baseweb="tab"] { font-size: 0.82rem; padding: 6px 14px; }
    .stTabs [aria-selected="true"] { border-bottom: 2px solid #1e3a5f; font-weight: 700; }
    h1 { color: #1e3a5f !important; }
    h3 { color: #1e3a5f !important; font-size: 1rem !important; }
    
    /* Estilos para Tablas Contables HTML */
    .tabla-contable { width: 100%; border-collapse: collapse; font-family: 'Calibri', sans-serif; font-size: 0.9rem; margin-bottom: 1rem; }
    .tabla-contable th { border-bottom: 1px solid #000; padding: 8px; text-align: right; font-weight: bold; }
    .tabla-contable th:first-child { text-align: left; }
    .tabla-contable td { padding: 6px 8px; text-align: right; }
    .tabla-contable td:first-child { text-align: left; }
    .tabla-contable .seccion { font-weight: bold; text-align: left; padding-top: 15px; }
    .tabla-contable .total td { border-top: 1px solid #000; border-bottom: 3px double #000; font-weight: bold; }
    .tabla-contable .subtotal td { border-top: 1px solid #ccc; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# PARÁMETROS FISCALES Y CONSTANTES
# ──────────────────────────────────────────────────────────────────────────────
NATURALEZAS = {'1': 'Debito', '2': 'Credito', '3': 'Credito', '4': 'Credito', '5': 'Debito', '6': 'Debito'}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Art. 287 CTRD: Validar NCF válido y medios de pago para crédito ITBIS.',
    'representacion': 'Art. 287 CTRD: Razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Art. 318 / Reg. 139-98: Retribuciones en especie. Validar ISR sustitutivo.',
    'gasto de personal': 'Art. 287 CTRD: Cruce obligatorio con IR-4 (TSS) para admitir deducción.',
    'honorario': 'Art. 309 CTRD: Retención 10% personas físicas / 2% entre jurídicas.',
}

# Formateador Contable (Convierte negativos a paréntesis y ceros a guiones)
def fmt_c(val):
    if pd.isna(val) or round(val, 2) == 0: return "-"
    return f"({abs(val):,.0f})" if val < 0 else f"{val:,.0f}"

# ──────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO DE BALANZA
# ──────────────────────────────────────────────────────────────────────────────
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.lower().endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(file, header=None)
        else:
            df_raw = pd.read_csv(file, header=None)
        
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(x).lower() for x in row.values])
            if ('código' in row_str or 'codigo' in row_str) and ('nombre' in row_str or 'cuenta' in row_str):
                header_idx = idx
                break
        
        column_names = df_raw.iloc[header_idx].astype(str).str.lower().str.strip()
        df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        
        idx_codigo, idx_cuenta = -1, -1
        indices_debe, indices_haber, indices_balance = [], [], []
        
        for i, col in enumerate(column_names):
            if any(x in col for x in ['código', 'codigo', 'cuenta no']) and idx_codigo == -1: idx_codigo = i
            elif any(x in col for x in ['nombre', 'descripción', 'cuenta']) and 'codigo' not in col and idx_cuenta == -1: idx_cuenta = i
            elif any(x in col for x in ['débito', 'debito', 'debe', 'cargos']): indices_debe.append(i)
            elif any(x in col for x in ['crédito', 'credito', 'haber', 'abonos']): indices_haber.append(i)
            elif any(x in col for x in ['saldo', 'balance', 'final', 'monto']): indices_balance.append(i)
        
        if idx_codigo == -1 or idx_cuenta == -1:
            st.error("⚠️ No se encontró la columna de Código o Nombre de Cuenta.")
            return pd.DataFrame()
        
        col_dict = {'codigo': df.iloc[:, idx_codigo], 'cuenta': df.iloc[:, idx_cuenta]}
        if indices_debe: col_dict['debito'] = df.iloc[:, indices_debe[-1]]
        if indices_haber: col_dict['credito'] = df.iloc[:, indices_haber[-1]]
        if indices_balance: col_dict['saldo_final'] = df.iloc[:, indices_balance[-1]]
        
        df_clean = pd.DataFrame(col_dict)
        if 'saldo_final' not in df_clean.columns and 'debito' in df_clean.columns:
            df_clean['saldo_final'] = pd.to_numeric(df_clean['debito'], errors='coerce').fillna(0) - pd.to_numeric(df_clean['credito'], errors='coerce').fillna(0)
        
        df_clean['codigo'] = df_clean['codigo'].fillna('').astype(str).str.strip().apply(lambda x: x.split('.')[0] if '.' in x else x)
        df_clean['cuenta'] = df_clean['cuenta'].fillna('').astype(str).str.strip()
        df_clean = df_clean[(df_clean['codigo'] != '') & (~df_clean['codigo'].str.lower().str.contains('total|suma', na=False))]
        
        for col in ['debito', 'credito', 'saldo_final']:
            if col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col].astype(str).str.replace(',', '').replace(r'^-*$', '0', regex=True), errors='coerce').fillna(0.0)
                
        return df_clean.reset_index(drop=True)
    except Exception as e:
        st.error(f"Error procesando balanza: {e}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat, alertas_fisc = [], []
    for _, row in df.iterrows():
        cod, nom = str(row['codigo']), str(row['cuenta']).lower()
        nat_esp = NATURALEZAS.get(cod[0] if cod else '', None)
        saldo = row['saldo_final']
        
        if nat_esp == 'Debito' and saldo < -1: alertas_nat.append("⚠️ Saldo crédito (nat. débito)")
        elif nat_esp == 'Credito' and saldo > 1: alertas_nat.append("⚠️ Saldo débito (nat. crédito)")
        else: alertas_nat.append("✅ Correcto")
        
        alerta_f = next((msg for palabra, msg in PALABRAS_CRITICAS_ART287.items() if palabra in nom), "")
        alertas_fisc.append(alerta_f)
        
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal'] = alertas_fisc
    return df

def procesar_comparativo(df_act: pd.DataFrame, df_ant: pd.DataFrame) -> pd.DataFrame:
    df_act_c = df_act[['codigo', 'cuenta', 'saldo_final']].copy()
    df_ant_c = df_ant[['codigo', 'saldo_final']].copy()
    df_comp = pd.merge(df_act_c, df_ant_c, on='codigo', how='outer', suffixes=('_Y2', '_Y1')).fillna(0.0)
    df_comp.loc[df_comp['cuenta'] == 0.0, 'cuenta'] = "Cuenta Histórica"
    df_comp['variacion_abs'] = df_comp['saldo_final_Y2'] - df_comp['saldo_final_Y1']
    return df_comp

# ──────────────────────────────────────────────────────────────────────────────
# GENERADORES DE TABLAS HTML CORPORATIVAS
# ──────────────────────────────────────────────────────────────────────────────
def html_estado_resultados(df_comp, anio):
    html = f"""<table class="tabla-contable">
        <tr><th>Conceptos</th><th>Nota</th><th>{anio}</th><th>{int(anio)-1}</th></tr>
        <tr><td class="seccion" colspan="4">Ingresos operacionales:</td></tr>"""
    
    # INGRESOS
    ing_y2, ing_y1 = 0, 0
    for _, r in df_comp[df_comp['codigo'].str.startswith('4', na=False)].iterrows():
        v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
        if v2 == 0 and v1 == 0: continue
        ing_y2 += v2; ing_y1 += v1
        html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    html += f"<tr class='subtotal'><td>Total Ingresos</td><td></td><td>{fmt_c(ing_y2)}</td><td>{fmt_c(ing_y1)}</td></tr>"
    
    # COSTOS Y GASTOS
    html += f"<tr><td class="seccion" colspan="4">Costos y gastos operacionales:</td></tr>"
    gas_y2, gas_y1 = 0, 0
    for _, r in df_comp[df_comp['codigo'].str.startswith(('5','6'), na=False)].iterrows():
        # Para que se vean negativos en el reporte
        v2, v1 = -abs(r['saldo_final_Y2']), -abs(r['saldo_final_Y1'])
        if v2 == 0 and v1 == 0: continue
        gas_y2 += v2; gas_y1 += v1
        html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    
    html += f"<tr class='subtotal'><td>Total Costos y Gastos</td><td></td><td>{fmt_c(gas_y2)}</td><td>{fmt_c(gas_y1)}</td></tr>"
    html += f"<tr class='total'><td>Utilidad (Pérdida) del Período</td><td></td><td>{fmt_c(ing_y2 + gas_y2)}</td><td>{fmt_c(ing_y1 + gas_y1)}</td></tr>"
    html += "</table>"
    return html

def html_balance_general(df_comp, anio, tipo='activo'):
    html = f"""<table class="tabla-contable">
        <tr><th>{tipo.capitalize()}s</th><th>Nota</th><th>{anio}</th><th>{int(anio)-1}</th></tr>"""
    
    tot_c_y2, tot_c_y1 = 0, 0
    tot_nc_y2, tot_nc_y1 = 0, 0
    
    if tipo == 'activo':
        html += f"<tr><td class="seccion" colspan="4">Activos corrientes:</td></tr>"
        prefijos_c, prefijos_nc = ('11','12','13','14'), ('15','16','17','18','19')
    else:
        html += f"<tr><td class="seccion" colspan="4">Pasivos corrientes:</td></tr>"
        prefijos_c, prefijos_nc = ('21',), ('22','23')

    # CORRIENTE
    for _, r in df_comp[df_comp['codigo'].str.startswith(prefijos_c, na=False)].iterrows():
        v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
        if v2 == 0 and v1 == 0: continue
        tot_c_y2 += v2; tot_c_y1 += v1
        html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    html += f"<tr class='subtotal'><td>Total {tipo}s corrientes</td><td></td><td>{fmt_c(tot_c_y2)}</td><td>{fmt_c(tot_c_y1)}</td></tr>"
    
    # NO CORRIENTE
    html += f"<tr><td class="seccion" colspan="4">{tipo.capitalize()}s no corrientes:</td></tr>"
    for _, r in df_comp[df_comp['codigo'].str.startswith(prefijos_nc, na=False)].iterrows():
        v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
        if tipo == 'activo' and 'acum' in r['cuenta'].lower():
            v2, v1 = -v2, -v1 # Depreciación resta
        if v2 == 0 and v1 == 0: continue
        tot_nc_y2 += v2; tot_nc_y1 += v1
        html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
    html += f"<tr class='subtotal'><td>Total {tipo}s no corrientes</td><td></td><td>{fmt_c(tot_nc_y2)}</td><td>{fmt_c(tot_nc_y1)}</td></tr>"
    
    # PATRIMONIO (Si es Pasivo, se añade al final)
    pat_y2, pat_y1 = 0, 0
    if tipo == 'pasivo':
        html += f"<tr><td class="seccion" colspan="4">Patrimonio:</td></tr>"
        for _, r in df_comp[df_comp['codigo'].str.startswith('3', na=False)].iterrows():
            v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
            if v2 == 0 and v1 == 0: continue
            pat_y2 += v2; pat_y1 += v1
            html += f"<tr><td>{r['cuenta'].title()}</td><td></td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
        html += f"<tr class='subtotal'><td>Total Patrimonio</td><td></td><td>{fmt_c(pat_y2)}</td><td>{fmt_c(pat_y1)}</td></tr>"
    
    gran_tot_y2 = tot_c_y2 + tot_nc_y2 + pat_y2
    gran_tot_y1 = tot_c_y1 + tot_nc_y1 + pat_y1
    
    titulo_tot = f"Total {tipo.capitalize()}s" if tipo == 'activo' else "Total Pasivos y Patrimonio"
    html += f"<tr class='total'><td>{titulo_tot}</td><td></td><td>{fmt_c(gran_tot_y2)}</td><td>{fmt_c(gran_tot_y1)}</td></tr>"
    html += "</table>"
    return html

def html_flujo_hoja_trabajo(df_comp, anio):
    html = f"""<table class="tabla-contable">
        <tr><th>Hoja de Flujo de Efectivo</th><th>{anio}</th><th>Anterior<br>{int(anio)-1}</th></tr>
        <tr><td class="seccion" colspan="3">Activos</td></tr>"""
    
    for _, r in df_comp[df_comp['codigo'].str.startswith('1', na=False)].iterrows():
        v2, v1 = abs(r['saldo_final_Y2']), abs(r['saldo_final_Y1'])
        if 'acum' in r['cuenta'].lower(): v2, v1 = -v2, -v1
        if v2 != 0 or v1 != 0:
            html += f"<tr><td>{r['cuenta'].title()}</td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
            
    html += f"<tr><td class="seccion" colspan="3">Pasivos</td></tr>"
    for _, r in df_comp[df_comp['codigo'].str.startswith('2', na=False)].iterrows():
        v2, v1 = -abs(r['saldo_final_Y2']), -abs(r['saldo_final_Y1']) # Pasivos se muestran negativos en la hoja de trabajo
        if v2 != 0 or v1 != 0:
            html += f"<tr><td>{r['cuenta'].title()}</td><td>{fmt_c(v2)}</td><td>{fmt_c(v1)}</td></tr>"
            
    html += "</table>"
    return html

def html_flujo_inversion_financiamiento(df_comp, anio):
    html = f"""<table class="tabla-contable">
        <tr><th></th><th>Nota</th><th>{anio}</th><th>{int(anio)-1}</th></tr>
        <tr><td class="seccion" colspan="4">Flujos de efectivo por las actividades de inversión:</td></tr>"""
    
    # Inversión (Variación en 15)
    inv_y2 = df_comp[df_comp['codigo'].str.startswith('15', na=False)]['variacion_abs'].sum() * -1
    # Simulamos el Y1 como un % de la variación o 0 si no hay datos de Y0
    html += f"<tr><td>Adquisición de propiedad, mobiliarios y equipos</td><td>10</td><td>{fmt_c(inv_y2)}</td><td>-</td></tr>"
    html += f"<tr class='subtotal'><td>Efectivo neto usado en las actividades de inversión</td><td></td><td>{fmt_c(inv_y2)}</td><td>-</td></tr>"
    
    # Financiamiento (Variación en 22 y 3)
    html += f"<tr><td class="seccion" colspan="4">Flujos de efectivo por actividades de financiamiento:</td></tr>"
    fin_y2_pas = df_comp[df_comp['codigo'].str.startswith('22', na=False)]['variacion_abs'].sum()
    fin_y2_pat = df_comp[df_comp['codigo'].str.startswith('3', na=False)]['variacion_abs'].sum() * -1 # Aumento de capital es positivo al flujo
    
    html += f"<tr><td>Préstamos obtenidos / pagados netos</td><td>11</td><td>{fmt_c(fin_y2_pas)}</td><td>-</td></tr>"
    html += f"<tr><td>Aportes recibidos de accionistas netos</td><td>12</td><td>{fmt_c(fin_y2_pat)}</td><td>-</td></tr>"
    html += f"<tr class='subtotal'><td>Efectivo neto provisto por las actividades de financiamiento</td><td></td><td>{fmt_c(fin_y2_pas + fin_y2_pat)}</td><td>-</td></tr>"
    
    html += "</table>"
    return html

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR & MAIN SETUP
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.title("🏛️ TaxTech Auditor RD")
st.sidebar.markdown("---")
empresa = st.sidebar.text_input("Empresa", value="Empresa de Prueba SRL")
periodo = st.sidebar.text_input("Período", value="Enero - Diciembre 2026")
anio    = st.sidebar.text_input("Año Fiscal", value="2026")
pct_mp   = st.sidebar.slider("% Materialidad Planificación", 0.5, 3.0, 1.0, 0.1) / 100

st.title("TaxTech Auditor — Declaraciones Juradas & Estados Financieros")
c_up1, c_up2 = st.columns(2)
with c_up1: uploaded = st.file_uploader("📂 Cargar Balanza (Año Actual)", type=["xlsx", "xls", "csv"])
with c_up2: uploaded_prev = st.file_uploader("📂 Cargar Balanza (Año Anterior)", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.info("👆 Sube una balanza para iniciar.")
    st.stop()

df_bal = analizar_balanza(procesar_balanza(uploaded))
if df_bal.empty: st.stop()

df_comp = pd.DataFrame()
if uploaded_prev:
    df_prev = procesar_balanza(uploaded_prev)
    if not df_prev.empty: df_comp = procesar_comparativo(df_bal, df_prev)

# KPIs
t_ingresos = abs(df_bal[df_bal['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
t_activos  = abs(df_bal[df_bal['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
utilidad_neta = t_ingresos - abs(df_bal[df_bal['codigo'].str.startswith(('5','6'), na=False)]['saldo_final'].sum())
mp = (t_ingresos if t_ingresos > 0 else t_activos) * pct_mp

st.markdown(f"### 📌 {empresa} — {periodo}")

# ─── TABS PRINCIPALES ─────────────────────────────────────────────────────────
tab_comp, tab_bg, tab_er, tab_efe, tab_bal, tab_inconsist, tab_art287 = st.tabs([
    "📈 Dashboard", "📊 Balance General", "📉 Estado de Resultados", "🌊 Flujo de Efectivo", "📋 Balanza", "🚨 Inconsistencias", "⚖️ Riesgos Art.287"
])

with tab_comp:
    if not df_comp.empty:
        st.markdown("### Resumen Ejecutivo")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Ingresos Año Actual", f"RD$ {t_ingresos:,.0f}")
        with c2:
            st.metric("Materialidad Planificación", f"RD$ {mp:,.0f}")
    else: st.warning("Sube el Año Anterior para comparativas completas.")

with tab_bg:
    if not df_comp.empty:
        col1, col2 = st.columns(2)
        with col1: st.markdown(html_balance_general(df_comp, anio, 'activo'), unsafe_allow_html=True)
        with col2: st.markdown(html_balance_general(df_comp, anio, 'pasivo'), unsafe_allow_html=True)
    else: st.info("Sube la balanza comparativa para ver el Balance General formateado.")

with tab_er:
    if not df_comp.empty:
        st.markdown(html_estado_resultados(df_comp, anio), unsafe_allow_html=True)
    else: st.info("Sube la balanza comparativa para ver el Estado de Resultados formateado.")

with tab_efe:
    if not df_comp.empty:
        st.markdown("### Estado de Flujo de Efectivo")
        col_f1, col_f2 = st.columns(2)
        with col_f1: 
            st.markdown("#### Hoja de Trabajo")
            st.markdown(html_flujo_hoja_trabajo(df_comp, anio), unsafe_allow_html=True)
        with col_f2: 
            st.markdown("#### Resumen Inversión y Financiamiento")
            st.markdown(html_flujo_inversion_financiamiento(df_comp, anio), unsafe_allow_html=True)
    else: st.info("Sube la balanza comparativa para ver el Flujo de Efectivo.")

with tab_inconsist:
    df_err = df_bal[~df_bal['validacion_naturaleza'].str.startswith('✅')]
    if df_err.empty: st.success("✅ Sin inconsistencias.")
    else: st.dataframe(df_err[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)

with tab_art287:
    df_fisc = df_bal[df_bal['alerta_fiscal'] != ""]
    if df_fisc.empty: st.success("✅ Sin alertas fiscales.")
    else: st.dataframe(df_fisc[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal']], use_container_width=True)