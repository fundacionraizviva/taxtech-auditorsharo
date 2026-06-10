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
    .fin-table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    .fin-table th { background: #1e3a5f; color: white; padding: 8px 12px; text-align: left; }
    .fin-table td { padding: 6px 12px; border-bottom: 1px solid #e2e8f0; }
    .fin-table tr:nth-child(even) { background: #f8fafc; }
    .fin-table .subtotal td { background: #dbeafe; font-weight: 600; }
    .fin-table .total td { background: #1e3a5f; color: white; font-weight: 700; }
    .fin-table .section-header td { background: #f1f5f9; font-weight: 700; color: #1e3a5f; letter-spacing: 0.04em; text-transform: uppercase; font-size: 0.78rem; }
    .fin-table .negative { color: #dc2626; }
    .fin-table .positive { color: #16a34a; }
    .stTabs [data-baseweb="tab"] { font-size: 0.82rem; padding: 6px 14px; }
    .stTabs [aria-selected="true"] { border-bottom: 2px solid #1e3a5f; font-weight: 700; }
    h1 { color: #1e3a5f !important; }
    h3 { color: #1e3a5f !important; font-size: 1rem !important; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# PARÁMETROS FISCALES RD 2026
# ──────────────────────────────────────────────────────────────────────────────
NATURALEZAS = {'1': 'Debito', '2': 'Credito', '3': 'Credito', '4': 'Credito', '5': 'Debito', '6': 'Debito'}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Art. 287 CTRD: Validar NCF válido y medios de pago para crédito ITBIS.',
    'representacion': 'Art. 287 CTRD: Razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Art. 318 / Reg. 139-98: Retribuciones en especie. Validar ISR sustitutivo.',
    'gasto de personal': 'Art. 287 CTRD: Cruce obligatorio con IR-4 (TSS) para admitir deducción.',
    'honorario': 'Art. 309 CTRD: Retención 10% personas físicas / 2% entre jurídicas.',
    'viaje': 'Art. 287 CTRD: Gastos de viaje. Documentación de viáticos y propósito del viaje.',
    'donacion': 'Art. 287 CTRD: Donaciones solo deducibles si beneficiario es Ley 122-05.',
    'multa': 'Art. 287 CTRD: Multas y recargos NO son gastos deducibles.',
    'perdida': 'Art. 287 CTRD: Validar naturaleza de la pérdida para deducibilidad.',
}

TASA_ITBIS = 0.18
TASA_SFS_PAT = 0.0709; TASA_AFP_PAT = 0.0710
TASA_SRL = 0.0120; TASA_INFOTEP = 0.0100
TASA_SFS_EMP = 0.0304; TASA_AFP_EMP = 0.0287
COSTO_PERCAPITA_2026 = 1691.38

# Excel styles
FILL_HDR   = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FILL_SEC   = PatternFill(start_color="EBF3FA", end_color="EBF3FA", fill_type="solid")
FNT_TITLE  = Font(name="Calibri", size=13, bold=True, color="1F497D")
FNT_HDR    = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FNT_BODY   = Font(name="Calibri", size=10)
FNT_BOLD   = Font(name="Calibri", size=10, bold=True)
FNT_SEC    = Font(name="Calibri", size=10, bold=True, color="1F497D")
THIN       = Side(border_style="thin", color="D9D9D9")
BRD        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_RD     = 'RD$ #,##0.00'
FMT_PCT    = '0.00%'

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS EXCEL & CÁLCULOS ESPECÍFICOS
# ──────────────────────────────────────────────────────────────────────────────
def xl_header(ws, title, subtitle, cols):
    ws["B2"] = title;   ws["B2"].font = FNT_TITLE
    ws["B3"] = subtitle; ws["B3"].font = Font(name="Calibri", size=10, italic=True)
    for i, h in enumerate(cols, 2):
        c = ws.cell(row=5, column=i, value=h)
        c.font = FNT_HDR; c.fill = FILL_HDR; c.border = BRD
        c.alignment = Alignment(horizontal="center")

def xl_col_widths(ws):
    for col in ws.columns:
        ltr = get_column_letter(col[0].column)
        if ltr != 'A':
            ws.column_dimensions[ltr].width = 45 if ltr in ['C', 'D'] else 22

def xl_money(cell, value):
    cell.value = value; cell.number_format = FMT_RD
    cell.alignment = Alignment(horizontal="right")
    cell.font = FNT_BODY; cell.border = BRD

def generar_nota_activos_fijos(df_comp):
    """Genera la estructura de la nota de PPE separando costo y depreciación."""
    categorias = {
        '151': 'Terrenos', '152': 'Edificaciones', '153': 'Mejoras en Prop.',
        '154': 'Maquinaria y Eq.', '155': 'Eq. de Transporte', '156': 'Mobiliarios',
        '157': 'Eq. Computación', '158': 'Construc. en Proceso'
    }
    data = []
    for cod, cat in categorias.items():
        mask_c = df_comp['codigo'].str.startswith(cod) & ~df_comp['cuenta'].str.lower().str.contains('acum', na=False)
        mask_d = df_comp['codigo'].str.startswith(cod) & df_comp['cuenta'].str.lower().str.contains('acum', na=False)
        
        c_ini = df_comp.loc[mask_c, 'saldo_final_Y1'].sum()
        c_fin = df_comp.loc[mask_c, 'saldo_final_Y2'].sum()
        d_ini = df_comp.loc[mask_d, 'saldo_final_Y1'].sum()
        d_fin = df_comp.loc[mask_d, 'saldo_final_Y2'].sum()
        
        d_ini = -abs(d_ini) if d_ini != 0 else 0
        d_fin = -abs(d_fin) if d_fin != 0 else 0
        
        if c_ini == 0 and c_fin == 0: continue
        
        data.append({
            'Categoría': cat,
            'Costo Inicial': c_ini,
            'Adiciones/Retiros': c_fin - c_ini,
            'Costo Final': c_fin,
            'Depr. Inicial': d_ini,
            'Gasto/Retiros Depr.': d_fin - d_ini,
            'Depr. Final': d_fin,
            'Balance Neto': c_fin + d_fin
        })
    return pd.DataFrame(data)

def generar_flujo_efectivo(df_comp, utilidad_neta):
    """Método Indirecto aproximado usando variaciones de balanza."""
    movimientos = [{'Concepto': 'Utilidad (Pérdida) Neta', 'Monto': utilidad_neta}]
    
    mask_d = df_comp['codigo'].str.startswith('15') & df_comp['cuenta'].str.lower().str.contains('acum', na=False)
    depr = abs(df_comp.loc[mask_d, 'variacion_abs'].sum())
    movimientos.append({'Concepto': '(+) Depreciación y Amortización', 'Monto': depr})
    
    var_cxc = df_comp.loc[df_comp['codigo'].str.startswith('12'), 'variacion_abs'].sum() * -1
    movimientos.append({'Concepto': 'Variación en Cuentas por Cobrar', 'Monto': var_cxc})
    
    var_inv = df_comp.loc[df_comp['codigo'].str.startswith('13'), 'variacion_abs'].sum() * -1
    movimientos.append({'Concepto': 'Variación en Inventarios', 'Monto': var_inv})
    
    var_cxp = df_comp.loc[df_comp['codigo'].str.startswith('21'), 'variacion_abs'].sum()
    movimientos.append({'Concepto': 'Variación en Cuentas por Pagar', 'Monto': var_cxp})
    
    return pd.DataFrame(movimientos)

def generar_cambios_patrimonio(df_comp):
    mask = df_comp['codigo'].str.startswith('3')
    df_pat = df_comp[mask][['codigo', 'cuenta', 'saldo_final_Y1', 'variacion_abs', 'saldo_final_Y2']].copy()
    df_pat['saldo_final_Y1'] = df_pat['saldo_final_Y1'] * -1
    df_pat['variacion_abs'] = df_pat['variacion_abs'] * -1
    df_pat['saldo_final_Y2'] = df_pat['saldo_final_Y2'] * -1
    df_pat.columns = ['Código', 'Cuenta', 'Saldo Inicial', 'Variación', 'Saldo Final']
    return df_pat

# ──────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO DE BALANZA (SOPORTE MULTI-NIVEL Y COLUMNAS DUPLICADAS)
# ──────────────────────────────────────────────────────────────────────────────
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.lower().endswith('.xlsx'):
            df_raw = pd.read_excel(file, engine='openpyxl', header=None)
        elif file.name.lower().endswith('.xls'):
            df_raw = pd.read_excel(file, engine='xlrd', header=None)
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
            if any(x in col for x in ['código', 'codigo', 'cuenta no', 'no.', 'cuenta_no', 'cta']) and idx_codigo == -1:
                idx_codigo = i
            elif any(x in col for x in ['nombre', 'descripción', 'descripcion', 'concepto']) and 'codigo' not in col and idx_cuenta == -1:
                idx_cuenta = i
            elif 'cuenta' in col and 'codigo' not in col and idx_cuenta == -1:
                idx_cuenta = i
            elif any(x in col for x in ['débito', 'debito', 'debe', 'cargos']):
                indices_debe.append(i)
            elif any(x in col for x in ['crédito', 'credito', 'haber', 'abonos']):
                indices_haber.append(i)
            elif any(x in col for x in ['saldo', 'balance', 'final', 'monto']):
                indices_balance.append(i)
        
        if idx_codigo == -1 or idx_cuenta == -1:
            st.error("⚠️ No se encontró la columna de Código o Nombre de Cuenta en el archivo.")
            return pd.DataFrame()
        
        col_dict = {
            'codigo': df.iloc[:, idx_codigo],
            'cuenta': df.iloc[:, idx_cuenta]
        }
        if indices_debe: col_dict['debito'] = df.iloc[:, indices_debe[-1]]
        if indices_haber: col_dict['credito'] = df.iloc[:, indices_haber[-1]]
        if indices_balance: col_dict['saldo_final'] = df.iloc[:, indices_balance[-1]]
        
        df_clean = pd.DataFrame(col_dict)
        
        if 'saldo_final' not in df_clean.columns and 'debito' in df_clean.columns and 'credito' in df_clean.columns:
            df_clean['saldo_final'] = pd.to_numeric(df_clean['debito'], errors='coerce').fillna(0) - pd.to_numeric(df_clean['credito'], errors='coerce').fillna(0)
        
        df_clean['codigo'] = df_clean['codigo'].fillna('').astype(str).str.strip()
        df_clean['codigo'] = df_clean['codigo'].apply(lambda x: x.split('.')[0] if '.' in x else x)
        df_clean['cuenta'] = df_clean['cuenta'].fillna('').astype(str).str.strip()
        
        df_clean = df_clean[(df_clean['codigo'] != '') & (df_clean['codigo'] != 'nan') & (~df_clean['codigo'].str.lower().str.contains('total|resultado|suma', na=False))]
        
        for col in ['debito', 'credito', 'saldo_final']:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].astype(str).str.strip().str.replace(',', '').replace(r'^-*$', '0', regex=True)
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0.0)
            else:
                df_clean[col] = 0.0
                
        return df_clean.reset_index(drop=True)
        
    except Exception as e:
        st.error(f"❌ Error al procesar la balanza analítica: {e}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat, alertas_fisc = [], []
    for _, row in df.iterrows():
        cod, nom = str(row['codigo']), str(row['cuenta']).lower()
        nat_esp = NATURALEZAS.get(cod[0] if cod else '', None)
        saldo = row['saldo_final']
        if nat_esp == 'Debito' and saldo < 0:
            alertas_nat.append("⚠️ Saldo crédito (nat. débito)")
        elif nat_esp == 'Credito' and saldo > 0:
            alertas_nat.append("⚠️ Saldo débito (nat. crédito)")
        else:
            alertas_nat.append("✅ Correcto")
        alerta_f = ""
        for palabra, msg in PALABRAS_CRITICAS_ART287.items():
            if palabra in nom:
                alerta_f = msg; break
        alertas_fisc.append(alerta_f)
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal'] = alertas_fisc
    return df

# ──────────────────────────────────────────────────────────────────────────────
# LÓGICA COMPARATIVA DE DOS AÑOS (EEFF CORPORATIVOS)
# ──────────────────────────────────────────────────────────────────────────────
def procesar_comparativo(df_act: pd.DataFrame, df_ant: pd.DataFrame) -> pd.DataFrame:
    try:
        df_act_clean = df_act[['codigo', 'cuenta', 'saldo_final']].copy()
        df_ant_clean = df_ant[['codigo', 'saldo_final']].copy()
        
        df_comp = pd.merge(df_act_clean, df_ant_clean, on='codigo', how='outer', suffixes=('_Y2', '_Y1'))
        
        if 'cuenta' not in df_comp.columns:
             df_comp['cuenta'] = "Cuenta Desconocida"
             
        df_comp['saldo_final_Y2'] = df_comp['saldo_final_Y2'].fillna(0.0)
        df_comp['saldo_final_Y1'] = df_comp['saldo_final_Y1'].fillna(0.0)
        
        df_comp['variacion_abs'] = df_comp['saldo_final_Y2'] - df_comp['saldo_final_Y1']
        df_comp['variacion_pct'] = np.where(df_comp['saldo_final_Y1'] != 0, 
                                            (df_comp['variacion_abs'] / df_comp['saldo_final_Y1'].replace(0, np.nan)), 0)
        
        return df_comp
    except Exception as e:
        st.error(f"❌ Error al procesar datos comparativos: {e}")
        return pd.DataFrame()

def exportar_reporte_corporativo(empresa, periodo, anio, df_comp):
    try:
        wb = openpyxl.Workbook()
        ws_dash = wb.active; ws_dash.title = "Dashboard Corporativo"
        ws_esf = wb.create_sheet("Estado de Situación")
        ws_er = wb.create_sheet("Estado de Resultados")
        ws_ecp = wb.create_sheet("Cambios en Patrimonio")
        ws_efe = wb.create_sheet("Flujo de Efectivo")

        def escribir_filas(ws, titulo, data_filas, row_start):
            xl_header(ws, f"{titulo} — {empresa.upper()}", f"Comparativo {anio} vs {int(anio)-1}", 
                      ["Código", "Cuenta", f"Año {anio}", f"Año {int(anio)-1}", "Variación RD$", "Variación %"])
            r = row_start
            for cod, cuenta, m2, m1, v_abs, v_pct in data_filas:
                ws.cell(row=r, column=2, value=cod).font = FNT_BODY
                ws.cell(row=r, column=3, value=cuenta).font = FNT_BODY
                xl_money(ws.cell(row=r, column=4), abs(m2))
                xl_money(ws.cell(row=r, column=5), abs(m1))
                xl_money(ws.cell(row=r, column=6), v_abs)
                c_pct = ws.cell(row=r, column=7, value=v_pct)
                c_pct.number_format = FMT_PCT; c_pct.border = BRD; c_pct.alignment = Alignment(horizontal="right")
                if r % 2 == 0:
                    for c_idx in range(2, 8): ws.cell(row=r, column=c_idx).fill = FILL_ZEBRA
                r += 1
            xl_col_widths(ws)
            ws.column_dimensions['D'].width = 20
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 20
            return r

        esf_data, er_data, ecp_data = [], [], []
        
        for _, row in df_comp.iterrows():
            cod = str(row['codigo'])
            if not cod: continue
            fila = (cod, row['cuenta'], row['saldo_final_Y2'], row['saldo_final_Y1'], row['variacion_abs'], row['variacion_pct'])
            prefijo = cod[0]
            if prefijo in ['1', '2', '3']:
                esf_data.append(fila)
                if prefijo == '3': ecp_data.append(fila)
            elif prefijo in ['4', '5', '6']:
                er_data.append(fila)

        esf_data = sorted(esf_data, key=lambda x: x[0])
        escribir_filas(ws_esf, "ESTADO DE SITUACIÓN FINANCIERA", esf_data, 6)
        
        er_data = sorted(er_data, key=lambda x: x[0])
        escribir_filas(ws_er, "ESTADO DE RESULTADOS", er_data, 6)
        
        ecp_data = sorted(ecp_data, key=lambda x: x[0])
        escribir_filas(ws_ecp, "ESTADO DE CAMBIOS EN EL PATRIMONIO", ecp_data, 6)

        xl_header(ws_efe, f"ESTADO DE FLUJO DE EFECTIVO (Borrador) — {empresa.upper()}", 
                  f"Período {anio} (Método Indirecto aproximado)", ["Concepto", "", "Monto RD$"])
        
        utilidad_neta = sum(r[2] for r in er_data if r[0].startswith('4')) - sum(abs(r[2]) for r in er_data if r[0].startswith(('5', '6')))
        var_act_op = sum(r[4] for r in esf_data if r[0].startswith('1') and not r[0].startswith('11')) * -1
        var_pas_op = sum(r[4] for r in esf_data if r[0].startswith('21'))
        flujo_operativo = utilidad_neta + var_act_op + var_pas_op
        var_inv = sum(r[4] for r in esf_data if r[0].startswith('15') or r[0].startswith('16')) * -1
        var_fin = sum(r[4] for r in esf_data if r[0].startswith('22') or r[0].startswith('3') and not 'resultado' in r[1].lower())
        
        efe_filas = [
            ("Actividades de Operación", ""),
            ("Utilidad Neta del Ejercicio", utilidad_neta),
            ("Variación Neta en Activos y Pasivos Operativos", var_act_op + var_pas_op),
            ("Efectivo Neto Provisto por Operaciones", flujo_operativo),
            ("Actividades de Inversión", ""),
            ("Variación Neta en Activos Fijos e Inversiones", var_inv),
            ("Actividades de Financiamiento", ""),
            ("Variación Neta en Deuda a Largo Plazo y Patrimonio", var_fin),
            ("INCREMENTO (DISMINUCIÓN) NETO DE EFECTIVO", flujo_operativo + var_inv + var_fin)
        ]
        
        r = 6
        for concepto, monto in efe_filas:
            ws_efe.cell(row=r, column=2, value=concepto).font = FNT_BOLD if monto == "" else FNT_BODY
            if monto != "": xl_money(ws_efe.cell(row=r, column=4), monto)
            r += 1
        xl_col_widths(ws_efe)

        ws_dash.sheet_view.showGridLines = False
        ws_dash["B2"] = f"DASHBOARD FINANCIERO — {empresa.upper()}"; ws_dash["B2"].font = Font(name="Calibri", size=18, bold=True, color="1F497D")
        
        ingresos_Y2 = sum(abs(r[2]) for r in er_data if r[0].startswith('4'))
        ingresos_Y1 = sum(abs(r[3]) for r in er_data if r[0].startswith('4'))
        activos_Y2  = sum(abs(r[2]) for r in esf_data if r[0].startswith('1'))
        activos_Y1  = sum(abs(r[3]) for r in esf_data if r[0].startswith('1'))
        
        ws_dash["B5"] = "Métrica"; ws_dash["C5"] = f"Año {anio}"; ws_dash["D5"] = f"Año {int(anio)-1}"
        kpis = [("Ingresos", ingresos_Y2, ingresos_Y1), ("Activos Totales", activos_Y2, activos_Y1)]
        
        r = 6
        for met, v2, v1 in kpis:
            ws_dash.cell(row=r, column=2, value=met)
            ws_dash.cell(row=r, column=3, value=v2)
            ws_dash.cell(row=r, column=4, value=v1)
            r += 1

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Crecimiento de Ingresos y Activos"
        data = Reference(ws_dash, min_col=3, min_row=5, max_col=4, max_row=7)
        cats = Reference(ws_dash, min_col=2, min_row=6, max_row=7)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 16
        chart.height = 8
        ws_dash.add_chart(chart, "F5")

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    except Exception as e:
        st.error(f"❌ Error al generar Excel Corporativo: {e}")
        return None

# ──────────────────────────────────────────────────────────────────────────────
# ESTADO FINANCIERO DESDE BALANZA (INDIVIDUAL)
# ──────────────────────────────────────────────────────────────────────────────
def fmt_rd(v): return f"RD$ {v:>14,.2f}"
def fmt_neg(v): return f"(RD$ {abs(v):>13,.2f})" if v < 0 else fmt_rd(v)

def generar_balance_general(df: pd.DataFrame) -> dict:
    result = {'activo_corriente': [], 'activo_no_corriente': [], 'pasivo_corriente': [], 'pasivo_no_corriente': [], 'patrimonio': []}
    for _, row in df.iterrows():
        cod = str(row['codigo'])
        if not cod or len(cod) == 0: continue
        saldo = abs(row['saldo_final']) if row['saldo_final'] != 0 else 0
        nombre = row['cuenta']
        entry = (nombre, cod, saldo)
        p = cod[0]
        if p == '1':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('1', '2', '3', '4'): result['activo_corriente'].append(entry)
            else: result['activo_no_corriente'].append(entry)
        elif p == '2':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('1', '2', '3'): result['pasivo_corriente'].append(entry)
            else: result['pasivo_no_corriente'].append(entry)
        elif p == '3':
            result['patrimonio'].append(entry)
    return result

def generar_estado_resultados(df: pd.DataFrame) -> dict:
    result = {'ingresos': [], 'costos': [], 'gastos_operacion': [], 'otros_ingresos': [], 'gastos_financieros': []}
    for _, row in df.iterrows():
        cod = str(row['codigo'])
        if not cod: continue
        p = cod[0]
        saldo = abs(row['saldo_final'])
        nombre = row['cuenta']
        entry = (nombre, cod, saldo)
        if p == '4':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('1', '2', '3', '4', '5'): result['ingresos'].append(entry)
            else: result['otros_ingresos'].append(entry)
        elif p == '5': result['costos'].append(entry)
        elif p == '6':
            sub = cod[1] if len(cod) > 1 else '0'
            if sub in ('3', '4'): result['gastos_financieros'].append(entry)
            else: result['gastos_operacion'].append(entry)
    return result

# ──────────────────────────────────────────────────────────────────────────────
# LLENADO IR-2 (Ajuste Patrimonial DGII)
# ──────────────────────────────────────────────────────────────────────────────
def calcular_casillas_ir2(df: pd.DataFrame) -> dict:
    def suma(prefijos):
        mask = df['codigo'].apply(lambda x: any(str(x).startswith(p) for p in prefijos))
        return abs(df.loc[mask, 'saldo_final'].sum())

    total_activos    = suma(['1'])
    total_pasivos    = suma(['2'])
    patrimonio       = suma(['3'])
    inventario       = suma(['13', '130', '131', '132', '133'])
    activo_fijo_cat1 = suma(['152', '153'])
    activo_fijo_cat2 = suma(['154', '155'])
    activo_fijo_cat3 = suma(['156', '157', '158'])
    otros_activos    = suma(['14', '16', '17', '18', '19'])

    saldo_act_fiscal = activo_fijo_cat1 + activo_fijo_cat2 + activo_fijo_cat3 + inventario + otros_activos
    saldo_pasivos    = total_pasivos
    patrimonio_fisc  = saldo_act_fiscal - saldo_pasivos
    total_no_monet   = activo_fijo_cat1 + activo_fijo_cat2 + activo_fijo_cat3 + inventario

    return {
        'cas_1':  total_activos, 'cas_27': total_pasivos, 'cas_31': saldo_pasivos,
        'cas_26': saldo_act_fiscal, 'cas_32': max(patrimonio_fisc, 0), 'cas_33': total_no_monet,
        'cas_34': min(max(patrimonio_fisc, 0), total_no_monet), 'cas_37': inventario,
        'cas_38': activo_fijo_cat1, 'cas_39': activo_fijo_cat2, 'cas_40': activo_fijo_cat3,
        'cas_49': total_no_monet,
    }

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.title("🏛️ TaxTech Auditor RD")
st.sidebar.markdown("---")
st.sidebar.title("Cliente")
empresa = st.sidebar.text_input("Empresa", value="Empresa de Prueba SRL")
rnc     = st.sidebar.text_input("RNC", value="1-31-12345-6")
periodo = st.sidebar.text_input("Período", value="Enero - Diciembre 2026")
anio    = st.sidebar.text_input("Año Fiscal", value="2026")

st.sidebar.markdown("---")
st.sidebar.title("Materialidad (NIA 320)")
tipo_ent = st.sidebar.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Zonas Francas", "Financieras"])
tasa_ref = 0.01 if tipo_ent == "Comercial / Servicios" else 0.005
pct_mp   = st.sidebar.slider("% Materialidad Planificación", 0.5, 3.0, tasa_ref * 100, 0.1) / 100
pct_me   = st.sidebar.slider("% Materialidad Ejecución (MP×)", 50, 75, 75, 5) / 100

# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────
st.title("TaxTech Auditor — Declaraciones Juradas & Estados Financieros")
st.caption("República Dominicana · DGII · TSS · Ciclo Fiscal 2025-2026")

c_up1, c_up2 = st.columns(2)
with c_up1:
    uploaded = st.file_uploader("📂 Cargar Balanza (Año Actual)", type=["xlsx", "xls", "csv"])
with c_up2:
    uploaded_prev = st.file_uploader("📂 Cargar Balanza (Año Anterior - Opcional)", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.info("👆 Sube una balanza de comprobación para iniciar el análisis.")
    st.stop()

df_bal = procesar_balanza(uploaded)
if df_bal.empty: st.stop()

df_bal = analizar_balanza(df_bal)

df_comp = pd.DataFrame()
if uploaded_prev:
    df_bal_prev = procesar_balanza(uploaded_prev)
    if not df_bal_prev.empty:
        df_comp = procesar_comparativo(df_bal, df_bal_prev)

t_ingresos = abs(df_bal[df_bal['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
t_activos  = abs(df_bal[df_bal['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
t_costos   = abs(df_bal[df_bal['codigo'].str.startswith('5', na=False)]['saldo_final'].sum())
t_gastos   = abs(df_bal[df_bal['codigo'].str.startswith('6', na=False)]['saldo_final'].sum())
t_pasivos  = abs(df_bal[df_bal['codigo'].str.startswith('2', na=False)]['saldo_final'].sum())
t_patrim   = abs(df_bal[df_bal['codigo'].str.startswith('3', na=False)]['saldo_final'].sum())
utilidad_bruta = t_ingresos - t_costos
utilidad_neta  = t_ingresos - t_costos - t_gastos
base_calc = t_ingresos if t_ingresos > 0 else t_activos
mp = base_calc * pct_mp; me = mp * pct_me

st.markdown(f"### 📌 {empresa} — {periodo}")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Ingresos",        f"RD$ {t_ingresos:,.0f}")
c2.metric("Costo de Ventas", f"RD$ {t_costos:,.0f}")
c3.metric("Utilidad Bruta",  f"RD$ {utilidad_bruta:,.0f}", delta=f"{(utilidad_bruta/t_ingresos*100 if t_ingresos else 0):.1f}%")
c4.metric("Utilidad Neta",   f"RD$ {utilidad_neta:,.0f}", delta=f"{(utilidad_neta/t_ingresos*100 if t_ingresos else 0):.1f}%")
c5.metric("MP (Materialidad)", f"RD$ {mp:,.0f}")
c6.metric("ME (Ejecución)",    f"RD$ {me:,.0f}")

st.markdown("---")

tab_comp, tab_bg, tab_er, tab_bal, tab_inconsist, tab_art287, tab_ir2, tab_it1, tab_tss, tab_ir17, tab_consol = st.tabs([
    "📈 EEFF Comparativos",
    "📊 Balance General",
    "📈 Estado de Resultados",
    "📋 Balanza",
    "🚨 Inconsistencias",
    "⚖️ Riesgos Art.287",
    "📝 IR-2 Borrador",
    "⚡ IT-1 (ITBIS)",
    "🏢 TSS / IR-3",
    "💸 IR-17",
    "🏛️ Consolidado",
])

bg = generar_balance_general(df_bal)
er = generar_estado_resultados(df_bal)
ir2_vals = calcular_casillas_ir2(df_bal)

# ── TAB: EEFF COMPARATIVOS Y DETALLES AÑADIDOS ───────────────────────────────
with tab_comp:
    if not df_comp.empty:
        st.markdown("### Dashboard Corporativo y Notas")
        
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            df_chart = pd.DataFrame({
                'Año': [f"{int(anio)-1}", f"{anio}"],
                'Ingresos': [sum(df_comp[df_comp['codigo'].str.startswith('4', na=False)]['saldo_final_Y1'].abs()), t_ingresos],
                'Activos': [sum(df_comp[df_comp['codigo'].str.startswith('1', na=False)]['saldo_final_Y1'].abs()), t_activos]
            })
            fig = px.bar(df_chart, x='Año', y=['Ingresos', 'Activos'], barmode='group', title="Evolución Ingresos vs Activos", color_discrete_sequence=['#1e3a5f', '#38bdf8'])
            st.plotly_chart(fig, use_container_width=True)
            
        with col_chart2:
            st.markdown("#### Exportar Paquete Financiero")
            st.info("Descarga los 4 Estados Financieros Básicos comparativos.")
            excel_corp_bytes = exportar_reporte_corporativo(empresa, periodo, anio, df_comp)
            if excel_corp_bytes:
                st.download_button("📥 Descargar Paquete Corporativo (Excel)", data=excel_corp_bytes, file_name=f"Reporte_{empresa}_{anio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
        st.markdown("---")
        
        # Integración de la Nota de Activos Fijos (PPE)
        st.subheader("📝 Nota: Propiedad, Planta y Equipo (Activos Fijos)")
        df_af = generar_nota_activos_fijos(df_comp)
        if not df_af.empty:
            st.dataframe(df_af.style.format({
                'Costo Inicial': 'RD$ {:,.2f}', 'Adiciones/Retiros': 'RD$ {:,.2f}', 'Costo Final': 'RD$ {:,.2f}',
                'Depr. Inicial': 'RD$ {:,.2f}', 'Gasto/Retiros Depr.': 'RD$ {:,.2f}', 'Depr. Final': 'RD$ {:,.2f}',
                'Balance Neto': 'RD$ {:,.2f}'
            }), use_container_width=True, hide_index=True)
        else:
            st.info("No se detectaron movimientos en Cuentas 15X de Activos Fijos.")
            
        st.markdown("---")
        
        # Integración de ECP y Flujo de Efectivo
        c_ecp, c_efe = st.columns(2)
        with c_ecp:
            st.subheader("💹 Cambios en el Patrimonio")
            df_ecp = generar_cambios_patrimonio(df_comp)
            st.dataframe(df_ecp.style.format({
                'Saldo Inicial': 'RD$ {:,.2f}', 'Variación': 'RD$ {:,.2f}', 'Saldo Final': 'RD$ {:,.2f}'
            }), use_container_width=True, hide_index=True)
            
        with c_efe:
            st.subheader("🌊 Flujo de Efectivo (Borrador)")
            df_flujo = generar_flujo_efectivo(df_comp, utilidad_neta)
            st.dataframe(df_flujo.style.format({'Monto': 'RD$ {:,.2f}'}), use_container_width=True, hide_index=True)
            total_operacion = df_flujo['Monto'].sum()
            st.metric("Flujo Neto de Actividades de Operación", f"RD$ {total_operacion:,.2f}")

    else:
        st.warning("⚠️ Para habilitar el Dashboard y los 4 Estados Financieros Comparativos, debes subir la **Balanza del Año Anterior** en la sección superior.")

# ── TAB: BALANCE GENERAL ──────────────────────────────────────────────────────
with tab_bg:
    st.markdown("### Balance General")
    col_act, col_pas = st.columns(2)
    with col_act:
        st.markdown("#### 🟦 ACTIVO")
        t_ac = sum(m for _, _, m in bg['activo_corriente'])
        t_anc = sum(m for _, _, m in bg['activo_no_corriente'])
        if bg['activo_corriente']: st.dataframe(pd.DataFrame(bg['activo_corriente'], columns=['Cuenta', 'Código', 'Monto']), use_container_width=True, hide_index=True)
        st.metric("Subtotal Activo", f"RD$ {t_ac + t_anc:,.2f}")
    with col_pas:
        st.markdown("#### 🟥 PASIVO & PATRIMONIO")
        t_pc  = sum(m for _, _, m in bg['pasivo_corriente'])
        t_pnc = sum(m for _, _, m in bg['pasivo_no_corriente'])
        t_pat = sum(m for _, _, m in bg['patrimonio'])
        if bg['pasivo_corriente']: st.dataframe(pd.DataFrame(bg['pasivo_corriente'], columns=['Cuenta', 'Código', 'Monto']), use_container_width=True, hide_index=True)
        st.metric("Subtotal Pasivo + Patrimonio", f"RD$ {t_pc + t_pnc + t_pat:,.2f}")
    diferencia = (t_ac + t_anc) - (t_pc + t_pnc + t_pat)
    if abs(diferencia) < 1: st.success(f"✅ Balanza CUADRADA")
    else: st.warning(f"⚠️ Diferencia de cuadre: RD$ {diferencia:,.2f}")

# ── OTROS TABS MANTENIDOS INTACTOS ────────────────────────────────────────────
with tab_er:
    st.markdown("### Estado de Resultados")
    st.metric("Utilidad / Pérdida del Período", f"RD$ {utilidad_neta:,.2f}")

with tab_inconsist:
    df_err = df_bal[~df_bal['validacion_naturaleza'].str.startswith('✅')]
    if df_err.empty: st.success("✅ Sin inconsistencias en naturaleza de saldos.")
    else: st.dataframe(df_err[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)

with tab_art287:
    df_fisc = df_bal[df_bal['alerta_fiscal'] != ""]
    if df_fisc.empty: st.success("✅ Sin alertas fiscales Art. 287 detectadas.")
    else: st.dataframe(df_fisc[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal']], use_container_width=True)

with tab_ir2:
    st.markdown("### 📝 IR-2 — Determinación Ajuste Fiscal Patrimonial")
    st.json(ir2_vals)

with tab_consol:
    isr_est = max(0, utilidad_neta) * 0.27
    st.error(f"**OBLIGACIONES FISCALES (ISR Estimado): RD$ {isr_est:,.2f}**")