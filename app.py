import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración inicial de la página de Streamlit
st.set_page_config(
    page_title="TaxTech Auditor - Análisis de Balanza & Riesgo Fiscal",
    layout="wide"
)

# ==========================================
# 1. PARÁMETROS FISCALES Y DE CONFIGURACIÓN (CTRD / TSS / ITBIS 2026)
# ==========================================
NATURALEZAS = {
    '1': 'Debito', '2': 'Credito', '3': 'Credito',
    '4': 'Credito', '5': 'Debito', '6': 'Debito'
}

PALABRAS_CRITICAS_ART287 = {
    'combustible': 'Riesgo Art. 287 CTRD: Validar deducibilidad, comprobantes con NCF válido y uso de medios de pago para crédito fiscal de ITBIS.',
    'representacion': 'Riesgo Art. 287 CTRD: Gastos de representación. Sujetos a criterios de razonabilidad, proporcionalidad y documentación fehaciente.',
    'retribucion': 'Riesgo Art. 318 CTRD / Reg. 139-98: Retribuciones en especie. Validar que la empresa efectúe el pago del ISR sustitutivo correspondiente.',
    'gasto de personal': 'Riesgo Art. 287 CTRD: Cruce obligatorio con la declaración jurada de TSS (Formulario IR-4) para admitir la deducción.',
    'honorario': 'Riesgo Art. 309 CTRD: Validar aplicación de retenciones fiscales (10% a personas físicas o 2% entre personas jurídicas).'
}

MAPEO_IR2 = {
    '11': 'Anexo A - Efectivo e Inversiones Temporales',
    '12': 'Anexo A - Cuentas por Cobrar (Neto)',
    '13': 'Anexo A - Inventarios',
    '15': 'Anexo A - Propiedad, Planta y Equipo (Neto)',
    '21': 'Anexo A - Pasivos Corrientes / Cuentas por Pagar',
    '31': 'Anexo A - Capital Social y Reservas',
    '41': 'Anexo B - Ingresos por Operaciones (Locales)',
    '51': 'Anexo B - Costo de Ventas / Servicios',
    '61': 'Anexo B - Gastos de Personal (TSS / Sueldos)',
    '62': 'Anexo B - Gastos Operativos y de Administración',
    '63': 'Anexo B - Gastos Financieros'
}

# Coeficientes Oficiales de Ley
TASA_ITBIS_GENERAL = 0.18
TASA_SFS_PATRONAL = 0.0709
TASA_AFP_PATRONAL = 0.0710
TASA_SRL_PROMEDIO = 0.0120
TASA_INFOTEP = 0.0100
TASA_SFS_EMPLEADO = 0.0304
TASA_AFP_EMPLEADO = 0.0287
COSTO_PER_CAPITA_2026 = 1691.38

# --- ESTILOS COMPARTIDOS PARA EXPORTACIÓN EXCEL ---
FILL_HEADER = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F497D")
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=11, bold=False)
FONT_BOLD = Font(name="Calibri", size=11, bold=True)
BORDER_THIN = Side(border_style="thin", color="D9D9D9")
CELL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)

def aplicar_estilos_base(ws, titulo, subtitulo):
    ws.views.sheetView[0].showGridLines = True
    ws["B2"] = titulo
    ws["B2"].font = FONT_TITLE
    ws["B3"] = subtitulo
    ws["B3"].font = Font(name="Calibri", size=11, italic=True)

def autoajustar_columnas(ws):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter != 'A':
            ws.column_dimensions[col_letter].width = 30 if col_letter in ['E', 'F'] else 50

# ==========================================
# 2. FUNCIONES DE LÓGICA DE NEGOCIO Y AUDITORÍA
# ==========================================
def procesar_balanza(file) -> pd.DataFrame:
    try:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl', dtype={'codigo': str, 'código': str})
        else:
            df = pd.read_csv(file, dtype={'codigo': str, 'código': str})
        
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        mapeo_columnas = {
            'código': 'codigo', 'cuenta': 'cuenta', 'nombre': 'cuenta', 'nombre de cuenta': 'cuenta',
            'débito': 'debito', 'crédito': 'credito', 'saldo final': 'saldo_final', 'saldo': 'saldo_final'
        }
        df = df.rename(columns=mapeo_columnas)
        
        columnas_requeridas = {'codigo', 'cuenta', 'debito', 'credito', 'saldo_final'}
        if not columnas_requeridas.issubset(set(df.columns)):
            st.error(f"⚠️ Estructura incorrecta en la balanza de comprobación.")
            return pd.DataFrame()
            
        df['codigo'] = df['codigo'].fillna('').astype(str).str.strip()
        df['codigo'] = df['codigo'].apply(lambda x: x.split('.')[0] if '.' in x else x)
        df['cuenta'] = df['cuenta'].fillna('').astype(str).str.strip()
        
        df = df[~df['codigo'].str.lower().str.contains('total|resultado|suma', na=False)]
        df = df[~df['cuenta'].str.lower().str.contains('total|resultado|suma', na=False)]
        df = df[df['codigo'] != '']
        
        for col in ['debito', 'credito', 'saldo_final']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
        return df
    except Exception as e:
        st.error(f"❌ Error al procesar archivo base: {str(e)}")
        return pd.DataFrame()

def analizar_balanza(df: pd.DataFrame) -> pd.DataFrame:
    alertas_nat = []
    alertas_fisc = []
    casillas_ir2 = []
    
    for _, row in df.iterrows():
        codigo_str = row['codigo']
        nombre_cuenta = str(row['cuenta']).strip().lower()
        
        if not codigo_str or any(keyword in nombre_cuenta for keyword in ['total', 'suma']):
            alertas_nat.append("Ignorado")
            alertas_fisc.append("Sin observaciones")
            casillas_ir2.append("No Aplica")
            continue
            
        primer_digito = codigo_str[0]
        primeros_dos = codigo_str[:2]
        nat_esperada = NATURALEZAS.get(primer_digito, None)
        saldo = row['saldo_final']
        
        if nat_esperada == 'Debito' and saldo < 0:
            alertas_nat.append("Saldo Crédito inusual (Naturaleza Débito)")
        elif nat_esperada == 'Credito' and saldo > 0:
            alertas_nat.append("Saldo Débito inusual (Naturaleza Crédito)")
        else:
            alertas_nat.append("Correcto")
            
        alerta_f = "Sin observaciones"
        for palabra, mensaje in PALABRAS_CRITICAS_ART287.items():
            if palabra in nombre_cuenta:
                alerta_f = mensaje
                break
        alertas_fisc.append(alerta_f)
        
        casilla = MAPEO_IR2.get(primeros_dos, MAPEO_IR2.get(primer_digito + '1', 'Otros Conceptos No Mapeados'))
        casillas_ir2.append(casilla)
        
    df['validacion_naturaleza'] = alertas_nat
    df['alerta_fiscal_rd'] = alertas_fisc
    df['casilla_ir2'] = casillas_ir2
    return df

# ==========================================
# 3. INTERFAZ DE USUARIO PRINCIPAL
# ==========================================

# --- BARRA LATERAL ---
st.sidebar.title("Configuración del Cliente")
empresa = st.sidebar.text_input("Nombre de la Empresa", value="Empresa de Prueba SRL")
periodo = st.sidebar.text_input("Período de Análisis", value="2026/05/30")

st.sidebar.markdown("---")
st.sidebar.title("Parámetros de Materialidad (NIA 320)")
tipo_entidad = st.sidebar.selectbox("Tipo de Entidad", ["Comercial / Servicios", "Zonas Francas", "Financieras"])

tasa_referencia = 0.01 if tipo_entidad == "Comercial / Servicios" else 0.005
porcentaje_mp = st.sidebar.slider("Porcentaje de Materialidad", 0.5, 3.0, tasa_referencia * 100, step=0.1) / 100
porcentaje_me = st.sidebar.slider("Porcentaje de Materialidad de Ejecución (ME)", 50, 75, 75, step=5) / 100

st.header("1. Carga de Balanza de Comprobación Base")
uploaded_file = st.file_uploader("Upload", type=["xlsx", "csv"], label_visibility="collapsed")

if uploaded_file is not None:
    df_balanza = procesar_balanza(uploaded_file)
    
    if not df_balanza.empty:
        df_balanza = analizar_balanza(df_balanza)
        
        total_activos = abs(df_balanza[df_balanza['codigo'].str.startswith('1', na=False)]['saldo_final'].sum())
        total_ingresos = abs(df_balanza[df_balanza['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
        
        base_calculo = total_ingresos if total_ingresos > 0 else total_activos
        mp = base_calculo * porcentaje_mp
        me = mp * porcentaje_me
        
        st.markdown("---")
        st.subheader(f"📌 Informe de Auditoría Analítica: {empresa} — Período: {periodo}")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Ingresos Declarados", f"RD$ {total_ingresos:,.2f}")
        c2.metric("Total Activos Registrados", f"RD$ {total_activos:,.2f}")
        c3.metric("Materialidad Planificación (MP)", f"RD$ {mp:,.2f}")
        c4.metric("Materialidad Ejecución (ME)", f"RD$ {me:,.2f}")
        
        st.markdown("---")
        
        # PESTAÑAS PRINCIPALES
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "📋 Balanza", "🚨 Inconsistencias", "🇩🇴 Riesgos Art. 287", 
            "📋 Borrador Anual IR-2", "⚡ Mensual (Liquidación IT-1)", 
            "🏢 Nómina y TSS (IR-3)", "💸 Liquidación IR-17", "🏛️ Consolidado Fiscal General"
        ])
        
        with tab1:
            st.dataframe(df_balanza, use_container_width=True)
            
        with tab2:
            df_errores = df_balanza[(df_balanza['validacion_naturaleza'] != "Correcto") & (df_balanza['validacion_naturaleza'] != "Ignorado")]
            if not df_errores.empty:
                st.error(f"Se detectaron {len(df_errores)} cuentas con saldos inconsistentes.")
                st.dataframe(df_errores[['codigo', 'cuenta', 'saldo_final', 'validacion_naturaleza']], use_container_width=True)
            else:
                st.success("✅ Excelente: No se han encontrado cuentas con inconsistencias en su saldo final.")
                
        with tab3:
            df_fiscal = df_balanza[df_balanza['alerta_fiscal_rd'] != "Sin observaciones"]
            if not df_fiscal.empty:
                st.warning(f"Atención: Se identificaron {len(df_fiscal)} cuentas expuestas a revisión fiscal.")
                st.dataframe(df_fiscal[['codigo', 'cuenta', 'saldo_final', 'alerta_fiscal_rd']], use_container_width=True)
            else:
                st.success("✅ Cumplimiento Inicial: No se detectaron alertas críticas.")
                
        with tab4:
            st.markdown("### 📋 Mapeo y Cruce Avanzado - Formulario Anual IR-2")
            with st.expander("📥 CARGAR ARCHIVO AUXILIAR DE ANEXOS IR-2 (DGII)", expanded=False):
                uploaded_anexos = st.file_uploader("Subir borrador de anexos (A, B, C, D) en Excel", type=["xlsx"], key="ir2_anexos")
            df_ir2 = df_balanza.groupby('casilla_ir2')['saldo_final'].sum().reset_index()
            df_ir2['saldo_final'] = df_ir2['saldo_final'].apply(lambda x: abs(x))
            df_ir2.columns = ['Renglón Formulario DGII', 'Monto Acumulado (RD$)']
            st.dataframe(df_ir2, use_container_width=True)
            
        # --- EXTRACCIÓN Y CÁLCULOS FISCALES DEL PERIODO ---
        monto_ingresos_gravados = abs(df_balanza[df_balanza['codigo'].str.startswith('4', na=False)]['saldo_final'].sum())
        itbis_ventas_generado = monto_ingresos_gravados * TASA_ITBIS_GENERAL
        compras_y_gastos_base = abs(df_balanza[(df_balanza['codigo'].str.startswith(('5', '6'), na=False)) & (~df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario|tss|infotep|percapita', na=False))]['saldo_final'].sum())
        itbis_soportado_compras = compras_y_gastos_base * TASA_ITBIS_GENERAL
        base_honorarios_fisicos = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('honorario', na=False)]['saldo_final'].sum())
        itbis_ret_100_fisicas = (base_honorarios_fisicos * TASA_ITBIS_GENERAL) * 1.00
        base_servicios_juridicas = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('servicio tecnico|consultoria|reparacion', na=False)]['saldo_final'].sum())
        itbis_ret_30_juridicas = (base_servicios_juridicas * TASA_ITBIS_GENERAL) * 0.30
        cuenta_ret_tarjeta = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('retencion tarjeta|adquirente|cardnet|azul', na=False)]['saldo_final'].sum())
        itbis_ret_tarjetas_2 = cuenta_ret_tarjeta if cuenta_ret_tarjeta > 0 else (monto_ingresos_gravados * 0.60) * 0.02
        neto_itbis_resultado = itbis_ventas_generado - (itbis_soportado_compras + itbis_ret_100_fisicas + itbis_ret_30_juridicas + itbis_ret_tarjetas_2)

        gasto_nominas_global = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('personal|sueldo|salario', na=False)]['saldo_final'].sum())
        gasto_per_capita_balanza = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('percapita|per capita|dependiente adicional', na=False)]['saldo_final'].sum())
        costo_patronal_total = (gasto_nominas_global * TASA_SFS_PATRONAL) + (gasto_nominas_global * TASA_AFP_PATRONAL) + (gasto_nominas_global * TASA_SRL_PROMEDIO) + (gasto_nominas_global * TASA_INFOTEP)
        retenciones_empleados_total = (gasto_nominas_global * TASA_SFS_EMPLEADO) + (gasto_nominas_global * TASA_AFP_EMPLEADO) + gasto_per_capita_balanza
        total_liquidacion_ir3_tss = costo_patronal_total + retenciones_empleados_total

        b_honorarios = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('honorario', na=False)]['saldo_final'].sum())
        b_reparaciones = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('reparacion|mantenimiento', na=False)]['saldo_final'].sum())
        b_vehiculos = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('vehiculo personal|combustible empleado', na=False)]['saldo_final'].sum())
        b_renta_vivienda = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('renta personal|alquiler personal|vivienda', na=False)]['saldo_final'].sum())
        b_otras_retribuciones = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('retribucion|especie', na=False) & ~df_balanza['cuenta'].str.lower().str.contains('vehiculo|renta|alquiler|vivienda', na=False)]['saldo_final'].sum())
        b_espana = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('españa|espana', na=False)]['saldo_final'].sum())
        b_canada = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('canada|canadá', na=False)]['saldo_final'].sum())
        b_exterior_general = abs(df_balanza[df_balanza['cuenta'].str.lower().str.contains('exterior|remesa|extranjero', na=False) & ~df_balanza['cuenta'].str.lower().str.contains('españa|espana|canada|canadá', na=False)]['saldo_final'].sum())
        
        r_honorarios = b_honorarios * 0.10
        r_reparaciones = b_reparaciones * 0.02
        r_vehiculos = b_vehiculos * 0.27
        r_renta = b_renta_vivienda * 0.27
        r_otras_ret = b_otras_retribuciones * 0.27
        r_espana = b_espana * 0.10
        r_canada = b_canada * 0.18
        r_exterior = b_exterior_general * 0.27
        total_ir17 = r_honorarios + r_reparaciones + r_vehiculos + r_renta + r_otras_ret + r_espana + r_canada + r_exterior

        with tab5:
            st.markdown("### 🇩🇴 Módulo de Liquidación de ITBIS (Anexo A + Formulario IT-1)")
            if neto_itbis_resultado > 0:
                st.error(f"🚨 **IMPUESTO NETO A PAGAR EN IT-1:** RD$ {neto_itbis_resultado:,.2f}")
            else:
                st.success(f"🎉 **SALDO A FAVOR COMPENSABLE:** RD$ {abs(neto_itbis_resultado):,.2f}")
            
            cit1, cit2, cit3, cit4 = st.columns(4)
            cit1.metric("ITBIS Ventas (Generado)", f"RD$ {itbis_ventas_generado:,.2f}")
            cit2.metric("ITBIS Soportado (Adelantos)", f"RD$ {itbis_soportado_compras:,.2f}")
            cit3.metric("Retenciones Sufridas", f"RD$ {itbis_ret_100_fisicas + itbis_ret_30_juridicas:,.2f}")
            cit4.metric("Retención Tarjeta (2%)", f"RD$ {itbis_ret_tarjetas_2:,.2f}")
            
            buffer_it1 = io.BytesIO()
            wb_it1 = openpyxl.Workbook()
            
            ws_anexo = wb_it1.active
            ws_anexo.title = "Anexo_A_IT1"
            aplicar_estilos_base(ws_anexo, f"TAXTECH AUDITOR RD — ANEXO A DEL IT-1: {empresa.upper()}", f"Desglose Analítico de Ingresos, Adelantos y Retenciones")
            
            headers_anexo = ["Casilla", "Descripción del Concepto Operativo", "Monto Base Imponible", "ITBIS Liquidado / Retenido"]
            for col_idx, h in enumerate(headers_anexo, start=2):
                cell = ws_anexo.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
            
            datos_anexo = [
                ("Casilla 1", "Ingresos por Operaciones Locales Gravadas (Ventas)", monto_ingresos_gravados, itbis_ventas_generado),
                ("Casilla 12", "ITBIS Pagado en Compras Locales (Adelantos del 606)", compras_y_gastos_base, itbis_soportado_compras),
                ("Casilla 22", "ITBIS Retenido por Servicios Profesionales de Personas Físicas (100%)", base_honorarios_fisicos, itbis_ret_100_fisicas),
                ("Casilla 23", "ITBIS Retenido entre Personas Jurídicas (30% Norma 02-05)", base_servicios_juridicas, itbis_ret_30_juridicas),
                ("Casilla 30", "Retención por Operaciones con Tarjetas de Crédito (2% Norma 08-04)", 0.00, itbis_ret_tarjetas_2)
            ]
            
            for idx, (cas, con, bas, imp) in enumerate(datos_anexo):
                r_idx = 6 + idx
                ws_anexo.cell(row=r_idx, column=2, value=cas).alignment = Alignment(horizontal="center")
                ws_anexo.cell(row=r_idx, column=3, value=con)
                c_bas = ws_anexo.cell(row=r_idx, column=4, value=bas)
                c_bas.number_format = 'RD$ #,##0.00'; c_bas.alignment = Alignment(horizontal="right")
                c_imp = ws_anexo.cell(row=r_idx, column=5, value=imp)
                c_imp.number_format = 'RD$ #,##0.00'; c_imp.alignment = Alignment(horizontal="right")
                
                for c in range(2, 6):
                    cell = ws_anexo.cell(row=r_idx, column=c)
                    cell.font = FONT_BODY; cell.border = CELL_BORDER
                    if idx % 2 == 1: cell.fill = FILL_ZEBRA
            autoajustar_columnas(ws_anexo)
            
            ws_form = wb_it1.create_sheet(title="Formulario_IT1")
            aplicar_estilos_base(ws_form, f"TAXTECH AUDITOR RD — FORMULARIO DEFINITIVO IT-1: {empresa.upper()}", f"Liquidación Final del Impuesto — Período: {periodo}")
            
            headers_form = ["Línea", "Renglón Final IT-1", "Fórmula de Amarre (DGII)", "Monto Relacionado"]
            for col_idx, h in enumerate(headers_form, start=2):
                cell = ws_form.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
            
            datos_form = [
                ("Línea 1", "Total ITBIS Bruto por Operaciones", "=Anexo_A_IT1!E6", "=Anexo_A_IT1!E6"),
                ("Línea 2", "(-) Menos ITBIS Deducible (Adelanto de Compras)", "=Anexo_A_IT1!E7", "=Anexo_A_IT1!E7"),
                ("Línea 3", "(-) Menos ITBIS Retenido por Personas Físicas", "=Anexo_A_IT1!E8", "=Anexo_A_IT1!E8"),
                ("Línea 4", "(-) Menos ITBIS Retenido por Personas Jurídicas", "=Anexo_A_IT1!E9", "=Anexo_A_IT1!E9"),
                ("Línea 5", "(-) Menos Retención por Tarjetas de Crédito", "=Anexo_A_IT1!E10", "=Anexo_A_IT1!E10"),
                ("TOTAL", "IMPUESTO NETO A PAGAR / SALDO A FAVOR", "=E6-SUM(E7:E10)", "=E6-SUM(E7:E10)")
            ]
            
            for idx, (lin, ren, for_am, val) in enumerate(datos_form):
                r_idx = 6 + idx
                ws_form.cell(row=r_idx, column=2, value=lin).alignment = Alignment(horizontal="center")
                ws_form.cell(row=r_idx, column=3, value=ren)
                ws_form.cell(row=r_idx, column=4, value=for_am).font = Font(name="Calibri", size=9, color="7F7F7F")
                c_val = ws_form.cell(row=r_idx, column=5, value=val)
                c_val.number_format = 'RD$ #,##0.00'; c_val.alignment = Alignment(horizontal="right")
                
                for c in range(2, 6):
                    cell = ws_form.cell(row=r_idx, column=c)
                    cell.font = FONT_BODY if lin != "TOTAL" else FONT_BOLD
                    cell.border = CELL_BORDER
                    if idx % 2 == 1 and lin != "TOTAL": cell.fill = FILL_ZEBRA
                    if lin == "TOTAL": cell.fill = FILL_TOTAL
            
            autoajustar_columnas(ws_form)
            wb_it1.save(buffer_it1)
            
            st.markdown(" ")
            st.download_button(
                label="📥 Descargar Libro Completo IT-1 (Anexo A + Formulario)", 
                data=buffer_it1.getvalue(), 
                file_name=f"Borrador_Completo_IT1_{empresa.replace(' ', '_')}.xlsx"
            )

        with tab6:
            st.markdown("### 🏢 Módulo de Conciliación y Liquidación TSS / INFOTEP / IR-3")
            
            # Subir detalles de empleados para la TSS (Cargador Auxiliar Operativo)
            uploaded_empleados = st.file_uploader("Subir listado auxiliar de empleados (Excel/CSV con salarios individuales)", type=["xlsx", "csv"], key="tss_empleados")
            if uploaded_empleados is not None:
                st.success("✅ Detalle de empleados cargado correctamente para cruce con IR-4.")
            
            cp1, cp2, cp3, cp4 = st.columns(4)
            cp1.metric("SFS Patronal (7.09%)", f"RD$ {gasto_nominas_global * TASA_SFS_PATRONAL:,.2f}")
            cp2.metric("AFP Patronal (7.10%)", f"RD$ {gasto_nominas_global * TASA_AFP_PATRONAL:,.2f}")
            cp3.metric("Seguro Riesgos Laborales (1.20%)", f"RD$ {gasto_nominas_global * TASA_SRL_PROMEDIO:,.2f}")
            cp4.metric("Aporte INFOTEP (1.00%)", f"RD$ {gasto_nominas_global * TASA_INFOTEP:,.2f}")
            
            buffer_tss = io.BytesIO()
            wb_tss = openpyxl.Workbook()
            ws_tss = wb_tss.active
            ws_tss.title = "Liquidación TSS"
            aplicar_estilos_base(ws_tss, f"TAXTECH AUDITOR RD — CONTROL DE NÓMINA: {empresa.upper()}", f"Carga Masiva de Aportes Patronales y Retenciones TSS / IR-3")
            
            headers_tss = ["Tipo Componente", "Concepto de Retención o Aporte", "Porcentaje", "Monto Autocalculado"]
            for col_idx, h in enumerate(headers_tss, start=2):
                cell = ws_tss.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
            
            datos_tss = [
                ("Patronal", "Seguro Familiar de Salud (SFS Patronal)", 0.0709, gasto_nominas_global * 0.0709),
                ("Patronal", "Fondo de Pensiones (AFP Patronal)", 0.0710, gasto_nominas_global * 0.0710),
                ("Patronal", "Seguro de Riesgos Laborales (SRL)", 0.0120, gasto_nominas_global * 0.0120),
                ("Patronal", "Aporte INFOTEP Obligatorio (1%)", 0.0100, gasto_nominas_global * 0.0100),
                ("Empleado", "SFS Retención Trabajador", 0.0304, gasto_nominas_global * 0.0304),
                ("Empleado", "AFP Retención Trabajador", 0.0287, gasto_nominas_global * 0.0287),
                ("Empleado", "Aporte Percápita Adicional (Dependientes)", 0.0000, gasto_per_capita_balanza),
                ("TOTAL", "TOTAL LIQUIDACIÓN MENSUAL DEL ARCHIVO TSS", 0.0000, f"=SUM(E6:E12)")
            ]
            
            for idx, (tipo, con, por, mnt) in enumerate(datos_tss):
                r_idx = 6 + idx
                ws_tss.cell(row=r_idx, column=2, value=tipo).alignment = Alignment(horizontal="center")
                ws_tss.cell(row=r_idx, column=3, value=con)
                c_por = ws_tss.cell(row=r_idx, column=4, value=por)
                c_por.number_format = '0.00%'; c_por.alignment = Alignment(horizontal="center")
                c_mnt = ws_tss.cell(row=r_idx, column=5, value=mnt)
                c_mnt.number_format = 'RD$ #,##0.00'; c_mnt.alignment = Alignment(horizontal="right")
                
                for c in range(2, 6):
                    cell = ws_tss.cell(row=r_idx, column=c)
                    cell.font = FONT_BODY if tipo != "TOTAL" else FONT_BOLD
                    cell.border = CELL_BORDER
                    if idx % 2 == 1 and tipo != "TOTAL": cell.fill = FILL_ZEBRA
                    if tipo == "TOTAL": cell.fill = FILL_TOTAL
            
            autoajustar_columnas(ws_tss)
            wb_tss.save(buffer_tss)
            st.download_button("📥 Descargar Plantilla Auxiliar TSS (Excel)", data=buffer_tss.getvalue(), file_name=f"Plantilla_TSS_{empresa.replace(' ', '_')}.xlsx")

        with tab7:
            st.markdown("### 💸 Formulario IR-17: Declaración Jurada de Otras Retenciones")
            st.error(f"💸 **TOTAL IMPUESTO A PAGAR (IR-17):** RD$ {total_ir17:,.2f}")
            
            df_ir17_pantalla = pd.DataFrame({
                'Casilla': ['Casilla 1', 'Casilla 2', 'Casilla 8', 'Casilla 9', 'Casilla 10', 'Casilla 15', 'Casilla 16', 'Casilla 17'],
                'Concepto Oficial': [
                    'Honorarios por Servicios Profesionales', 'Otras Rentas / Servicios Técnicos y Reparaciones',
                    'Retribuciones Complementarias - Vehículos', 'Retribuciones Complementarias - Alquileres',
                    'Retribuciones Complementarias - Otros', 'Remesas - España', 'Remesas - Canadá', 'Remesas - General'
                ],
                'Tasa': ['10%', '2%', '27%', '27%', '27%', '10%', '18%', '27%'],
                'Base Imponible': [b_honorarios, b_reparaciones, b_vehiculos, b_renta_vivienda, b_otras_retribuciones, b_espana, b_canada, b_exterior_general],
                'Impuesto Retenido': [r_honorarios, r_reparaciones, r_vehiculos, r_renta, r_otras_ret, r_espana, r_canada, r_exterior]
            })
            st.dataframe(df_ir17_pantalla.style.format({'Base Imponible': 'RD$ {:,.2f}', 'Impuesto Retenido': 'RD$ {:,.2f}'}), use_container_width=True, hide_index=True)

            # --- CORRECCIÓN DEFINITIVA DE NOMBRE DEL LIBRO (OPENPYXL) ---
            buffer_ir17 = io.BytesIO()
            wb_ir17 = openpyxl.Workbook()
            ws_ir17 = wb_ir17.active
            ws_ir17.title = "Borrador IR-17"
            aplicar_estilos_base(ws_ir17, f"TAXTECH AUDITOR RD — RESUMEN IR-17: {empresa.upper()}", f"Borrador de Casillas Oficiales del Formulario IR-17 de la DGII")
            
            headers_ir17 = ["Casilla", "Concepto Detallado (DGII)", "Tasa", "Monto Base Imponible", "Impuesto Retenido"]
            for col_idx, h in enumerate(headers_ir17, start=2):
                cell = ws_ir17.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
                cell.alignment = Alignment(horizontal="center" if col_idx in [2,4] else "left")
            
            datos_ir17 = [
                ("Casilla 1", "Honorarios por Servicios Profesionales (Persona Física)", 0.10, b_honorarios),
                ("Casilla 2", "Otras Rentas / Servicios Técnicos y Reparaciones", 0.02, b_reparaciones),
                ("Casilla 8", "Retribuciones Complementarias - Asignación de Vehículos / Combustible", 0.27, b_vehiculos),
                ("Casilla 9", "Retribuciones Complementarias - Pago de Alquileres / Renta de Vivienda", 0.27, b_renta_vivienda),
                ("Casilla 10", "Retribuciones Complementarias - Otros Beneficios en Especie", 0.27, b_otras_retribuciones),
                ("Casilla 15", "Remesas al Exterior - Convenio Doble Imposición España", 0.10, b_espana),
                ("Casilla 16", "Remesas al Exterior - Convenio Doble Imposición Canadá", 0.18, b_canada),
                ("Casilla 17", "Remesas al Exterior - Otras Rentas de Fuente Dominicana (General)", 0.27, b_exterior_general)
            ]
            
            for idx, (cas, con, tas, bas) in enumerate(datos_ir17):
                r_idx = 6 + idx
                ws_ir17.cell(row=r_idx, column=2, value=cas).alignment = Alignment(horizontal="center")
                ws_ir17.cell(row=r_idx, column=3, value=con)
                c_tas = ws_ir17.cell(row=r_idx, column=4, value=tas)
                c_tas.number_format = '0%'; c_tas.alignment = Alignment(horizontal="center")
                c_bas = ws_ir17.cell(row=r_idx, column=5, value=bas)
                c_bas.number_format = 'RD$ #,##0.00'; c_bas.alignment = Alignment(horizontal="right")
                c_imp = ws_ir17.cell(row=r_idx, column=6, value=f"=E{r_idx}*D{r_idx}")
                c_imp.number_format = 'RD$ #,##0.00'; c_imp.alignment = Alignment(horizontal="right")
                
                for c in range(2, 7):
                    cell = ws_ir17.cell(row=r_idx, column=c)
                    cell.font = FONT_BODY; cell.border = CELL_BORDER
                    if idx % 2 == 1: cell.fill = FILL_ZEBRA
            
            tot_row = 14
            ws_ir17.cell(row=tot_row, column=2, value="TOTAL").font = FONT_BOLD
            ws_ir17.cell(row=tot_row, column=2).fill = FILL_TOTAL; ws_ir17.cell(row=tot_row, column=2).border = CELL_BORDER; ws_ir17.cell(row=tot_row, column=2).alignment = Alignment(horizontal="center")
            ws_ir17.cell(row=tot_row, column=3, value="Liquidación General de Retenciones").font = FONT_BOLD
            ws_ir17.cell(row=tot_row, column=3).fill = FILL_TOTAL; ws_ir17.cell(row=tot_row, column=3).border = CELL_BORDER
            ws_ir17.cell(row=tot_row, column=4, value="").fill = FILL_TOTAL; ws_ir17.cell(row=tot_row, column=4).border = CELL_BORDER
            
            t_bas = ws_ir17.cell(row=tot_row, column=5, value=f"=SUM(E6:E13)")
            t_bas.number_format = 'RD$ #,##0.00'; t_bas.font = FONT_BOLD; t_bas.fill = FILL_TOTAL; t_bas.border = CELL_BORDER
            t_imp = ws_ir17.cell(row=tot_row, column=6, value=f"=SUM(F6:F13)")
            t_imp.number_format = 'RD$ #,##0.00'; t_imp.font = FONT_BOLD; t_imp.fill = FILL_TOTAL; t_imp.border = CELL_BORDER
            
            autoajustar_columnas(ws_ir17)
            wb_ir17.save(buffer_ir17)
            st.download_button("📥 Descargar Borrador Resumido IR-17 (Excel)", data=buffer_ir17.getvalue(), file_name=f"Borrador_Resumido_IR17_{empresa.replace(' ', '_')}.xlsx")
            
        with tab8:
            st.markdown("### 🏛️ Consolidado Fiscal General del Periodo")
            itbis_caja = neto_itbis_resultado if neto_itbis_resultado > 0 else 0.0
            gran_total_periodo_pagar = itbis_caja + total_liquidacion_ir3_tss + total_ir17
            
            st.warning(f"🏦 **EFECTIVO TOTAL ESTIMADO A TRANSFERIR (DGII / TSS):** RD$ {gran_total_periodo_pagar:,.2f}")
            
            buffer_con = io.BytesIO()
            wb_con = openpyxl.Workbook()
            ws_con = wb_con.active
            ws_con.title = "Consolidado Fiscal"
            aplicar_estilos_base(ws_con, f"TAXTECH AUDITOR RD — VOLANTE DE CONSOLIDACIÓN GENERAL", f"Resumen Maestro de Obligaciones Liquidadas — Período Fiscal: {periodo}")
            
            headers_con = ["Formulario Oficial", "Origen / Módulo", "Estado de Cuenta", "Monto Neto Del Periodo"]
            for col_idx, h in enumerate(headers_con, start=2):
                cell = ws_con.cell(row=5, column=col_idx, value=h)
                cell.font = FONT_HEADER; cell.fill = FILL_HEADER; cell.border = CELL_BORDER
            
            datos_con = [
                ("Formulario IT-1", "Módulo de ITBIS / 606 / 607", "Saldo a Pagar" if neto_itbis_resultado > 0 else "Saldo a Favor", itbis_caja),
                ("Tesorería TSS", "Seguridad Social Patronal", "Costo Operativo", costo_patronal_total),
                ("Formulario IR-3", "Retenciones Empleados Nómina", "Pasivo de Retención", retenciones_empleados_total),
                ("Formulario IR-17", "Retenciones Locales / Exterior", "Impuesto Retenido", total_ir17),
                ("TOTAL GENERAL", "FONDO LIQUID
