# TaxTech Auditor RD 🏛️
### Análisis de Balanza, Estados Financieros y Declaraciones Juradas

---

## Requisitos
- Python 3.9 o superior (https://python.org)

## Instalación y arranque

```bash
python launch.py
```

Eso es todo. El script instala las dependencias automáticamente y abre el navegador.

---

## Qué puedes hacer

| Tab | Función |
|-----|---------|
| **Balance General** | Activo / Pasivo / Patrimonio desde la balanza. Detecta cuadre. |
| **Estado de Resultados** | P&L con márgenes bruto, operacional y neto. ISR estimado. |
| **Balanza** | Vista cruda + descarga con análisis |
| **Inconsistencias** | Cuentas con saldo contrario a su naturaleza contable |
| **Riesgos Art. 287** | Alertas por palabras clave fiscales (combustible, honorarios, etc.) |
| **IR-2 Borrador** | Casillas del Ajuste Patrimonial calculadas desde la balanza. Sube el IR-2.xls oficial para cruzar. |
| **IT-1 (ITBIS)** | Liquidación mensual. Sube los formatos 606/607 para más precisión. |
| **TSS / IR-3** | Costo de nómina, aportes patronales y retenciones empleados |
| **IR-17** | Otras retenciones (honorarios, reparaciones, retribuciones, remesas) |
| **Consolidado** | Resumen de todas las obligaciones + ratio Deuda/Activo |

---

## Formato de la Balanza

Columnas mínimas (nombres flexibles, detección automática):

| Código | Nombre/Cuenta | Débito | Crédito | Saldo |
|--------|---------------|--------|---------|-------|
| 1101   | Caja           | 500000 | 0       | 500000|

El sistema detecta variaciones: `codigo`, `cta`, `cuenta no`, etc.

---

## Exportar

El botón **"Descargar Estados Financieros + IR-2 (Excel)"** genera un libro con:
1. Hoja: Balance General formateado
2. Hoja: Estado de Resultados
3. Hoja: Borrador IR-2 con todas las casillas

---

*Desarrollado para el ciclo fiscal dominicano (DGII / TSS / 2025-2026)*
