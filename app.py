
import streamlit as st
import pandas as pd
from logic import process_reconciliation
import io

st.set_page_config(
    page_title="Conciliación Bancaria",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS ---
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 3em;
        background-color: #2e7d32;
        color: white;
        font-weight: bold;
        border: none;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #1b5e20;
        border: none;
        color: white;
    }
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border: 1px solid #eee;
    }
    h1, h2, h3 {
        color: #1e3d59;
    }
    .stInfo {
        background-color: #e3f2fd;
        border-color: #2196f3;
        color: #0d47a1;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏦 Conciliación Bancaria Automática")
st.markdown("""
Esta herramienta realiza la **conciliación bancaria** completa, identificando diferencias **temporales** y **permanentes** 
entre el saldo del Libro y el saldo del Extracto Bancario.
""")

# --- Sidebar Instructions ---
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/000000/bank.png", width=80)
    st.title("Instrucciones")
    st.info("""
    1. **Sube el Libro**: Excel de tu sistema contable con las transacciones registradas.
    2. **Sube el Extracto**: Excel del banco con los movimientos bancarios.
    3. **Procesa**: El sistema identificará:
       - ✅ **Items que coinciden**
       - 🕐 **Diferencias Temporales** (depósitos en tránsito, cheques pendientes)
       - ⚠️ **Diferencias Permanentes** (comisiones, impuestos, errores)
    4. **Descarga**: Obtén el reporte de conciliación completo.
    """)
    
    st.divider()
    st.subheader("📚 ¿Qué son las diferencias?")
    
    with st.expander("🕐 Diferencias Temporales"):
        st.markdown("""
        Se ajustan con el paso del tiempo **sin necesidad de asiento contable**:
        - Depósitos en tránsito
        - Cheques pendientes de acreditación
        - Pagos registrados en diferentes períodos
        """)
    
    with st.expander("⚠️ Diferencias Permanentes"):
        st.markdown("""
        Requieren **ajuste contable** por errores u omisiones:
        - Comisiones bancarias
        - Impuestos y percepciones
        - Cheques rechazados no registrados
        - Débitos automáticos
        - Acreditaciones mal cargadas
        """)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Subir 'Libro Banco' (Excel)")
    libro_file = st.file_uploader("Cargar archivo del Libro", type=["xlsx", "xls"], key="libro")
    if libro_file:
        try:
            df_preview = pd.read_excel(libro_file, header=1, nrows=3)
            st.caption("Vista previa del Libro:")
            st.dataframe(df_preview, use_container_width=True)
        except Exception:
            st.error("Error al leer la vista previa del Libro.")

with col2:
    st.subheader("2. Subir 'Extracto Bancario' (Excel)")
    extracto_file = st.file_uploader("Cargar archivo del Banco", type=["xlsx", "xls"], key="extracto")
    if extracto_file:
        try:
            df_preview_b = pd.read_excel(extracto_file, nrows=3)
            st.caption("Vista previa del Extracto:")
            st.dataframe(df_preview_b, use_container_width=True)
        except Exception:
            st.error("Error al leer la vista previa del Banco.")

if libro_file and extracto_file:
    if st.button("🔄 Ejecutar Conciliación Bancaria"):
        with st.spinner("Procesando conciliación bancaria..."):
            try:
                output_excel, summary = process_reconciliation(libro_file, extracto_file)
                
                st.success("✅ ¡Conciliación completada!")
                
                # --- Show Summary ---
                st.divider()
                st.header("📊 Resumen de Conciliación")
                
                col_a, col_b, col_c = st.columns(3)
                with col_a:
                    st.metric(
                        "💰 Saldo Final Banco", 
                        f"${summary['saldo_final_banco']:,.2f}",
                        help="Saldo según extracto bancario"
                    )
                with col_b:
                    st.metric(
                        "📚 Saldo Final Libro", 
                        f"${summary['saldo_final_libro']:,.2f}",
                        help="Saldo según registros contables"
                    )
                with col_c:
                    diferencia = summary['diferencia_total']
                    st.metric(
                        "📊 Diferencia Total", 
                        f"${abs(diferencia):,.2f}",
                        delta=f"{'Faltante' if diferencia < 0 else 'Excedente'}",
                        delta_color="inverse" if diferencia < 0 else "normal",
                        help="Diferencia entre Libro y Banco"
                    )
                
                st.divider()
                
                # --- Categories Breakdown ---
                st.subheader("🔍 Análisis de Diferencias")
                
                col_x, col_y, col_z = st.columns(3)
                
                with col_x:
                    st.metric(
                        "✅ Items Coincidentes",
                        summary['items_coinciden'],
                        help="Transacciones que coinciden en ambos registros"
                    )
                
                with col_y:
                    st.metric(
                        "🕐 Diferencias Temporales",
                        summary['diferencias_temporales_count'],
                        f"${summary['diferencias_temporales_monto']:,.2f}",
                        help="Se ajustan sin asiento contable"
                    )
                
                with col_z:
                    st.metric(
                        "⚠️ Diferencias Permanentes",
                        summary['diferencias_permanentes_count'],
                        f"${summary['diferencias_permanentes_monto']:,.2f}",
                        help="Requieren ajuste contable"
                    )
                
                st.divider()
                
                # --- Explanation ---
                st.subheader("💡 ¿Qué se hizo?")
                st.info(f"""
                **Proceso de Conciliación:**
                
                1. **Coincidencias**: Se identificaron **{summary['items_coinciden']}** transacciones que coinciden en ambos registros:
                   - 🧾 **Por Cheque**: {summary.get('matches_cheque', 0)} items
                   - 🆔 **Por CUIT + Monto**: {summary.get('matches_cuit', 0)} items
                   - 🗓️ **Por Monto + Fecha**: {summary.get('matches_fuzzy', 0)} items
                
                2. **Diferencias Temporales** ({summary['diferencias_temporales_count']} items):
                   - Depósitos en tránsito (registrados en Libro, aún no acreditados por el Banco)
                   - Cheques pendientes de presentación
                   - **Ajuste**: Sin necesidad de asiento contable, se regularán con el tiempo
                
                3. **Diferencias Permanentes** ({summary['diferencias_permanentes_count']} items):
                   - Comisiones bancarias, impuestos, débitos automáticos
                   - Acreditaciones u omisiones de registro
                   - **Ajuste**: Requieren asiento contable para corregir
                
                4. **Resultado**: La diferencia total de **${abs(summary['diferencia_total']):,.2f}** se explica por la suma 
                   de diferencias temporales y permanentes.
                """)
                
                if abs(summary['diferencia_total']) < 1.0:
                    st.success("✨ ¡Perfecta conciliación! La diferencia es menor a $1.00")
                elif abs(summary['diferencia_total']) > 10000:
                    st.warning("⚠️ Diferencia significativa detectada. Revisa el detalle en el reporte Excel.")
                
                # --- Download Button ---
                st.divider()
                st.subheader("📥 Descargar Reporte")
                st.markdown("""
                El Excel incluye:
                - **Conciliación Bancaria**: Estado de conciliación con saldos acumulados
                - **Items Coincidentes**: Transacciones que coinciden
                - **Diferencias Temporales**: Detalle de ajustes sin asiento contable
                - **Diferencias Permanentes**: Detalle de ajustes con asiento contable
                - **Resumen**: Métricas y totales
                """)
                
                st.download_button(
                    label="📊 Descargar Reporte Completo (XLSX)",
                    data=output_excel,
                    file_name="Conciliacion_Bancaria.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Error durante la conciliación: {e}")
                st.exception(e)

else:
    st.info("👆 Por favor sube ambos archivos para comenzar la conciliación bancaria.")
