
import streamlit as st
import pandas as pd
from logic import process_reconciliation
import io

st.set_page_config(
    page_title="Tuchi | Conciliador de Cheques",
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

st.title("🏦 Tuchi: Conciliación Automática")
st.markdown("""
Esta herramienta cruza los datos del **Libro** con el **Extracto Bancario** utilizando los números de cheques.
""")

# --- Sidebar Instructions ---
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/000000/bank.png", width=80)
    st.title("Instrucciones")
    st.info("""
    1. **Sube el Libro**: El Excel de tu sistema con los cheques entre paréntesis, ej: `(123456)`.
    2. **Sube el Extracto**: El Excel del banco (Galicia).
    3. **Previsualiza**: Asegúrate de que los datos se lean correctamente.
    4. **Procesa**: Cruza la información y descarga el reporte.
    """)
    
    st.divider()
    st.subheader("📥 ¿No tienes el formato?")
    
    # Simple template generator
    template_data = io.BytesIO()
    with pd.ExcelWriter(template_data, engine='xlsxwriter') as writer:
        # Sample Libro
        pd.DataFrame({
            'Fecha Pago ': ['2026-01-01'],
            'Concepto': ['Cheques de terceros (123)(456)'],
            'Ingreso': ['1.000,00']
        }).to_excel(writer, sheet_name='Libro', index=False, startrow=1)
        # Sample Extracto
        pd.DataFrame({
            'Fecha': ['2026-01-02', '2026-01-02'],
            'Numero de Comprobante': [123, 456],
            'Creditos': [400, 600],
            'Descripcion': ['Deposito cheque', 'Deposito cheque']
        }).to_excel(writer, sheet_name='Banco', index=False)
    template_data.seek(0)
    
    st.download_button(
        "Descargar Plantilla de Ejemplo",
        data=template_data,
        file_name="Plantilla_Tuchi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Subir 'Libro' (Excel)")
    libro_file = st.file_uploader("Cargar archivo del Libro", type=["xlsx", "xls"], key="libro")
    if libro_file:
        try:
            # Quick preview
            df_preview = pd.read_excel(libro_file, header=1, nrows=3)
            st.caption("Vista previa del Libro:")
            st.dataframe(df_preview, use_container_width=True)
        except Exception:
            st.error("Error al leer la vista previa del Libro.")

with col2:
    st.subheader("2. Subir 'Extracto' (Excel)")
    extracto_file = st.file_uploader("Cargar archivo del Banco", type=["xlsx", "xls"], key="extracto")
    if extracto_file:
        try:
            # Quick preview
            df_preview_b = pd.read_excel(extracto_file, nrows=3)
            st.caption("Vista previa del Extracto:")
            st.dataframe(df_preview_b, use_container_width=True)
        except Exception:
            st.error("Error al leer la vista previa del Banco.")

if libro_file and extracto_file:
    if st.button("🔄 Ejecutar Conciliación"):
        with st.spinner("Procesando archivos y buscando coincidencias..."):
            try:
                # Process only if files are uploaded
                output_excel, summary = process_reconciliation(libro_file, extracto_file)
                
                st.success("✅ ¡Proceso completado!")
                
                # --- Show Explanations & Stats ---
                st.divider()
                st.header("📊 Resumen del Análisis")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("Registros en Libro Procesados", summary['total_registros_procesados'])
                m2.metric("Con Cheques Identificados", summary['registros_con_cheques'])
                m3.metric("Conciliados Exitosamente", summary['registros_conciliados_ok'])
                
                st.divider()
                st.subheader("💡 ¿Qué se hizo?")
                st.info(f"""
                1. **Lectura**: Se procesaron **{summary['total_registros_procesados']}** líneas del Libro.
                2. **Extracción**: Se detectaron grupos de cheques en **{summary['registros_con_cheques']}** registros utilizando los números entre paréntesis `(xxxx)`.
                3. **Cruce**: Cada cheque individual fue buscado en el Extracto Bancario.
                4. **Validación**: Se sumaron los montos del banco para cada grupo y se compararon con tu registro original.
                   - **{summary['registros_conciliados_ok']}** registros coincidieron exactamente (o con diferencia < $1).
                   - Monto Total Analizado: **${summary['monto_total_libro_analizado']:,.2f}**
                   - Monto Total Conciliado: **${summary['monto_total_conciliado']:,.2f}**
                """)
                
                if summary['diferencia_global'] != 0:
                    st.warning(f"⚠️ Existe una diferencia global no conciliada de: ${summary['diferencia_global']:,.2f}")
                
                # --- Download Button ---
                st.subheader("📥 Descargar Resultados")
                st.markdown("Descarga el Excel con el detalle fila por fila, incluyendo las columnas 'Estado' y 'Diferencia'.")
                
                st.download_button(
                    label="Descargar Reporte Completo (XLSX)",
                    data=output_excel,
                    file_name="Resultado_Conciliacion.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Ocurrió un error :/: {e}")
                st.exception(e)

else:
    st.info("👆 Por favor sube ambos archivos para comenzar.")

