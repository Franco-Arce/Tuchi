
import streamlit as st
import pandas as pd
from logic import process_reconciliation
import io

st.set_page_config(page_title="Conciliador de Cheques", layout="wide")

st.title("🏦 Conciliación Automática de Cheques")
st.markdown("""
Esta herramienta cruza los datos del **Libro** (Tu sistema) con el **Extracto Bancario** (Galicia) utilizando los números de cheques.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Subir 'Libro' (Excel)")
    libro_file = st.file_uploader("Cargar archivo del Libro", type=["xlsx", "xls"], key="libro")

with col2:
    st.subheader("2. Subir 'Extracto' (Excel)")
    extracto_file = st.file_uploader("Cargar archivo del Banco", type=["xlsx", "xls"], key="extracto")

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

