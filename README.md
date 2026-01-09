# 🏦 Tuchi: Conciliador de Cheques

**Tuchi** es una herramienta de automatización diseñada para facilitar la conciliación bancaria entre los registros internos corporativos (Libro) y los extractos bancarios (ej. Banco Galicia).

## 🚀 Funcionalidades

- **Extracción Automática**: Detecta números de cheques dentro de descripciones complejas usando expresiones regulares.
- **Cruce de Datos Inteligente**: Maneja registros que contienen múltiples cheques agrupados, buscando coincidencias individuales en el banco.
- **Validación de Montos**: Compara la suma de los créditos bancarios contra el monto registrado en el libro, identificando discrepancias de centavos.
- **Reportes en Excel**: Genera un archivo `.xlsx` detallado con el estado de cada transacción (Conciliado OK, Diferencia de Monto, No Encontrado).
- **Interfaz Intuitiva**: Construido con Streamlit para una experiencia de usuario fluida y visual.

## 🛠️ Instalación

1. **Clonar el repositorio**:
   ```bash
   git clone https://github.com/TU_USUARIO/Tuchi.git
   cd Tuchi
   ```

2. **Crear entorno virtual**:
   ```bash
   python -m venv venv
   source venv/bin/scripts/activate  # En Windows: venv\Scripts\activate
   ```

3. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Uso

1. Ejecuta la aplicación:
   ```bash
   streamlit run app.py
   ```
2. Sube el archivo de **Libro** (asegúrate de que los cheques estén entre paréntesis, ej: `(123456)`).
3. Sube el archivo de **Extracto Bancario**.
4. Haz clic en **Ejecutar Conciliación**.
5. Descarga el reporte generado.

## 📁 Estructura del Proyecto

- `app.py`: Interfaz de usuario y lógica de presentación.
- `logic.py`: Motor de procesamiento y lógica de conciliación.
- `requirements.txt`: Librerías necesarias (Pandas, Streamlit, Openpyxl, XlsxWriter).

---
Desarrollado con ❤️ para simplificar las finanzas.
