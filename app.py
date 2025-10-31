import streamlit as st
import pandas as pd
import subprocess
import os
import datetime as dt
import io
import plotly.express as px


# --- Sección para actualizar la hoja principal ---
st.header("Actualizar Hoja Principal")
uploaded_file = st.file_uploader("Sube la nueva Hoja principal.xlsm", type=["xlsm"])

if uploaded_file:
    # Guardar el archivo subido
    with open("Hoja principal.xlsm", "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("Archivo actualizado correctamente.")

    # Ejecutar el script para generar Resultados.xlsx
    with st.spinner("Procesando datos..."):
        subprocess.run(["python", "final.py"])
    st.success("Datos procesados. Recargando dashboard...")

# --- Cargar datos procesados ---
@st.cache_data
def load_data():
    if os.path.exists("Resultados.xlsx"):
        df = pd.read_excel("Resultados.xlsx", sheet_name="resultado final")
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        if "fecha_expiracion" in df.columns:
            df["fecha_expiracion"] = pd.to_datetime(df["fecha_expiracion"], errors="coerce")
        today = pd.Timestamp.today()
        if "status" not in df.columns:
            df["status"] = df["fecha_expiracion"].apply(
                lambda x: "VENCIDO" if x < today
                else "POR VENCER" if x < today + pd.Timedelta(days=30)
                else "VIGENTE"
            )
        return df
    else:
        return pd.DataFrame()

df = load_data()

if df.empty:
    st.warning("No hay datos procesados aún. Sube la hoja principal para comenzar.")
    st.stop()

# --- Sidebar de filtros ---
st.sidebar.header("Filtros")
sqe_list = sorted(df["sqe"].dropna().unique())
selected_sqe = st.sidebar.selectbox("Selecciona un SQE:", sqe_list)
proveedores = sorted(df.loc[df["sqe"] == selected_sqe, "supplier_name"].dropna().unique())
selected_proveedor = st.sidebar.selectbox("Filtra por proveedor (opcional):", ["Todos"] + list(proveedores))
status_options = ["Todos", "VENCIDO", "POR VENCER", "VIGENTE"]
selected_status = st.sidebar.radio("Estatus:", status_options)

# --- Filtrado dinámico ---
filtered_df = df[df["sqe"] == selected_sqe]
if selected_proveedor != "Todos":
    filtered_df = filtered_df[filtered_df["supplier_name"] == selected_proveedor]
if selected_status != "Todos":
    filtered_df = filtered_df[filtered_df["status"] == selected_status]

st.image("sensaa.png", width=450)

# --- Tabs principales ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Dashboard principal", "Piezas por SQE", "Piezas sin proveedor", "Piezas sin SQE", "Bill of Material"
])

# Tab 1: Dashboard principal
with tab1:
    st.title("📊 Dashboard General")
    st.subheader("Distribución total de piezas por estatus")
    status_counts_total = df["status"].value_counts().reset_index()
    status_counts_total.columns = ["status", "count"]
    fig_total = px.pie(status_counts_total, names="status", values="count", title="Distribución total de estatus")
    st.plotly_chart(fig_total, use_container_width=True)
    st.divider()
    st.subheader("📋 Tabla completa de todas las piezas")
    st.dataframe(df, use_container_width=True)

# Tab 2: Piezas por SQE
with tab2:
    st.title("📊 Piezas por SQE")
    st.subheader(f"Responsable: {selected_sqe}")
    col1, col2, col3 = st.columns(3)
    col1.metric("🔴 Vencidos", (filtered_df["status"] == "VENCIDO").sum())
    col2.metric("🟡 Por vencer", (filtered_df["status"] == "POR VENCER").sum())
    col3.metric("🟢 Vigentes", (filtered_df["status"] == "VIGENTE").sum())
    st.divider()
    st.subheader("📦 Detalle de proveedores y piezas")
    st.dataframe(filtered_df, use_container_width=True)

# Tab 3: Piezas sin proveedor
with tab3:
    st.subheader("📋 Piezas sin proveedor")
    sin_proveedor = df[df["supplier_name"].isna() | (df["supplier_name"] == "")]
    st.dataframe(sin_proveedor if not sin_proveedor.empty else pd.DataFrame({"Mensaje": ["No hay piezas sin proveedor."]}))

# Tab 4: Piezas con proveedor pero sin SQE
with tab4:
    st.subheader("📋 Piezas con proveedor pero sin SQE")
    sin_sqe = df[df["supplier_name"].notna() & (df["sqe"].isna() | (df["sqe"] == ""))]
    st.dataframe(sin_sqe if not sin_sqe.empty else pd.DataFrame({"Mensaje": ["No hay piezas con proveedor pero sin SQE."]}))

# Tab 5: Bill of Material
with tab5:
    st.title("📦 Verificación por Bill of Material")
    uploaded_bom = st.file_uploader("Subir archivo Excel BOM", type=["xlsx"])
    if uploaded_bom:
        bom_df = pd.read_excel(uploaded_bom)
        st.write("Vista previa del archivo:")
        st.dataframe(bom_df.head(), use_container_width=True)
        column_selected = st.selectbox("Selecciona la columna con las piezas:", bom_df.columns)
        piezas_lista = bom_df[column_selected].dropna().unique().tolist()
        resultado = df[df["item"].isin(piezas_lista)]
        st.subheader("Resultado de la verificación:")
        st.dataframe(resultado if not resultado.empty else pd.DataFrame({"Mensaje": ["No se encontraron coincidencias."]}))

# --- Exportar resultados filtrados ---
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    filtered_df.to_excel(writer, index=False, sheet_name="Datos Filtrados")
output.seek(0)
st.download_button(
    label="💾 Descargar tabla filtrada a Excel",
    data=output,
    file_name=f"seguimiento_{selected_sqe}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)