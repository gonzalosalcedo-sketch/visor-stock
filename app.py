import streamlit as st
import pandas as pd
import re
import numpy as np
import io
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA (MODO MODERNO) ---
st.set_page_config(
    page_title="Gestor de Stock de Cañerías",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS para modernizar
st.markdown("""
<style>
    .metric-card {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    }
    .stDataFrame { border: 1px solid #e0e0e0; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR: CONFIGURACIÓN Y CARGA ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2821/2821637.png", width=60) # Icono generico
    st.title("Panel de Control")
    
    st.subheader("1. Base de Datos (ASME)")
    archivo_asme = "base_datos.xlsm"
    uploaded_asme = None
    use_local_asme = False
    
    try:
        f = open(archivo_asme, "r")
        st.success(f"✅ BD Interna: {archivo_asme}")
        use_local_asme = True
    except FileNotFoundError:
        st.warning("⚠️ Carga 'base_datos.xlsm'")
        uploaded_asme = st.file_uploader("Subir Tabla ASME", type=["xlsx", "xlsm"])

    st.markdown("---")
    st.subheader("2. Archivo del Proveedor")
    uploaded_file = st.file_uploader("Sube el Excel semanal", type=["xlsx"])
    
    # Selector de fecha (Por defecto HOY)
    fecha_stock = st.date_input("Fecha del Stock", datetime.date.today())

    st.markdown("---")
    st.subheader("3. Precios (USD/Kg)")
    # Editor de precios en sidebar
    data_precios = {
        "Min (mm)": [1.0, 6.01, 9.01],
        "Max (mm)": [6.0, 9.0, 999.0],
        "Piso": [0.90, 1.00, 1.15],
        "Techo": [1.15, 1.25, 1.50]
    }
    df_precios_editor = st.data_editor(
        pd.DataFrame(data_precios), 
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
    )

# --- FUNCIONES ROBUSTAS ---
def limpiar_numero(valor):
    """Convierte texto sucio (ej: 1.000,50) a float (1000.50)"""
    if pd.isna(valor): return 0.0
    val_str = str(valor).strip()
    if val_str == '-' or val_str == '': return 0.0
    
    # Si tiene puntos y comas, asumimos formato europeo (1.000,00)
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str: # Solo comas (10,5)
        val_str = val_str.replace(',', '.')
    
    # Extraer solo numeros y punto
    val_str = re.sub(r'[^\d.]', '', val_str)
    try:
        return float(val_str)
    except:
        return 0.0

def cargar_asme_engine(file):
    try:
        df_raw = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=None, engine='openpyxl')
        fila_header = -1
        for i in range(15):
            fila_texto = df_raw.iloc[i].astype(str).str.lower().tolist()
            if any("scheduleno" in x.replace(" ","") for x in fila_texto):
                fila_header = i
                break
        if fila_header == -1: fila_header = 2 
        
        df_asme = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=fila_header, engine='openpyxl')
        
        # Francotirador de columnas
        def buscar(keys, cols):
            for c in cols:
                if all(k.lower() in str(c).lower() for k in keys): return c
            return None
            
        c_dn = buscar(['NPS'], df_asme.columns) or buscar(['Diameter', 'in.'], df_asme.columns)
        c_sch = buscar(['ScheduleNo'], df_asme.columns) or buscar(['Schedule'], df_asme.columns)
        c_esp = buscar(['Thickness', 'mm'], df_asme.columns)
        c_peso = buscar(['Mass', 'kg/m'], df_asme.columns)
        
        if not all([c_dn, c_sch, c_esp, c_peso]): return None, "Faltan columnas clave en ASME"
        
        df_out = df_asme[[c_dn, c_sch, c_esp, c_peso]].copy()
        df_out.columns = ['DN', 'Schedule', 'Espesor_mm', 'Peso_Kg_m']
        df_out = df_out.dropna(subset=['DN'])
        # Limpieza agresiva de strings
        df_out['Schedule'] = df_out['Schedule'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_out['DN'] = df_out['DN'].astype(str).str.replace('.0', '', regex=False).str.strip()
        return df_out, "OK"
    except Exception as e: return None, str(e)

def procesar_regex(texto):
    d = {"DN": None, "SCH": None, "Espesor": None, "Tira_Mts": 6.0} 
    t = str(texto)
    # Regex
    m_dn = re.search(r'(\d+\s?\d*/?\d*)"', t)
    if m_dn: d["DN"] = m_dn.group(1)
    
    m_mts = re.search(r'(\d+(?:[.,]\d+)?)\s*Mts', t, re.IGNORECASE)
    if m_mts: d["Tira_Mts"] = float(m_mts.group(1).replace(',', '.'))
    
    m_sch = re.search(r'Sch\s?(\d+|STD|XS)', t, re.IGNORECASE)
    if m_sch: d["SCH"] = m_sch.group(1)
    
    if not d["SCH"]:
        m_esp = re.search(r'[xX]\s*(\d+[,.]\d+)\s*[xX]', t)
        if m_esp: d["Espesor"] = float(m_esp.group(1).replace(',', '.'))
    return pd.Series(d)

def asignar_tipo(desc):
    d = str(desc).lower()
    if "iso" in d and "negra" in d: return "3. CCC ISO Negra"
    if "iso" in d and "galva" in d: return "5. CCC ISO Galvanizada"
    if "astm" in d and "a-53" in d: return "1. CCC ASTM A53 / API 5L"
    if "galva" in d: return "7. CCC ASTM A-53 Galva"
    return "⚠️ MANUAL"

def get_precio(esp, df_conf):
    if pd.isna(esp) or esp == 0: return 0.0, 0.0
    for _, r in df_conf.iterrows():
        if r['Min (mm)'] <= esp <= r['Max (mm)']:
            return r['Piso'], r['Techo']
    return 0.0, 0.0

# --- PROCESO PRINCIPAL ---
st.title("🏭 Visor de Stock Inteligente")

if uploaded_file:
    # --- 1. PRE-CARGA Y SELECCIÓN DE COLUMNAS (SOLUCIÓN A CEROS) ---
    df_preview = pd.read_excel(uploaded_file)
    cols_excel = df_preview.columns.tolist()
    
    # Intentamos adivinar columnas
    def encontrar_idx(lista, pistas):
        for i, col in enumerate(lista):
            if any(p in str(col).lower() for p in pistas): return i
        return 0

    idx_desc = encontrar_idx(cols_excel, ['descrip', 'articulo'])
    idx_cant = encontrar_idx(cols_excel, ['dispo', 'cant', 'stock', 'saldo'])

    with st.expander("🛠️ Ajuste de Columnas (Si ves ceros, revisa esto)", expanded=False):
        c1, c2 = st.columns(2)
        col_desc = c1.selectbox("Columna Descripción:", cols_excel, index=idx_desc)
        col_cant = c2.selectbox("Columna Cantidad (Disponible):", cols_excel, index=idx_cant)

    # --- BOTÓN DE PROCESAR ---
    if st.session_state.get('data_procesada') is None:
        if st.button("🚀 PROCESAR DATOS", type="primary", use_container_width=True):
            fuente_asme = archivo_asme if use_local_asme else uploaded_asme
            if not fuente_asme: st.error("Falta ASME"); st.stop()
            
            df_asme, err = cargar_asme_engine(fuente_asme)
            if df_asme is None: st.error(err); st.stop()
            
            with st.spinner("Analizando y calculando..."):
                # Proceso
                df_work = df_preview.copy()
                datos_ext = df_work[col_desc].apply(procesar_regex)
                df_final = pd.concat([df_work, datos_ext], axis=1)
                
                # Cruce ASME
                pesos, esps = [], []
                for _, r in df_final.iterrows():
                    dn, sch, esp = str(r['DN']).strip(), r['SCH'], r['Espesor']
                    sch = str(sch).strip() if pd.notna(sch) else None
                    match = pd.DataFrame()
                    if dn != "None":
                        base = df_asme[df_asme['DN'] == dn]
                        if not base.empty:
                            if sch: match = base[base['Schedule'] == sch]
                            elif pd.notna(esp):
                                be = pd.to_numeric(base['Espesor_mm'], errors='coerce')
                                match = base[np.isclose(be, esp, atol=0.2)]
                    if not match.empty:
                        pesos.append(match.iloc[0]['Peso_Kg_m'])
                        esps.append(match.iloc[0]['Espesor_mm'])
                    else:
                        pesos.append(0); esps.append(esp if pd.notna(esp) else 0)
                
                df_final['Peso_Unitario_Kg_m'] = pd.to_numeric(pesos, errors='coerce').fillna(0)
                df_final['Espesor_Final'] = pd.to_numeric(esps, errors='coerce').fillna(0)
                df_final['TIPO_SISTEMA'] = df_final[col_desc].apply(asignar_tipo)
                
                # CÁLCULOS (Solución Ceros: Limpieza explicita)
                df_final['Cant_Clean'] = df_final[col_cant].apply(limpiar_numero)
                df_final['Stock_Total_Mts'] = df_final['Cant_Clean'] * df_final['Tira_Mts']
                df_final['Stock_Total_Kgs'] = df_final['Stock_Total_Mts'] * df_final['Peso_Unitario_Kg_m']
                
                # Precios
                precios = df_final['Espesor_Final'].apply(lambda x: get_precio(x, df_precios_editor))
                df_final['Px_Piso_Kg_USD'] = precios.apply(lambda x: x[0])
                df_final['Px_Techo_Kg_USD'] = precios.apply(lambda x: x[1])
                df_final['Px_Piso_Mt_USD'] = df_final['Px_Piso_Kg_USD'] * df_final['Peso_Unitario_Kg_m']
                df_final['Px_Techo_Mt_USD'] = df_final['Px_Techo_Kg_USD'] * df_final['Peso_Unitario_Kg_m']
                
                # Columnas finales ordenadas
                cols_view = [col_desc, 'TIPO_SISTEMA', 'DN', 'SCH', 'Espesor_Final', 'Tira_Mts',
                             'Stock_Total_Mts', 'Stock_Total_Kgs', 
                             'Px_Piso_Kg_USD', 'Px_Techo_Kg_USD', 'Px_Piso_Mt_USD']
                
                st.session_state['data_procesada'] = df_final[cols_view]
                st.rerun()

# --- VISUALIZACIÓN Y FILTROS ---
if st.session_state.get('data_procesada') is not None:
    df = st.session_state['data_procesada']
    
    st.markdown("### 🔎 Filtros Dinámicos")
    
    # Fila 1 de Filtros
    c1, c2, c3, c4 = st.columns(4)
    tipos_disp = ["Todos"] + sorted(list(df['TIPO_SISTEMA'].unique()))
    dn_disp = ["Todos"] + sorted(list(df['DN'].dropna().unique()))
    tira_disp = ["Todas"] + sorted(list(df['Tira_Mts'].astype(str).unique()))
    
    f_tipo = c1.selectbox("Tipo:", tipos_disp)
    f_dn = c2.selectbox("DN:", dn_disp)
    f_tira = c3.selectbox("Tira (Mts):", tira_disp)
    
    # Fila 2 de Filtros (SCH y Espesor)
    c5, c6, c7 = st.columns([1,1,2])
    sch_disp = ["Todos"] + sorted(list(df['SCH'].dropna().astype(str).unique()))
    f_sch = c5.selectbox("SCH:", sch_disp)
    
    # Aplicar Filtros
    df_filt = df.copy()
    if f_tipo != "Todos": df_filt = df_filt[df_filt['TIPO_SISTEMA'] == f_tipo]
    if f_dn != "Todos": df_filt = df_filt[df_filt['DN'] == f_dn]
    if f_tira != "Todas": df_filt = df_filt[df_filt['Tira_Mts'].astype(str) == f_tira]
    if f_sch != "Todos": df_filt = df_filt[df_filt['SCH'].astype(str) == f_sch]

    st.divider()

    # --- KPIS MODERNOS ---
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📅 Fecha Stock", str(fecha_stock))
    k2.metric("📦 Tubos Listados", len(df_filt))
    k3.metric("📏 Total Metros", f"{df_filt['Stock_Total_Mts'].sum():,.0f} m")
    k4.metric("⚖️ Total Kilos", f"{df_filt['Stock_Total_Kgs'].sum():,.0f} kg")

    st.divider()
    
    # --- TABLA EDITABLE (Solución a "Manual") ---
    st.markdown("### 📝 Planilla de Stock (Editable)")
    st.info("💡 Puedes editar la columna **TIPO_SISTEMA** directamente aquí abajo.")

    # Configuración de columnas para el editor
    column_config = {
        "Stock_Total_Kgs": st.column_config.NumberColumn("Stock Kg", format="%.0f kg"),
        "Stock_Total_Mts": st.column_config.NumberColumn("Stock Mts", format="%.0f m"),
        "Px_Piso_Kg_USD": st.column_config.NumberColumn("Piso USD/kg", format="$ %.2f"),
        "Px_Techo_Kg_USD": st.column_config.NumberColumn("Techo USD/kg", format="$ %.2f"),
        "TIPO_SISTEMA": st.column_config.SelectboxColumn(
            "Tipo Sistema",
            options=[
                "1. CCC ASTM A53 / API 5L",
                "2. CSC ASTM A106 / API 5L",
                "3. CCC ISO Negra",
                "4. CCC A795 Negra",
                "5. CCC ISO Galvanizada",
                "7. CCC ASTM A-53 Galva",
                "⚠️ MANUAL"
            ],
            required=True
        )
    }

    df_editado = st.data_editor(
        df_filt,
        column_config=column_config,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        key="editor_datos"
    )

    # --- DESCARGA ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_editado.to_excel(writer, index=False, sheet_name=f'Stock_{fecha_stock}')
    
    st.download_button(
        label="📥 Descargar Excel Filtrado y Editado",
        data=buffer.getvalue(),
        file_name=f"Stock_Procesado_{fecha_stock}.xlsx",
        mime="application/vnd.ms-excel",
        type="primary"
    )

    # Botón de Reset
    if st.button("🔄 Cargar Nuevo Archivo"):
        st.session_state['data_procesada'] = None
        st.rerun()

else:
    st.info("👈 Sube el archivo del proveedor en el menú lateral para comenzar.")