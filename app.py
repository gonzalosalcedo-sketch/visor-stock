import streamlit as st
import pandas as pd
import re
import numpy as np
import io
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Gestor de Stock de Cañerías",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS MODERNOS ---
st.markdown("""
<style>
    /* Tarjetas KPI */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e6e6e6;
        padding: 15px 10px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    /* Encabezados */
    h1, h2, h3 {
        color: #2c3e50;
    }
    /* Tabla */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR: CONFIGURACIÓN ---
with st.sidebar:
    st.title("🎛️ Panel de Control")
    
    st.markdown("### 1. Base de Datos (ASME)")
    archivo_asme = "base_datos.xlsm"
    uploaded_asme = None
    use_local_asme = False
    
    try:
        f = open(archivo_asme, "r")
        st.success(f"✅ BD Interna Detectada")
        use_local_asme = True
    except FileNotFoundError:
        st.warning("⚠️ No se detectó BD local")
        uploaded_asme = st.file_uploader("Subir Tabla ASME", type=["xlsx", "xlsm"])

    st.markdown("---")
    st.markdown("### 2. Archivo Proveedor")
    uploaded_file = st.file_uploader("Sube el Excel semanal", type=["xlsx"])
    
    # Selector de fecha
    fecha_stock = st.date_input("Fecha del Reporte", datetime.date.today())

    st.markdown("---")
    st.markdown("### 3. Config. Precios (USD/Kg)")
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

# --- FUNCIONES DE LÓGICA (BLINDADAS) ---

def cargar_asme_engine(file):
    try:
        # Lectura inteligente del encabezado
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
        
        # Limpieza agresiva para evitar errores de cruce
        df_out['Schedule'] = df_out['Schedule'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_out['DN'] = df_out['DN'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # Asegurar que peso y espesor sean numéricos en la base
        df_out['Peso_Kg_m'] = pd.to_numeric(df_out['Peso_Kg_m'], errors='coerce')
        df_out['Espesor_mm'] = pd.to_numeric(df_out['Espesor_mm'], errors='coerce')
        
        return df_out, "OK"
    except Exception as e: return None, str(e)

def procesar_regex(texto):
    d = {"DN": None, "SCH": None, "Espesor": None, "Tira_Mts": 6.0} 
    t = str(texto)
    # Regex DN
    m_dn = re.search(r'(\d+\s?\d*/?\d*)"', t)
    if m_dn: d["DN"] = m_dn.group(1)
    
    # Regex Tira
    m_mts = re.search(r'(\d+(?:[.,]\d+)?)\s*Mts', t, re.IGNORECASE)
    if m_mts: d["Tira_Mts"] = float(m_mts.group(1).replace(',', '.'))
    
    # Regex SCH
    m_sch = re.search(r'Sch\s?(\d+|STD|XS)', t, re.IGNORECASE)
    if m_sch: d["SCH"] = m_sch.group(1)
    
    # Regex Espesor (Solo si no hay SCH)
    if not d["SCH"]:
        m_esp = re.search(r'[xX]\s*(\d+[,.]\d+)\s*[xX]', t)
        if m_esp: d["Espesor"] = float(m_esp.group(1).replace(',', '.'))
    return pd.Series(d)

def asignar_tipo_inteligente(desc):
    d = str(desc).lower()
    # Lógica de asignación basada en tus reglas
    if "iso" in d and "negra" in d: return "3. CCC ISO Negra"
    if "iso" in d and "galva" in d: return "5. CCC ISO Galvanizada"
    if "astm" in d and "a-53" in d: return "1. CCC ASTM A53 / API 5L"
    if "galva" in d and "bsp" in d: return "8. CCC ASTM A-53 Galva R65 H/BSP"
    if "galva" in d: return "7. CCC ASTM A-53 Galva"
    if "a795" in d and "ranurada" in d: return "4a. CCC A795 Negra - Ranurada"
    if "a795" in d: return "4. CCC A795 Negra"
    return "⚠️ MANUAL"

def get_precio(esp, df_conf):
    if pd.isna(esp) or esp == 0: return 0.0, 0.0
    for _, r in df_conf.iterrows():
        if r['Min (mm)'] <= esp <= r['Max (mm)']:
            return r['Piso'], r['Techo']
    return 0.0, 0.0

def safe_float(val):
    """Convierte cualquier cosa a float o devuelve 0.0"""
    try:
        if pd.isna(val): return 0.0
        return float(val)
    except:
        return 0.0

# --- INTERFAZ PRINCIPAL ---
st.title("🏭 Visor de Stock Inteligente")

if uploaded_file:
    # --- 1. PRE-CARGA Y SELECCIÓN DE COLUMNAS ---
    df_preview = pd.read_excel(uploaded_file)
    cols_excel = df_preview.columns.tolist()
    
    # Auto-detectar columnas
    idx_desc = next((i for i, c in enumerate(cols_excel) if 'descrip' in str(c).lower()), 0)
    idx_cant = next((i for i, c in enumerate(cols_excel) if any(x in str(c).lower() for x in ['dispo', 'cant', 'stock'])), 2)

    with st.expander("🛠️ Ajuste de Columnas (Revisa si ves ceros)", expanded=False):
        c1, c2 = st.columns(2)
        col_desc = c1.selectbox("Columna Descripción:", cols_excel, index=idx_desc)
        col_cant = c2.selectbox("Columna Cantidad:", cols_excel, index=idx_cant)

    # --- BOTÓN DE PROCESAR ---
    if st.session_state.get('data_procesada') is None:
        if st.button("🚀 PROCESAR DATOS", type="primary", use_container_width=True):
            fuente_asme = archivo_asme if use_local_asme else uploaded_asme
            if not fuente_asme: st.error("Falta ASME"); st.stop()
            
            df_asme, err = cargar_asme_engine(fuente_asme)
            if df_asme is None: st.error(err); st.stop()
            
            with st.spinner("Procesando lógica de negocio..."):
                # 1. Regex
                df_work = df_preview.copy()
                datos_ext = df_work[col_desc].apply(procesar_regex)
                df_final = pd.concat([df_work, datos_ext], axis=1)
                
                # 2. Cruce ASME (FIX DEL ERROR ATTRIBUTEERROR)
                pesos_vals = []
                esps_vals = []
                
                for _, r in df_final.iterrows():
                    dn = str(r['DN']).strip()
                    sch = str(r['SCH']).strip() if pd.notna(r['SCH']) else None
                    esp = r['Espesor']
                    
                    match = pd.DataFrame()
                    if dn != "None":
                        base = df_asme[df_asme['DN'] == dn]
                        if not base.empty:
                            if sch: match = base[base['Schedule'] == sch]
                            elif pd.notna(esp):
                                be = pd.to_numeric(base['Espesor_mm'], errors='coerce')
                                match = base[np.isclose(be, esp, atol=0.2)]
                    
                    # Extracción segura de valores
                    if not match.empty:
                        # Usamos .values[0] y safe_float para evitar errores de objetos pandas
                        p_val = safe_float(match['Peso_Kg_m'].values[0])
                        e_val = safe_float(match['Espesor_mm'].values[0])
                        pesos_vals.append(p_val)
                        esps_vals.append(e_val)
                    else:
                        pesos_vals.append(0.0)
                        esps_vals.append(safe_float(esp))
                
                # Asignación limpia
                df_final['Peso_Unitario_Kg_m'] = pesos_vals
                df_final['Espesor_Final'] = esps_vals
                
                # 3. Clasificación
                df_final['TIPO_SISTEMA'] = df_final[col_desc].apply(asignar_tipo_inteligente)
                
                # 4. Cálculos Matemáticos
                # Limpieza de la columna cantidad (quita letras, puntos de mil, etc)
                def clean_cant(x):
                    s = str(x).replace('.', '').replace(',', '.') # Asume formato 1.000,00
                    return safe_float(re.sub(r'[^\d.]', '', s))
                
                df_final['Cant_Clean'] = df_final[col_cant].apply(clean_cant)
                df_final['Stock_Total_Mts'] = df_final['Cant_Clean'] * df_final['Tira_Mts']
                df_final['Stock_Total_Kgs'] = df_final['Stock_Total_Mts'] * df_final['Peso_Unitario_Kg_m']
                
                # 5. Precios
                precios = df_final['Espesor_Final'].apply(lambda x: get_precio(x, df_precios_editor))
                df_final['Px_Piso_Kg_USD'] = precios.apply(lambda x: x[0])
                df_final['Px_Techo_Kg_USD'] = precios.apply(lambda x: x[1])
                df_final['Px_Piso_Mt_USD'] = df_final['Px_Piso_Kg_USD'] * df_final['Peso_Unitario_Kg_m']
                df_final['Px_Techo_Mt_USD'] = df_final['Px_Techo_Kg_USD'] * df_final['Peso_Unitario_Kg_m']
                
                # Guardar resultado
                cols_view = [col_desc, 'TIPO_SISTEMA', 'DN', 'SCH', 'Espesor_Final', 'Tira_Mts',
                             'Stock_Total_Mts', 'Stock_Total_Kgs', 
                             'Px_Piso_Kg_USD', 'Px_Techo_Kg_USD', 'Px_Piso_Mt_USD']
                
                st.session_state['data_procesada'] = df_final[cols_view]
                st.rerun()

# --- VISUALIZACIÓN Y FILTROS ---
if st.session_state.get('data_procesada') is not None:
    df = st.session_state['data_procesada']
    
    st.markdown("### 🔎 Filtros de Búsqueda")
    
    # Fila 1 Filtros
    c1, c2, c3, c4 = st.columns(4)
    tipos_disp = ["Todos"] + sorted(list(df['TIPO_SISTEMA'].unique()))
    dn_disp = ["Todos"] + sorted(list(df['DN'].dropna().unique()))
    tira_disp = ["Todas"] + sorted(list(df['Tira_Mts'].astype(str).unique()))
    
    f_tipo = c1.selectbox("Tipo:", tipos_disp)
    f_dn = c2.selectbox("DN:", dn_disp)
    f_tira = c3.selectbox("Largo Tira:", tira_disp)
    
    # Fila 2 Filtros (SCH y ESPESOR NUEVO)
    c5, c6, c7 = st.columns([1,1,2])
    sch_disp = ["Todos"] + sorted(list(df['SCH'].dropna().astype(str).unique()))
    # Crear bins o lista de espesores únicos para filtrar
    esp_disp = ["Todos"] + sorted(list(df['Espesor_Final'].unique()))
    
    f_sch = c5.selectbox("SCH:", sch_disp)
    f_esp = c6.selectbox("Espesor (mm):", esp_disp)
    
    # Aplicar Filtros
    df_filt = df.copy()
    if f_tipo != "Todos": df_filt = df_filt[df_filt['TIPO_SISTEMA'] == f_tipo]
    if f_dn != "Todos": df_filt = df_filt[df_filt['DN'] == f_dn]
    if f_tira != "Todas": df_filt = df_filt[df_filt['Tira_Mts'].astype(str) == f_tira]
    if f_sch != "Todos": df_filt = df_filt[df_filt['SCH'].astype(str) == f_sch]
    if f_esp != "Todos": df_filt = df_filt[df_filt['Espesor_Final'] == f_esp]

    st.divider()

    # --- KPIS MODERNOS ---
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📅 Fecha Stock", str(fecha_stock))
    k2.metric("📦 Tubos Listados", len(df_filt))
    k3.metric("📏 Total Metros", f"{df_filt['Stock_Total_Mts'].sum():,.0f} m")
    k4.metric("⚖️ Total Kilos", f"{df_filt['Stock_Total_Kgs'].sum():,.0f} kg")

    st.divider()
    
    # --- TABLA EDITABLE CON LISTA CORRECTA ---
    st.markdown("### 📝 Planilla Interactiva")
    st.caption("Selecciona una celda en 'Tipo Sistema' para corregir manualmente.")

    # Opciones basadas en tu imagen
    opciones_tipo = [
        "1. CCC ASTM A53 / API 5L",
        "2. CSC ASTM A106 / API 5L",
        "3. CCC ISO Negra",
        "4. CCC A795 Negra",
        "4a. CCC A795 Negra - Ranurada",
        "5. CCC ISO Galvanizada",
        "6. CCC ASTM A-53 Gr.B Galva/H/NPT",
        "7. CCC ASTM A-53 Galva",
        "8. CCC ASTM A-53 Galva R65 H/BSP",
        "9. CCC ASTM A-53 Galva H/NPT CUPLA",
        "⚠️ MANUAL"
    ]

    column_config = {
        "Stock_Total_Kgs": st.column_config.NumberColumn("Stock Kg", format="%.0f kg"),
        "Stock_Total_Mts": st.column_config.NumberColumn("Stock Mts", format="%.0f m"),
        "Px_Piso_Kg_USD": st.column_config.NumberColumn("Piso USD/kg", format="$ %.2f"),
        "Px_Techo_Kg_USD": st.column_config.NumberColumn("Techo USD/kg", format="$ %.2f"),
        "Px_Piso_Mt_USD": st.column_config.NumberColumn("Piso USD/m", format="$ %.2f"),
        "TIPO_SISTEMA": st.column_config.SelectboxColumn(
            "Tipo Sistema",
            options=opciones_tipo,
            required=True,
            width="medium"
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
        label="📥 Descargar Excel Final",
        data=buffer.getvalue(),
        file_name=f"Stock_Procesado_{fecha_stock}.xlsx",
        mime="application/vnd.ms-excel",
        type="primary"
    )

    if st.button("🔄 Reiniciar / Cargar Nuevo"):
        st.session_state['data_procesada'] = None
        st.rerun()

else:
    st.info("👈 Sube el archivo para comenzar.")