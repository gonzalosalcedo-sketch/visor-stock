import streamlit as st
import pandas as pd
import re
import numpy as np
import io
import datetime

# --- CONFIGURACIÓN VISUAL ---
st.set_page_config(page_title="Gestor de Stock", page_icon="🏭", layout="wide")
st.markdown("""
<style>
    .metric-card { background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 0.5rem; padding: 1rem; text-align: center; }
    .stDataFrame { border: 1px solid #e0e0e0; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("🎛️ Panel de Control")
    
    # 1. ASME
    st.subheader("1. Base de Datos")
    archivo_asme = "base_datos.xlsm"
    uploaded_asme = None
    use_local_asme = False
    try:
        f = open(archivo_asme, "r"); use_local_asme = True
        st.success("✅ BD Local Conectada")
    except:
        uploaded_asme = st.file_uploader("Subir Tabla ASME", type=["xlsx", "xlsm"])

    # 2. PROVEEDOR
    st.markdown("---")
    st.subheader("2. Archivo Semanal")
    uploaded_file = st.file_uploader("Sube el Excel aquí", type=["xlsx"])
    fecha_stock = st.date_input("Fecha Stock", datetime.date.today())

    # 3. PRECIOS
    st.markdown("---")
    st.subheader("3. Precios (USD)")
    df_precios_editor = st.data_editor(
        pd.DataFrame({
            "Min (mm)": [1.0, 6.01, 9.01], "Max (mm)": [6.0, 9.0, 999.0],
            "Piso": [0.90, 1.00, 1.15], "Techo": [1.15, 1.25, 1.50]
        }), num_rows="dynamic", hide_index=True
    )

# --- FUNCIONES MATEMÁTICAS AVANZADAS ---

def fraction_to_float(text):
    """Convierte '2 1/2', '1/2', '2' a float 2.5, 0.5, 2.0"""
    if pd.isna(text): return None
    t = str(text).strip().replace('"', '')
    try:
        # Intento 1: Es un número normal (ej: 2.5)
        return float(t)
    except:
        # Intento 2: Es una fracción (ej: 2 1/2)
        try:
            if ' ' in t: # Caso mixto "2 1/2"
                parts = t.split()
                if len(parts) == 2:
                    whole = float(parts[0])
                    num, den = parts[1].split('/')
                    return whole + (float(num)/float(den))
            elif '/' in t: # Caso simple "1/2"
                num, den = t.split('/')
                return float(num)/float(den)
        except:
            return None
    return None

def cargar_asme_engine(file):
    try:
        # Carga tolerante a encabezados
        df_raw = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=None, engine='openpyxl')
        fila_header = 2 # Default
        for i in range(15):
            row_str = df_raw.iloc[i].astype(str).str.lower().tolist()
            if any("scheduleno" in str(x).replace(" ","") for x in row_str):
                fila_header = i; break
        
        df_asme = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=fila_header, engine='openpyxl')
        
        # Mapeo inteligente
        def fcol(keys, cols):
            for c in cols:
                if any(k in str(c).lower() for k in keys): return c
            return None
            
        c_dn = fcol(['nps', 'diameter'], df_asme.columns)
        c_sch = fcol(['scheduleno', 'schedule'], df_asme.columns)
        c_esp = fcol(['thickness', 'mm'], df_asme.columns)
        c_peso = fcol(['mass', 'kg/m'], df_asme.columns)
        
        if not all([c_dn, c_sch, c_esp, c_peso]): 
            # Intento de búsqueda fallback más agresivo
            cols_str = [str(c).lower() for c in df_asme.columns]
            return None, f"No encontré columnas. Vi esto: {cols_str}"
        
        df = df_asme[[c_dn, c_sch, c_esp, c_peso]].copy()
        df.columns = ['DN_Raw', 'SCH', 'Espesor_mm', 'Peso_Kg_m']
        df = df.dropna(subset=['DN_Raw'])
        
        # LIMPIEZA CLAVE: Convertir todo DN a número decimal para cruzar
        df['DN_Float'] = df['DN_Raw'].apply(fraction_to_float)
        df['SCH'] = df['SCH'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # Asegurar números
        df['Peso_Kg_m'] = pd.to_numeric(df['Peso_Kg_m'], errors='coerce')
        df['Espesor_mm'] = pd.to_numeric(df['Espesor_mm'], errors='coerce')
        
        return df, "OK"
    except Exception as e: return None, str(e)

def procesar_regex(texto):
    d = {"DN_Txt": None, "SCH": None, "Esp": None, "Tira": 6.0} 
    t = str(texto)
    # 1. DN: Busca patron numero + comilla (ej: 2 1/2")
    m_dn = re.search(r'(\d+\s?\d*/?\d*)"', t)
    if m_dn: d["DN_Txt"] = m_dn.group(1) # Guarda "2 1/2"
    
    # 2. Tira
    m_mts = re.search(r'(\d+(?:[.,]\d+)?)\s*Mts', t, re.IGNORECASE)
    if m_mts: d["Tira"] = float(m_mts.group(1).replace(',', '.'))
    
    # 3. SCH
    m_sch = re.search(r'Sch\s?(\d+|STD|XS)', t, re.IGNORECASE)
    if m_sch: d["SCH"] = m_sch.group(1)
    
    # 4. Espesor (si no hay SCH)
    if not d["SCH"]:
        m_esp = re.search(r'[xX]\s*(\d+[,.]\d+)\s*[xX]', t)
        if m_esp: d["Esp"] = float(m_esp.group(1).replace(',', '.'))
    return pd.Series(d)

def safe_clean_qty(val):
    """Limpia cantidades de Excel de forma segura"""
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    # Si es texto, limpiamos
    s = str(val).strip()
    if s == '-' or s == '': return 0.0
    # Caso europeo 1.000,00 -> 1000.00
    if '.' in s and ',' in s: s = s.replace('.', '').replace(',', '.')
    elif ',' in s: s = s.replace(',', '.')
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

# --- MAIN ---
st.title("🏭 Visor de Stock V7.1")

if uploaded_file:
    df_prev = pd.read_excel(uploaded_file)
    cols = df_prev.columns.tolist()
    
    # Buscadores de columnas
    i_desc = next((i for i,c in enumerate(cols) if 'descrip' in str(c).lower()), 0)
    i_cant = next((i for i,c in enumerate(cols) if any(x in str(c).lower() for x in ['dispo','cant','stock'])), 2)

    with st.expander("🛠️ Configurar Columnas", expanded=False):
        c1, c2 = st.columns(2)
        col_desc = c1.selectbox("Columna Descripción:", cols, index=i_desc)
        col_cant = c2.selectbox("Columna Cantidad:", cols, index=i_cant)

    if st.button("🚀 PROCESAR AHORA", type="primary"):
        fuente = archivo_asme if use_local_asme else uploaded_asme
        if not fuente: st.error("Falta ASME"); st.stop()
        
        df_asme, msg = cargar_asme_engine(fuente)
        if df_asme is None: st.error(msg); st.stop()

        with st.spinner("Analizando datos..."):
            df_final = df_prev.copy()
            
            # 1. Extracción Regex
            extracted = df_final[col_desc].apply(procesar_regex)
            df_final = pd.concat([df_final, extracted], axis=1)
            
            # 2. Conversión DN a Float para cruce (Magia V7)
            df_final['DN_Float'] = df_final['DN_Txt'].apply(fraction_to_float)

            # 3. Cruce con ASME
            pesos, espesores_fin = [], []
            matches_count = 0
            
            for _, r in df_final.iterrows():
                dn_val = r['DN_Float'] # Float (ej: 2.5)
                sch = str(r['SCH']).strip() if pd.notna(r['SCH']) else None
                esp_manual = r['Esp']
                
                match = pd.DataFrame()
                if pd.notna(dn_val):
                    # Filtramos ASME por valor numérico (2.5 == 2.5)
                    # Usamos tolerancia pequeña
                    base = df_asme[np.isclose(df_asme['DN_Float'], dn_val, atol=0.01)]
                    
                    if not base.empty:
                        if sch: 
                            match = base[base['SCH'] == sch]
                        elif pd.notna(esp_manual):
                            # Buscar espesor mas cercano
                            match = base[np.isclose(base['Espesor_mm'], esp_manual, atol=0.25)]
                
                if not match.empty:
                    pesos.append(match.iloc[0]['Peso_Kg_m'])
                    espesores_fin.append(match.iloc[0]['Espesor_mm'])
                    matches_count += 1
                else:
                    pesos.append(0.0)
                    espesores_fin.append(esp_manual if pd.notna(esp_manual) else 0.0)
            
            df_final['Peso_Unitario'] = pd.Series(pesos).fillna(0.0)
            df_final['Espesor_Final'] = pd.Series(espesores_fin).fillna(0.0)
            
            # 4. Asignar Tipos
            def get_tipo(txt):
                t = str(txt).lower()
                if "iso" in t and "negra" in t: return "3. CCC ISO Negra"
                if "iso" in t and "galva" in t: return "5. CCC ISO Galvanizada"
                if "astm" in t and "a-53" in t: return "1. CCC ASTM A53 / API 5L"
                if "galva" in t and "bsp" in t: return "8. CCC ASTM A-53 Galva R65 H/BSP"
                if "galva" in t: return "7. CCC ASTM A-53 Galva"
                if "a795" in t and "ranurada" in t: return "4a. CCC A795 Negra - Ranurada"
                if "a795" in t: return "4. CCC A795 Negra"
                return "⚠️ MANUAL"
            
            df_final['TIPO_SISTEMA'] = df_final[col_desc].apply(get_tipo)
            
            # 5. Cálculos Finales
            df_final['Cant_Clean'] = df_final[col_cant].apply(safe_clean_qty)
            df_final['Stock_Mts'] = df_final['Cant_Clean'] * df_final['Tira']
            df_final['Stock_Kgs'] = df_final['Stock_Mts'] * df_final['Peso_Unitario']
            
            # 6. Precios
            def calc_px(esp, conf):
                if esp <= 0: return 0.0, 0.0
                for _, x in conf.iterrows():
                    if x['Min (mm)'] <= esp <= x['Max (mm)']: return x['Piso'], x['Techo']
                return 0.0, 0.0
            
            pxs = df_final['Espesor_Final'].apply(lambda x: calc_px(x, df_precios_editor))
            df_final['Px_Piso_Kg'] = pxs.apply(lambda x: x[0])
            df_final['Px_Techo_Kg'] = pxs.apply(lambda x: x[1])
            df_final['Px_Piso_Mt'] = df_final['Px_Piso_Kg'] * df_final['Peso_Unitario']
            
            # Guardar
            cols_ok = [col_desc, 'TIPO_SISTEMA', 'DN_Txt', 'SCH', 'Espesor_Final', 
                       'Stock_Mts', 'Stock_Kgs', 'Px_Piso_Kg', 'Px_Techo_Kg', 'Px_Piso_Mt']
            st.session_state['data'] = df_final[cols_ok]
            
            if matches_count == 0:
                st.warning("⚠️ OJO: No se encontraron coincidencias con ASME. Revisa que el DN se extraiga bien.")
            else:
                st.success(f"✅ Procesado: {matches_count} artículos cruzados correctamente con tabla ASME.")
            st.rerun()

# --- VISUALIZADOR ---
if st.session_state.get('data') is not None:
    df = st.session_state['data']
    
    # Filtros
    st.markdown("### 🔎 Explorador")
    c1, c2, c3, c4 = st.columns(4)
    df['DN_Txt'] = df['DN_Txt'].fillna("?")
    df['SCH'] = df['SCH'].fillna("?")
    
    f_tipo = c1.selectbox("Tipo", ["Todos"] + sorted(df['TIPO_SISTEMA'].unique()))
    f_dn = c2.selectbox("DN", ["Todos"] + sorted(df['DN_Txt'].astype(str).unique()))
    f_sch = c3.selectbox("SCH", ["Todos"] + sorted(df['SCH'].astype(str).unique()))
    f_esp = c4.selectbox("Espesor", ["Todos"] + sorted(df['Espesor_Final'].unique()))
    
    df_show = df.copy()
    if f_tipo != "Todos": df_show = df_show[df_show['TIPO_SISTEMA'] == f_tipo]
    if f_dn != "Todos": df_show = df_show[df_show['DN_Txt'] == f_dn]
    if f_sch != "Todos": df_show = df_show[df_show['SCH'] == f_sch]
    if f_esp != "Todos": df_show = df_show[df_show['Espesor_Final'] == f_esp]
    
    st.divider()
    
    # KPIs
    k1, k2, k3 = st.columns(3)
    k1.metric("📦 Items", len(df_show))
    k2.metric("📏 Metros Totales", f"{df_show['Stock_Mts'].sum():,.0f} m")
    k3.metric("⚖️ Kilos Totales", f"{df_show['Stock_Kgs'].sum():,.0f} kg")
    
    # Tabla Editable
    st.markdown("### 📝 Planilla Final")
    opciones_tipo = [
        "1. CCC ASTM A53 / API 5L", "2. CSC ASTM A106 / API 5L", "3. CCC ISO Negra",
        "4. CCC A795 Negra", "4a. CCC A795 Negra - Ranurada", "5. CCC ISO Galvanizada",
        "6. CCC ASTM A-53 Gr.B Galva/H/NPT", "7. CCC ASTM A-53 Galva",
        "8. CCC ASTM A-53 Galva R65 H/BSP", "9. CCC ASTM A-53 Galva H/NPT CUPLA", "⚠️ MANUAL"
    ]
    
    df_edit = st.data_editor(
        df_show,
        column_config={
            "TIPO_SISTEMA": st.column_config.SelectboxColumn("Tipo", options=opciones_tipo, required=True),
            "Px_Piso_Kg": st.column_config.NumberColumn("USD/Kg Piso", format="$ %.2f"),
            "Px_Techo_Kg": st.column_config.NumberColumn("USD/Kg Techo", format="$ %.2f"),
            "Px_Piso_Mt": st.column_config.NumberColumn("USD/Mt Piso", format="$ %.2f"),
            "Stock_Kgs": st.column_config.NumberColumn("Total Kg", format="%.0f"),
        },
        use_container_width=True, hide_index=True
    )
    
    # Descarga
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_edit.to_excel(writer, index=False)
    st.download_button("📥 Descargar Excel", buffer.getvalue(), f"Stock_{fecha_stock}.xlsx", "application/vnd.ms-excel", type="primary")
    
    if st.button("Reiniciar"):
        st.session_state['data'] = None; st.rerun()