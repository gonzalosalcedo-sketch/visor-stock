import streamlit as st
import pandas as pd
import re
import numpy as np
import io
import datetime

st.set_page_config(page_title="Gestor Stock - V9 Final", layout="wide", page_icon="🏭")

# --- ESTILOS ---
st.markdown("""
<style>
    .stAlert { padding: 0.5rem; }
    .metric-card { background:#f9f9f9; border:1px solid #ddd; padding:10px; border-radius:5px; text-align:center; }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.title("🎛️ Panel de Control")
    
    # 1. ASME
    st.header("1. Base de Datos")
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
    st.header("2. Archivo Semanal")
    uploaded_file = st.file_uploader("Subir Excel Proveedor", type=["xlsx"])
    fecha_stock = st.date_input("Fecha", datetime.date.today())

    # 3. PRECIOS
    st.markdown("---")
    st.header("3. Precios")
    df_precios_editor = st.data_editor(
        pd.DataFrame({
            "Min (mm)": [1.0, 6.01, 9.01], "Max (mm)": [6.0, 9.0, 999.0],
            "Piso": [0.90, 1.00, 1.15], "Techo": [1.15, 1.25, 1.50]
        }), num_rows="dynamic", hide_index=True
    )

# --- FUNCIONES DE CEREBRO ---

def find_header_row(file):
    """Busca en qué fila empiezan los títulos reales (Descripción, Stock, etc)"""
    df_temp = pd.read_excel(file, header=None, nrows=15)
    for i, row in df_temp.iterrows():
        row_str = row.astype(str).str.lower().tolist()
        # Si la fila tiene palabras clave como 'descrip' o 'articulo' Y 'stock' o 'cant'
        if any('descrip' in x for x in row_str) or (any('articulo' in x for x in row_str) and any('stock' in x for x in row_str)):
            return i
    return 0 # Si no encuentra, asume fila 0

def fraction_to_float(text):
    """Convierte '2 1/2' -> 2.5 de forma robusta"""
    if pd.isna(text): return None
    t = str(text).replace('"', '').replace("'", "").strip()
    try:
        return float(t)
    except:
        try:
            if ' ' in t: parts = t.split(); return float(parts[0]) + eval(parts[1])
            if '/' in t: return eval(t)
        except: return None
    return None

def extract_regex(desc):
    """Extrae datos del texto de descripción"""
    d = {"DN_Txt": None, "SCH": None, "Esp": None, "Tira": 6.0}
    t = str(desc)
    
    # DN: Busca numero antes de comilla o espacio (2" o 2 1/2)
    m_dn = re.search(r'(\d+[\s-]?\d*/?\d*)(?="|\s|$)', t)
    if m_dn: d["DN_Txt"] = m_dn.group(1).replace('-', ' ')
    
    # SCH
    m_sch = re.search(r'Sch\.?\s?(\d+|STD|XS)', t, re.IGNORECASE)
    if m_sch: d["SCH"] = m_sch.group(1)
    
    # Tira
    m_mts = re.search(r'(\d+(?:[.,]\d+)?)\s*Mts', t, re.IGNORECASE)
    if m_mts: d["Tira"] = float(m_mts.group(1).replace(',', '.'))
    
    # Espesor Manual (si dice x 6.35 x)
    m_esp = re.search(r'[xX]\s*(\d+[,.]\d+)\s*[xX]', t)
    if m_esp: d["Esp"] = float(m_esp.group(1).replace(',', '.'))
    
    return pd.Series(d)

def load_asme_engine(file):
    try:
        # Busca header ASME
        df_raw = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=None, engine='openpyxl')
        h_idx = 2
        for i in range(15):
            r = df_raw.iloc[i].astype(str).str.lower().tolist()
            if any("schedule" in x for x in r): h_idx = i; break
            
        df = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=h_idx, engine='openpyxl')
        
        # Busca columnas
        cols = [str(c) for c in df.columns]
        c_dn = next((c for c in cols if 'nps' in c.lower() or 'diameter' in c.lower()), None)
        c_sch = next((c for c in cols if 'schedule' in c.lower()), None)
        c_esp = next((c for c in cols if 'thickness' in c.lower() and 'mm' in c.lower()), None)
        c_peso = next((c for c in cols if 'mass' in c.lower()), None)
        
        if not all([c_dn, c_sch, c_esp, c_peso]): return None, "Faltan columnas en ASME"
        
        df = df[[c_dn, c_sch, c_esp, c_peso]].dropna()
        df.columns = ['DN', 'SCH', 'Esp', 'Peso']
        
        # Limpieza matematica
        df['DN_Float'] = df['DN'].apply(fraction_to_float)
        df['Esp'] = pd.to_numeric(df['Esp'], errors='coerce')
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce')
        df['SCH'] = df['SCH'].astype(str).str.replace('.0','').str.strip()
        
        return df, "OK"
    except Exception as e: return None, str(e)

# --- APP PRINCIPAL ---
st.title("🏭 Visor de Stock V9 (Auto-Detect)")

if uploaded_file:
    # 1. DETECCIÓN AUTOMÁTICA DE ENCABEZADOS
    header_row = find_header_row(uploaded_file)
    if header_row > 0:
        st.info(f"💡 Detecté que la tabla empieza en la fila {header_row+1}. Ajustando automáticamente...")
    
    # Recargar con el header correcto
    uploaded_file.seek(0)
    df_prov = pd.read_excel(uploaded_file, header=header_row)
    
    # 2. SELECTOR DE COLUMNAS (CRÍTICO)
    cols = df_prov.columns.tolist()
    
    # Intentar adivinar la columna descripcion (evitar 'Codigo')
    idx_desc = 0
    for i, c in enumerate(cols):
        c_lower = str(c).lower()
        if 'descrip' in c_lower: idx_desc = i; break
    
    # Intentar adivinar cantidad
    idx_cant = 2
    for i, c in enumerate(cols):
        c_lower = str(c).lower()
        if any(x in c_lower for x in ['disp', 'stock', 'cant', 'saldo']): idx_cant = i; break

    st.markdown("### 🛠️ Paso 1: Confirma las Columnas")
    c1, c2 = st.columns(2)
    col_desc = c1.selectbox("¿Cuál es la DESCRIPCIÓN? (Ej: Tubo ASTM...)", cols, index=idx_desc)
    col_cant = c2.selectbox("¿Cuál es la CANTIDAD?", cols, index=idx_cant)
    
    # Muestra de seguridad
    st.caption(f"Ejemplo de lo que voy a leer en **{col_desc}**: _{df_prov[col_desc].iloc[0]}_")
    
    if st.button("🚀 PROCESAR AHORA", type="primary"):
        # Cargar ASME
        fuente = archivo_asme if use_local_asme else uploaded_asme
        if not fuente: st.error("Falta ASME"); st.stop()
        
        df_asme, msg = load_asme_engine(fuente)
        if df_asme is None: st.error(msg); st.stop()
        
        with st.spinner("Analizando..."):
            df_final = df_prov.copy()
            
            # 1. Extraer Info
            extracted = df_final[col_desc].apply(extract_regex)
            df_final = pd.concat([df_final, extracted], axis=1)
            df_final['DN_Float'] = df_final['DN_Txt'].apply(fraction_to_float)
            
            # 2. Cruce ASME
            pesos, esps_fin = [], []
            matches = 0
            
            for _, r in df_final.iterrows():
                dn_val = r['DN_Float']
                sch = str(r['SCH']).strip() if pd.notna(r['SCH']) else None
                esp_man = r['Esp']
                
                match = pd.DataFrame()
                if pd.notna(dn_val):
                    # Filtro DN (tolerancia)
                    base = df_asme[np.isclose(df_asme['DN_Float'], dn_val, atol=0.05)]
                    
                    if not base.empty:
                        if sch: match = base[base['SCH'] == sch]
                        elif pd.notna(esp_man): match = base[np.isclose(base['Esp'], esp_man, atol=0.25)]
                        
                        # Fallback: si encontró DN pero no SCH/Esp, tomar promedio (mejor que 0)
                        if match.empty: match = base.head(1) 
                
                if not match.empty:
                    pesos.append(float(match.iloc[0]['Peso']))
                    esps_fin.append(float(match.iloc[0]['Esp']))
                    matches += 1
                else:
                    pesos.append(0.0)
                    esps_fin.append(float(esp_man) if pd.notna(esp_man) else 0.0)
            
            df_final['Peso_Unitario'] = pesos
            df_final['Espesor_Final'] = esps_fin
            
            # 3. Cálculos
            def clean_num(x):
                s = str(x).strip()
                if '.' in s and ',' in s: s = s.replace('.','').replace(',','.') # 1.000,00 -> 1000.00
                elif ',' in s: s = s.replace(',','.')
                try: return float(re.sub(r'[^\d.]', '', s))
                except: return 0.0
                
            df_final['Cant_Clean'] = df_final[col_cant].apply(clean_num)
            df_final['Stock_Mts'] = df_final['Cant_Clean'] * df_final['Tira']
            df_final['Stock_Kgs'] = df_final['Stock_Mts'] * df_final['Peso_Unitario']
            
            # 4. Tipos y Precios
            def get_tipo(t):
                t = str(t).lower()
                if "iso" in t: return "3. CCC ISO Negra" if "negra" in t else "5. CCC ISO Galvanizada"
                if "astm" in t and "a-53" in t: return "1. CCC ASTM A53 / API 5L"
                return "⚠️ MANUAL"
            
            df_final['TIPO_SISTEMA'] = df_final[col_desc].apply(get_tipo)
            
            def get_px(esp, df_p):
                if esp <= 0: return 0.0, 0.0
                for _, r in df_p.iterrows():
                    if r['Min (mm)'] <= esp <= r['Max (mm)']: return r['Piso'], r['Techo']
                return 0.0, 0.0
            
            pxs = df_final['Espesor_Final'].apply(lambda x: get_px(x, df_precios_editor))
            df_final['Piso_Kg'] = pxs.apply(lambda x: x[0])
            df_final['Techo_Kg'] = pxs.apply(lambda x: x[1])
            df_final['Piso_Mt'] = df_final['Piso_Kg'] * df_final['Peso_Unitario']
            
            # Guardar en sesión
            cols_ok = [col_desc, 'TIPO_SISTEMA', 'DN_Txt', 'SCH', 'Espesor_Final', 'Stock_Mts', 'Stock_Kgs', 'Piso_Kg', 'Techo_Kg', 'Piso_Mt']
            st.session_state['data'] = df_final[cols_ok]
            
            if matches > 0: st.success(f"✅ ¡Éxito! {matches} Artículos procesados correctamente.")
            else: st.error("⚠️ Alerta: No se cruzaron datos. Revisa si la columna Descripción es la correcta.")
            st.rerun()

# --- VISUALIZADOR ---
if st.session_state.get('data') is not None:
    df = st.session_state['data']
    
    st.markdown("### 🔎 Resultados")
    
    # Filtros
    c1, c2, c3 = st.columns(3)
    f_tipo = c1.selectbox("Tipo", ["Todos"] + sorted(df['TIPO_SISTEMA'].unique()))
    f_dn = c2.selectbox("DN", ["Todos"] + sorted(df['DN_Txt'].fillna("?").unique()))
    
    df_view = df.copy()
    if f_tipo != "Todos": df_view = df_view[df_view['TIPO_SISTEMA'] == f_tipo]
    if f_dn != "Todos": df_view = df_view[df_view['DN_Txt'] == f_dn]
    
    # KPIs
    k1, k2, k3 = st.columns(3)
    k1.metric("📦 Items", len(df_view))
    k2.metric("📏 Metros", f"{df_view['Stock_Mts'].sum():,.0f}")
    k3.metric("⚖️ Kilos", f"{df_view['Stock_Kgs'].sum():,.0f}")
    
    # Tabla
    opciones_tipo = [
        "1. CCC ASTM A53 / API 5L", "2. CSC ASTM A106 / API 5L", "3. CCC ISO Negra",
        "4. CCC A795 Negra", "4a. CCC A795 Negra - Ranurada", "5. CCC ISO Galvanizada",
        "6. CCC ASTM A-53 Gr.B Galva/H/NPT", "7. CCC ASTM A-53 Galva",
        "8. CCC ASTM A-53 Galva R65 H/BSP", "9. CCC ASTM A-53 Galva H/NPT CUPLA", "⚠️ MANUAL"
    ]
    
    df_edit = st.data_editor(
        df_view,
        column_config={
            "TIPO_SISTEMA": st.column_config.SelectboxColumn("Tipo", options=opciones_tipo, required=True),
            "Piso_Kg": st.column_config.NumberColumn("USD/Kg Piso", format="$ %.2f"),
            "Techo_Kg": st.column_config.NumberColumn("USD/Kg Techo", format="$ %.2f"),
            "Stock_Kgs": st.column_config.NumberColumn("Total Kg", format="%.0f"),
        },
        use_container_width=True, hide_index=True
    )
    
    # Descarga
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_edit.to_excel(writer, index=False)
    st.download_button("📥 Descargar Excel", buffer.getvalue(), f"Stock_{fecha_stock}.xlsx", "application/vnd.ms-excel", type="primary")
    
    if st.button("🔄 Reiniciar"): st.session_state['data'] = None; st.rerun()

else:
    st.info("👆 Sube tu archivo. El sistema detectará dónde empieza la tabla automáticamente.")