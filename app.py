import streamlit as st
import pandas as pd
import re
import numpy as np
import io
import datetime

st.set_page_config(page_title="Gestor Stock - Modo Diagnóstico", layout="wide")

st.markdown("""
<style>
    .stAlert { padding: 0.5rem; margin-bottom: 1rem; }
    .debug-box { background: #f0f2f6; padding: 10px; border-radius: 5px; font-family: monospace; font-size: 12px; }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Archivos")
    
    # Carga ASME
    archivo_asme = "base_datos.xlsm"
    uploaded_asme = None
    use_local_asme = False
    try:
        f = open(archivo_asme, "r"); use_local_asme = True
        st.success("✅ Base de datos (Local) detectada")
    except:
        uploaded_asme = st.file_uploader("Subir Tabla ASME", type=["xlsx", "xlsm"])

    # Carga Proveedor
    st.markdown("---")
    uploaded_file = st.file_uploader("Subir Excel Proveedor", type=["xlsx"])
    fecha_stock = st.date_input("Fecha", datetime.date.today())

    # Precios
    st.markdown("---")
    st.caption("Configuración de Precios")
    df_precios_editor = st.data_editor(
        pd.DataFrame({
            "Min (mm)": [1.0, 6.01, 9.01], "Max (mm)": [6.0, 9.0, 999.0],
            "Piso": [0.90, 1.00, 1.15], "Techo": [1.15, 1.25, 1.50]
        }), num_rows="dynamic", hide_index=True
    )

# --- FUNCIONES ---
def clean_dn_text(text):
    """Limpia el texto del DN para intentar entenderlo"""
    if pd.isna(text): return None
    # Quita comillas y espacios extra
    return str(text).replace('"', '').replace("'", "").strip()

def fraction_to_float(text):
    """Convierte 2 1/2 -> 2.5"""
    t = clean_dn_text(text)
    if not t: return None
    try:
        return float(t) # Caso 2.5
    except:
        try:
            if ' ' in t: # Caso 2 1/2
                parts = t.split()
                return float(parts[0]) + eval(parts[1])
            if '/' in t: # Caso 1/2
                return eval(t)
        except: return None
    return None

def extract_info(desc):
    """Extrae DN, SCH y Tira"""
    d = {"DN_Original": None, "SCH": None, "Tira": 6.0}
    t = str(desc)
    
    # 1. DN: Buscar Numero + comilla (2") o Numero + espacio + fraccion (2 1/2")
    # Regex flexible: captura "2 1/2", "2-1/2", "2", "2.5" antes de una comilla o Mts
    match_dn = re.search(r'(\d+[\s-]?\d*/?\d*)(?="|\s|$)', t)
    if match_dn: d["DN_Original"] = match_dn.group(1).replace('-', ' ')
    
    # 2. SCH
    match_sch = re.search(r'Sch\.?\s?(\d+|STD|XS)', t, re.IGNORECASE)
    if match_sch: d["SCH"] = match_sch.group(1)
    
    # 3. Tira
    match_tira = re.search(r'(\d+(?:[.,]\d+)?)\s*Mts', t, re.IGNORECASE)
    if match_tira: d["Tira"] = float(match_tira.group(1).replace(',', '.'))
    
    # 4. Espesor suelto
    match_esp = re.search(r'[xX]\s*(\d+[,.]\d+)\s*[xX]', t)
    if match_esp: d["Esp_Manual"] = float(match_esp.group(1).replace(',', '.'))
    else: d["Esp_Manual"] = None
    
    return pd.Series(d)

def load_asme(file):
    try:
        # Busca el encabezado
        df_raw = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=None, engine='openpyxl')
        header_idx = 2
        for i in range(15):
            row = df_raw.iloc[i].astype(str).str.lower().tolist()
            if any("schedule" in x for x in row): header_idx = i; break
            
        df = pd.read_excel(file, sheet_name="TABLA ASME-B36.10M", header=header_idx, engine='openpyxl')
        
        # Mapeo agresivo
        cols = [str(c) for c in df.columns]
        c_dn = next((c for c in cols if 'nps' in c.lower() or 'diameter' in c.lower()), None)
        c_sch = next((c for c in cols if 'schedule' in c.lower()), None)
        c_esp = next((c for c in cols if 'thickness' in c.lower() and 'mm' in c.lower()), None)
        c_peso = next((c for c in cols if 'mass' in c.lower() and 'kg' in c.lower()), None)
        
        if not all([c_dn, c_sch, c_esp, c_peso]):
            return None, f"No encontré columnas clave. Detectadas: {cols}"
            
        df = df[[c_dn, c_sch, c_esp, c_peso]].dropna()
        df.columns = ['DN', 'SCH', 'Esp', 'Peso']
        
        # Convertir a numeros puros
        df['DN_Float'] = df['DN'].apply(fraction_to_float)
        df['Esp'] = pd.to_numeric(df['Esp'], errors='coerce')
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce')
        df['SCH'] = df['SCH'].astype(str).str.replace('.0', '').str.strip()
        
        return df, "OK"
    except Exception as e: return None, str(e)

# --- MAIN ---
st.title("🏭 Visor de Stock (Modo Diagnóstico)")

if uploaded_file:
    df_prov = pd.read_excel(uploaded_file)
    st.info("📂 Archivo cargado. Por favor, verifica las columnas abajo:")
    
    cols = df_prov.columns.tolist()
    c1, c2 = st.columns(2)
    col_desc = c1.selectbox("Columna DESCRIPCIÓN:", cols, index=next((i for i,c in enumerate(cols) if 'desc' in str(c).lower()), 0))
    col_cant = c2.selectbox("Columna CANTIDAD:", cols, index=next((i for i,c in enumerate(cols) if 'disp' in str(c).lower()), 2))
    
    # Paso 1: Extracción
    st.markdown("### 🕵️ Paso 1: ¿Qué estoy leyendo?")
    extracted = df_prov[col_desc].apply(extract_info)
    df_combined = pd.concat([df_prov, extracted], axis=1)
    df_combined['DN_Float'] = df_combined['DN_Original'].apply(fraction_to_float)
    
    # Muestra visual para el usuario
    st.dataframe(df_combined[[col_desc, 'DN_Original', 'DN_Float', 'SCH', 'Tira']].head(5), use_container_width=True)
    st.caption("👆 Si la columna 'DN_Float' está vacía o dice NaN, el problema es que no entiendo cómo está escrito el diámetro.")
    
    # Paso 2: Carga ASME
    fuente = archivo_asme if use_local_asme else uploaded_asme
    if not fuente: st.error("❌ Falta ASME"); st.stop()
    
    df_asme, msg = load_asme(fuente)
    if df_asme is None: st.error(f"Error ASME: {msg}"); st.stop()
    
    # Paso 3: Cruce
    st.markdown("### 🔗 Paso 2: Intentando conectar con ASME...")
    
    matches = 0
    resultados = []
    
    for idx, row in df_combined.iterrows():
        dn_val = row['DN_Float']
        sch = str(row['SCH']).strip() if pd.notna(row['SCH']) else None
        esp_manual = row['Esp_Manual']
        
        peso = 0.0
        esp_final = 0.0
        
        if pd.notna(dn_val):
            # Filtro por DN (Tolerancia 0.05 para errores de redondeo)
            base = df_asme[np.isclose(df_asme['DN_Float'], dn_val, atol=0.05)]
            
            if not base.empty:
                match = pd.DataFrame()
                if sch:
                    match = base[base['SCH'] == sch]
                elif pd.notna(esp_manual):
                    match = base[np.isclose(base['Esp'], esp_manual, atol=0.3)]
                
                if not match.empty:
                    peso = match.iloc[0]['Peso']
                    esp_final = match.iloc[0]['Esp']
                    matches += 1
                elif not base.empty:
                    # Encontró el DN pero no el SCH/Espesor, usamos promedio o el primero para no devolver cero
                    # (Fallback de emergencia)
                    peso = base.iloc[0]['Peso'] 
                    esp_final = base.iloc[0]['Esp']
        
        resultados.append({'Peso_Unitario': peso, 'Espesor_Final': esp_final})
    
    df_res = pd.DataFrame(resultados)
    df_final = pd.concat([df_combined.reset_index(drop=True), df_res], axis=1)
    
    if matches == 0:
        st.error("🚨 CRÍTICO: No logré cruzar NINGÚN artículo. Revisa 'DN_Float' en la tabla de arriba.")
    else:
        st.success(f"✅ ¡Éxito! Crucé {matches} artículos correctamente.")
    
    # Paso 4: Cálculos
    def clean_num(x):
        try: return float(str(x).replace('.', '').replace(',', '.')) # Formato europeo 1.000,00
        except: return 0.0
        
    df_final['Cant_Clean'] = pd.to_numeric(df_final[col_cant], errors='coerce').fillna(0)
    # Si dio todo cero, intentamos limpieza agresiva europea
    if df_final['Cant_Clean'].sum() == 0:
         df_final['Cant_Clean'] = df_final[col_cant].apply(clean_num)
         
    df_final['Stock_Mts'] = df_final['Cant_Clean'] * df_final['Tira']
    df_final['Stock_Kgs'] = df_final['Stock_Mts'] * df_final['Peso_Unitario']
    
    # Tipos
    def get_tipo(t):
        t = str(t).lower()
        if "iso" in t: return "3. CCC ISO Negra" if "negra" in t else "5. CCC ISO Galvanizada"
        if "a-53" in t: return "1. CCC ASTM A53"
        return "⚠️ MANUAL"
    df_final['TIPO_SISTEMA'] = df_final[col_desc].apply(get_tipo)

    # Precios
    def get_px(esp, df_p):
        for _, r in df_p.iterrows():
            if r['Min (mm)'] <= esp <= r['Max (mm)']: return r['Piso'], r['Techo']
        return 0.0, 0.0
    
    pxs = df_final['Espesor_Final'].apply(lambda x: get_px(x, df_precios_editor))
    df_final['Piso_Kg'] = pxs.apply(lambda x: x[0])
    df_final['Techo_Kg'] = pxs.apply(lambda x: x[1])
    
    # VISUALIZACIÓN FINAL
    st.markdown("### 📊 Resultado Final")
    cols_ver = [col_desc, 'DN_Original', 'Espesor_Final', 'Stock_Mts', 'Stock_Kgs', 'Piso_Kg', 'Techo_Kg']
    st.dataframe(df_final[cols_ver].head(50))
    
    # Descarga
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False)
    st.download_button("📥 Descargar Excel Operativo", buffer.getvalue(), "Stock_Final.xlsx")

else:
    st.info("Esperando archivo...")