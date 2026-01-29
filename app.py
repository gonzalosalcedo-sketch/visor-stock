import streamlit as st
import pandas as pd
import re
import numpy as np
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Sistema de Stock Inteligente", layout="wide")

st.title("🏭 Sistema de Actualización de Stock y Precios")
st.markdown("---")

# --- BARRA LATERAL: CONFIGURACIÓN ---
with st.sidebar:
    st.header("⚙️ Configuración")
    
    st.subheader("1. Tabla ASME (Base de Datos)")
    # Intenta cargar el archivo local, si no pide subirlo
    archivo_asme = "base_datos.xlsm"
    uploaded_asme = None
    
    try:
        # Intento leer local para ver si existe en el repositorio
        f = open(archivo_asme, "r")
        st.success(f"✅ Base de datos '{archivo_asme}' cargada del sistema.")
        use_local_asme = True
    except FileNotFoundError:
        st.warning("⚠️ No encuentro 'base_datos.xlsm'. Súbelo aquí:")
        uploaded_asme = st.file_uploader("Cargar Tabla ASME", type=["xlsx", "xlsm"])
        use_local_asme = False

    st.markdown("---")
    st.subheader("2. Matriz de Precios (USD/Kg)")
    st.info("Edita los precios aquí. Se aplicarán al procesar.")
    
    # Datos iniciales de precios (Tu configuración por defecto)
    data_precios = {
        "Espesor Min (mm)": [1.0, 6.01, 9.01],
        "Espesor Max (mm)": [6.0, 9.0, 999.0],
        "Px Piso (USD)": [0.90, 1.00, 1.15],
        "Px Techo (USD)": [1.15, 1.25, 1.50]
    }
    df_precios_config = pd.DataFrame(data_precios)
    df_precios_editor = st.data_editor(df_precios_config, num_rows="dynamic")

# --- FUNCIONES DE LÓGICA (EL CEREBRO) ---
def cargar_asme(file_path_or_buffer):
    try:
        # 1. Detectar encabezado real (Estrategia Francotirador)
        df_raw = pd.read_excel(file_path_or_buffer, sheet_name="TABLA ASME-B36.10M", header=None, engine='openpyxl')
        fila_header = -1
        for i in range(15):
            fila_texto = df_raw.iloc[i].astype(str).str.lower().tolist()
            if any("scheduleno" in x.replace(" ","") for x in fila_texto):
                fila_header = i
                break
        
        if fila_header == -1: fila_header = 2 # Fallback
        
        # 2. Cargar con encabezado correcto
        df_asme = pd.read_excel(file_path_or_buffer, sheet_name="TABLA ASME-B36.10M", header=fila_header, engine='openpyxl')
        
        # 3. Buscar columnas clave
        def buscar_col(keywords, cols):
            for col in cols:
                if all(k.lower() in str(col).lower() for k in keywords): return col
            return None
            
        col_dn = buscar_col(['NPS'], df_asme.columns) or buscar_col(['Diameter', 'in.'], df_asme.columns)
        col_sch = buscar_col(['ScheduleNo'], df_asme.columns) or buscar_col(['Schedule'], df_asme.columns)
        col_esp = buscar_col(['Thickness', 'mm'], df_asme.columns)
        col_peso = buscar_col(['Mass', 'kg/m'], df_asme.columns)
        
        if not all([col_dn, col_sch, col_esp, col_peso]):
            return None, "Error de Mapeo: No encuentro columnas NPS, SCH, Thickness(mm) o Mass(kg/m)"
            
        df_asme = df_asme[[col_dn, col_sch, col_esp, col_peso]].copy()
        df_asme.columns = ['DN', 'Schedule', 'Espesor_mm', 'Peso_Kg_m']
        df_asme = df_asme.dropna(subset=['DN'])
        df_asme['Schedule'] = df_asme['Schedule'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_asme['DN'] = df_asme['DN'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        return df_asme, "OK"
    except Exception as e:
        return None, str(e)

def procesar_descripcion(texto):
    datos = {"DN": None, "SCH": None, "Espesor": None, "Tira_Mts": 6.0} 
    texto = str(texto)
    match_dn = re.search(r'(\d+\s?\d*/?\d*)"', texto)
    if match_dn: datos["DN"] = match_dn.group(1)
    match_tira = re.search(r'(\d+(?:[.,]\d+)?)\s*Mts', texto, re.IGNORECASE)
    if match_tira: datos["Tira_Mts"] = float(match_tira.group(1).replace(',', '.'))
    match_sch = re.search(r'Sch\s?(\d+|STD|XS)', texto, re.IGNORECASE)
    if match_sch: datos["SCH"] = match_sch.group(1)
    if not datos["SCH"]:
        match_esp = re.search(r'[xX]\s*(\d+[,.]\d+)\s*[xX]', texto)
        if match_esp: datos["Espesor"] = float(match_esp.group(1).replace(',', '.'))
    return pd.Series(datos)

def asignar_tipo(desc):
    desc = str(desc).lower()
    if "iso" in desc and "negra" in desc: return "3. CCC ISO Negra"
    if "iso" in desc and "galva" in desc: return "5. CCC ISO Galvanizada"
    if "astm" in desc and "a-53" in desc: return "1. CCC ASTM A53 / API 5L"
    if "galva" in desc: return "7. CCC ASTM A-53 Galva"
    return "⚠️ MANUAL"

def obtener_precio(espesor, df_config):
    if pd.isna(espesor) or espesor == 0: return 0.0, 0.0
    for _, row in df_config.iterrows():
        if row['Espesor Min (mm)'] <= espesor <= row['Espesor Max (mm)']:
            return row['Px Piso (USD)'], row['Px Techo (USD)']
    return 0.0, 0.0

# --- INTERFAZ PRINCIPAL ---

st.write("### 1. Carga de Archivo Semanal")
uploaded_file = st.file_uploader("Sube el Excel que envió el proveedor", type=["xlsx"])

if uploaded_file:
    # Botón para iniciar proceso
    if st.button("🚀 PROCESAR STOCK", type="primary"):
        with st.spinner('Analizando descripciones, cruzando con ASME y calculando precios...'):
            try:
                # 1. Cargar Proveedor
                df_prov = pd.read_excel(uploaded_file)
                
                # 2. Cargar ASME
                fuente_asme = archivo_asme if use_local_asme else uploaded_asme
                if not fuente_asme:
                    st.error("❌ Falta la Tabla ASME. Súbela en la barra lateral.")
                    st.stop()
                    
                df_asme, msg_asme = cargar_asme(fuente_asme)
                if df_asme is None:
                    st.error(f"❌ Error leyendo ASME: {msg_asme}")
                    st.stop()

                # 3. Procesamiento
                # Buscar columna descripción
                col_desc = next((c for c in df_prov.columns if 'descrip' in c.lower()), df_prov.columns[1])
                
                # Extracción Regex
                datos_ext = df_prov[col_desc].apply(procesar_descripcion)
                df_final = pd.concat([df_prov, datos_ext], axis=1)
                
                # Cruce ASME
                pesos, esps = [], []
                for _, row in df_final.iterrows():
                    dn, sch, esp = str(row['DN']).strip(), row['SCH'], row['Espesor']
                    sch = str(sch).strip() if pd.notna(sch) else None
                    
                    match = pd.DataFrame()
                    if dn != "None":
                        base = df_asme[df_asme['DN'] == dn]
                        if not base.empty:
                            if sch: match = base[base['Schedule'] == sch]
                            elif pd.notna(esp):
                                base_esp = pd.to_numeric(base['Espesor_mm'], errors='coerce')
                                match = base[np.isclose(base_esp, esp, atol=0.2)]
                    
                    if not match.empty:
                        pesos.append(pd.to_numeric(match.iloc[0]['Peso_Kg_m'], errors='coerce'))
                        esps.append(pd.to_numeric(match.iloc[0]['Espesor_mm'], errors='coerce'))
                    else:
                        pesos.append(0)
                        esps.append(esp if pd.notna(esp) else 0)
                
                df_final['Peso_Unitario_Kg_m'] = pd.Series(pesos).fillna(0)
                df_final['Espesor_Final'] = pd.Series(esps).fillna(0)
                
                # Clasificación y Cálculos
                df_final['TIPO_SISTEMA'] = df_final[col_desc].apply(asignar_tipo)
                
                col_cant = next((c for c in df_prov.columns if 'disp' in c.lower() or 'cant' in c.lower()), df_prov.columns[2])
                
                # Limpieza numérica vital
                c_clean = pd.to_numeric(df_final[col_cant], errors='coerce').fillna(0)
                t_clean = pd.to_numeric(df_final['Tira_Mts'], errors='coerce').fillna(0)
                p_clean = pd.to_numeric(df_final['Peso_Unitario_Kg_m'], errors='coerce').fillna(0)
                
                df_final['Stock_Total_Mts'] = c_clean * t_clean
                df_final['Stock_Total_Kgs'] = df_final['Stock_Total_Mts'] * p_clean
                
                # Precios Dinámicos (Usando la tabla de la barra lateral)
                precios = df_final['Espesor_Final'].apply(lambda x: obtener_precio(x, df_precios_editor))
                df_final['Px_Piso_Kg_USD'] = precios.apply(lambda x: x[0])
                df_final['Px_Techo_Kg_USD'] = precios.apply(lambda x: x[1])
                df_final['Px_Piso_Mt_USD'] = df_final['Px_Piso_Kg_USD'] * df_final['Peso_Unitario_Kg_m']
                df_final['Px_Techo_Mt_USD'] = df_final['Px_Techo_Kg_USD'] * df_final['Peso_Unitario_Kg_m']

                # Guardar en sesión para no perderlo al filtrar
                st.session_state['df_resultado'] = df_final
                st.success("✅ ¡Procesamiento Exitoso!")
                
            except Exception as e:
                st.error(f"Ocurrió un error: {e}")

# --- VISUALIZACIÓN DE RESULTADOS ---
if 'df_resultado' in st.session_state:
    df = st.session_state['df_resultado']
    
    st.write("### 2. Visor de Stock y Precios")
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        filtro_tipo = st.multiselect("Filtrar por Tipo:", options=df['TIPO_SISTEMA'].unique())
    with col2:
        filtro_dn = st.multiselect("Filtrar por DN:", options=df['DN'].unique())
    with col3:
        filtro_sch = st.multiselect("Filtrar por SCH:", options=df['SCH'].dropna().unique())
        
    # Aplicar filtros
    df_view = df.copy()
    if filtro_tipo: df_view = df_view[df_view['TIPO_SISTEMA'].isin(filtro_tipo)]
    if filtro_dn: df_view = df_view[df_view['DN'].isin(filtro_dn)]
    if filtro_sch: df_view = df_view[df_view['SCH'].isin(filtro_sch)]
    
    # KPIs
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Stock Total (Kgs)", f"{df_view['Stock_Total_Kgs'].sum():,.0f}")
    kpi2.metric("Stock Total (Mts)", f"{df_view['Stock_Total_Mts'].sum():,.0f}")
    kpi3.metric("Items Listados", len(df_view))
    
    # Tabla
    cols_mostrar = ['Descripción', 'TIPO_SISTEMA', 'DN', 'SCH', 'Espesor_Final', 
                    'Stock_Total_Kgs', 'Stock_Total_Mts', 
                    'Px_Piso_Kg_USD', 'Px_Techo_Kg_USD', 'Px_Piso_Mt_USD']
    # Aseguramos que existan las columnas para mostrar
    cols_existentes = [c for c in cols_mostrar if c in df_view.columns]
    
    st.dataframe(df_view[cols_existentes].style.format({
        'Stock_Total_Kgs': '{:,.0f}', 
        'Px_Piso_Kg_USD': '${:.2f}',
        'Px_Techo_Kg_USD': '${:.2f}',
        'Px_Piso_Mt_USD': '${:.2f}'
    }), use_container_width=True)
    
    # Botón Descarga
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_view.to_excel(writer, index=False, sheet_name='Stock_Procesado')
    
    st.download_button(
        label="📥 Descargar Excel Procesado",
        data=buffer.getvalue(),
        file_name="Stock_Con_Precios.xlsx",
        mime="application/vnd.ms-excel"
    )