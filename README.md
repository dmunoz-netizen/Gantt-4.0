import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta, time
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

# Configuración de la página
st.set_page_config(page_title="APS Inyección", layout="wide")
# --- AQUÍ ESTÁ TU NUEVO TÍTULO PRO MAX ---
st.title("🏭 Planificador de Planta APS - Moldeo por Inyección DMUNOZ 4.0 PRO MAX")

# ==========================================
# 1. PARÁMETROS EN LA BARRA LATERAL (CON MEMORIA)
# ==========================================
st.sidebar.header("⚙️ Parámetros de Planta")
TIEMPO_CAMBIO_MOLDE = st.sidebar.number_input("🔧 Cambio de Molde (min):", value=180, step=15)
TIEMPO_CAMBIO_COLOR = st.sidebar.number_input("🎨 Purga por Color (min):", value=30, step=5)

st.sidebar.write("---")
st.sidebar.write("### 📅 Arranque del Plan")

if "fecha_inicial" not in st.session_state:
    st.session_state.fecha_inicial = datetime.now().date()
if "hora_inicial" not in st.session_state:
    st.session_state.hora_inicial = time(7, 30) 

FECHA_ARRANQUE = st.sidebar.date_input("Fecha de Inicio del Plan:", value=st.session_state.fecha_inicial)
HORA_ARRANQUE = st.sidebar.time_input("Hora de Inicio del Plan:", value=st.session_state.hora_inicial)
FECHA_HORA_COMBINADA = datetime.combine(FECHA_ARRANQUE, HORA_ARRANQUE)

RUTA_EXCEL = r"G:\Mi unidad\programacion.xlsx"

# ==========================================
# 2. LECTOR DE EXCEL "A PRUEBA DE BALAS"
# ==========================================
@st.cache_data(ttl=1)
def cargar_datos_excel(ruta):
    try:
        df = pd.read_excel(ruta, engine='openpyxl')
        
        def safe_col(keyword):
            for col in df.columns:
                if keyword.lower() in str(col).lower().replace(' ', '').replace('.', '').replace('\n', ''):
                    return col
            return None

        if not safe_col('numsap') and not safe_col('maquina'):
            df = pd.read_excel(ruta, skiprows=1, engine='openpyxl')
            st.session_state.fila_salto_excel = 1
        else:
            st.session_state.fila_salto_excel = 0

        col_maquina = safe_col('maquina')
        col_sap = safe_col('numsap') or safe_col('sap')
        col_detalle = safe_col('detalle') or safe_col('producto')
        col_codigo = safe_col('codigo')
        col_faltante = safe_col('faltante')
        col_ciclo = safe_col('ciclo')

        if not all([col_maquina, col_sap, col_faltante]):
            st.error(f"Faltan columnas clave en el Excel. Detectadas: {list(df.columns)}")
            return pd.DataFrame()

        df_limpio = pd.DataFrame()
        df_limpio['Maquina'] = df[col_maquina].astype(str).str.strip()
        df_limpio['OF'] = df[col_sap].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_limpio['Producto'] = df[col_detalle].astype(str).str.strip() if col_detalle else "N/A"
        df_limpio['Codigo_Articulo'] = df[col_codigo].astype(str).str.strip() if col_codigo else "N/A"
        df_limpio['Un_Faltante'] = pd.to_numeric(df[col_faltante], errors='coerce').fillna(0)
        df_limpio['Ciclo_Seg'] = pd.to_numeric(df[col_ciclo], errors='coerce').fillna(60)

        df_limpio = df_limpio.dropna(subset=['Maquina', 'OF'])
        df_limpio = df_limpio[(df_limpio['Maquina'] != 'nan') & (df_limpio['OF'] != 'nan')]
        df_limpio = df_limpio[~df_limpio['Maquina'].str.contains('maquina|máquina|total|id', case=False, na=False)]
        df_limpio = df_limpio[df_limpio['Un_Faltante'] > 0]

        return df_limpio.reset_index(drop=True)

    except Exception as e:
        st.error(f"⚠️ Error leyendo Excel: {e}")
        return pd.DataFrame()

if 'df_planta' not in st.session_state or st.sidebar.button("🔄 Recargar Excel Original"):
    st.session_state.df_planta = cargar_datos_excel(RUTA_EXCEL)

df_base = st.session_state.df_planta.copy()

# ==========================================
# 3. FILTRO MÚLTIPLE DE MÁQUINAS
# ==========================================
df_filtrado = pd.DataFrame()
lista_maquinas = []
if not df_base.empty:
    st.sidebar.write("---")
    st.sidebar.write("### 🔍 Filtro de Inyectoras")
    lista_maquinas = sorted(df_base['Maquina'].unique())
    
    maquinas_seleccionadas = st.sidebar.multiselect(
        "🏭 Selecciona las máquinas a planificar:",
        options=lista_maquinas,
        default=lista_maquinas,
        key="filtro_maestro_1" 
    )
    df_filtrado = df_base[df_base['Maquina'].isin(maquinas_seleccionadas)].copy()
else:
    st.warning("⚠️ Esperando conexión con el archivo de Excel...")

# ==========================================
# 4. CUADRÍCULA INTELIGENTE CON NUMERACIÓN AUTOMÁTICA
# ==========================================
if not df_filtrado.empty:
    
    st.write("---")
    with st.expander("🛑 Insertar Tiempo Muerto / Parada (Sin abrir Excel)"):
        col_p1, col_p2, col_p3, col_p4 = st.columns([2, 1, 3, 2])
        maq_parada = col_p1.selectbox("Máquina a detener:", options=lista_maquinas)
        horas_parada = col_p2.number_input("Horas de parada:", min_value=0.5, value=2.0, step=0.5)
        motivo_parada = col_p3.text_input("Motivo / Nota:", value="Mantenimiento Preventivo")
        
        col_p4.write("") 
        if col_p4.button("➕ Inyectar al Plan", use_container_width=True):
            try:
                df_excel_temp = pd.read_excel(RUTA_EXCEL, engine='openpyxl')
                skip_val = st.session_state.get('fila_salto_excel', 0)
                if skip_val == 1:
                    df_excel_temp = pd.read_excel(RUTA_EXCEL, skiprows=1, engine='openpyxl')
                
                def get_real_col(kw):
                    for c in df_excel_temp.columns:
                        if kw.lower() in str(c).lower().replace(' ', '').replace('.', '').replace('\n', ''): return c
                    return None
                
                c_maq = get_real_col('maquina')
                c_sap = get_real_col('numsap') or get_real_col('sap')
                c_det = get_real_col('detalle') or get_real_col('producto')
                c_cod = get_real_col('codigo')
                c_fal = get_real_col('faltante')
                c_cic = get_real_col('ciclo')
                
                nueva_fila = {col: '' for col in df_excel_temp.columns}
                id_unico = f"PARADA-{datetime.now().strftime('%H%M%S')}"
                
                if c_maq: nueva_fila[c_maq] = maq_parada
                if c_sap: nueva_fila[c_sap] = id_unico
                if c_det: nueva_fila[c_det] = motivo_parada
                if c_cod: nueva_fila[c_cod] = "PARADA-GEN"
                if c_fal: nueva_fila[c_fal] = horas_parada
                if c_cic: nueva_fila[c_cic] = 3600
                
                df_excel_temp = pd.concat([df_excel_temp, pd.DataFrame([nueva_fila])], ignore_index=True)
                
                if skip_val == 1:
                    with pd.ExcelWriter(RUTA_EXCEL, engine='openpyxl', mode='w') as writer:
                        df_excel_temp.to_excel(writer, sheet_name='Sheet1', startrow=1, index=False)
                else:
                    df_excel_temp.to_excel(RUTA_EXCEL, index=False)
                
                st.success("✅ ¡Parada inyectada en el Excel! Recargando...")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Error: {e}")

    if 'Prioridad' not in df_filtrado.columns:
        df_filtrado.insert(0, 'Prioridad', '')
    
    st.write("### 🎛️ Cola de Secuenciación (Planificador APS)")
    st.info("💡 **TRUCO DE GESTIÓN:** ¡Arrastra y suelta! Usa el ícono de los puntitos. La 'Prioridad' (10, 20, 30...) se recalculará automáticamente según el orden visual que dejes.")
    
    gb = GridOptionsBuilder.from_dataframe(df_filtrado)
    gb.configure_default_column(editable=False, groupable=True)
    
    gb.configure_column(
        "Prioridad", 
        rowDrag=True, 
        editable=False, 
        valueGetter="node.rowIndex * 10 + 10", 
        cellStyle={'backgroundColor': '#e8f4f8', 'fontWeight': 'bold', 'color': '#0056b3'}, 
        pinned='left', 
        width=120
    )
    
    gb.configure_selection('multiple', use_checkbox=True, header_checkbox=True)
    
    gridOptions = gb.build()
    gridOptions['rowDragManaged'] = True 
    gridOptions['animateRows'] = True
    
    grid_response = AgGrid(
        df_filtrado, 
        gridOptions=gridOptions, 
        theme="alpine", 
        fit_columns_on_grid_load=True,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED, 
        update_mode=GridUpdateMode.VALUE_CHANGED | GridUpdateMode.SELECTION_CHANGED,
        height=320,
        allow_unsafe_jscode=True 
    )
    
    df_grid = pd.DataFrame(grid_response['data'])
    
    col_b1, col_b2 = st.columns(2)

    # --- BOTÓN 1: GUARDAR ORDEN ---
    if col_b1.button("💾 1. APLICAR NUEVO ORDEN EN EXCEL", type="primary", use_container_width=True):
        if not df_grid.empty and 'OF' in df_grid.columns:
            
            df_grid['Prioridad'] = range(10, len(df_grid)*10 + 1, 10)
            df_grid_clean = df_grid.drop(columns=['Prioridad'])
            
            maquinas_editadas = df_grid_clean['Maquina'].unique()
            df_resto = st.session_state.df_planta[~st.session_state.df_planta['Maquina'].isin(maquinas_editadas)]
            st.session_state.df_planta = pd.concat([df_grid_clean, df_resto], ignore_index=True)
            
            orden_nuevo_of = [str(x) for x in df_grid_clean['OF'].tolist()]
            
            try:
                df_excel_original = pd.read_excel(RUTA_EXCEL, engine='openpyxl')
                skip = st.session_state.get('fila_salto_excel', 0)
                
                df_excel_original.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df_excel_original.columns]
                
                def get_real_col(df_temp, keyword):
                    for col in df_temp.columns:
                        if keyword.lower() in str(col).lower(): return col
                    return None

                if skip == 1:
                    df_datos_reales = pd.read_excel(RUTA_EXCEL, skiprows=1, engine='openpyxl')
                    df_datos_reales.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df_datos_reales.columns]
                else:
                    df_datos_reales = df_excel_original.copy()
                
                real_sap = get_real_col(df_datos_reales, 'numsap') or get_real_col(df_datos_reales, 'sap')
                real_maq = get_real_col(df_datos_reales, 'maquina')

                df_datos_reales['of_str'] = df_datos_reales[real_sap].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df_datos_reales['maquina_str'] = df_datos_reales[real_maq].astype(str).str.strip()
                orden_dict = {of_val: idx for idx, of_val in enumerate(orden_nuevo_of)}
                
                mascara_afectada = df_datos_reales['maquina_str'].isin(maquinas_editadas)
                df_afectado = df_datos_reales[mascara_afectada].copy()
                df_no_afectado = df_datos_reales[~mascara_afectada].copy()
                
                df_afectado['sort_key'] = df_afectado['of_str'].map(orden_dict)
                df_afectado['sort_key'] = df_afectado['sort_key'].fillna(999)
                df_afectado = df_afectado.sort_values('sort_key')
                
                df_excel_nuevo = pd.concat([df_afectado, df_no_afectado], ignore_index=True)
                df_excel_nuevo = df_excel_nuevo.drop(['of_str', 'maquina_str', 'sort_key'], axis=1, errors='ignore')
                
                if skip == 1:
                    with pd.ExcelWriter(RUTA_EXCEL, engine='openpyxl', mode='w') as writer:
                        df_excel_nuevo.to_excel(writer, sheet_name='Sheet1', startrow=1, index=False)
                else:
                    df_excel_nuevo.to_excel(RUTA_EXCEL, index=False)
                
                st.success("🚀 ¡Orden guardado con éxito! Gantt y secuencias automatizadas.")
                st.cache_data.clear()
                st.rerun()

            except PermissionError:
                st.error("❌ Archivo abierto en tu PC. Ciérralo en Windows para permitir que Python guarde los cambios.")
            except Exception as e:
                st.error(f"⚠️ Error al guardar: {str(e)}")

    # --- BOTÓN 2: BORRAR SELECCIÓN ---
    if col_b2.button("🗑️ 2. BORRAR SELECCIÓN DEL EXCEL", type="secondary", use_container_width=True):
        filas_seleccionadas = grid_response.get('selected_rows', [])
        
        ofs_a_borrar = []
        if isinstance(filas_seleccionadas, pd.DataFrame) and not filas_seleccionadas.empty:
            ofs_a_borrar = [str(x) for x in filas_seleccionadas['OF'].tolist()]
        elif isinstance(filas_seleccionadas, list) and len(filas_seleccionadas) > 0:
            ofs_a_borrar = [str(r.get('OF')) for r in filas_seleccionadas if 'OF' in r]
            
        if ofs_a_borrar:
            try:
                st.session_state.df_planta = st.session_state.df_planta[~st.session_state.df_planta['OF'].isin(ofs_a_borrar)]
                
                df_excel_original = pd.read_excel(RUTA_EXCEL, engine='openpyxl')
                skip = st.session_state.get('fila_salto_excel', 0)
                
                def get_real_col(df_temp, keyword):
                    for col in df_temp.columns:
                        if keyword.lower() in str(col).lower(): return col
                    return None

                if skip == 1:
                    df_datos_reales = pd.read_excel(RUTA_EXCEL, skiprows=1, engine='openpyxl')
                    df_datos_reales.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df_datos_reales.columns]
                else:
                    df_datos_reales = df_excel_original.copy()
                
                real_sap = get_real_col(df_datos_reales, 'numsap') or get_real_col(df_datos_reales, 'sap')
                df_datos_reales['of_str'] = df_datos_reales[real_sap].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                
                df_excel_nuevo = df_datos_reales[~df_datos_reales['of_str'].isin(ofs_a_borrar)].copy()
                df_excel_nuevo = df_excel_nuevo.drop(['of_str'], axis=1, errors='ignore')
                
                if skip == 1:
                    with pd.ExcelWriter(RUTA_EXCEL, engine='openpyxl', mode='w') as writer:
                        df_excel_nuevo.to_excel(writer, sheet_name='Sheet1', startrow=1, index=False)
                else:
                    df_excel_nuevo.to_excel(RUTA_EXCEL, index=False)
                
                st.success(f"🗑️ Se eliminaron {len(ofs_a_borrar)} registro(s) del Excel exitosamente.")
                st.cache_data.clear()
                st.rerun()

            except PermissionError:
                st.error("❌ El Excel está abierto en tu PC. Ciérralo e intenta de nuevo.")
            except Exception as e:
                st.error(f"⚠️ Error al borrar: {e}")
        else:
            st.warning("⚠️ Selecciona al menos una fila marcando la casilla en la tabla antes de borrar.")

# ==========================================
# 5. MOTOR LÓGICO DE GANTT (CON PARADAS)
# ==========================================
df_calculado = []
if not df_filtrado.empty:
    df_gantt_origen = st.session_state.df_planta[st.session_state.df_planta['Maquina'].isin(maquinas_seleccionadas)]
    
    for maquina, grupo in df_gantt_origen.groupby("Maquina", sort=False):
        tiempo_actual = FECHA_HORA_COMBINADA  
        ultimo_molde = None
        ultimo_color = None
        
        for _, row in grupo.iterrows():
            
            es_parada = str(row["OF"]).upper().startswith(("PARADA", "MANT", "PAUSA", "INACTIVO"))
            
            if es_parada:
                if float(row["Un_Faltante"]) > 0 and float(row["Ciclo_Seg"]) > 0:
                    horas_produccion = (float(row["Ciclo_Seg"]) * float(row["Un_Faltante"])) / 3600
                    fin_prod = tiempo_actual + timedelta(hours=horas_produccion)
                    
                    df_calculado.append({
                        "Máquina": maquina, 
                        "Tarea": f"🛑 PARADA | {row['Producto'][:30]}", 
                        "Inicio": tiempo_actual, 
                        "Fin": fin_prod, 
                        "Tipo": "Tiempo Inactivo / Parada", 
                        "OF": str(row["OF"]),
                        "Cantidad": row["Un_Faltante"],
                        "Ciclo": row["Ciclo_Seg"]
                    })
                    tiempo_actual = fin_prod
                continue 
            
            molde_actual = str(row["Codigo_Articulo"])[:4]
            partes_codigo = str(row["Codigo_Articulo"]).split('-')
            color_actual = partes_codigo[-1] if len(partes_codigo) > 1 else "STD"

            if ultimo_molde is not None and molde_actual != ultimo_molde:
                fin_setup = tiempo_actual + timedelta(minutes=TIEMPO_CAMBIO_MOLDE)
                df_calculado.append({
                    "Máquina": maquina, 
                    "Tarea": f"🔧 CAMBIO MOLDE: {ultimo_molde} ➡️ {molde_actual}", 
                    "Inicio": tiempo_actual, 
                    "Fin": fin_setup, 
                    "Tipo": "Cambio de Molde", 
                    "OF": "SETUP",
                    "Cantidad": 0,
                    "Ciclo": 0
                })
                tiempo_actual = fin_setup
                ultimo_color = None 

            elif ultimo_color is not None and color_actual != ultimo_color:
                fin_color = tiempo_actual + timedelta(minutes=TIEMPO_CAMBIO_COLOR)
                df_calculado.append({
                    "Máquina": maquina, 
                    "Tarea": f"🎨 PURGA COLOR: {ultimo_color} ➡️ {color_actual}", 
                    "Inicio": tiempo_actual, 
                    "Fin": fin_color, 
                    "Tipo": "Cambio de Color", 
                    "OF": "PURGA",
                    "Cantidad": 0,
                    "Ciclo": 0
                })
                tiempo_actual = fin_color

            if float(row["Un_Faltante"]) > 0 and float(row["Ciclo_Seg"]) > 0:
                horas_produccion = (float(row["Ciclo_Seg"]) * float(row["Un_Faltante"])) / 3600
                fin_prod = tiempo_actual + timedelta(hours=horas_produccion)
                
                df_calculado.append({
                    "Máquina": maquina, 
                    "Tarea": f"📦 OF {row['OF']} | {row['Producto'][:30]}...", 
                    "Inicio": tiempo_actual, 
                    "Fin": fin_prod, 
                    "Tipo": "Producción Activa", 
                    "OF": str(row["OF"]),
                    "Cantidad": row["Un_Faltante"],
                    "Ciclo": row["Ciclo_Seg"]
                })
                
                tiempo_actual = fin_prod
                ultimo_molde = molde_actual
                ultimo_color = color_actual

    df_gantt = pd.DataFrame(df_calculado)

    # ==========================================
    # 6. GANTT INTERACTIVO
    # ==========================================
    if not df_gantt.empty:
        st.write("---")
        st.write("### 🔎 Rango de Vista del Gráfico y Análisis de Montadores")
        
        col_z1, col_z2, col_z3, col_z4 = st.columns(4)
        if "zoom_fd1" not in st.session_state:
            st.session_state.zoom_fd1 = FECHA_ARRANQUE
            st.session_state.zoom_fh1 = HORA_ARRANQUE
            st.session_state.zoom_fd2 = FECHA_ARRANQUE + timedelta(days=2)
            st.session_state.zoom_fh2 = HORA_ARRANQUE

        vista_fecha_inicio = col_z1.date_input("Ver desde (Fecha):", key="zoom_fd1")
        vista_hora_inicio = col_z2.time_input("Ver desde (Hora):", key="zoom_fh1")
        vista_fecha_fin = col_z3.date_input("Ver hasta (Fecha):", key="zoom_fd2")
        vista_hora_fin = col_z4.time_input("Ver hasta (Hora):", key="zoom_fh2")
        
        rango_inicio_dt = pd.to_datetime(datetime.combine(vista_fecha_inicio, vista_hora_inicio))
        rango_fin_dt = pd.to_datetime(datetime.combine(vista_fecha_fin, vista_hora_fin))

        st.write("### 📊 Línea de Tiempo de Producción Detallada")
        
        altura_dinamica = max(350, len(df_gantt["Máquina"].unique()) * 75) 
        fig = px.timeline(
            df_gantt, 
            x_start="Inicio", 
            x_end="Fin", 
            y="Máquina", 
            color="Tipo", 
            text="OF", 
            hover_name="Tarea",
            hover_data={"Cantidad": True, "Ciclo": True, "Inicio": "|%d %b %H:%M", "Fin": "|%d %b %H:%M"},
            color_discrete_map={
                "Producción Activa": "#1f77b4", 
                "Cambio de Molde": "#d62728", 
                "Cambio de Color": "#ff7f0e",
                "Tiempo Inactivo / Parada": "#808080" 
            }
        )

        ahora = datetime.now()
        fig.add_shape(type="line", x0=ahora, x1=ahora, y0=0, y1=1, xref="x", yref="paper", line=dict(color="red", width=3, dash="dash"))
        fig.add_annotation(x=ahora, y=1.03, yref="paper", xref="x", text="<b>⏰ TIEMPO REAL AHORA</b>", showarrow=False, font=dict(color="red", size=11))

        fig.update_layout(yaxis_title="Inyectoras", height=altura_dinamica, margin=dict(l=10, r=10, t=40, b=10))
        fig.update_xaxes(range=[rango_inicio_dt, rango_fin_dt], showgrid=True, tickformat="%d %b\n%H:%M", rangeslider_visible=True)
        st.plotly_chart(fig, use_container_width=True)

        # ==========================================
        # 7. ANÁLISIS DE SATURACIÓN OPERATIVA EN TURNO
        # ==========================================
        st.write("---")
        st.write("### 🛠️ Análisis de Cargas Sincronizado - Montadores de Molde")
        
        df_moldes = df_gantt[df_gantt["Tipo"] == "Cambio de Molde"].copy()
        
        if not df_moldes.empty:
            df_moldes['Inicio'] = pd.to_datetime(df_moldes['Inicio'])
            df_moldes['Fin'] = pd.to_datetime(df_moldes['Fin'])

            cambios_en_turno = df_moldes[
                (df_moldes['Inicio'] >= rango_inicio_dt) & 
                (df_moldes['Inicio'] <= rango_fin_dt)
            ].copy()

            eventos = []
            for _, row in cambios_en_turno.iterrows():
                eventos.append((row["Inicio"], 1))   
                eventos.append((row["Fin"], -1))     

            eventos.sort(key=lambda x: (x[0], x[1]))
            max_pico = 0
            actual = 0
            tiempo_pico = None

            for tiempo, cambio in eventos:
                actual += cambio
                if actual > max_pico:
                    max_pico = actual
                    tiempo_pico = tiempo

            col1, col2, col3 = st.columns(3)
            col1.metric("Moldes que inician en el rango", len(cambios_en_turno))
            col2.metric("Pico de Personal en Simultáneo", f"{max_pico} montadores")
            col3.metric("Momento crítico estimado", tiempo_pico.strftime('%d %b - %H:%M') if tiempo_pico else "N/A")

            with st.expander("📋 Desplegar lista auditora de montajes del rango"):
                if not cambios_en_turno.empty:
                    df_auditor = cambios_en_turno[['Máquina', 'Tarea', 'Inicio', 'Fin']].sort_values('Inicio')
                    df_auditor['Inicio'] = df_auditor['Inicio'].dt.strftime('%d/%m %H:%M')
                    df_auditor['Fin'] = df_auditor['Fin'].dt.strftime('%d/%m %H:%M')
                    st.dataframe(df_auditor, use_container_width=True, hide_index=True)
                else:
                    st.info("Sin registros de cambios para mostrar en este bloque.")
        else:
            st.info("✅ No existen órdenes de cambio de molde programadas en este horizonte.")
