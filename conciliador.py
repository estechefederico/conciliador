import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Conciliador PRO", layout="wide")
st.title("⚡ Conciliador PRO (Motor de Reglas + Autoguardado)")

# --- 0. FUNCIONES DE AUTOGUARDADO (NUEVO) ---
ARCHIVO_BANCO = "backup_banco.pkl"
ARCHIVO_LIBRO = "backup_libro.pkl"
ARCHIVO_CONCIL = "backup_conciliados.pkl"

def guardar_backup():
    """Guarda el estado actual en archivos invisibles de la PC"""
    try:
        if st.session_state.df_banco is not None:
            st.session_state.df_banco.to_pickle(ARCHIVO_BANCO)
        if st.session_state.df_libro is not None:
            st.session_state.df_libro.to_pickle(ARCHIVO_LIBRO)
        if st.session_state.df_conciliados is not None:
            st.session_state.df_conciliados.to_pickle(ARCHIVO_CONCIL)
    except Exception as e:
        st.toast(f"Error guardando backup: {e}")

def cargar_backup():
    """Recupera los archivos guardados tras un corte de luz"""
    try:
        if os.path.exists(ARCHIVO_BANCO): st.session_state.df_banco = pd.read_pickle(ARCHIVO_BANCO)
        if os.path.exists(ARCHIVO_LIBRO): st.session_state.df_libro = pd.read_pickle(ARCHIVO_LIBRO)
        if os.path.exists(ARCHIVO_CONCIL): st.session_state.df_conciliados = pd.read_pickle(ARCHIVO_CONCIL)
        st.success("¡Sesión recuperada con éxito!")
    except Exception as e:
        st.error(f"No se pudo recuperar la sesión: {e}")

def borrar_backup():
    """Limpia los archivos de rescate al reiniciar"""
    for archivo in [ARCHIVO_BANCO, ARCHIVO_LIBRO, ARCHIVO_CONCIL]:
        if os.path.exists(archivo):
            try: os.remove(archivo)
            except: pass

# --- 1. FUNCIONES BASE ---
def limpiar_monto_v15(serie, formato_elegido):
    if pd.api.types.is_numeric_dtype(serie): return serie
    serie = serie.astype(str).str.replace('$', '', regex=False).str.replace(' ', '', regex=False)
    serie = serie.str.replace('USD', '', regex=False).str.replace('ARS', '', regex=False)
    if formato_elegido == "Mercado Pago (Puntos: -1181.67)": serie = serie.str.replace(',', '', regex=False)
    else: serie = serie.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(serie, errors='coerce').fillna(0)

def unificar_archivos(lista_archivos):
    df_total = pd.DataFrame()
    for archivo in lista_archivos:
        try:
            if archivo.name.endswith('.csv'): df_temp = pd.read_csv(archivo, on_bad_lines='skip', engine='python')
            else: df_temp = pd.read_excel(archivo)
            df_total = pd.concat([df_total, df_temp], ignore_index=True)
        except Exception as e: st.error(f"Error técnico leyendo {archivo.name}: {e}")
    return df_total

def normalizar_df(df, origen):
    if "Conciliar" not in df.columns: df.insert(0, "Conciliar", False)
    df["Origen_Dato"] = origen
    df["_ID_Interno"] = range(1, len(df) + 1)
    return df

def generar_excel_completo(df_conciliados, df_pendiente_banco, df_pendiente_libro):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_conciliados.empty: df_conciliados.to_excel(writer, index=False, sheet_name='1. Conciliados')
        cols_borrar = ['Conciliar', '_ID_Interno', 'Ayuda']
        if df_pendiente_banco is not None: df_pendiente_banco.drop(columns=cols_borrar, errors='ignore').to_excel(writer, index=False, sheet_name='2. Pendientes Banco')
        if df_pendiente_libro is not None: df_pendiente_libro.drop(columns=cols_borrar, errors='ignore').to_excel(writer, index=False, sheet_name='3. Pendientes Libro')
    return output.getvalue()

# --- 2. MOTOR DE REGLAS (NÚCLEO) ---
def get_clean_val_v15(val, tipo, formato_elegido):
    if pd.isna(val): return None if tipo == "Fecha Exacta" else ""
    
    if tipo == "Monto Exacto":
        s = str(val).replace('$', '').replace(' ', '').replace('USD', '').replace('ARS', '')
        if formato_elegido == "Mercado Pago (Puntos: -1181.67)": s = s.replace(',', '')
        else: s = s.replace('.', '').replace(',', '.')
        try: return round(float(s), 2)
        except: return 0.0
        
    elif tipo == "Fecha Exacta":
        try:
            d = pd.to_datetime(val, dayfirst=True)
            if pd.isna(d): return None
            return d.date()
        except: return None
            
    else: 
        s = str(val).strip().lower()
        if s.endswith('.0'): s = s[:-2] 
        return s if s not in ['nan', 'nat'] else ""

def es_valida(tupla_valores):
    for v in tupla_valores:
        if v != "" and v != 0.0 and v is not None: return True
    return False

def ejecutar_auto_conciliacion(reglas, col_mb, col_ml, formato_elegido):
    df_b = st.session_state.df_banco.copy()
    df_l = st.session_state.df_libro.copy()
    
    match_base = datetime.now().strftime("%Y%m%d%H%M")
    parejas, ids_b_borrar, ids_l_borrar = [], [], []
    lista_b = df_b.to_dict('records')
    lista_l = df_l.to_dict('records')
    usados_l = set()
    contador = 0
    
    for item_b in lista_b:
        tup_b = tuple(get_clean_val_v15(item_b.get(cb), tipo, formato_elegido) for cb, cl, tipo in reglas)
        if not es_valida(tup_b): continue
        
        for idx_l, item_l in enumerate(lista_l):
            if idx_l not in usados_l:
                tup_l = tuple(get_clean_val_v15(item_l.get(cl), tipo, formato_elegido) for cb, cl, tipo in reglas)
                
                if tup_b == tup_l:
                    usados_l.add(idx_l)
                    match_id = f"{match_base}_{contador}"
                    
                    rb, rl = pd.DataFrame([item_b]), pd.DataFrame([item_l])
                    rb['ID_Cruce'], rl['ID_Cruce'] = match_id, match_id
                    rb['Monto_Op'] = item_b[col_mb]
                    rl['Monto_Op'] = item_l[col_ml]
                    
                    parejas.extend([rb, rl])
                    ids_b_borrar.append(item_b['_ID_Interno'])
                    ids_l_borrar.append(item_l['_ID_Interno'])
                    contador += 1
                    break

    if contador > 0:
        nuevo_match = pd.concat(parejas, ignore_index=True)
        st.session_state.df_conciliados = pd.concat([st.session_state.df_conciliados, nuevo_match], ignore_index=True)
        st.session_state.df_banco = st.session_state.df_banco[~st.session_state.df_banco['_ID_Interno'].isin(ids_b_borrar)]
        st.session_state.df_libro = st.session_state.df_libro[~st.session_state.df_libro['_ID_Interno'].isin(ids_l_borrar)]
        guardar_backup() # <-- AUTOGUARDADO TRAS CONCILIAR
        return contador
    return 0

# --- 3. CONFIGURACIÓN E INICIALIZACIÓN ---
if 'df_banco' not in st.session_state: st.session_state.df_banco = None
if 'df_libro' not in st.session_state: st.session_state.df_libro = None
if 'df_conciliados' not in st.session_state: st.session_state.df_conciliados = pd.DataFrame()
if 'num_reglas' not in st.session_state: st.session_state.num_reglas = 1

with st.sidebar:
    st.header("🔧 Formato Base")
    opcion_formato = st.radio("Lectura de Decimales:", ("Excel Arg (-1.181,67)", "Mercado Pago (Puntos: -1181.67)"), index=1)
    st.divider()
    st.metric("Conciliados (Parejas)", len(st.session_state.df_conciliados) // 2)
    
    if st.session_state.df_banco is not None:
        excel = generar_excel_completo(st.session_state.df_conciliados, st.session_state.df_banco, st.session_state.df_libro)
        st.download_button("📥 Descargar Excel", data=excel, file_name='Conciliacion.xlsx')
    
    if st.button("⚠️ Reiniciar Todo (Borrar Progreso)"):
        st.session_state.clear()
        borrar_backup() # <-- LIMPIA BACKUPS AL REINICIAR
        st.rerun()

# --- 4. CARGA DE ARCHIVOS / RECUPERACIÓN ---
if st.session_state.df_banco is None:
    # AVISO DE RECUPERACIÓN SI HAY ARCHIVOS GUARDADOS
    if os.path.exists(ARCHIVO_BANCO) or os.path.exists(ARCHIVO_LIBRO):
        st.warning("⚠️ Parece que se cerró el programa sin descargar el Excel final.")
        if st.button("🆘 Recuperar Trabajo Anterior", type="primary", use_container_width=True):
            cargar_backup()
            st.rerun()
        st.markdown("---")

    c1, c2 = st.columns(2)
    files_b = c1.file_uploader("Banco (Uno o Varios)", accept_multiple_files=True)
    files_l = c2.file_uploader("Libro (Uno o Varios)", accept_multiple_files=True)

    if st.button("Procesar Archivos", type="primary"):
        if files_b and files_l:
            d1 = unificar_archivos(files_b)
            d2 = unificar_archivos(files_l)
            if not d1.empty and not d2.empty:
                st.session_state.df_banco = normalizar_df(d1, "BANCO")
                st.session_state.df_libro = normalizar_df(d2, "LIBRO")
                guardar_backup() # <-- AUTOGUARDADO TRAS CARGAR ARCHIVOS
                st.rerun()

# --- 5. ÁREA DE TRABAJO ---
else:
    ignorar = ['Conciliar', '_ID_Interno', 'Origen_Dato', 'Ayuda']
    cols_b = [c for c in st.session_state.df_banco.columns if c not in ignorar]
    cols_l = [c for c in st.session_state.df_libro.columns if c not in ignorar]

    with st.expander("💰 1. Seleccionar Importes (Para la suma inferior)", expanded=False):
        c1, c2 = st.columns(2)
        col_mb = c1.selectbox("Importe Banco", cols_b, index=len(cols_b)-1)
        col_ml = c2.selectbox("Importe Libro", cols_l, index=len(cols_l)-1)
    
    st.session_state.df_banco[col_mb] = limpiar_monto_v15(st.session_state.df_banco[col_mb], opcion_formato)
    st.session_state.df_libro[col_ml] = limpiar_monto_v15(st.session_state.df_libro[col_ml], opcion_formato)

    st.markdown("### ⚙️ 2. Reglas de Conciliación Estricta")
    reglas = []
    for i in range(st.session_state.num_reglas):
        cx1, cx2, cx3 = st.columns([2, 2, 1])
        cb = cx1.selectbox(f"Columna Banco (Regla {i+1})", cols_b, key=f"cb_{i}")
        cl = cx2.selectbox(f"Columna Libro (Regla {i+1})", cols_l, key=f"cl_{i}")
        tipo = cx3.selectbox("¿Qué tipo de dato es?", ["Monto Exacto", "Texto Exacto", "Fecha Exacta"], key=f"tipo_{i}")
        reglas.append((cb, cl, tipo))
        
    c_btn1, c_btn2, c_btn3 = st.columns([1, 1, 3])
    if c_btn1.button("➕ Agregar Condición"):
        st.session_state.num_reglas += 1
        st.rerun()
    if c_btn2.button("➖ Quitar Condición") and st.session_state.num_reglas > 1:
        st.session_state.num_reglas -= 1
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button(f"🚀 Ejecutar Conciliación ({len(reglas)} Regla/s)", type="primary"):
        cantidad = ejecutar_auto_conciliacion(reglas, col_mb, col_ml, opcion_formato)
        if cantidad > 0: st.success(f"¡Éxito! Se conciliaron {cantidad} partidas.")
        else: st.warning("No se encontraron coincidencias bajo estas reglas.")
        st.rerun()

    st.divider()

    # --- SEMÁFORO ---
    libro_tuples = set()
    for _, row in st.session_state.df_libro.iterrows():
        t = tuple(get_clean_val_v15(row.get(cl), tipo, opcion_formato) for cb, cl, tipo in reglas)
        if es_valida(t): libro_tuples.add(t)
            
    banco_tuples = set()
    for _, row in st.session_state.df_banco.iterrows():
        t = tuple(get_clean_val_v15(row.get(cb), tipo, opcion_formato) for cb, cl, tipo in reglas)
        if es_valida(t): banco_tuples.add(t)

    def sem_banco(row):
        t = tuple(get_clean_val_v15(row.get(cb), tipo, opcion_formato) for cb, cl, tipo in reglas)
        return "🟢" if t in libro_tuples and es_valida(t) else ""

    def sem_libro(row):
        t = tuple(get_clean_val_v15(row.get(cl), tipo, opcion_formato) for cb, cl, tipo in reglas)
        return "🟢" if t in banco_tuples and es_valida(t) else ""

    st.session_state.df_banco['Ayuda'] = st.session_state.df_banco.apply(sem_banco, axis=1)
    st.session_state.df_libro['Ayuda'] = st.session_state.df_libro.apply(sem_libro, axis=1)

    ord_b = ["Conciliar", "Ayuda"] + cols_b
    ord_l = ["Conciliar", "Ayuda"] + cols_l

    c_izq, c_der = st.columns(2)
    with c_izq:
        st.subheader(f"🏦 Banco ({len(st.session_state.df_banco)})")
        ed_b = st.data_editor(st.session_state.df_banco, key="ed_b", hide_index=True, height=400, column_order=ord_b, column_config={"Conciliar": st.column_config.CheckboxColumn(width="small"), "Ayuda": st.column_config.Column(width="small")})

    with c_der:
        st.subheader(f"📖 Libro ({len(st.session_state.df_libro)})")
        ed_l = st.data_editor(st.session_state.df_libro, key="ed_l", hide_index=True, height=400, column_order=ord_l, column_config={"Conciliar": st.column_config.CheckboxColumn(width="small"), "Ayuda": st.column_config.Column(width="small")})

    sel_b = ed_b[ed_b["Conciliar"]]
    sel_l = ed_l[ed_l["Conciliar"]]
    sb = sel_b[col_mb].sum()
    sl = sel_l[col_ml].sum()
    dif = sb - sl

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Suma Banco Tildado", f"${sb:,.2f}")
    m2.metric("Suma Libro Tildado", f"${sl:,.2f}")
    m3.metric("Diferencia", f"${dif:,.2f}", delta_color="normal" if abs(dif)<0.01 else "inverse")
    
    if m4.button("✨ CONCILIAR MANUAL (Lo seleccionado)", disabled=not(abs(dif)<0.01 and (len(sel_b)>0 or len(sel_l)>0)), use_container_width=True):
        mid = datetime.now().strftime("%Y%m%d%H%M%S")
        b_s, l_s = sel_b.copy(), sel_l.copy()
        b_s['ID_Cruce'], l_s['ID_Cruce'] = mid, mid
        b_s['Monto_Op'] = b_s[col_mb]
        l_s['Monto_Op'] = l_s[col_ml]
        full = pd.concat([b_s, l_s], ignore_index=True)
        st.session_state.df_conciliados = pd.concat([st.session_state.df_conciliados, full], ignore_index=True)
        ids_b, ids_l = sel_b['_ID_Interno'], sel_l['_ID_Interno']
        st.session_state.df_banco = st.session_state.df_banco[~st.session_state.df_banco['_ID_Interno'].isin(ids_b)]
        st.session_state.df_libro = st.session_state.df_libro[~st.session_state.df_libro['_ID_Interno'].isin(ids_l)]
        guardar_backup() # <-- AUTOGUARDADO TRAS CONCILIAR MANUAL
        st.rerun()
