import streamlit as st
import pandas as pd
import os
from PIL import Image
from pypdf import PdfReader
from servicios.exportador import generar_excel, generar_pdf
from servicios.whatsapp_service import generar_enlace_whatsapp

# ===================================================================
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS (CSS)
# ===================================================================
st.set_page_config(
    page_title="Sala de Autogobierno - Comuna Socialista El Paraíso",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Estilos Generales del Tema */
    .main { 
        background-color: #F1F5F9; 
    }
    
    /* Sidebar Profesional */
    [data-testid="stSidebar"] {
        background-color: #0F172A;
        color: #F8FAFC;
    }
    [data-testid="stSidebar"] .stRadio label {
        color: #E2E8F0 !important;
        font-weight: 500;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {
        color: #F8FAFC !important;
    }

    /* Tarjetas de Métricas Elegantes */
    .stMetric { 
        background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
        padding: 18px; 
        border-radius: 12px; 
        border-left: 5px solid #0284C7;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .stMetric label {
        color: #64748B !important;
        font-weight: 600;
    }
    
    /* Botones de Acción Personalizados */
    .stButton>button {
        background-color: #0284C7;
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: 600;
    }
    .stButton>button:hover {
        background-color: #0369A1;
        color: white;
    }

    /* Pestañas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #E2E8F0;
        border-radius: 6px 6px 0px 0px;
        color: #334155;
        font-weight: 600;
        padding: 10px 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0284C7 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)


# ===================================================================
# 2. FUNCIONES DE CARGA Y GESTIÓN DE ARCHIVOS EXCEL / PDF
# ===================================================================
RUTAS = {
    "INCIDENCIAS": "data/Incidencias.xlsx",
    "PROYECTOS": "data/Proyectos.xlsx",
    "ESTRUCTURAS": "data/Estructuras.xlsx",
    "ACA": "data/ACA_Agenda.xlsx",
    "CONSEJOS_COMUNALES": "data/Consejos_Comunales.xlsx",
    "VOCEROS": "data/Voceros.xlsx",
    "LOGISTICA": "data/Logistica.xlsx",
    "UBCH_DATOS": "data/UBCH_Datos.xlsx",
    "CCBI": "data/CCBI_Miembros.xlsx",
    "CBBI": "data/CBBI_Miembros.xlsx",
    "LOGO": "data/logo_comuna.png",
    "UBCH_JEFE": "data/UBCH_Jefe.xlsx",
    "UBCH_ESTRUCTURA": "data/UBCH_Estructura_Miembros.xlsx"   
}
data_ubch_jefe = pd.DataFrame(columns=["Nombre UBCH", "Nombre y Apellido", "Cédula", "Teléfono"])
data_ubch_estru = pd.DataFrame(columns=["Nombre UBCH", "Cargo", "Nombre y Apellido", "Cédula", "Teléfono"])
data_ubch_detalles = pd.DataFrame()
data_estructuras = pd.DataFrame()
data_ccbi = pd.DataFrame()
data_cbbi = pd.DataFrame()

# Intentar cargar los archivos reales si existen
try:
    if "UBCH_JEFE" in RUTAS:
        data_ubch_jefe = pd.read_excel(RUTAS["UBCH_JEFE"])
except Exception:
    pass

try:
    if "UBCH_ESTRUCTURA" in RUTAS:
        data_ubch_estru = pd.read_excel(RUTAS["UBCH_ESTRUCTURA"])
except Exception:
    pass
def cargar_datos_incidencias():
    asegurar_directorio()
    if os.path.exists(RUTAS["INCIDENCIAS"]):
        df = pd.read_excel(RUTAS["INCIDENCIAS"])
        if "Nivel de Severidad" in df.columns and "Severidad" not in df.columns:
            df.rename(columns={"Nivel de Severidad": "Severidad"}, inplace=True)
    else:
        df = pd.DataFrame({
            "ID": [101, 102],
            "Zona / Sector": ["Sector Norte", "Sector Sur"],
            "Tipo de Incidencia": ["Servicios Públicos", "Vialidad"],
            "Severidad": ["Alta", "Media"],
            "Estado": ["Pendiente", "En Proceso"],
            "Observaciones": ["Falla en tubería", "Bacheo pendiente"]
        })
        df.to_excel(RUTAS["INCIDENCIAS"], index=False)
    for col in df.columns:
        df[col] = df[col].astype(str)
    return df

def cargar_datos_proyectos():
    asegurar_directorio()
    if os.path.exists(RUTAS["PROYECTOS"]):
        df = pd.read_excel(RUTAS["PROYECTOS"])
    else:
        df = pd.DataFrame({
            "ID Proyecto": ["PRY-001", "PRY-002"],
            "Nombre del Proyecto": ["Rehabilitación de Cancha", "Alumbrado Público LED"],
            "Sector": ["Sector Norte", "Centro Comunal"],
            "Presupuesto ($)": [4500.0, 2800.0],
            "Estatus": ["En Ejecución", "En Planificación"],
            "Responsable": ["Consejo Comunal A", "Mesa Técnica de Energía"]
        })
        df.to_excel(RUTAS["PROYECTOS"], index=False)
    for col in ["ID Proyecto", "Nombre del Proyecto", "Sector", "Estatus", "Responsable"]:
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

def cargar_datos_estructuras():
    asegurar_directorio()
    if os.path.exists(RUTAS["ESTRUCTURAS"]):
        df = pd.read_excel(RUTAS["ESTRUCTURAS"])
    else:
        df = pd.DataFrame({
            "Estructura": ["Sala de Gobierno", "UBCH", "Consejos Comunales"],
            "Nombre y Apellido": ["Juan Pérez", "María Rodríguez", "Carlos Gómez"],
            "Cédula": ["V-12345678", "V-87654321", "V-11223344"],
            "Cargo / Responsabilidad": ["Vocero Principal", "Jefe de UBCH", "Vocero Finanzas"],
            "Teléfono": ["04121112233", "04142223344", "04163334455"],
            "Comunidad / Sector": ["Centro", "Sector Norte", "Sector Sur"]
        })
        df.to_excel(RUTAS["ESTRUCTURAS"], index=False)
    for col in df.columns:
        df[col] = df[col].astype(str)
    return df

def cargar_datos_aca():
    asegurar_directorio()
    if os.path.exists(RUTAS["ACA"]):
        df = pd.read_excel(RUTAS["ACA"])
    else:
        df = pd.DataFrame({
            "Comunidad": ["Comunidad El Paraíso Norte"],
            "Gabinete": ["Servicios Públicos"],
            "Área / Comité": ["Agua Potable"],
            "Descripción": ["Sustitución de tuberías"],
            "Nudo Crítico": ["Baja presión de agua"],
            "Potencialidades": ["Mano de obra comunitaria"],
            "Qué Haremos (Autogestión)": ["Zanjado"],
            "Con Ayuda de Otros (Congestión)": ["Suministro de tubos"],
            "Programación (Gestión)": ["Q3 2026"],
            "Responsables": ["Mesa Técnica de Agua"]
        })
        df.to_excel(RUTAS["ACA"], index=False)
    for col in df.columns:
        df[col] = df[col].astype(str)
    return df

def cargar_datos_consejos_comunales():
    asegurar_directorio()
    if os.path.exists(RUTAS["CONSEJOS_COMUNALES"]):
        df_cc = pd.read_excel(RUTAS["CONSEJOS_COMUNALES"])
    else:
        df_cc = pd.DataFrame({
            "Código SITUR": ["CC-010203001", "CC-010203002"],
            "Nombre del Consejo Comunal": ["CC Vencedores del Norte", "CC Centro Comunal"],
            "Comunidad": ["Sector Norte", "Centro Comunal"],
            "Fecha Elección": ["2023-05-15", "2022-11-10"],
            "Fecha Vencimiento": ["2025-05-15", "2024-11-10"],
            "Estatus Vocería": ["Vigente", "Vencido"]
        })
        df_cc.to_excel(RUTAS["CONSEJOS_COMUNALES"], index=False)
        
    if os.path.exists(RUTAS["VOCEROS"]):
        df_voceros = pd.read_excel(RUTAS["VOCEROS"])
        if "Condición" not in df_voceros.columns:
            df_voceros["Condición"] = "Principal"
    else:
        df_voceros = pd.DataFrame({
            "Código SITUR": ["CC-010203001", "CC-010203002"],
            "Nombre y Apellido": ["Pedro Infante", "José Félix Ribas"],
            "Cédula": ["V-15987654", "V-14567890"],
            "Teléfono": ["04129998877", "04167776655"],
            "Vocería / Comité": ["Unidad Administrativa", "Contraloría Social"],
            "Condición": ["Principal", "Principal"]
        })
        df_voceros.to_excel(RUTAS["VOCEROS"], index=False)
        
    for col in df_cc.columns:
        df_cc[col] = df_cc[col].astype(str)
    for col in df_voceros.columns:
        df_voceros[col] = df_voceros[col].astype(str)
    return df_cc, df_voceros

def cargar_datos_logistica():
    asegurar_directorio()
    if os.path.exists(RUTAS["LOGISTICA"]):
        df = pd.read_excel(RUTAS["LOGISTICA"])
    else:
        df = pd.DataFrame({
            "Recurso": ["Combustible (L)", "Raciones"],
            "Disponible": [1200, 450],
            "Nivel_Critico": ["No", "No"]
        })
        df.to_excel(RUTAS["LOGISTICA"], index=False)
    for col in df.columns:
        df[col] = df[col].astype(str)
    return df

def cargar_datos_ubch_extra():
    asegurar_directorio()
    if os.path.exists(RUTAS["UBCH_DATOS"]):
        df_ubch = pd.read_excel(RUTAS["UBCH_DATOS"])
    else:
        df_ubch = pd.DataFrame({
            "Nombre UBCH": ["UBCH Centro Educativo El Paraíso", "UBCH Grupo Escolar Simón Bolívar"],
            "Consejos Comunales Asociados": ["CC Vencedores del Norte, CC Centro Comunal", "CC Sur Unido"],
            "Zonas en Agregación": ["Sector Las Acacias, Calle Bolívar", "Sector La Línea"],
            "Cantidad Votantes": [2450, 1890],
            "Cantidad Mesas": [4, 3],
            "Jefe o Jefa de Ubch" : ["Tibisay Rodriguez"],
            "Cedula de Identidad" : ["10288206"],
            "Telefono" : ["04122148330"],
        })
        df_ubch.to_excel(RUTAS["UBCH_DATOS"], index=False)

    # CCBI (Comando de Comunidad de Base Integral)
    if os.path.exists(RUTAS["CCBI"]):
        df_ccbi = pd.read_excel(RUTAS["CCBI"])
    else:
        df_ccbi = pd.DataFrame({
            "Nombre UBCH": ["UBCH Centro Educativo El Paraíso", "UBCH Grupo Escolar Simón Bolívar"],
            "Nombre y Apellido": ["Carmen Teresa", "Ruperto Medina"],
            "Cédula": ["V-11111111", "V-22222222"],
            "Teléfono": ["04121234567", "04149876543"],
            "Comunidad": ["Comunidad El Paraíso Norte", "Comunidad El Paraíso Sur"]
        })
        df_ccbi.to_excel(RUTAS["CCBI"], index=False)

    # CBBI (Comité Bolivariano de Base Integral)
    if os.path.exists(RUTAS["CBBI"]):
        df_cbbi = pd.read_excel(RUTAS["CBBI"])
    else:
        df_cbbi = pd.DataFrame({
            "Nombre UBCH": ["UBCH Centro Educativo El Paraíso", "UBCH Centro Educativo El Paraíso"],
            "Nombre y Apellido": ["Ana Soto", "Luis Díaz"],
            "Cédula": ["V-33333333", "V-44444444"],
            "Teléfono": ["04165554433", "04241112233"],
            "Nombre Calle": ["Calle Los Mangos", "Calle El Progreso"]
        })
        df_cbbi.to_excel(RUTAS["CBBI"], index=False)

    for col in df_ubch.columns:
        df_ubch[col] = df_ubch[col].astype(str)
    for col in df_ccbi.columns:
        df_ccbi[col] = df_ccbi[col].astype(str)
    for col in df_cbbi.columns:
        df_cbbi[col] = df_cbbi[col].astype(str)

    return df_ubch, df_ccbi, df_cbbi

# Carga global de datos
data_incidencias = cargar_datos_incidencias()
data_proyectos = cargar_datos_proyectos()
data_estructuras = cargar_datos_estructuras()
data_aca = cargar_datos_aca()
data_cc, data_voceros = cargar_datos_consejos_comunales()
data_logistica = cargar_datos_logistica()
data_ubch_detalles, data_ccbi, data_cbbi = cargar_datos_ubch_extra()


# ===================================================================
# 3. BARRA LATERAL DE NAVEGACIÓN Y LOGO
# ===================================================================
asegurar_directorio()
if os.path.exists(RUTAS["LOGO"]):
    st.sidebar.image(RUTAS["LOGO"], use_container_width=True)

st.sidebar.title("🏛️ Sala de Autogobierno")
st.sidebar.caption("Comuna Socialista El Paraíso")

opcion = st.sidebar.radio(
    "Módulos de Gestión", 
    [
        "🏠 Inicio (Resumen General)", 
        "🏗️ Proyectos Comunitarios", 
        "🏛️ Estructuras",
        "🎯 ACA (Agenda Concreta de Acción)",
        "🏡 Consejos Comunales",
        "🚨 Incidencias", 
        "📦 Logística"
    ]
)

st.sidebar.divider()
st.sidebar.subheader("⚙️ Configuración Visual")
with st.sidebar.expander("🖼️ Cargar / Cambiar Logo Comunal"):
    logo_subido = st.file_uploader("Seleccione una imagen (PNG / JPG):", type=["png", "jpg", "jpeg"])
    if logo_subido is not None:
        img = Image.open(logo_subido)
        img.save(RUTAS["LOGO"])
        st.success("✅ Logo actualizado correctamente.")
        st.rerun()


# ===================================================================
# 4. LÓGICA PRINCIPAL
# ===================================================================

# --- OPCIÓN 1: INICIO Y RESUMEN GENERAL ---
if opcion == "🏠 Inicio (Resumen General)":
    st.title("📊 Sala Situacional - Comuna Socialista El Paraíso")
    st.caption("Consolidado estadístico e indicadores clave de gestión comunitaria y electoral.")
    st.divider()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total_proyectos = len(data_proyectos)
    proyectos_ejecucion = len(data_proyectos[data_proyectos["Estatus"] == "En Ejecución"]) if "Estatus" in data_proyectos.columns else 0
    total_incidencias = len(data_incidencias)
    total_cc = len(data_cc)
    total_miembros_estructuras = len(data_estructuras) + len(data_voceros) + len(data_ccbi) + len(data_cbbi)
    
    col1.metric("Proyectos", total_proyectos, delta=f"{proyectos_ejecucion} en ejecución")
    col2.metric("Incidencias", total_incidencias, delta="Registradas")
    col3.metric("Consejos Comunales", total_cc, delta="Registrados")
    col4.metric("Agendas ACA", len(data_aca), delta="Acciones")
    col5.metric("Fuerza Popular", total_miembros_estructuras, delta="Miembros Activos")
    
    st.divider()
    
    st.subheader("🗳️ Resumen Detallado por UBCH (Ámbito, Votantes, Mesas, CCBI y CBBI)")
    if not data_ubch_detalles.empty:
        tabla_resumen_ubch = []
        for _, ubch_row in data_ubch_detalles.iterrows():
            nombre_u = ubch_row.get("Nombre UBCH", "")
            cc_asoc = ubch_row.get("Consejos Comunales Asociados", "N/A")
            zonas_agreg = ubch_row.get("Zonas en Agregación", "N/A")
            try:
                votantes = int(float(ubch_row.get("Cantidad Votantes", 0)))
            except:
                votantes = 0
            try:
                mesas = int(float(ubch_row.get("Cantidad Mesas", 0)))
            except:
                mesas = 0
            
            # Buscar CCBI (Comando de Comunidad de Base Integral)
            ccbi_miembros = data_ccbi[data_ccbi["Nombre UBCH"] == nombre_u]
            str_ccbi = ", ".join([f"{row.get('Nombre y Apellido', '')} ({row.get('Comunidad', '')})" for _, row in ccbi_miembros.iterrows()]) if not ccbi_miembros.empty else "Sin asignar"
            
            # Contar CBBI (Comité Bolivariano de Base Integral)
            cbbi_miembros = data_cbbi[data_cbbi["Nombre UBCH"] == nombre_u]
            cant_cbbi = len(cbbi_miembros)

            tabla_resumen_ubch.append({
                "UBCH": nombre_u,
                "Consejos Comunales en Ámbito": cc_asoc,
                "Zonas en Agregación": zonas_agreg,
                "Votantes": votantes,
                "Mesas": mesas,
                "CCBI (Comando de Comunidad)": str_ccbi,
                "Total CBBI (Comités Base)": cant_cbbi
            })
        
        df_res_ubch_final = pd.DataFrame(tabla_resumen_ubch)
        st.dataframe(df_res_ubch_final, use_container_width=True)
    else:
        st.info("No hay datos de UBCH registrados.")

    st.divider()
    st.subheader("🏛️ Resumen de Estructuras del Poder Popular")
    c_est1, c_est2 = st.columns([1, 2])
    
    with c_est1:
        st.markdown("**Distribución por Instancia:**")
        conteo_est = data_estructuras["Estructura"].value_counts().to_dict() if "Estructura" in data_estructuras.columns else {}
        conteo_est["Voceros de CC (Automáticos)"] = len(data_voceros)
        conteo_est["Miembros CCBI"] = len(data_ccbi)
        conteo_est["Miembros CBBI"] = len(data_cbbi)
        
        for inst, cant in conteo_est.items():
            st.write(f"- **{inst}:** {cant} integrantes")
            
    with c_est2:
        df_est_grafico = pd.DataFrame(list(conteo_est.items()), columns=["Estructura", "Cantidad"])
        st.bar_chart(df_est_grafico.set_index("Estructura"))


# --- OPCIÓN 2: PROYECTOS COMUNITARIOS ---
elif opcion == "🏗️ Proyectos Comunitarios":
    st.title("🏗️ Módulo de Proyectos Comunitarios")
    
    st.sidebar.divider()
    st.sidebar.subheader("🔍 Filtros de Proyectos")
    sectores_unicos = ["Todos"] + sorted(list(data_proyectos["Sector"].dropna().unique())) if "Sector" in data_proyectos.columns else ["Todos"]
    estatus_unicos = ["Todos"] + sorted(list(data_proyectos["Estatus"].dropna().unique())) if "Estatus" in data_proyectos.columns else ["Todos"]
    
    filtro_sector = st.sidebar.selectbox("Filtrar por Sector:", sectores_unicos)
    filtro_estatus = st.sidebar.selectbox("Filtrar por Estatus:", estatus_unicos)
    
    data_proyectos_filtrada = data_proyectos.copy()
    if filtro_sector != "Todos":
        data_proyectos_filtrada = data_proyectos_filtrada[data_proyectos_filtrada["Sector"] == filtro_sector]
    if filtro_estatus != "Todos":
        data_proyectos_filtrada = data_proyectos_filtrada[data_proyectos_filtrada["Estatus"] == filtro_estatus]
        
    tab1, tab2, tab3 = st.tabs(["📋 Lista de Proyectos", "➕ Registrar Nuevo Proyecto", "✏️ Editar / Eliminar"])
    
    with tab1:
        st.subheader("Proyectos en Registro")
        st.dataframe(data_proyectos_filtrada, use_container_width=True)
        
    with tab2:
        st.subheader("Formulario de Registro de Proyectos")
        with st.form("form_proyecto", clear_on_submit=True):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                nombre_p = st.text_input("Nombre del Proyecto:")
                sector_p = st.text_input("Sector / Comunidad:")
                presupuesto_p = st.number_input("Presupuesto Estimado ($):", min_value=0.0, step=100.0)
            with col_p2:
                estatus_p = st.selectbox("Estatus Inicial:", ["En Planificación", "En Ejecución", "Detenido", "Concluido"])
                responsable_p = st.text_input("Responsable / Organización:")
            
            btn_guardar_p = st.form_submit_button("💾 Guardar Proyecto")
            if btn_guardar_p:
                if not nombre_p.strip() or not sector_p.strip():
                    st.error("⚠️ El Nombre del Proyecto y Sector son obligatorios.")
                else:
                    nuevo_id_p = f"PRY-00{len(data_proyectos) + 1}"
                    nueva_fila_p = pd.DataFrame([{
                        "ID Proyecto": str(nuevo_id_p),
                        "Nombre del Proyecto": str(nombre_p),
                        "Sector": str(sector_p),
                        "Presupuesto ($)": float(presupuesto_p),
                        "Estatus": str(estatus_p),
                        "Responsable": str(responsable_p)
                    }])
                    df_act = pd.concat([data_proyectos, nueva_fila_p], ignore_index=True)
                    df_act.to_excel(RUTAS["PROYECTOS"], index=False)
                    st.success("✅ Proyecto guardado exitosamente.")
                    st.rerun()

    with tab3:
        st.subheader("Modificar o Eliminar Registros de Proyectos")
        if not data_proyectos.empty:
            id_p_sel = st.selectbox("Seleccione el Proyecto a modificar:", options=data_proyectos["ID Proyecto"].tolist())
            idx_p = data_proyectos[data_proyectos["ID Proyecto"] == id_p_sel].index[0]
            fila_p = data_proyectos.loc[idx_p]

            with st.form("form_edit_proyecto"):
                c_ep1, c_ep2 = st.columns(2)
                with c_ep1:
                    e_nombre_p = st.text_input("Nombre del Proyecto:", value=str(fila_p.get("Nombre del Proyecto", "")))
                    e_sector_p = st.text_input("Sector / Comunidad:", value=str(fila_p.get("Sector", "")))
                    try:
                        val_pres = float(fila_p.get("Presupuesto ($)", 0.0))
                    except:
                        val_pres = 0.0
                    e_presupuesto_p = st.number_input("Presupuesto ($):", value=val_pres, min_value=0.0, step=100.0)
                with c_ep2:
                    opciones_estatus = ["En Planificación", "En Ejecución", "Detenido", "Concluido"]
                    estatus_val = str(fila_p.get("Estatus", "En Planificación"))
                    idx_est = opciones_estatus.index(estatus_val) if estatus_val in opciones_estatus else 0
                    e_estatus_p = st.selectbox("Estatus:", opciones_estatus, index=idx_est)
                    e_responsable_p = st.text_input("Responsable:", value=str(fila_p.get("Responsable", "")))

                col_btn_p1, col_btn_p2 = st.columns(2)
                btn_actualizar_p = col_btn_p1.form_submit_button("💾 Guardar Cambios")
                btn_eliminar_p = col_btn_p2.form_submit_button("🗑️ Eliminar Proyecto")

                if btn_actualizar_p:
                    data_proyectos.loc[idx_p, "Nombre del Proyecto"] = str(e_nombre_p)
                    data_proyectos.loc[idx_p, "Sector"] = str(e_sector_p)
                    data_proyectos.loc[idx_p, "Presupuesto ($)"] = float(e_presupuesto_p)
                    data_proyectos.loc[idx_p, "Estatus"] = str(e_estatus_p)
                    data_proyectos.loc[idx_p, "Responsable"] = str(e_responsable_p)
                    data_proyectos.to_excel(RUTAS["PROYECTOS"], index=False)
                    st.success("✅ Proyecto actualizado correctamente.")
                    st.rerun()

                if btn_eliminar_p:
                    df_p_nuevo = data_proyectos.drop(idx_p)
                    df_p_nuevo.to_excel(RUTAS["PROYECTOS"], index=False)
                    st.success("🗑️ Proyecto eliminado.")
                    st.rerun()

    df_para_exportar = data_proyectos_filtrada


# --- OPCIÓN 3: ESTRUCTURAS (CON UBCH, CCBI Y CBBI) ---
elif opcion == "🏛️ Estructuras":
    st.title("🏛️ Módulo de Estructuras del Poder Popular")
    st.caption("Gestión e integración de miembros en las diferentes instancias organizativas (Salas de Gobierno, UBCH, Comando CCBI, Comité CBBI, Consejos Comunales).")
    
    pestañas_est = ["Sala de Gobierno", "UBCH", "Consejos Comunales", "Comunidades"]
    tabs = st.tabs([f"📌 {p}" for p in pestañas_est])
    
    # 1. Sala de Gobierno
    with tabs[0]:
        st.subheader("Estructura: Sala de Gobierno")
        df_sub = data_estructuras[data_estructuras["Estructura"] == "Sala de Gobierno"]
        if not df_sub.empty:
            for idx_m, row_m in df_sub.iterrows():
                c_m1, c_m2, c_m3, c_m4, c_m5, c_m6 = st.columns([2, 1.5, 2, 1.5, 1.5, 1.5])
                c_m1.write(f"**{row_m.get('Nombre y Apellido', '')}**")
                c_m2.write(row_m.get('Cédula', ''))
                c_m3.write(row_m.get('Cargo / Responsabilidad', ''))
                c_m4.write(row_m.get('Comunidad / Sector', ''))
                c_m5.write(row_m.get('Teléfono', ''))
                
                telf_limpio = "".join(filter(str.isdigit, str(row_m.get('Teléfono', ''))))
                if telf_limpio.startswith("0"): 
                    telf_limpio = "58" + telf_limpio[1:]
                elif not telf_limpio.startswith("58") and len(telf_limpio) == 10: 
                    telf_limpio = "58" + telf_limpio
                
                url_miembro = generar_enlace_whatsapp(telf_limpio, f"Saludos {row_m.get('Nombre y Apellido', '')}, te contactamos de la Sala de Autogobierno.")
                c_m6.markdown(f'<a href="{url_miembro}" target="_blank"><button style="background-color:#25D366;color:white;border:none;padding:4px 10px;border-radius:5px;cursor:pointer;font-size:12px;">📲 WhatsApp</button></a>', unsafe_allow_html=True)
                st.divider()
        else:
            st.info("No hay integrantes registrados en Sala de Gobierno.")
            
        with st.form("form_est_saldelgob"):
            st.markdown("**➕ Registrar en Sala de Gobierno**")
            n_nom = st.text_input("Nombre y Apellido:")
            n_ced = st.text_input("Cédula:")
            n_car = st.text_input("Cargo / Responsabilidad:")
            n_tel = st.text_input("Teléfono:")
            n_com = st.text_input("Comunidad / Sector:")
            if st.form_submit_button("💾 Guardar Integrante"):
                if n_nom and n_ced:
                    nueva_fila = pd.DataFrame([{
                        "Estructura": "Sala de Gobierno", "Nombre y Apellido": str(n_nom), "Cédula": str(n_ced),
                        "Cargo / Responsabilidad": str(n_car), "Teléfono": str(n_tel), "Comunidad / Sector": str(n_com)
                    }])
                    pd.concat([data_estructuras, nueva_fila], ignore_index=True).to_excel(RUTAS["ESTRUCTURAS"], index=False)
                    st.success("✅ Guardado con éxito.")
                    st.rerun()

    # 2. UBCH (Con Ámbito, Consejos Comunales, Zonas, Votantes, Mesas, CCBI y CBBI)
    with tabs[1]:
        st.subheader("Estructura: Unidades de Batalla Bolívar-Chávez (UBCH)")
        st.caption("Gestión de Jefes de UBCH, equipos de estructura, CCBI y CBBI con enlaces directos a WhatsApp.")

        ubch_tab_ver, ubch_tab_reg_jefe, ubch_tab_reg_est, ubch_tab_ccbi_reg, ubch_tab_cbbi_reg = st.tabs([
            "📋 Ver UBCH y Equipos", 
            "➕ Registrar / Editar Jefe UBCH", 
            "➕ Registrar Miembro de Estructura UBCH", 
            "➕ Registrar Miembro CCBI", 
            "➕ Registrar Miembro CBBI (Calle)"
        ])

        with ubch_tab_ver:
            if not data_ubch_detalles.empty:
                for idx_u, ubch_row in data_ubch_detalles.iterrows():
                    nombre_ubch = ubch_row.get("Nombre UBCH", f"UBCH {idx_u+1}")
                    cc_asociados = ubch_row.get("Consejos Comunales Asociados", "N/A")
                    zonas_agreg = ubch_row.get("Zonas en Agregación", "N/A")
                    try:
                        cant_vot = int(float(ubch_row.get("Cantidad Votantes", 0)))
                    except:
                        cant_vot = 0
                    try:
                        cant_mesas = int(float(ubch_row.get("Cantidad Mesas", 0)))
                    except:
                        cant_mesas = 0

                    with st.expander(f"📍 UBCH: {nombre_ubch} (Votantes: {cant_vot} | Mesas: {cant_mesas})", expanded=False):
                        c_u1, c_u2 = st.columns(2)
                        with c_u1:
                            st.write(f"**Consejos Comunales en su Ámbito:** {cc_asociados}")
                        with c_u2:
                            st.write(f"**Zonas en Agregación:** {zonas_agreg}")

                        st.divider()
                        st.markdown("#### 👑 Jefe de UBCH")

                        jefe_filtrado = pd.DataFrame()
                        if 'data_ubch_jefe' in globals() and data_ubch_jefe is not None and not data_ubch_jefe.empty:
                            if "Nombre UBCH" in data_ubch_jefe.columns:
                                jefe_filtrado = data_ubch_jefe[data_ubch_jefe["Nombre UBCH"] == nombre_ubch]

                        if not jefe_filtrado.empty:
                            for _, jr in jefe_filtrado.iterrows():
                                j1, j2, j3, j4 = st.columns([2, 1.2, 1.5, 1.5])
                                j1.write(f"**{jr.get('Nombre y Apellido', '')}**")
                                j2.write(f"C.I: {jr.get('Cédula', '')}")
                                j3.write(jr.get('Teléfono', ''))

                                t_jefe = "".join(filter(str.isdigit, str(jr.get('Teléfono', ''))))
                                if t_jefe.startswith("0"): 
                                    t_jefe = "58" + t_jefe[1:]
                                elif not t_jefe.startswith("58") and len(t_jefe) == 10: 
                                    t_jefe = "58" + t_jefe
                                    
                                url_jefe = generar_enlace_whatsapp(t_jefe, f"Saludos camarada {jr.get('Nombre y Apellido', '')}, Jefe de la UBCH {nombre_ubch}.")
                                j4.markdown(f'<a href="{url_jefe}" target="_blank"><button style="background-color:#25D366;color:white;border:none;padding:4px 10px;border-radius:5px;cursor:pointer;font-size:12px;">📲 WhatsApp</button></a>', unsafe_allow_html=True)
                        else:
                            st.info("No hay Jefe de UBCH asignado.")

                        st.divider()
                        st.markdown("#### 👥 Estructura Operativa de la UBCH")

                        est_filtrada = pd.DataFrame()
                        if 'data_ubch_estru' in globals() and data_ubch_estru is not None and not data_ubch_estru.empty:
                            if "Nombre UBCH" in data_ubch_estru.columns:
                                est_filtrada = data_ubch_estru[data_ubch_estru["Nombre UBCH"] == nombre_ubch]

                        if not est_filtrada.empty:
                            for _, er in est_filtrada.iterrows():
                                e1, e2, e3, e4, e5 = st.columns([1.5, 2, 1.2, 1.5, 1.5])
                                e1.write(f"**{er.get('Cargo', '')}**")
                                e2.write(er.get('Nombre y Apellido', ''))
                                e3.write(f"C.I: {er.get('Cédula', '')}")
                                e4.write(er.get('Teléfono', ''))

                                t_est = "".join(filter(str.isdigit, str(er.get('Teléfono', ''))))
                                if t_est.startswith("0"): 
                                    t_est = "58" + t_est[1:]
                                elif not t_est.startswith("58") and len(t_est) == 10: 
                                    t_est = "58" + t_est
                                    
                                url_est = generar_enlace_whatsapp(t_est, f"Saludos {er.get('Nombre y Apellido', '')}, integrante de la estructura de la UBCH {nombre_ubch}.")
                                e5.markdown(f'<a href="{url_est}" target="_blank"><button style="background-color:#25D366;color:white;border:none;padding:4px 10px;border-radius:5px;cursor:pointer;font-size:12px;">📲 WhatsApp</button></a>', unsafe_allow_html=True)
                        else:
                            st.info("No hay miembros de la estructura registrados para esta UBCH.")
            else:
                st.info("No hay UBCH registradas.")

        with ubch_tab_reg_jefe:
            st.subheader("➕ Registrar o Actualizar Jefe de UBCH")
            with st.form("form_reg_jefe_ubch", clear_on_submit=True):
                ubch_opciones = data_ubch_detalles["Nombre UBCH"].tolist() if not data_ubch_detalles.empty else []
                sel_u_jefe = st.selectbox("Seleccione la UBCH:", options=ubch_opciones)
                j_nombre = st.text_input("Nombre y Apellido del Jefe:")
                j_cedula = st.text_input("Cédula de Identidad:")
                j_telefono = st.text_input("Teléfono (Móvil):")

                if st.form_submit_button("💾 Guardar Jefe de UBCH"):
                    if j_nombre and sel_u_jefe:
                        df_j_act = data_ubch_jefe[data_ubch_jefe["Nombre UBCH"] != sel_u_jefe] if not data_ubch_jefe.empty else pd.DataFrame()
                        nuevo_jefe = pd.DataFrame({
                            "Nombre UBCH": [str(sel_u_jefe)],
                            "Nombre y Apellido": [str(j_nombre)],
                            "Cédula": [str(j_cedula)],
                            "Teléfono": [str(j_telefono)]
                        })
                        pd.concat([df_j_act, nuevo_jefe], ignore_index=True).to_excel(RUTAS["UBCH_JEFE"], index=False)
                        st.success("✅ Jefe de UBCH registrado con éxito.")
                        st.rerun()

        with ubch_tab_reg_est:
            st.subheader("➕ Registrar Miembro en la Estructura de la UBCH")
            with st.form("form_reg_est_ubch", clear_on_submit=True):
                ubch_opciones2 = data_ubch_detalles["Nombre UBCH"].tolist() if not data_ubch_detalles.empty else []
                sel_u_est = st.selectbox("Seleccione la UBCH correspondiente:", options=ubch_opciones2)
                e_cargo = st.selectbox("Cargo / Responsabilidad:", [
                    "Movilización", "Técnica Electoral", "Ideología y Formación", 
                    "Comunas y Movimientos Sociales", "Mujeres", "Juventud", 
                    "Agitacion, Propaganda y Comunicacion"
                ])
                e_nombre = st.text_input("Nombre y Apellido:")
                e_cedula = st.text_input("Cédula:")
                e_telefono = st.text_input("Teléfono:")

                submitted_est = st.form_submit_button("💾 Guardar Miembro de Estructura")
                if submitted_est:
                    if e_nombre and sel_u_est:
                        nuevo_miembro_est = pd.DataFrame({
                            "Nombre UBCH": [str(sel_u_est)],
                            "Cargo": [str(e_cargo)],
                            "Nombre y Apellido": [str(e_nombre)],
                            "Cédula": [str(e_cedula)],
                            "Teléfono": [str(e_telefono)]
                        })
                        df_concat_est = pd.concat([data_ubch_estru, nuevo_miembro_est], ignore_index=True)
                        df_concat_est.to_excel(RUTAS["UBCH_ESTRUCTURA"], index=False)
                        st.success("✅ Miembro de estructura de UBCH guardado con éxito.")
                        st.rerun()
      else:
                st.info("No hay UBCH registradas.")
        with ubch_tab_ccbi_reg:
            st.subheader("➕ Registrar Miembro CCBI")
            with st.form("form_reg_ccbi", clear_on_submit=True):
                ubch_opciones3 = data_ubch_detalles["Nombre UBCH"].tolist() if not data_ubch_detalles.empty else []
                sel_u_ccbi = st.selectbox("Seleccione la UBCH para CCBI:", options=ubch_opciones3)
                ccbi_nom = st.text_input("Nombre y Apellido (CCBI):")
                ccbi_ced = st.text_input("Cédula (CCBI):")
                ccbi_com = st.text_input("Comunidad (CCBI):")
                ccbi_tel = st.text_input("Teléfono (CCBI):")

                if st.form_submit_button("💾 Guardar Miembro CCBI"):
                    if ccbi_nom and sel_u_ccbi:
                        nuevo_ccbi = pd.DataFrame({
                            "Nombre UBCH": [str(sel_u_ccbi)],
                            "Nombre y Apellido": [str(ccbi_nom)],
                            "Cédula": [str(ccbi_ced)],
                            "Comunidad": [str(ccbi_com)],
                            "Teléfono": [str(ccbi_tel)]
                        })
                        df_concat_ccbi = pd.concat([data_ccbi, nuevo_ccbi], ignore_index=True)
                        if "CCBI" in RUTAS:
                            df_concat_ccbi.to_excel(RUTAS["CCBI"], index=False)
                        st.success("✅ Miembro CCBI guardado con éxito.")
                        st.rerun()

        with ubch_tab_cbbi_reg:
            st.subheader("➕ Registrar Miembro CBBI (Calle)")
            with st.form("form_reg_cbbi", clear_on_submit=True):
                ubch_opciones4 = data_ubch_detalles["Nombre UBCH"].tolist() if not data_ubch_detalles.empty else []
                sel_u_cbbi = st.selectbox("Seleccione la UBCH para CBBI:", options=ubch_opciones4)
                cbbi_nom = st.text_input("Nombre y Apellido (CBBI):")
                cbbi_ced = st.text_input("Cédula (CBBI):")
                cbbi_cal = st.text_input("Nombre de Calle (CBBI):")
                cbbi_tel = st.text_input("Teléfono (CBBI):")

                if st.form_submit_button("💾 Guardar Miembro CBBI"):
                    if cbbi_nom and sel_u_cbbi:
                        nuevo_cbbi = pd.DataFrame({
                            "Nombre UBCH": [str(sel_u_cbbi)],
                            "Nombre y Apellido": [str(cbbi_nom)],
                            "Cédula": [str(cbbi_ced)],
                            "Nombre Calle": [str(cbbi_cal)],
                            "Teléfono": [str(cbbi_tel)]
                        })
                        df_concat_cbbi = pd.concat([data_cbbi, nuevo_cbbi], ignore_index=True)
                        if "CBBI" in RUTAS:
                            df_concat_cbbi.to_excel(RUTAS["CBBI"], index=False)
                        st.success("✅ Miembro CBBI guardado con éxito.")
                        st.rerun()
# --- OPCIÓN 4: ACA (AGENDA CONCRETA DE ACCIÓN) ---
elif opcion == "🎯 ACA (Agenda Concreta de Acción)":
    st.subheader("Agenda Concreta de Acción")
    st.caption("Planificación estratégica comunal, diagnósticos, nudos críticos y mapas de solución.")
    
    tab_resumen, tab_comunidad, tab_carga, tab_edicion = st.tabs([
        "📊 Entrada Principal (Resumen)", 
        "🏘️ ACA por Comunidad", 
        "➕ Registrar nueva ACA",
        "✏️ Editar / Eliminar ACA"
    ])
    
    with tab_resumen:
        st.subheader("📌 Resumen Consolidado del Diagnóstico Comunal")
        col_a1, col_a2, col_a3 = st.columns(3)
        col_a1.metric("Agendas Registradas", len(data_aca))
        col_a2.metric("Comunidades Levantadas", data_aca["Comunidad"].nunique() if "Comunidad" in data_aca.columns else 0)
        col_a3.metric("Nudos Críticos", len(data_aca))
        st.dataframe(data_aca, use_container_width=True)

    with tab_comunidad:
        st.subheader("🏘️ Consulta de ACA por Comunidad")
        comunidades = ["Todas"] + list(data_aca["Comunidad"].dropna().unique()) if "Comunidad" in data_aca.columns else ["Todas"]
        com_sel = st.selectbox("Seleccione la Comunidad:", comunidades)
        df_aca_filtrado = data_aca[data_aca["Comunidad"] == com_sel] if com_sel != "Todas" else data_aca
        for _, row in df_aca_filtrado.iterrows():
            with st.expander(f"🔴 {row.get('Gabinete', '')} - {row.get('Área / Comité', '')} | {row.get('Comunidad', '')}"):
                st.write(f"**Nudo Crítico:** {row.get('Nudo Crítico', '')}")
                st.write(f"**Autogestión:** {row.get('Qué Haremos (Autogestión)', '')}")

    with tab_carga:
        st.subheader("➕ Cargar Agenda Concreta de Acción (A.C.A.)")
        with st.form("form_aca", clear_on_submit=True):
            comunidad_aca = st.text_input("Nombre de la Comunidad:")
            gabinete_aca = st.selectbox("Gabinete:", ["Servicios Públicos", "Economía Productiva", "Infraestructura", "Social", "Seguridad y Defensa"])
            area_aca = st.text_input("Área / Comité:")
            desc_aca = st.text_area("Descripción:")
            nudo_aca = st.text_area("Nudo Crítico:")
            potencial_aca = st.text_area("Potencialidades:")
            autogestion_aca = st.text_area("Autogestión:")
            congestion_aca = st.text_area("Congestión:")
            prog_aca = st.text_input("Programación:")
            resp_aca = st.text_input("Responsables:")
            
            if st.form_submit_button("💾 Guardar Agenda ACA"):
                if comunidad_aca and nudo_aca:
                    nueva_aca = pd.DataFrame([{
                        "Comunidad": str(comunidad_aca), "Gabinete": str(gabinete_aca), "Área / Comité": str(area_aca),
                        "Descripción": str(desc_aca), "Nudo Crítico": str(nudo_aca), "Potencialidades": str(potencial_aca),
                        "Qué Haremos (Autogestión)": str(autogestion_aca), "Con Ayuda de Otros (Congestión)": str(congestion_aca),
                        "Programación (Gestión)": str(prog_aca), "Responsables": str(resp_aca)
                    }])
                    pd.concat([data_aca, nueva_aca], ignore_index=True).to_excel(RUTAS["ACA"], index=False)
                    st.success("✅ Guardado exitosamente.")
                    st.rerun()

    with tab_edicion:
        st.subheader("✏️ Modificar o Eliminar una Agenda ACA")
        if not data_aca.empty:
            opciones_aca = [f"{row.get('Comunidad', '')} - {row.get('Gabinete', '')}" for _, row in data_aca.iterrows()]
            aca_sel = st.selectbox("Seleccione entrada ACA:", options=opciones_aca)
            idx_aca = opciones_aca.index(aca_sel)
            if st.button("🗑️ Eliminar Agenda Seleccionada"):
                data_aca.drop(idx_aca).to_excel(RUTAS["ACA"], index=False)
                st.success("🗑️ Eliminado.")
                st.rerun()

    df_para_exportar = data_aca


# --- OPCIÓN 5: CONSEJOS COMUNALES ---
elif opcion == "🏡 Consejos Comunales":
    st.title("🏡 Módulo de Consejos Comunales y Vocerías")
    st.caption("Registro oficial de Códigos SITUR, vigencias y voceros electos con soporte de importación masiva.")
    
    tab_cc_lista, tab_cc_registro, tab_cc_masiva, tab_cc_editar = st.tabs([
        "📋 Lista de Consejos Comunales", 
        "➕ Registrar Consejo Comunal",
        "📥 Carga Masiva (Excel/PDF)",
        "✏️ Editar / Eliminar CC o Vocero"
    ])
    
    with tab_cc_lista:
        st.subheader("Directorio de Consejos Comunales")
        st.dataframe(data_cc, use_container_width=True)
        
        situr_seleccionado = st.selectbox(
            "Seleccione un Consejo Comunal para consultar sus voceros electos:",
            options=data_cc["Código SITUR"].unique() if "Código SITUR" in data_cc.columns else []
        )
        if situr_seleccionado:
            info_cc = data_cc[data_cc["Código SITUR"] == situr_seleccionado].iloc[0]
            with st.expander(f"📌 Detalle de Voceros - {info_cc.get('Nombre del Consejo Comunal', situr_seleccionado)}", expanded=True):
                voceros_filtrados = data_voceros[data_voceros["Código SITUR"] == situr_seleccionado] if "Código SITUR" in data_voceros.columns else pd.DataFrame()
                if not voceros_filtrados.empty:
                    st.dataframe(voceros_filtrados[["Nombre y Apellido", "Cédula", "Teléfono", "Vocería / Comité", "Condición"]], use_container_width=True)
                else:
                    st.info("No hay voceros registrados para este Consejo Comunal.")

    with tab_cc_registro:
        st.subheader("➕ Registrar Nuevo Consejo Comunal")
        with st.form("form_nuevo_cc", clear_on_submit=True):
            c_situr = st.text_input("Código SITUR:")
            c_nombre = st.text_input("Nombre del Consejo Comunal:")
            c_comunidad = st.text_input("Comunidad / Sector:")
            c_f_elec = st.date_input("Fecha de Elección:")
            c_f_venc = st.date_input("Fecha de Vencimiento:")
            c_estatus = st.selectbox("Estatus Vocería:", ["Vigente", "Vencido", "En Reestructuración"])
            
            if st.form_submit_button("💾 Guardar Consejo Comunal"):
                if c_situr and c_nombre:
                    nuevo_cc_row = pd.DataFrame([{
                        "Código SITUR": str(c_situr), "Nombre del Consejo Comunal": str(c_nombre),
                        "Comunidad": str(c_comunidad), "Fecha Elección": str(c_f_elec),
                        "Fecha Vencimiento": str(c_f_venc), "Estatus Vocería": str(c_estatus)
                    }])
                    pd.concat([data_cc, nuevo_cc_row], ignore_index=True).to_excel(RUTAS["CONSEJOS_COMUNALES"], index=False)
                    st.success("✅ Consejo Comunal registrado.")
                    st.rerun()

    with tab_cc_masiva:
        st.subheader("📥 Carga Masiva de Consejos Comunales o Voceros")
        tipo_carga = st.radio("Tipo de información:", ["Consejos Comunales", "Voceros (por Código SITUR)"])
        archivo_cc_subido = st.file_uploader("Seleccione archivo:", type=["xlsx", "csv", "pdf"])
        if archivo_cc_subido is not None:
            try:
                if archivo_cc_subido.name.endswith(".pdf"):
                    reader_cc = PdfReader(archivo_cc_subido)
                    filas_pdf = []
                    for page in reader_cc.pages:
                        texto_pag = page.extract_text()
                        if texto_pag:
                            for linea in texto_pag.split("\n"):
                                partes = [p.strip() for p in linea.split(",") if p.strip()]
                                if len(partes) >= 2:
                                    filas_pdf.append({"Código SITUR": partes[0], "Nombre del Consejo Comunal": partes[1]})
                    df_archivo_procesado = pd.DataFrame(filas_pdf)
                else:
                    df_archivo_procesado = pd.read_excel(archivo_cc_subido, dtype=str)
                
                if not df_archivo_procesado.empty and st.button("🚀 Confirmar Importación"):
                    if tipo_carga == "Consejos Comunales":
                        pd.concat([data_cc, df_archivo_procesado], ignore_index=True).to_excel(RUTAS["CONSEJOS_COMUNALES"], index=False)
                    else:
                        pd.concat([data_voceros, df_archivo_procesado], ignore_index=True).to_excel(RUTAS["VOCEROS"], index=False)
                    st.success("✅ Importado con éxito.")
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

    with tab_cc_editar:
        st.subheader("✏️ Gestión de Modificaciones")
        if not data_cc.empty and st.button("🗑️ Eliminar Consejo Seleccionado"):
            situr_edit = st.selectbox("SITUR:", options=data_cc["Código SITUR"].tolist())
            idx_cc = data_cc[data_cc["Código SITUR"] == situr_edit].index[0]
            data_cc.drop(idx_cc).to_excel(RUTAS["CONSEJOS_COMUNALES"], index=False)
            st.success("Eliminado.")
            st.rerun()

    df_para_exportar = data_cc


# --- OPCIÓN 6: INCIDENCIAS ---
elif opcion == "🚨 Incidencias":
    st.title("🚨 Módulo de Gestión de Incidencias")
    tab1, tab2, tab3 = st.tabs(["📋 Registros Actuales", "➕ Agregar", "✏️ Editar / Eliminar"])
    
    with tab1:
        st.dataframe(data_incidencias, use_container_width=True)
        
    with tab2:
        with st.form("form_incidencia", clear_on_submit=True):
            zona = st.text_input("Zona / Sector:")
            tipo = st.selectbox("Tipo:", ["Servicios Públicos", "Vialidad", "Infraestructura", "Salud", "Otro"])
            severidad = st.select_slider("Severidad:", options=["Baja", "Media", "Alta", "Crítica"])
            estado = st.selectbox("Estado:", ["Pendiente", "En Proceso", "Resuelto"])
            observaciones = st.text_area("Observaciones:")
            
            if st.form_submit_button("💾 Guardar Incidencia"):
                if zona:
                    nueva_fila = pd.DataFrame([{
                        "ID": len(data_incidencias) + 101, "Zona / Sector": str(zona), "Tipo de Incidencia": str(tipo),
                        "Severidad": str(severidad), "Estado": str(estado), "Observaciones": str(observaciones)
                    }])
                    pd.concat([data_incidencias, nueva_fila], ignore_index=True).to_excel(RUTAS["INCIDENCIAS"], index=False)
                    st.success("✅ Guardado.")
                    st.rerun()

    with tab3:
        if not data_incidencias.empty and st.button("🗑️ Eliminar Incidencia"):
            id_sel = st.selectbox("ID:", options=data_incidencias["ID"].tolist())
            idx_inc = data_incidencias[data_incidencias["ID"] == id_sel].index[0]
            data_incidencias.drop(idx_inc).to_excel(RUTAS["INCIDENCIAS"], index=False)
            st.success("Eliminado.")
            st.rerun()

    df_para_exportar = data_incidencias


# --- OPCIÓN 7: LOGÍSTICA ---
elif opcion == "📦 Logística":
    st.title("📦 Módulo de Recursos Logísticos")
    tab1, tab2, tab3 = st.tabs(["📋 Recursos", "➕ Agregar", "✏️ Editar / Eliminar"])
    
    with tab1:
        st.dataframe(data_logistica, use_container_width=True)
    with tab2:
        with st.form("form_logistica", clear_on_submit=True):
            n_recurso = st.text_input("Recurso:")
            n_disp = st.number_input("Disponible:", min_value=0, step=1)
            n_critico = st.selectbox("Nivel Crítico:", ["No", "Sí"])
            if st.form_submit_button("💾 Guardar"):
                if n_recurso:
                    pd.concat([data_logistica, pd.DataFrame([{"Recurso": str(n_recurso), "Disponible": int(n_disp), "Nivel_Critico": str(n_critico)}])], ignore_index=True).to_excel(RUTAS["LOGISTICA"], index=False)
                    st.success("Guardado.")
                    st.rerun()
    with tab3:
        if not data_logistica.empty and st.button("🗑️ Eliminar Recurso"):
            rec_sel = st.selectbox("Recurso:", options=data_logistica["Recurso"].tolist())
            idx_l = data_logistica[data_logistica["Recurso"] == rec_sel].index[0]
            data_logistica.drop(idx_l).to_excel(RUTAS["LOGISTICA"], index=False)
            st.success("Eliminado.")
            st.rerun()

    df_para_exportar = data_logistica


# ===================================================================
# 5. EXPORTACIÓN
# ===================================================================
st.divider()
st.subheader("📥 Exportación e Integraciones")

col_exp1, col_exp2, col_exp3 = st.columns(3)
with col_exp1:
    st.download_button(
        label="📄 Descargar en Excel",
        data=generar_excel(df_para_exportar),
        file_name=f"reporte_{opcion.lower().replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with col_exp2:
    st.download_button(
        label="🔴 Descargar en PDF",
        data=generar_pdf(opcion, df_para_exportar),
        file_name=f"reporte_{opcion.lower().replace(' ', '_')}.pdf",
        mime="application/pdf"
    )
with col_exp3:
    telefono = st.text_input("Número de destino:", value="584120000000")
    resumen_texto = f"Reporte de {opcion}:\nGenerado desde la Sala Situacional de la Comuna Socialista El Paraíso."
    url_ws = generar_enlace_whatsapp(telefono, resumen_texto)
    st.markdown(f'<a href="{url_ws}" target="_blank"><button style="background-color:#25D366;color:white;border:none;padding:8px 16px;border-radius:5px;cursor:pointer;">📲 Compartir en WhatsApp</button></a>', unsafe_allow_html=True)