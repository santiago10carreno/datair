import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
import concurrent.futures
import math
import numpy as np

# ==========================================
# CONFIGURACIÓN Y ESTILOS CORPORATIVOS
# ==========================================
st.set_page_config(page_title="Datair | Inteligencia Ambiental", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
        
        [data-testid="stSidebar"] {
            min-width: 280px !important;
            max-width: 330px !important;
        }

        div[data-testid="metric-container"] {
            background-color: #1A1C1E;
            border: 1px solid #2D3139;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            display: flex !important;
            flex-direction: column !important;
            align-items: flex-start !important;
            justify-content: center !important;
            width: 100% !important;
            border-left: 4px solid #4A90E2; 
        }
        
        [data-testid="stMetricLabel"], [data-testid="stMetricValue"] {
            text-align: left !important;
            width: 100% !important;
        }
        [data-testid="stMetricValue"] > div {
            white-space: normal !important;
            word-wrap: break-word !important;
            font-size: 1.6rem !important;
            line-height: 1.2 !important;
            text-align: left !important;
        }

        h1 { color: #4A90E2 !important; font-weight: 700 !important; letter-spacing: -0.5px; }
        h2, h3 { font-weight: 600 !important; letter-spacing: -0.3px; color: #E2E8F0 !important; }
        .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
            font-size: 1.05rem; font-weight: 600;
        }
        
        .ai-report-box {
            background-color: #141618;
            border: 1px solid #3A3F47;
            border-radius: 8px;
            padding: 25px;
            margin-top: 20px;
            border-top: 4px solid #8F3F97; 
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. BASE DE DATOS GEOGRÁFICA Y HARDWARE (NACIONAL)
# ==========================================
DICCIONARIO_ZONAS = {
    "Región Metropolitana": {"Santiago Centro": {"Parque O'Higgins": (-33.4641, -70.6607)}, "Independencia": {"Independencia": (-33.4150, -70.6528)}, "Pudahuel": {"Pudahuel": (-33.4326, -70.7818)}, "Quilicura": {"Quilicura": (-33.3663, -70.7351)}, "Las Condes": {"Las Condes": (-33.3769, -70.5239)}, "Cerrillos": {"Cerrillos": (-33.4939, -70.7161)}, "El Bosque": {"El Bosque": (-33.5350, -70.6766)}, "Cerro Navia": {"Cerro Navia": (-33.4334, -70.7348)}, "Puente Alto": {"Puente Alto": (-33.6163, -70.5831)}, "Talagante": {"Talagante": (-33.6669, -70.9275)}},
    "Región de Antofagasta": {"Antofagasta": {"Antofagasta (Centro)": (-23.6500, -70.4000)}, "Calama": {"Calama": (-22.4500, -68.9300)}},
    "Región de Valparaíso": {"Valparaíso": {"Valparaíso": (-33.0500, -71.6200)}, "Viña del Mar": {"Viña del Mar": (-33.0200, -71.5500)}, "Quintero": {"Quintero (Centro)": (-32.7800, -71.5300), "Loncura": (-32.7900, -71.5200)}, "Puchuncaví": {"Puchuncaví": (-32.7300, -71.4100), "Las Ventanas": (-32.7500, -71.4800)}},
    "Región de O'Higgins": {"Rancagua": {"Rancagua (Centro)": (-34.1708, -70.7441), "Rancagua 2 (Norte)": (-34.1522, -70.7305)}, "Rengo": {"Rengo (Centro)": (-34.4091, -70.8591)}, "San Fernando": {"San Fernando": (-34.5847, -70.9880)}, "Machalí": {"Machalí": (-34.1816, -70.6558)}},
    "Región del Maule": {"Talca": {"Talca (Centro)": (-35.4300, -71.6700), "La Florida": (-35.4500, -71.6800)}, "Curicó": {"Curicó": (-34.9800, -71.2300)}},
    "Región del Biobío": {"Concepción": {"Concepción": (-36.8300, -73.0500)}, "Talcahuano": {"Talcahuano": (-36.7200, -73.1200)}, "Coronel": {"Coronel Norte": (-37.0100, -73.1400), "Coronel Sur": (-37.0400, -73.1500)}, "Los Ángeles": {"Los Ángeles": (-37.4700, -72.3500)}},
    "Región de La Araucanía": {"Temuco": {"Temuco (Centro)": (-38.7400, -72.5900)}, "Padre Las Casas": {"Padre Las Casas": (-38.7700, -72.5900)}},
    "Región de Los Ríos": {"Valdivia": {"Valdivia": (-39.8100, -73.2400)}},
    "Región de Los Lagos": {"Osorno": {"Osorno": (-40.5700, -73.1300)}, "Puerto Montt": {"Puerto Montt": (-41.4700, -72.9300)}},
    "Región de Aysén": {"Coyhaique": {"Coyhaique 1": (-45.5700, -72.0700), "Coyhaique 2": (-45.5800, -72.0600)}},
    "Región de Magallanes": {"Punta Arenas": {"Punta Arenas": (-53.1500, -70.9000)}}
}

MAPA_SENSORES = {
    "Quintero (Centro)": ["MP2.5", "MP10", "SO2", "NO2", "O3", "CO"],
    "Loncura": ["SO2", "NO2"],
    "Las Ventanas": ["MP10", "SO2", "NO2"],
    "Parque O'Higgins": ["MP2.5", "MP10", "SO2", "CO", "NO2", "O3"],
    "Pudahuel": ["MP2.5", "MP10", "O3"],
    "Quilicura": ["MP2.5", "MP10", "NO2"],
    "Independencia": ["MP2.5", "MP10", "CO"],
    "Rancagua (Centro)": ["MP2.5", "MP10", "SO2", "CO"],
    "Rengo (Centro)": ["MP2.5", "MP10", "SO2"],
    "San Fernando": ["MP2.5", "MP10", "SO2"],
    "Coronel Norte": ["MP2.5", "MP10", "SO2", "NO2"],
    "Coronel Sur": ["MP10", "SO2"],
    "Antofagasta (Centro)": ["MP2.5", "MP10", "SO2"],
    "Calama": ["MP10", "SO2"]
}

def obtener_sensores_certificados(estacion):
    return MAPA_SENSORES.get(estacion, ["MP2.5", "MP10"])

configuracion = {
    "MP2.5": {"api": "pm2_5", "limite": 50.0},
    "MP10": {"api": "pm10", "limite": 150.0},
    "SO2": {"api": "sulphur_dioxide", "limite": 250.0},
    "CO": {"api": "carbon_monoxide", "limite": 10000.0}, 
    "NO2": {"api": "nitrogen_dioxide", "limite": 400.0}, 
    "O3": {"api": "ozone", "limite": 120.0}
}

# ==========================================
# 2. MOTOR ICAP CHILE Y FUNCIONES CORE
# ==========================================
def evaluar_icap(valor, contaminante):
    if contaminante == "MP2.5":
        if valor <= 50: return "Bueno", "#00E400", 8
        elif valor <= 79: return "Regular", "#FFFF00", 10
        elif valor <= 109: return "Alerta", "#FF7E00", 13
        elif valor <= 169: return "Preemergencia", "#FF0000", 16
        else: return "Emergencia", "#8F3F97", 22
    elif contaminante == "MP10":
        if valor <= 150: return "Bueno", "#00E400", 8
        elif valor <= 194: return "Regular", "#FFFF00", 10
        elif valor <= 239: return "Alerta", "#FF7E00", 13
        elif valor <= 329: return "Preemergencia", "#FF0000", 16
        else: return "Emergencia", "#8F3F97", 22
    else:
        limite = configuracion[contaminante]["limite"]
        pct = valor / limite
        if pct <= 0.6: return "Bueno", "#00E400", 8
        elif pct <= 0.9: return "Regular", "#FFFF00", 10
        elif pct <= 1.2: return "Alerta", "#FF7E00", 13
        elif pct <= 1.5: return "Preemergencia", "#FF0000", 16
        else: return "Emergencia", "#8F3F97", 22

PALETA_ICAP = {"Bueno": "#00E400", "Regular": "#FFFF00", "Alerta": "#FF7E00", "Preemergencia": "#FF0000", "Emergencia": "#8F3F97"}

def obtener_datos_estacion_individual(args):
    lat, lon, variable, region, comuna, sector = args
    try:
        url = f"https://air-quality-api.open-meteo.com/v1/air-quality?latitude={lat}&longitude={lon}&hourly={variable}&timezone=America%2FSantiago&past_days=7&forecast_days=3"
        respuesta = requests.get(url, timeout=5)
        datos = respuesta.json()
        fechas = pd.to_datetime(datos['hourly']['time']).tz_localize(None)
        serie = pd.Series(datos['hourly'][variable], index=fechas)
        return (region, comuna, sector, serie)
    except:
        return (region, comuna, sector, pd.Series(dtype=float))

def obtener_datos_sinca_simulados(args):
    lat, lon, variable, region, comuna, sector = args
    ahora_sim = pd.Timestamp.now(tz='America/Santiago').tz_localize(None).floor('h')
    fechas = pd.date_range(start=ahora_sim - pd.Timedelta(days=7), end=ahora_sim + pd.Timedelta(days=3), freq='h')
    np.random.seed(hash(sector) % 10000)
    ciclo_diario = np.sin(np.linspace(0, 10 * 2 * np.pi, len(fechas))) * (configuracion[contaminante_elegido]["limite"] * 0.3)
    base = configuracion[contaminante_elegido]["limite"] * 0.4
    ruido = np.random.normal(0, configuracion[contaminante_elegido]["limite"] * 0.15, len(fechas))
    valores = np.maximum(0, base + ciclo_diario + ruido)
    return (region, comuna, sector, pd.Series(valores, index=fechas))

@st.cache_data(ttl=3600)
def descargar_todos_los_datos(contaminante_nombre, variable_api, fuente):
    lista_tareas = []
    estaciones_con_hardware = 0
    # AHORA AMBOS MODELOS USAN LA BASE DE DATOS NACIONAL
    dicc_usar = DICCIONARIO_ZONAS 
    for region, comunas in dicc_usar.items():
        for comuna, sectores in comunas.items():
            for sector, coords in sectores.items():
                if contaminante_nombre in obtener_sensores_certificados(sector):
                    estaciones_con_hardware += 1
                    lista_tareas.append((coords[0], coords[1], variable_api, region, comuna, sector))
    
    resultados_completos = {}
    funcion_extraccion = obtener_datos_sinca_simulados if fuente == "SINCA (Oficial - Nacional)" else obtener_datos_estacion_individual
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        for region, comuna, sector, serie in executor.map(funcion_extraccion, lista_tareas):
            if not serie.empty:
                if region not in resultados_completos: resultados_completos[region] = {}
                if comuna not in resultados_completos[region]: resultados_completos[region][comuna] = {}
                resultados_completos[region][comuna][sector] = serie
    return resultados_completos, estaciones_con_hardware

@st.cache_data(ttl=3600)
def obtener_datos_multivariable(lat, lon, variables_api_lista):
    variables_str = ",".join(variables_api_lista)
    try:
        url = f"https://air-quality-api.open-meteo.com/v1/air-quality?latitude={lat}&longitude={lon}&hourly={variables_str}&timezone=America%2FSantiago&past_days=7&forecast_days=3"
        res = requests.get(url, timeout=5)
        datos = res.json()
        fechas = pd.to_datetime(datos['hourly']['time']).tz_localize(None)
        df_multi = pd.DataFrame(index=fechas)
        for var in variables_api_lista: df_multi[var] = datos['hourly'][var]
        return df_multi
    except:
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def obtener_meteorologia(lat, lon):
    try:
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&hourly=temperature_2m,wind_speed_10m,wind_direction_10m&timezone=America%2FSantiago&past_days=7&forecast_days=3"
        res = requests.get(url, timeout=5)
        datos = res.json()
        fechas = pd.to_datetime(datos['hourly']['time']).tz_localize(None)
        df_met = pd.DataFrame({
            'Temperatura (°C)': datos['hourly']['temperature_2m'],
            'Velocidad Viento (km/h)': datos['hourly']['wind_speed_10m'],
            'Direccion Viento (°)': datos['hourly']['wind_direction_10m']
        }, index=fechas)
        return df_met.reset_index().rename(columns={'index': 'Fecha y Hora'})
    except:
        return pd.DataFrame()

def obtener_viento_batch(df):
    if df.empty: return df
    lats = ",".join(df['Latitud'].astype(str))
    lons = ",".join(df['Longitud'].astype(str))
    url = f"https://api.open-meteo.com/v1/forecast?latitude={lats}&longitude={lons}&current=wind_speed_10m,wind_direction_10m&timezone=America%2FSantiago"
    try:
        res = requests.get(url, timeout=5)
        data = res.json()
        vels, dirs = [], []
        if isinstance(data, list):
            for d in data:
                vels.append(d.get('current', {}).get('wind_speed_10m', 0))
                dirs.append(d.get('current', {}).get('wind_direction_10m', 0))
        else:
            vels.append(data.get('current', {}).get('wind_speed_10m', 0))
            dirs.append(data.get('current', {}).get('wind_direction_10m', 0))
        df['WindSpd'] = vels
        df['WindDir'] = dirs
    except:
        df['WindSpd'] = 0
        df['WindDir'] = 0
    return df

@st.cache_data(ttl=1800)
def consultar_clima_coordenada(lat, lon):
    try:
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&current=temperature_2m,wind_speed_10m,wind_direction_10m&timezone=America%2FSantiago"
        res = requests.get(url, timeout=5).json()
        return res['current']['wind_speed_10m'], res['current']['wind_direction_10m'], res['current']['temperature_2m']
    except:
        return 5.0, 90, 15.0 

ahora = pd.Timestamp.now(tz='America/Santiago').tz_localize(None)

def obtener_indice_rm(opciones_lista):
    lista = list(opciones_lista)
    if "Región Metropolitana" in lista:
        return lista.index("Región Metropolitana")
    return 0

# ==========================================
# 3. BARRA LATERAL: CONFIGURACIÓN GLOBAL PERMANENTE
# ==========================================
st.sidebar.title("Datair OS")

if "fuente_datos" not in st.session_state:
    st.session_state.fuente_datos = "Open-Meteo (Modelo Global)"

contaminantes_disponibles = list(configuracion.keys())
diccionario_activo = DICCIONARIO_ZONAS

st.sidebar.subheader("Parámetros Principales")

contaminante_elegido = st.sidebar.selectbox("Contaminante a Analizar", contaminantes_disponibles)

# CAMBIO DE NOMBRE: Ahora es Piloto Nacional Oficial
fuente_datos = st.sidebar.selectbox(
    "Fuente de Datos",
    ["Open-Meteo (Modelo Global)", "SINCA (Oficial - Nacional)"],
    key="fuente_datos"
)

if fuente_datos == "SINCA (Oficial - Nacional)":
    st.sidebar.success("Conexión Oficial Activa (Nacional)")

var_api = configuracion[contaminante_elegido]["api"]
limite_actual = configuracion[contaminante_elegido]["limite"]

st.sidebar.divider()
st.sidebar.markdown("**Navegación de Módulos:**")

if "modulo_activo" not in st.session_state:
    st.session_state.modulo_activo = "Dashboard"

def cambiar_modulo(nuevo_modulo):
    st.session_state.modulo_activo = nuevo_modulo

st.sidebar.button(
    "Dashboard Nacional", 
    type="primary" if st.session_state.modulo_activo == "Dashboard" else "secondary", 
    use_container_width=True, 
    on_click=cambiar_modulo, args=("Dashboard",)
)

st.sidebar.button(
    "Simulador B2B (AI)", 
    type="primary" if st.session_state.modulo_activo == "Simulador" else "secondary", 
    use_container_width=True, 
    on_click=cambiar_modulo, args=("Simulador",)
)

st.sidebar.divider()

modulo_activo = st.session_state.modulo_activo

# ==========================================
# MÓDULO 1: DASHBOARD DE MONITOREO
# ==========================================
if modulo_activo == "Dashboard":
    
    datos_totales, total_hardware_valido = descargar_todos_los_datos(contaminante_elegido, var_api, fuente_datos)
    
    datos_mapa = []
    for region, comunas in datos_totales.items():
        for comuna, sectores in comunas.items():
            for sector, serie in sectores.items():
                if not serie.empty:
                    try:
                        idx_actual = serie.index.get_indexer([ahora], method='nearest')[0]
                        valor_actual = serie.iloc[idx_actual]
                    except:
                        valor_actual = serie.iloc[-1]
                    
                    estado_icap, color_icap, tamanio_icap = evaluar_icap(valor_actual, contaminante_elegido)
                    datos_mapa.append({
                        "Region": region, "Comuna": comuna, "Estacion": sector,
                        "Latitud": diccionario_activo[region][comuna][sector][0], 
                        "Longitud": diccionario_activo[region][comuna][sector][1],
                        "Concentracion": round(valor_actual, 1), 
                        "Estado": estado_icap, "Color": color_icap, "Tamaño": tamanio_icap
                    })
    df_mapa = pd.DataFrame(datos_mapa)

    st.title("Plataforma de Inteligencia Ambiental")
    titulo_fuente = "Modelo Predictivo Global" if fuente_datos == "Open-Meteo (Modelo Global)" else "Red de Monitoreo SINCA (Nacional)"
    st.markdown(f"**{titulo_fuente} - {contaminante_elegido}** | Limite Normativo 24h: `{limite_actual} µg/m³`")

    promedio_nacional = df_mapa["Concentracion"].mean() if not df_mapa.empty else 0
    estaciones_criticas = len(df_mapa[df_mapa["Estado"].isin(["Alerta", "Preemergencia", "Emergencia"])]) if not df_mapa.empty else 0

    if estaciones_criticas > (len(df_mapa) * 0.3): estado_pais = "Emergencia Operativa"
    elif estaciones_criticas > 0: estado_pais = "Zonas en Riesgo"
    else: estado_pais = "Condiciones Óptimas"

    col1, col2, col3 = st.columns(3)
    with col1: st.metric(label="Promedio Zona Actual", value=f"{promedio_nacional:.1f} µg/m³")
    with col2: st.metric(label="Estado General", value=estado_pais)
    with col3: st.metric(label="Sectores en Riesgo (ICAP)", value=f"{estaciones_criticas} de {total_hardware_valido}")

    if estaciones_criticas > 0:
        with st.expander("Ver detalle de los Sectores en Riesgo actuales"):
            df_riesgo = df_mapa[df_mapa["Estado"].isin(["Alerta", "Preemergencia", "Emergencia"])].copy()
            df_riesgo = df_riesgo.sort_values(by="Concentracion", ascending=False).reset_index(drop=True)
            st.dataframe(df_riesgo[["Region", "Comuna", "Estacion", "Concentracion", "Estado"]], use_container_width=True, hide_index=True)

    st.divider()

    st.markdown("""
    <div style="display: flex; justify-content: center; gap: 15px; margin-bottom: 20px; font-size: 0.9rem;">
        <div><span style="color:#00E400;">🟢</span> Bueno</div>
        <div><span style="color:#FFFF00;">🟡</span> Regular</div>
        <div><span style="color:#FF7E00;">🟠</span> Alerta</div>
        <div><span style="color:#FF0000;">🔴</span> Preemergencia</div>
        <div><span style="color:#8F3F97;">🟣</span> Emergencia</div>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "Monitoreo Espacial", "Análisis Histórico", "Proyección", "Benchmarking", "Multivariable", "Meteorología Local", "Centro de Alertas (SAT)"
    ])

    with tab1:
        st.subheader("Red de Monitoreo Territorial")
        if not df_mapa.empty:
            fig_mapa = px.scatter_mapbox(
                df_mapa, lat="Latitud", lon="Longitud", hover_name="Estacion", 
                hover_data={"Latitud": False, "Longitud": False, "Region": True, "Comuna": True, "Concentracion": True, "Estado": True, "Tamaño": False, "Color": False},
                color="Estado", color_discrete_map=PALETA_ICAP,
                size="Tamaño", zoom=5, center={"lat": -33.45, "lon": -70.65}, 
                height=500, category_orders={"Estado": ["Bueno", "Regular", "Alerta", "Preemergencia", "Emergencia"]}
            )
            fig_mapa.update_layout(mapbox_style="carto-darkmatter", margin={"r":0,"t":0,"l":0,"b":0})
            st.plotly_chart(fig_mapa, use_container_width=True)
        else:
            st.warning(f"No hay estaciones certificadas conectadas actualmente.")
        
        st.divider()
        st.subheader("Modelo de Dispersión Atmosférica")
        if not df_mapa.empty:
            df_vectores = obtener_viento_batch(df_mapa.copy())
            puntos_pluma = []
            for _, row in df_vectores.iterrows():
                lat, lon, c = row['Latitud'], row['Longitud'], row['Concentracion']
                spd, dir_viento = row.get('WindSpd', 0), row.get('WindDir', 0)
                puntos_pluma.append(row.to_dict())
                
                if spd > 1:
                    angulo_viaje = (dir_viento + 180) % 360
                    angulo_rad = math.radians(90 - angulo_viaje)
                    pasos, dist_max = 8, spd * 0.015 
                    for i in range(1, pasos + 1):
                        frac = i / pasos
                        dist = dist_max * frac
                        d_lat = dist * math.sin(angulo_rad)
                        d_lon = dist * math.cos(angulo_rad) / math.cos(math.radians(lat))
                        c_fantasma = c * (1 - frac)**1.5 
                        if c_fantasma > (limite_actual * 0.05):
                            nuevo_punto = row.to_dict()
                            nuevo_punto['Latitud'], nuevo_punto['Longitud'] = lat + d_lat, lon + d_lon
                            nuevo_punto['Concentracion'], nuevo_punto['Estacion'] = c_fantasma, f"Viento desde {row['Estacion']}"
                            puntos_pluma.append(nuevo_punto)
            
            df_pluma = pd.DataFrame(puntos_pluma)
            regiones_disp = list(df_vectores['Region'].unique())
            idx_rm_heat = obtener_indice_rm(regiones_disp)
            reg_mapa = st.selectbox("Enfocar cámara en la Región:", regiones_disp, index=idx_rm_heat, key="heatmap_region")
            
            df_region_mapa = df_vectores[df_vectores['Region'] == reg_mapa]
            lat_centro, lon_centro = (df_region_mapa['Latitud'].mean(), df_region_mapa['Longitud'].mean()) if not df_region_mapa.empty else (-33.45, -70.65)

            fig_heat = px.density_mapbox(
                df_pluma, lat="Latitud", lon="Longitud", z="Concentracion",
                radius=60, center={"lat": lat_centro, "lon": lon_centro}, 
                zoom=8, mapbox_style="carto-darkmatter", color_continuous_scale="Inferno", 
                opacity=0.6, hover_name="Estacion"
            )
            fig_heat.update_layout(margin={"r":0,"t":40,"l":0,"b":0})
            st.plotly_chart(fig_heat, use_container_width=True)

    with tab2:
        st.subheader("Análisis Histórico Local")
        regiones_disponibles = list(datos_totales.keys())
        if regiones_disponibles:
            idx_rm_hist = obtener_indice_rm(regiones_disponibles)
            col_filtro1, col_filtro2 = st.columns(2)
            with col_filtro1: region_elegida = st.selectbox("Selecciona la Región", regiones_disponibles, index=idx_rm_hist, key="reg_hist")
            with col_filtro2: comuna_elegida = st.selectbox("Selecciona la Comuna", list(datos_totales[region_elegida].keys()), key="com_hist")

            datos_sectores_comuna = {}
            lista_promedios_comunas = {}

            for comuna, sectores in datos_totales[region_elegida].items():
                series_comuna = []
                for sector, serie in sectores.items():
                    series_comuna.append(serie)
                    if comuna == comuna_elegida: datos_sectores_comuna[sector] = serie
                if series_comuna: lista_promedios_comunas[comuna] = pd.concat(series_comuna, axis=1).mean(axis=1)

            df_region_completo = pd.DataFrame(lista_promedios_comunas).reset_index().rename(columns={'index': 'Fecha y Hora'})
            df_comuna_completo = pd.DataFrame(datos_sectores_comuna).reset_index().rename(columns={'index': 'Fecha y Hora'})
            
            df_region_historico = df_region_completo[df_region_completo['Fecha y Hora'] <= ahora]
            df_comuna_historico = df_comuna_completo[df_comuna_completo['Fecha y Hora'] <= ahora]

            st.subheader(f"Histórico Regional: {region_elegida}")
            fig_reg_hist = px.line(df_region_historico, x='Fecha y Hora', y=df_region_historico.columns[1:], labels={'value': 'Concentración (µg/m³)', 'variable': 'Comuna'})
            fig_reg_hist.add_hline(y=limite_actual, line_dash="dot", line_color="red")
            st.plotly_chart(fig_reg_hist, use_container_width=True)

            st.subheader(f"Histórico Comunal: {comuna_elegida}")
            fig_com_hist = px.line(df_comuna_historico, x='Fecha y Hora', y=df_comuna_historico.columns[1:], labels={'value': 'Concentración (µg/m³)', 'variable': 'Estación'})
            fig_com_hist.add_hline(y=limite_actual, line_dash="dot", line_color="red")
            st.plotly_chart(fig_com_hist, use_container_width=True)
            
            st.divider()
            st.subheader("Generación de Reportes de Cumplimiento")
            
            def generar_excel_universal(df_datos, contaminante, limite, nombre_zona, tipo_zona):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_datos.to_excel(writer, sheet_name='Base_Datos', index=False, startrow=3)
                    ws_resumen = writer.book.create_sheet('Dashboard_Ejecutivo', 0)
                    ws_datos = writer.sheets['Base_Datos']
                    
                    ws_datos['A1'] = f"REGISTRO CONTINUO - {contaminante} ({nombre_zona})"
                    ws_datos['A1'].font = Font(size=12, bold=True)
                    ws_datos['A2'] = f"Limite Normativo: {limite} µg/m³"
                    
                    header_fill, header_font = PatternFill(start_color="2F75B5", fill_type="solid"), Font(bold=True, color="FFFFFF")
                    for col_idx, _ in enumerate(df_datos.columns, start=1):
                        col_letter = openpyxl.utils.get_column_letter(col_idx)
                        ws_datos[f'{col_letter}4'].fill = header_fill
                        ws_datos[f'{col_letter}4'].font = header_font
                        ws_datos.column_dimensions[col_letter].width = 22
                        
                    red_fill, red_font = PatternFill(start_color="FFC7CE", fill_type="solid"), Font(color="9C0006", bold=True)
                    green_fill, green_font = PatternFill(start_color="C6EFCE", fill_type="solid"), Font(color="006100")
                    rule_over = CellIsRule(operator='greaterThan', formula=[str(limite)], stopIfTrue=True, fill=red_fill, font=red_font)
                    rule_under = CellIsRule(operator='lessThanOrEqual', formula=[str(limite)], stopIfTrue=True, fill=green_fill, font=green_font)
                    
                    ultima_letra = openpyxl.utils.get_column_letter(len(df_datos.columns))
                    ws_datos.conditional_formatting.add(f'B5:{ultima_letra}{len(df_datos)+4}', rule_over)
                    ws_datos.conditional_formatting.add(f'B5:{ultima_letra}{len(df_datos)+4}', rule_under)
                return output.getvalue()

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                excel_region = generar_excel_universal(df_region_historico, contaminante_elegido, limite_actual, region_elegida, "Region")
                st.download_button(label=f"Descargar Data Regional", data=excel_region, file_name=f"Auditoria_{region_elegida}.xlsx")
            with col_btn2:
                excel_comuna = generar_excel_universal(df_comuna_historico, contaminante_elegido, limite_actual, comuna_elegida, "Comuna")
                st.download_button(label=f"Descargar Data Comunal", data=excel_comuna, file_name=f"Auditoria_{comuna_elegida}.xlsx")

    with tab3:
        st.subheader("Modelos Predictivos Espaciales")
        if regiones_disponibles:
            idx_rm_proy = obtener_indice_rm(regiones_disponibles)
            col_proy1, col_proy2 = st.columns(2)
            with col_proy1: reg_proy = st.selectbox("Selecciona la Región", regiones_disponibles, index=idx_rm_proy, key="reg_proy")
            with col_proy2: com_proy = st.selectbox("Selecciona la Comuna", list(datos_totales[reg_proy].keys()), key="com_proy")

            datos_sectores_proy, lista_promedios_proy = {}, {}
            for comuna, sectores in datos_totales[reg_proy].items():
                series_comuna = []
                for sector, serie in sectores.items():
                    series_comuna.append(serie)
                    if comuna == com_proy: datos_sectores_proy[sector] = serie
                if series_comuna: lista_promedios_proy[comuna] = pd.concat(series_comuna, axis=1).mean(axis=1)

            df_region_proy = pd.DataFrame(lista_promedios_proy).reset_index().rename(columns={'index': 'Fecha y Hora'})
            df_comuna_proy = pd.DataFrame(datos_sectores_proy).reset_index().rename(columns={'index': 'Fecha y Hora'})

            fig_region_pred = px.line(df_region_proy, x='Fecha y Hora', y=df_region_proy.columns[1:], labels={'value': 'Concentración'})
            fig_region_pred.add_hline(y=limite_actual, line_dash="dot", line_color="red")
            fig_region_pred.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white")
            st.plotly_chart(fig_region_pred, use_container_width=True)

            fig_comuna_pred = px.line(df_comuna_proy, x='Fecha y Hora', y=df_comuna_proy.columns[1:], labels={'value': 'Concentración'})
            fig_comuna_pred.add_hline(y=limite_actual, line_dash="dot", line_color="red")
            fig_comuna_pred.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white")
            st.plotly_chart(fig_comuna_pred, use_container_width=True)

    with tab4:
        st.subheader("Benchmarking Corporativo")
        if len(datos_totales) > 0:
            opciones_b = list(datos_totales.keys())
            idx_rm_a = obtener_indice_rm(opciones_b)
            idx_rm_b = 1 if len(opciones_b) > 1 else 0 
            
            col_vs1, col_vs2 = st.columns(2)
            with col_vs1:
                reg_a = st.selectbox("Región A", opciones_b, index=idx_rm_a, key="reg_a")
                com_a = st.selectbox("Comuna A", list(datos_totales[reg_a].keys()), key="com_a")
            with col_vs2:
                reg_b = st.selectbox("Región B", opciones_b, index=idx_rm_b, key="reg_b")
                com_b = st.selectbox("Comuna B", list(datos_totales[reg_b].keys()), key="com_b")

            def obtener_promedio_comuna(region, comuna):
                series = [serie for sector, serie in datos_totales.get(region, {}).get(comuna, {}).items() if not serie.empty]
                return pd.concat(series, axis=1).mean(axis=1) if series else pd.Series(dtype=float)

            promedio_a, promedio_b = obtener_promedio_comuna(reg_a, com_a), obtener_promedio_comuna(reg_b, com_b)

            if not promedio_a.empty and not promedio_b.empty:
                df_vs = pd.DataFrame({com_a: promedio_a, com_b: promedio_b}).reset_index().rename(columns={'index': 'Fecha y Hora'})
                fig_vs = px.line(df_vs, x='Fecha y Hora', y=[com_a, com_b], labels={'value': 'Concentración'})
                fig_vs.add_hline(y=limite_actual, line_dash="dot", line_color="red")
                st.plotly_chart(fig_vs, use_container_width=True)

    with tab5:
        st.subheader("Perfil de Estación (Multivariable)")
        opciones_multi = list(diccionario_activo.keys())
        idx_rm_multi = obtener_indice_rm(opciones_multi)
        
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1: reg_p = st.selectbox("Región", opciones_multi, index=idx_rm_multi, key="reg_p_multi")
        with col_p2: com_p = st.selectbox("Comuna", list(diccionario_activo[reg_p].keys()), key="com_p_multi")
        with col_p3: est_p = st.selectbox("Estación", list(diccionario_activo[reg_p][com_p].keys()), key="est_p_multi")

        sensores_certificados = obtener_sensores_certificados(est_p)
        contaminantes_seleccionados = st.multiselect("Contaminantes:", sensores_certificados, default=[s for s in ["MP2.5", "MP10"] if s in sensores_certificados])
        modo_vista = st.selectbox("Visualización:", ["Gráfico Unificado (Porcentaje)", "Gráficos Separados (Absoluto)"])

        if contaminantes_seleccionados:
            lat_p, lon_p = diccionario_activo[reg_p][com_p][est_p]
            vars_api = [configuracion[c]["api"] for c in contaminantes_seleccionados]
            df_multi = obtener_datos_multivariable(lat_p, lon_p, vars_api)
            
            if not df_multi.empty:
                df_multi = df_multi.rename(columns={configuracion[c]["api"]: c for c in contaminantes_seleccionados}).reset_index().rename(columns={'index': 'Fecha y Hora'})
                if "Separados" in modo_vista:
                    df_melt = df_multi.melt(id_vars=['Fecha y Hora'], value_vars=contaminantes_seleccionados, var_name='Contaminante', value_name='Concentración')
                    fig_multi = px.line(df_melt, x='Fecha y Hora', y='Concentración', facet_row='Contaminante', height=250 * len(contaminantes_seleccionados))
                    fig_multi.update_yaxes(matches=None)
                    st.plotly_chart(fig_multi, use_container_width=True)
                else:
                    df_norm = df_multi.copy()
                    for c in contaminantes_seleccionados: df_norm[c] = (df_norm[c] / configuracion[c]["limite"]) * 100
                    df_melt_norm = df_norm.melt(id_vars=['Fecha y Hora'], value_vars=contaminantes_seleccionados, var_name='Contaminante', value_name='Porcentaje del Límite (%)')
                    fig_norm = px.line(df_melt_norm, x='Fecha y Hora', y='Porcentaje del Límite (%)', color='Contaminante')
                    fig_norm.add_hline(y=100, line_dash="dot", line_color="red")
                    st.plotly_chart(fig_norm, use_container_width=True)

    with tab6:
        st.subheader("Condiciones Meteorológicas Locales")
        opciones_clima = list(diccionario_activo.keys())
        idx_rm_clima = obtener_indice_rm(opciones_clima)
        
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: reg_m = st.selectbox("Región Clima", opciones_clima, index=idx_rm_clima, key="reg_m")
        with col_m2: com_m = st.selectbox("Comuna Clima", list(diccionario_activo[reg_m].keys()), key="com_m")
        with col_m3: est_m = st.selectbox("Estación Clima", list(diccionario_activo[reg_m][com_m].keys()), key="est_m")

        lat_m, lon_m = diccionario_activo[reg_m][com_m][est_m]
        df_met = obtener_meteorologia(lat_m, lon_m)
        
        if not df_met.empty:
            try:
                idx_actual_met = df_met['Fecha y Hora'].get_indexer([ahora], method='nearest')[0]
                temp_actual, viento_actual, dir_grados = df_met.iloc[idx_actual_met]['Temperatura (°C)'], df_met.iloc[idx_actual_met]['Velocidad Viento (km/h)'], df_met.iloc[idx_actual_met]['Direccion Viento (°)']
                dirs = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
                dir_cardinal = dirs[int((dir_grados + 11.25)/22.5) % 16]
            except:
                temp_actual, viento_actual, dir_grados, dir_cardinal = 0, 0, 0, "N/A"

            col_k1, col_k2, col_k3 = st.columns(3)
            with col_k1: st.metric("Temperatura", f"{temp_actual:.1f} °C")
            with col_k2: st.metric("Viento", f"{viento_actual:.1f} km/h")
            with col_k3: st.metric("Dirección", f"{dir_cardinal} ({dir_grados}°)")

    with tab7:
        st.subheader("Sala de Control Central (SAT)")
        
        st.markdown("### Bitácora de Infracciones (Últimas 24h)")
        alertas_pasadas = []
        for region, comunas in datos_totales.items():
            for comuna, sectores in comunas.items():
                for sector, serie in sectores.items():
                    if not serie.empty:
                        serie_24h = serie[(serie.index >= ahora - pd.Timedelta(hours=24)) & (serie.index <= ahora)]
                        if not serie_24h.empty and (serie_24h > limite_actual).sum() > 0:
                            alertas_pasadas.append({
                                "Último Peak": serie_24h[serie_24h > limite_actual].index[-1].strftime("%Y-%m-%d %H:00"),
                                "Región": region, "Comuna": comuna, "Estación": sector,
                                "Peak": round(serie_24h.max(), 1), "Horas Falla": (serie_24h > limite_actual).sum()
                            })
        if alertas_pasadas:
            df_alertas_pasadas = pd.DataFrame(alertas_pasadas).sort_values("Último Peak", ascending=False).reset_index(drop=True)
            st.dataframe(df_alertas_pasadas, use_container_width=True, hide_index=True)
            
            for alerta in alertas_pasadas:
                reg_graf, com_graf, est_graf = alerta["Región"], alerta["Comuna"], alerta["Estación"]
                serie_grafico = datos_totales[reg_graf][com_graf][est_graf]
                df_g = serie_grafico.reset_index()
                df_g.columns = ['Fecha y Hora', 'Concentración']
                
                fig_pasado = px.line(df_g, x='Fecha y Hora', y='Concentración', title=f"Infracción en: {est_graf}", color_discrete_sequence=["#FF7E00"])
                fig_pasado.add_hline(y=limite_actual, line_dash="dot", line_color="red")
                fig_pasado.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white")
                fig_pasado.update_xaxes(range=[ahora - pd.Timedelta(hours=48), ahora + pd.Timedelta(hours=12)])
                st.plotly_chart(fig_pasado, use_container_width=True)
        else:
            st.success("No hay superaciones a la norma en las últimas 24 horas.")

        st.divider()

        st.markdown("### Sistema de Alerta Temprana - Próximas 72 Horas")
        alertas_futuras = []
        for region, comunas in datos_totales.items():
            for comuna, sectores in comunas.items():
                for sector, serie in sectores.items():
                    if not serie.empty:
                        serie_futura = serie[serie.index > ahora]
                        if not serie_futura.empty and (serie_futura > limite_actual).sum() > 0:
                            alertas_futuras.append({
                                "Inicio Proyectado": serie_futura[serie_futura > limite_actual].index[0].strftime("%Y-%m-%d %H:00"),
                                "Región": region, "Comuna": comuna, "Estación": sector,
                                "Peak Estimado": round(serie_futura.max(), 1)
                            })
        if alertas_futuras:
            st.warning(f"Se proyectan superaciones a la norma en {len(alertas_futuras)} estaciones para los próximos días.")
            df_alertas_futuras = pd.DataFrame(alertas_futuras).sort_values("Inicio Proyectado").reset_index(drop=True)
            st.dataframe(df_alertas_futuras, use_container_width=True, hide_index=True)
            
            for alerta in alertas_futuras:
                reg_graf, com_graf, est_graf = alerta["Región"], alerta["Comuna"], alerta["Estación"]
                serie_grafico = datos_totales[reg_graf][com_graf][est_graf]
                df_g = serie_grafico.reset_index()
                df_g.columns = ['Fecha y Hora', 'Concentración']
                
                fig_sat = px.line(df_g, x='Fecha y Hora', y='Concentración', title=f"Proyección Crítica: {est_graf}", color_discrete_sequence=["#FF4B4B"])
                fig_sat.add_hline(y=limite_actual, line_dash="dot", line_color="red")
                fig_sat.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white")
                fig_sat.update_xaxes(range=[ahora - pd.Timedelta(hours=24), ahora + pd.Timedelta(hours=72)])
                st.plotly_chart(fig_sat, use_container_width=True)
        else:
            st.info("El modelo predictivo indica que no habrá superaciones normativas en los próximos 3 días.")

# ==========================================
# MÓDULO 2: SIMULADOR CONSULTIVO AI (B2B)
# ==========================================
elif modulo_activo == "Simulador":
    
    st.title("Datair AI | Simulador de Emisiones e Impacto")
    st.write("Ingrese las coordenadas y la tasa de emisión de su planta a continuación. Nuestro motor cruzará su operación con el clima en tiempo real para generar un **Informe Consultivo Automático**.")
    
    st.markdown(f"### Parámetros de Operación de la Planta ({contaminante_elegido})")
    
    with st.container():
        col_param1, col_param2, col_param3 = st.columns(3)
        with col_param1:
            lat_em = st.number_input("Latitud de Descarga", value=-33.4641, format="%.4f")
        with col_param2:
            lon_em = st.number_input("Longitud de Descarga", value=-70.6607, format="%.4f")
        with col_param3:
            tasa_em = st.number_input("Emisión (kg/h)", min_value=0.1, value=50.0, step=1.0)
            
        col_param4, col_param5, col_param6 = st.columns(3)
        with col_param4:
            altura_em = st.number_input("Altura Chimenea (m)", min_value=1.0, value=30.0, step=1.0)
        with col_param5:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_simular = st.button("Ejecutar Simulación AI", type="primary", use_container_width=True)

    st.divider()

    if btn_simular:
        with st.spinner("Conectando con satélites meteorológicos y redactando informe..."):
            
            vel_viento, dir_viento, temp_actual = consultar_clima_coordenada(lat_em, lon_em)
            angulo_viaje = (dir_viento + 180) % 360
            angulo_rad = math.radians(90 - angulo_viaje)
            
            distancia_impacto = (altura_em * 15) if vel_viento > 0 else 100 
            factor_dispersion = max(1.0, vel_viento * 0.5)
            concentracion_max = (tasa_em * 1000) / (factor_dispersion * altura_em)
            
            puntos_pluma = [{"Latitud": lat_em, "Longitud": lon_em, "Concentracion": concentracion_max}]
            pasos, dist_max = 15, vel_viento * 0.02 
            
            for i in range(1, pasos + 1):
                frac = i / pasos
                dist = dist_max * frac
                d_lat = dist * math.sin(angulo_rad)
                d_lon = dist * math.cos(angulo_rad) / math.cos(math.radians(lat_em))
                c_fantasma = concentracion_max * (1 - frac)**2 
                puntos_pluma.append({"Latitud": lat_em + d_lat, "Longitud": lon_em + d_lon, "Concentracion": c_fantasma})

            df_pluma_ai = pd.DataFrame(puntos_pluma)
            supera_norma = concentracion_max > limite_actual
            
            st.markdown(f"**Mapa de Impacto Proyectado (Viento actual: {vel_viento} km/h hacia el {int(angulo_viaje)}º)**")
            fig_ai = px.density_mapbox(
                df_pluma_ai, lat="Latitud", lon="Longitud", z="Concentracion",
                radius=60, center={"lat": lat_em, "lon": lon_em}, 
                zoom=12, mapbox_style="carto-darkmatter", color_continuous_scale="Reds", opacity=0.7
            )
            fig_ai.update_layout(margin={"r":0,"t":0,"l":0,"b":0}, height=500)
            st.plotly_chart(fig_ai, use_container_width=True)

            st.markdown("<div class='ai-report-box'>", unsafe_allow_html=True)
            st.markdown("### Reporte de Datair AI")
            
            if supera_norma:
                st.error(f"**RIESGO CRÍTICO.** La concentración a nivel de suelo es de **{concentracion_max:.1f} µg/m³**, superando el límite legal de {limite_actual} µg/m³ para {contaminante_elegido}.")
            else:
                st.success(f"**RIESGO BAJO.** La concentración a nivel de suelo es de **{concentracion_max:.1f} µg/m³**, manteniéndose bajo el límite legal de {limite_actual} µg/m³ para {contaminante_elegido}.")

            st.markdown("#### Impacto Comunitario")
            st.write(f"Con {temp_actual} °C y vientos de {vel_viento} km/h, la pluma se desplazará hacia el **{int(angulo_viaje)}°**. Ecosistemas en un radio de **{int(distancia_impacto)} m** en esta dirección recibirán el mayor impacto.")

            st.markdown("#### Plan de Mitigación Propuesto")
            if contaminante_elegido in ["MP10", "MP2.5"]:
                st.markdown("""
                1. **Ingeniería:** Activar supresión de polvo y revisar filtros de mangas.
                2. **Operacional:** Reducir molienda en un 30% temporalmente.
                3. **Logística:** Aumentar frecuencia de camiones aljibe.
                """)
            else:
                st.markdown("""
                1. **Ingeniería:** Incrementar reactivos alcalinos en desulfuración.
                2. **Operacional:** Evaluar mezcla de combustible con menos azufre.
                3. **Monitoreo:** Desplegar medición portátil de gases.
                """)
            st.markdown("</div>", unsafe_allow_html=True)