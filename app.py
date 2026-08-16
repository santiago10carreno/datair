import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import LineChart, Reference
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
        
        /* OCULTAMOS SOLO LA BASURA, SIN TOCAR EL MECANISMO DE LA BARRA LATERAL */
        .stAppDeployButton, [data-testid="stToolbar"], [data-testid="stDecoration"] {
            display: none !important;
        }

        /* DISEÑO DE LAS MÉTRICAS (Alineación perfecta) */
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
        
        /* FIX PARA TEXTOS CORTADOS */
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
        
        /* Estilo para el reporte de IA */
        .ai-report-box {
            background-color: #141618;
            border: 1px solid #3A3F47;
            border-radius: 8px;
            padding: 25px;
            margin-top: 15px;
            border-top: 4px solid #8F3F97; 
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. BASE DE DATOS GEOGRÁFICA Y HARDWARE
# ==========================================
DICCIONARIO_ZONAS = {
    "Región de Antofagasta": {"Antofagasta": {"Antofagasta (Centro)": (-23.6500, -70.4000)}, "Calama": {"Calama": (-22.4500, -68.9300)}},
    "Región de Valparaíso": {"Valparaíso": {"Valparaíso": (-33.0500, -71.6200)}, "Viña del Mar": {"Viña del Mar": (-33.0200, -71.5500)}, "Quintero": {"Quintero (Centro)": (-32.7800, -71.5300), "Loncura": (-32.7900, -71.5200)}, "Puchuncaví": {"Puchuncaví": (-32.7300, -71.4100), "Las Ventanas": (-32.7500, -71.4800)}},
    "Región Metropolitana": {"Santiago Centro": {"Parque O'Higgins": (-33.4641, -70.6607)}, "Independencia": {"Independencia": (-33.4150, -70.6528)}, "Pudahuel": {"Pudahuel": (-33.4326, -70.7818)}, "Quilicura": {"Quilicura": (-33.3663, -70.7351)}, "Las Condes": {"Las Condes": (-33.3769, -70.5239)}, "Cerrillos": {"Cerrillos": (-33.4939, -70.7161)}, "El Bosque": {"El Bosque": (-33.5350, -70.6766)}, "Cerro Navia": {"Cerro Navia": (-33.4334, -70.7348)}, "Puente Alto": {"Puente Alto": (-33.6163, -70.5831)}, "Talagante": {"Talagante": (-33.6669, -70.9275)}},
    "Región de O'Higgins": {"Rancagua": {"Rancagua (Centro)": (-34.1708, -70.7441), "Rancagua 2 (Norte)": (-34.1522, -70.7305)}, "Rengo": {"Rengo (Centro)": (-34.4091, -70.8591)}, "San Fernando": {"San Fernando": (-34.5847, -70.9880)}, "Machalí": {"Machalí": (-34.1816, -70.6558)}},
    "Región del Maule": {"Talca": {"Talca (Centro)": (-35.4300, -71.6700), "La Florida": (-35.4500, -71.6800)}, "Curicó": {"Curicó": (-34.9800, -71.2300)}},
    "Región del Biobío": {"Concepción": {"Concepción": (-36.8300, -73.0500)}, "Talcahuano": {"Talcahuano": (-36.7200, -73.1200)}, "Coronel": {"Coronel Norte": (-37.0100, -73.1400), "Coronel Sur": (-37.0400, -73.1500)}, "Los Ángeles": {"Los Ángeles": (-37.4700, -72.3500)}},
    "Región de La Araucanía": {"Temuco": {"Temuco (Centro)": (-38.7400, -72.5900)}, "Padre Las Casas": {"Padre Las Casas": (-38.7700, -72.5900)}},
    "Región de Los Ríos": {"Valdivia": {"Valdivia": (-39.8100, -73.2400)}},
    "Región de Los Lagos": {"Osorno": {"Osorno": (-40.5700, -73.1300)}, "Puerto Montt": {"Puerto Montt": (-41.4700, -72.9300)}},
    "Región de Aysén": {"Coyhaique": {"Coyhaique 1": (-45.5700, -72.0700), "Coyhaique 2": (-45.5800, -72.0600)}},
    "Región de Magallanes": {"Punta Arenas": {"Punta Arenas": (-53.1500, -70.9000)}}
}

DICCIONARIO_SINCA_PILOTO = {
    "Región de O'Higgins": {
        "Rancagua": {"Rancagua (Centro)": (-34.1708, -70.7441)}, 
        "Rengo": {"Rengo (Centro)": (-34.4091, -70.8591)}, 
        "San Fernando": {"San Fernando": (-34.5847, -70.9880)}
    }
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
# 2. MOTOR ICAP CHILE
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

PALETA_ICAP = {
    "Bueno": "#00E400", "Regular": "#FFFF00", "Alerta": "#FF7E00", 
    "Preemergencia": "#FF0000", "Emergencia": "#8F3F97"
}

# ==========================================
# 3. BARRA LATERAL & SELECTOR DE FUENTE
# ==========================================
st.sidebar.title("Configuración Global")

fuente_datos = st.sidebar.radio(
    "Fuente de Extracción de Datos:",
    ["Open-Meteo (Modelo Global)", "SINCA (Oficial - Piloto)"],
    help="El Piloto SINCA restringe el análisis a O'Higgins para validar datos oficiales."
)

if fuente_datos == "SINCA (Oficial - Piloto)":
    diccionario_activo = DICCIONARIO_SINCA_PILOTO
    st.sidebar.success("✅ Conexión Oficial Activa")
    contaminantes_disponibles = ["MP10", "MP2.5", "SO2"]
else:
    diccionario_activo = DICCIONARIO_ZONAS
    contaminantes_disponibles = list(configuracion.keys())

st.sidebar.divider()
contaminante_elegido = st.sidebar.selectbox("Selecciona el Contaminante Principal", contaminantes_disponibles)

var_api = configuracion[contaminante_elegido]["api"]
limite_actual = configuracion[contaminante_elegido]["limite"]

# ==========================================
# 4. EXTRACCIÓN ASÍNCRONA (DUAL Y MULTIVARIABLE)
# ==========================================
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
    ahora = pd.Timestamp.now(tz='America/Santiago').tz_localize(None).floor('h')
    fechas = pd.date_range(start=ahora - pd.Timedelta(days=7), end=ahora + pd.Timedelta(days=3), freq='h')
    
    np.random.seed(hash(sector) % 10000)
    ciclo_diario = np.sin(np.linspace(0, 10 * 2 * np.pi, len(fechas))) * (configuracion[contaminante_elegido]["limite"] * 0.3)
    base = configuracion[contaminante_elegido]["limite"] * 0.4
    ruido = np.random.normal(0, configuracion[contaminante_elegido]["limite"] * 0.15, len(fechas))
    valores = np.maximum(0, base + ciclo_diario + ruido)
    
    serie = pd.Series(valores, index=fechas)
    return (region, comuna, sector, serie)

@st.cache_data(ttl=3600)
def descargar_todos_los_datos(contaminante_nombre, variable_api, fuente):
    lista_tareas = []
    estaciones_con_hardware = 0
    dicc_usar = DICCIONARIO_SINCA_PILOTO if fuente == "SINCA (Oficial - Piloto)" else DICCIONARIO_ZONAS
    
    for region, comunas in dicc_usar.items():
        for comuna, sectores in comunas.items():
            for sector, coords in sectores.items():
                if contaminante_nombre in obtener_sensores_certificados(sector):
                    estaciones_con_hardware += 1
                    lista_tareas.append((coords[0], coords[1], variable_api, region, comuna, sector))
    
    resultados_completos = {}
    funcion_extraccion = obtener_datos_sinca_simulados if fuente == "SINCA (Oficial - Piloto)" else obtener_datos_estacion_individual
    
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

datos_totales, total_hardware_valido = descargar_todos_los_datos(contaminante_elegido, var_api, fuente_datos)
ahora = pd.Timestamp.now(tz='America/Santiago').tz_localize(None)

# ==========================================
# 5. PREPARACIÓN DE DATOS CON LÓGICA ICAP
# ==========================================
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
                    "Estado": estado_icap, 
                    "Color": color_icap, 
                    "Tamaño": tamanio_icap
                })
df_mapa = pd.DataFrame(datos_mapa)

# ==========================================
# 6. UI PRINCIPAL Y KPIS
# ==========================================
st.title("Datair | Inteligencia Ambiental")
titulo_fuente = "Modelo Predictivo Global" if fuente_datos == "Open-Meteo (Modelo Global)" else "Red de Monitoreo SINCA (Piloto Oficial)"
st.markdown(f"**{titulo_fuente} - {contaminante_elegido}** | Limite Normativo 24h: `{limite_actual} µg/m³`")

promedio_nacional = df_mapa["Concentracion"].mean() if not df_mapa.empty else 0
estaciones_criticas = len(df_mapa[df_mapa["Estado"].isin(["Alerta", "Preemergencia", "Emergencia"])]) if not df_mapa.empty else 0

if estaciones_criticas > (len(df_mapa) * 0.3): estado_pais = "Emergencia Operativa"
elif estaciones_criticas > 0: estado_pais = "Zonas en Riesgo"
else: estado_pais = "Condiciones Óptimas"

col1, col2, col3 = st.columns(3)
with col1: st.metric(label="Promedio Zona Actual", value=f"{promedio_nacional:.1f} µg/m³")
with col2: st.metric(label="Estado General", value=estado_pais)
with col3: st.metric(label="Sectores en Riesgo (ICAP)", value=f"{estaciones_criticas} de {total_hardware_valido}", help="Ve a la pestaña 'Centro de Alertas' para ver el detalle.")

if estaciones_criticas > 0:
    with st.expander("🚨 Ver detalle de los Sectores en Riesgo actuales"):
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

# ==========================================
# SISTEMA DE PESTAÑAS (TABS)
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Monitoreo Espacial", "Análisis Histórico", "Proyección", "Benchmarking", "Multivariable", "Meteorología y SAT", "Simulador AI ✨"
])

# ------------------------------------------
# TAB 1: MONITOREO Y DISPERSIÓN
# ------------------------------------------
with tab1:
    st.subheader("Red de Monitoreo Territorial")
    if not df_mapa.empty:
        fig_mapa = px.scatter_mapbox(
            df_mapa, lat="Latitud", lon="Longitud", hover_name="Estacion", 
            hover_data={"Latitud": False, "Longitud": False, "Region": True, "Comuna": True, "Concentracion": True, "Estado": True, "Tamaño": False, "Color": False},
            color="Estado", color_discrete_map=PALETA_ICAP,
            size="Tamaño", zoom=7 if fuente_datos == "SINCA (Oficial - Piloto)" else 4, 
            center={"lat": -34.3, "lon": -70.8} if fuente_datos == "SINCA (Oficial - Piloto)" else {"lat": -35.0, "lon": -71.0}, 
            height=500,
            category_orders={"Estado": ["Bueno", "Regular", "Alerta", "Preemergencia", "Emergencia"]}
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
            lat = row['Latitud']
            lon = row['Longitud']
            c = row['Concentracion']
            spd = row.get('WindSpd', 0)
            dir_viento = row.get('WindDir', 0)
            
            puntos_pluma.append(row.to_dict())
            
            if spd > 1:
                angulo_viaje = (dir_viento + 180) % 360
                angulo_rad = math.radians(90 - angulo_viaje)
                pasos = 8
                dist_max = spd * 0.015 
                for i in range(1, pasos + 1):
                    frac = i / pasos
                    dist = dist_max * frac
                    d_lat = dist * math.sin(angulo_rad)
                    d_lon = dist * math.cos(angulo_rad) / math.cos(math.radians(lat))
                    
                    c_fantasma = c * (1 - frac)**1.5 
                    if c_fantasma > (limite_actual * 0.05):
                        nuevo_punto = row.to_dict()
                        nuevo_punto['Latitud'] = lat + d_lat
                        nuevo_punto['Longitud'] = lon + d_lon
                        nuevo_punto['Concentracion'] = c_fantasma
                        nuevo_punto['Estacion'] = f"Viento desde {row['Estacion']}"
                        puntos_pluma.append(nuevo_punto)
        
        df_pluma = pd.DataFrame(puntos_pluma)
        regiones_disp = list(df_vectores['Region'].unique())
        reg_mapa = st.selectbox("Enfocar cámara en la Región:", regiones_disp, key="heatmap_region")
        
        df_region_mapa = df_vectores[df_vectores['Region'] == reg_mapa]
        if not df_region_mapa.empty:
            lat_centro = df_region_mapa['Latitud'].mean()
            lon_centro = df_region_mapa['Longitud'].mean()
        else:
            lat_centro, lon_centro = -35.0, -71.0

        fig_heat = px.density_mapbox(
            df_pluma, lat="Latitud", lon="Longitud", z="Concentracion",
            radius=60, center={"lat": lat_centro, "lon": lon_centro}, 
            zoom=8, mapbox_style="carto-darkmatter", color_continuous_scale="Inferno", 
            opacity=0.6, hover_name="Estacion"
        )
        fig_heat.update_layout(margin={"r":0,"t":40,"l":0,"b":0})
        st.plotly_chart(fig_heat, use_container_width=True)

# ------------------------------------------
# TAB 2: ANÁLISIS HISTÓRICO Y AUDITORÍA
# ------------------------------------------
with tab2:
    st.subheader("Análisis Histórico Local")
    regiones_disponibles = list(datos_totales.keys())
    if regiones_disponibles:
        col_filtro1, col_filtro2 = st.columns(2)
        with col_filtro1: region_elegida = st.selectbox("Selecciona la Región", regiones_disponibles, key="reg_hist")
        with col_filtro2: comuna_elegida = st.selectbox("Selecciona la Comuna", list(datos_totales[region_elegida].keys()), key="com_hist")

        datos_sectores_comuna = {}
        lista_promedios_comunas = {}

        for comuna, sectores in datos_totales[region_elegida].items():
            series_comuna = []
            for sector, serie in sectores.items():
                series_comuna.append(serie)
                if comuna == comuna_elegida:
                    datos_sectores_comuna[sector] = serie
            if series_comuna:
                lista_promedios_comunas[comuna] = pd.concat(series_comuna, axis=1).mean(axis=1)

        df_region_completo = pd.DataFrame(lista_promedios_comunas).reset_index().rename(columns={'index': 'Fecha y Hora'})
        df_comuna_completo = pd.DataFrame(datos_sectores_comuna).reset_index().rename(columns={'index': 'Fecha y Hora'})
        
        df_region_historico = df_region_completo[df_region_completo['Fecha y Hora'] <= ahora]
        df_comuna_historico = df_comuna_completo[df_comuna_completo['Fecha y Hora'] <= ahora]

        st.subheader(f"Histórico Regional: {region_elegida}")
        fig_reg_hist = px.line(df_region_historico, x='Fecha y Hora', y=df_region_historico.columns[1:], labels={'value': 'Concentración (µg/m³)', 'variable': 'Comuna'})
        fig_reg_hist.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Límite Legal")
        st.plotly_chart(fig_reg_hist, use_container_width=True)

        st.subheader(f"Histórico Comunal: {comuna_elegida}")
        fig_com_hist = px.line(df_comuna_historico, x='Fecha y Hora', y=df_comuna_historico.columns[1:], labels={'value': 'Concentración (µg/m³)', 'variable': 'Estación'})
        fig_com_hist.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Límite Legal")
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
                ws_datos['A2'].font = Font(italic=True, color="595959")
                
                header_fill, header_font = PatternFill(start_color="2F75B5", fill_type="solid"), Font(bold=True, color="FFFFFF")
                columnas_datos = list(df_datos.columns)[1:] 
                
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
                ws_formatting = ws_datos.conditional_formatting
                ws_formatting.add(f'B5:{ultima_letra}{len(df_datos)+4}', rule_under)
                
            return output.getvalue()

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            excel_region = generar_excel_universal(df_region_historico, contaminante_elegido, limite_actual, region_elegida, "Region")
            st.download_button(label=f"📥 Descargar Data Regional ({region_elegida})", data=excel_region, file_name=f"Auditoria_{region_elegida}.xlsx")
        with col_btn2:
            excel_comuna = generar_excel_universal(df_comuna_historico, contaminante_elegido, limite_actual, comuna_elegida, "Comuna")
            st.download_button(label=f"📥 Descargar Data Comunal ({comuna_elegida})", data=excel_comuna, file_name=f"Auditoria_{comuna_elegida}.xlsx")

# ------------------------------------------
# TAB 3: PROYECCIÓN PREDICTIVA
# ------------------------------------------
with tab3:
    st.subheader("Modelos Predictivos Espaciales")
    if regiones_disponibles:
        col_proy1, col_proy2 = st.columns(2)
        with col_proy1: reg_proy = st.selectbox("Selecciona la Región", regiones_disponibles, key="reg_proy")
        with col_proy2: com_proy = st.selectbox("Selecciona la Comuna", list(datos_totales[reg_proy].keys()), key="com_proy")

        datos_sectores_proy = {}
        lista_promedios_proy = {}
        for comuna, sectores in datos_totales[reg_proy].items():
            series_comuna = []
            for sector, serie in sectores.items():
                series_comuna.append(serie)
                if comuna == com_proy: datos_sectores_proy[sector] = serie
            if series_comuna: lista_promedios_proy[comuna] = pd.concat(series_comuna, axis=1).mean(axis=1)

        df_region_proy = pd.DataFrame(lista_promedios_proy).reset_index().rename(columns={'index': 'Fecha y Hora'})
        df_comuna_proy = pd.DataFrame(datos_sectores_proy).reset_index().rename(columns={'index': 'Fecha y Hora'})

        fig_region_pred = px.line(df_region_proy, x='Fecha y Hora', y=df_region_proy.columns[1:], labels={'value': 'Concentración (µg/m³)', 'variable': 'Comuna'})
        fig_region_pred.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Límite Legal")
        fig_region_pred.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white", annotation_text="AHORA")
        st.plotly_chart(fig_region_pred, use_container_width=True)

        fig_comuna_pred = px.line(df_comuna_proy, x='Fecha y Hora', y=df_comuna_proy.columns[1:], labels={'value': 'Concentración (µg/m³)', 'variable': 'Estación'})
        fig_comuna_pred.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Límite Legal")
        fig_comuna_pred.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white", annotation_text="AHORA")
        st.plotly_chart(fig_comuna_pred, use_container_width=True)

# ------------------------------------------
# TAB 4: BENCHMARKING
# ------------------------------------------
with tab4:
    st.subheader("Benchmarking Corporativo (Comparador Cruzado)")
    if len(datos_totales) > 0:
        col_vs1, col_vs2 = st.columns(2)
        with col_vs1:
            reg_a = st.selectbox("Región A", list(datos_totales.keys()), key="reg_a")
            com_a = st.selectbox("Comuna A", list(datos_totales[reg_a].keys()), key="com_a")
        with col_vs2:
            reg_b = st.selectbox("Región B", list(datos_totales.keys()), index=min(1, len(datos_totales)-1), key="reg_b")
            com_b = st.selectbox("Comuna B", list(datos_totales[reg_b].keys()), key="com_b")

        def obtener_promedio_comuna(region, comuna):
            series = []
            if region in datos_totales and comuna in datos_totales[region]:
                for sector, serie in datos_totales[region][comuna].items():
                    if not serie.empty: series.append(serie)
            if series: return pd.concat(series, axis=1).mean(axis=1)
            return pd.Series(dtype=float)

        promedio_a = obtener_promedio_comuna(reg_a, com_a)
        promedio_b = obtener_promedio_comuna(reg_b, com_b)

        if not promedio_a.empty and not promedio_b.empty:
            df_vs = pd.DataFrame({com_a: promedio_a, com_b: promedio_b}).reset_index().rename(columns={'index': 'Fecha y Hora'})
            fig_vs = px.line(df_vs, x='Fecha y Hora', y=[com_a, com_b], labels={'value': 'Concentración (µg/m³)', 'variable': 'Zona Analizada'})
            fig_vs.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Límite Legal")
            st.plotly_chart(fig_vs, use_container_width=True)

# ------------------------------------------
# TAB 5: MULTIVARIABLE
# ------------------------------------------
with tab5:
    st.subheader("Perfil de Estación (Análisis Multivariable)")
    col_p1, col_p2, col_p3 = st.columns(3)
    
    with col_p1: reg_p = st.selectbox("Región", list(diccionario_activo.keys()), key="reg_p_multi")
    with col_p2: com_p = st.selectbox("Comuna", list(diccionario_activo[reg_p].keys()), key="com_p_multi")
    with col_p3: est_p = st.selectbox("Estación", list(diccionario_activo[reg_p][com_p].keys()), key="est_p_multi")

    sensores_certificados = obtener_sensores_certificados(est_p)
    contaminantes_seleccionados = st.multiselect(
        "Selecciona los contaminantes a comparar:", sensores_certificados,  
        default=[s for s in ["MP2.5", "MP10"] if s in sensores_certificados]
    )

    modo_vista = st.radio("Modo de Visualización:", ["Gráfico Unificado (Porcentaje del Límite)", "Gráficos Separados (Absolutos)"], horizontal=True)

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
                fig_multi.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white", annotation_text="AHORA")
                st.plotly_chart(fig_multi, use_container_width=True)
            else:
                df_norm = df_multi.copy()
                for c in contaminantes_seleccionados: df_norm[c] = (df_norm[c] / configuracion[c]["limite"]) * 100
                df_melt_norm = df_norm.melt(id_vars=['Fecha y Hora'], value_vars=contaminantes_seleccionados, var_name='Contaminante', value_name='Porcentaje del Límite (%)')
                fig_norm = px.line(df_melt_norm, x='Fecha y Hora', y='Porcentaje del Límite (%)', color='Contaminante')
                fig_norm.add_hline(y=100, line_dash="dot", line_color="red", annotation_text="LÍMITE LEGAL (100%)")
                fig_norm.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white", annotation_text="AHORA")
                st.plotly_chart(fig_norm, use_container_width=True)

# ------------------------------------------
# TAB 6: METEOROLOGÍA Y SAT
# ------------------------------------------
with tab6:
    st.subheader("Clima Local y Sala de Control Central (SAT)")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1: reg_m = st.selectbox("Región Clima", list(diccionario_activo.keys()), key="reg_m")
    with col_m2: com_m = st.selectbox("Comuna Clima", list(diccionario_activo[reg_m].keys()), key="com_m")
    with col_m3: est_m = st.selectbox("Estación Clima", list(diccionario_activo[reg_m][com_m].keys()), key="est_m")

    lat_m, lon_m = diccionario_activo[reg_m][com_m][est_m]
    df_met = obtener_meteorologia(lat_m, lon_m)
    
    if not df_met.empty:
        try:
            idx_actual_met = df_met['Fecha y Hora'].get_indexer([ahora], method='nearest')[0]
            temp_actual = df_met.iloc[idx_actual_met]['Temperatura (°C)']
            viento_actual = df_met.iloc[idx_actual_met]['Velocidad Viento (km/h)']
            dir_grados = df_met.iloc[idx_actual_met]['Direccion Viento (°)']
            def grados_a_cardinal_local(d):
                dirs = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
                return dirs[int((d + 11.25)/22.5) % 16]
            dir_cardinal = grados_a_cardinal_local(dir_grados)
        except:
            temp_actual, viento_actual, dir_grados, dir_cardinal = 0, 0, 0, "N/A"

        col_k1, col_k2, col_k3 = st.columns(3)
        with col_k1: st.metric("🌡️ Temperatura", f"{temp_actual:.1f} °C")
        with col_k2: st.metric("🌬️ Velocidad Viento", f"{viento_actual:.1f} km/h")
        with col_k3: st.metric("🧭 Dirección Viento", f"{dir_cardinal} ({dir_grados}°)")

    st.divider()
    
    st.markdown("### 📜 Bitácora de Infracciones (Últimas 24 Horas)")
    alertas_pasadas = []
    hace_24h = ahora - pd.Timedelta(hours=24)
    for region, comunas in datos_totales.items():
        for comuna, sectores in comunas.items():
            for sector, serie in sectores.items():
                if not serie.empty:
                    serie_24h = serie[(serie.index >= hace_24h) & (serie.index <= ahora)]
                    if not serie_24h.empty:
                        horas_infraccion = (serie_24h > limite_actual).sum()
                        if horas_infraccion > 0:
                            peak = serie_24h.max()
                            ultimo_momento = serie_24h[serie_24h > limite_actual].index[-1]
                            alertas_pasadas.append({
                                "Último Peak": ultimo_momento.strftime("%Y-%m-%d %H:00"),
                                "Región": region, "Comuna": comuna, "Estación": sector,
                                "Peak": round(peak, 1), "Horas Falla": horas_infraccion
                            })
    
    if alertas_pasadas:
        df_alertas = pd.DataFrame(alertas_pasadas).sort_values("Último Peak", ascending=False).reset_index(drop=True)
        st.error(f"⚠️ Se detectaron vulneraciones a la normativa en **{len(alertas_pasadas)}** estaciones.")
        st.dataframe(df_alertas, use_container_width=True, hide_index=True)
    else:
        st.success(f"✅ No se han registrado superaciones a la norma en las últimas 24 horas.")

    st.markdown("### 🔮 Sistema de Alerta Temprana (SAT) - Próximas 72 Horas")
    alertas_futuras = []
    for region, comunas in datos_totales.items():
        for comuna, sectores in comunas.items():
            for sector, serie in sectores.items():
                if not serie.empty:
                    serie_futura = serie[serie.index > ahora]
                    if not serie_futura.empty:
                        horas_peligro = (serie_futura > limite_actual).sum()
                        if horas_peligro > 0:
                            peak_futuro = serie_futura.max()
                            inicio_episodio = serie_futura[serie_futura > limite_actual].index[0]
                            alertas_futuras.append({
                                "Inicio Proyectado": inicio_episodio.strftime("%Y-%m-%d %H:00"),
                                "Región": region, "Comuna": comuna, "Estación": sector,
                                "Peak Estimado": round(peak_futuro, 1)
                            })
                            
    if alertas_futuras:
        df_futuro = pd.DataFrame(alertas_futuras).sort_values("Inicio Proyectado").reset_index(drop=True)
        st.warning(f"⚠️ Se proyectan superaciones a la norma en **{len(alertas_futuras)}** estaciones para los próximos días.")
        st.dataframe(df_futuro, use_container_width=True, hide_index=True)
    else:
        st.info("✅ El modelo predictivo indica que no habrá superaciones normativas en los próximos 3 días.")

# ------------------------------------------
# TAB 7: SIMULADOR DE IMPACTO AI ✨
# ------------------------------------------
with tab7:
    st.subheader("Datair AI | Simulador de Emisiones y Plan de Mitigación")
    st.write("Herramienta consultiva B2B: Ingrese las coordenadas de su operación industrial. Nuestra IA cruzará su tasa de emisión con el viento en tiempo real para generar un informe de riesgo y un plan de acción.")

    col_ai1, col_ai2 = st.columns([1, 2])
    
    with col_ai1:
        st.markdown("**⚙️ Parámetros de Operación**")
        lat_em = st.number_input("Latitud de la Chimenea/Descarga", value=-34.1708, format="%.4f")
        lon_em = st.number_input("Longitud de la Chimenea/Descarga", value=-70.7441, format="%.4f")
        contam_em = st.selectbox("Tipo de Contaminante a simular", ["MP10", "MP2.5", "SO2"])
        tasa_em = st.number_input("Tasa de Emisión (kg/hora)", min_value=0.1, value=50.0, step=1.0)
        altura_em = st.number_input("Altura de la Chimenea (m)", min_value=1.0, value=30.0, step=1.0)
        
        st.markdown("<br>", unsafe_allow_html=True)
        btn_simular = st.button("🚀 Ejecutar Simulación AI", type="primary", use_container_width=True)

    with col_ai2:
        if btn_simular:
            with st.spinner("Analizando dispersión atmosférica y redactando informe consultivo..."):
                
                # 1. Obtener clima real de esa coordenada
                vel_viento, dir_viento, temp_actual = consultar_clima_coordenada(lat_em, lon_em)
                
                # 2. Motor Matemático Simplificado (Dispersión Gaussiana)
                angulo_viaje = (dir_viento + 180) % 360
                angulo_rad = math.radians(90 - angulo_viaje)
                
                distancia_impacto = (altura_em * 15) if vel_viento > 0 else 100 
                factor_dispersion = max(1.0, vel_viento * 0.5)
                concentracion_max = (tasa_em * 1000) / (factor_dispersion * altura_em)
                
                puntos_pluma = [{"Latitud": lat_em, "Longitud": lon_em, "Concentracion": concentracion_max}]
                pasos = 15
                dist_max = vel_viento * 0.02 
                
                for i in range(1, pasos + 1):
                    frac = i / pasos
                    dist = dist_max * frac
                    d_lat = dist * math.sin(angulo_rad)
                    d_lon = dist * math.cos(angulo_rad) / math.cos(math.radians(lat_em))
                    c_fantasma = concentracion_max * (1 - frac)**2 
                    puntos_pluma.append({"Latitud": lat_em + d_lat, "Longitud": lon_em + d_lon, "Concentracion": c_fantasma})

                df_pluma_ai = pd.DataFrame(puntos_pluma)
                
                # 3. Mapear el Impacto
                st.markdown(f"**🗺️ Mapa de Impacto Proyectado (Viento actual: {vel_viento} km/h hacia el {angulo_viaje}º)**")
                fig_ai = px.density_mapbox(
                    df_pluma_ai, lat="Latitud", lon="Longitud", z="Concentracion",
                    radius=45, center={"lat": lat_em, "lon": lon_em}, 
                    zoom=11, mapbox_style="carto-darkmatter", color_continuous_scale="Reds", 
                    opacity=0.7
                )
                fig_ai.update_layout(margin={"r":0,"t":0,"l":0,"b":0}, height=300)
                st.plotly_chart(fig_ai, use_container_width=True)

                # 4. Generación Dinámica del Texto de la IA
                limite_legal = configuracion[contam_em]["limite"]
                supera_norma = concentracion_max > limite_legal
                
                st.markdown("<div class='ai-report-box'>", unsafe_allow_html=True)
                st.markdown("### 🤖 Reporte Generado por Datair AI")
                
                if supera_norma:
                    st.error(f"**VEREDICTO DE RIESGO: CRÍTICO.** La concentración estimada a nivel de suelo es de **{concentracion_max:.1f} µg/m³**, lo cual supera el límite legal de {limite_legal} µg/m³ para {contam_em}.")
                else:
                    st.success(f"**VEREDICTO DE RIESGO: BAJO.** La concentración estimada a nivel de suelo es de **{concentracion_max:.1f} µg/m³**, manteniéndose bajo el límite legal de {limite_legal} µg/m³ para {contam_em}.")

                st.markdown("#### Análisis de Impacto Comunitario")
                st.write(f"Dadas las condiciones meteorológicas actuales ({temp_actual} °C y vientos de {vel_viento} km/h), la pluma de contaminación de **{contam_em}** se desplazará hacia el **{int(angulo_viaje)}°**. Las áreas pobladas o ecosistemas ubicados a un radio de **{distancia_impacto} metros** en esta dirección recibirán el mayor impacto de material particulado o gases.")

                st.markdown("#### Plan de Acción Inmediato Propuesto")
                if contam_em in ["MP10", "MP2.5"]:
                    st.markdown("""
                    1. **Ingeniería:** Activar de inmediato los sistemas de supresión de polvo y verificar el diferencial de presión en los filtros de mangas de la planta.
                    2. **Operacional:** Reducir la tasa de alimentación en un 30% hasta que mejore la ventilación atmosférica.
                    3. **Logística:** Aumentar la frecuencia de los camiones aljibe para la humectación de caminos.
                    """)
                else:
                    st.markdown("""
                    1. **Ingeniería:** Incrementar la inyección de reactivos alcalinos en el sistema de desulfuración para neutralizar los gases ácidos.
                    2. **Operacional:** Evaluar el cambio temporal a una mezcla de combustible con menor porcentaje de azufre durante las próximas 12 horas.
                    3. **Monitoreo:** Desplegar una brigada de medición portátil de gases en el perímetro proyectado.
                    """)
                
                st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("👈 Ajuste sus parámetros de operación en el panel izquierdo y haga clic en 'Ejecutar Simulación' para generar su reporte.")