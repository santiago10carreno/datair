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

# ==========================================
# CONFIGURACIÓN Y ESTILOS CORPORATIVOS
# ==========================================
st.set_page_config(page_title="Datair | Inteligencia Ambiental", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* NUEVO: Hacer la barra lateral más angosta */
        section[data-testid="stSidebar"] {
            width: 200px !important;
            min-width: 200px !important;
        }
        
        div[data-testid="metric-container"] {
            background-color: #1A1C1E;
            border: 1px solid #2D3139;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        h1 { color: #4A90E2 !important; font-weight: 700 !important; letter-spacing: -0.5px; }
        h2, h3 { font-weight: 600 !important; letter-spacing: -0.3px; color: #E2E8F0 !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. BASE DE DATOS GEOGRÁFICA
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

TOTAL_ESTACIONES = sum(len(sectores) for comunas in DICCIONARIO_ZONAS.values() for sectores in comunas.values())

configuracion = {
    "MP2.5": {"api": "pm2_5", "limite": 50.0},
    "MP10": {"api": "pm10", "limite": 150.0},
    "SO2": {"api": "sulphur_dioxide", "limite": 250.0},
    "CO": {"api": "carbon_monoxide", "limite": 10000.0}, 
    "NO2": {"api": "nitrogen_dioxide", "limite": 400.0}, 
    "O3": {"api": "ozone", "limite": 120.0}
}

# ==========================================
# 2. BARRA LATERAL (PORTAL DE CLIENTE)
# ==========================================
st.sidebar.title("🏢 Portal de Empresa")
cliente_activo = st.sidebar.selectbox("Sesion iniciada como:", [
    "Demo Publica", 
    "Agroindustria Valle Central", 
    "Minera Los Andes",
    "Constructora Metropolitana"
])

st.sidebar.divider()

st.sidebar.title("⚙️ Configuracion Global")
contaminante_elegido = st.sidebar.selectbox("Selecciona el Contaminante", list(configuracion.keys()))
var_api = configuracion[contaminante_elegido]["api"]
limite_actual = configuracion[contaminante_elegido]["limite"]

# ==========================================
# 3. EXTRACCIÓN ASÍNCRONA
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

@st.cache_data(ttl=3600)
def descargar_todos_los_datos(variable_api):
    lista_tareas = []
    for region, comunas in DICCIONARIO_ZONAS.items():
        for comuna, sectores in comunas.items():
            for sector, coords in sectores.items():
                lista_tareas.append((coords[0], coords[1], variable_api, region, comuna, sector))
    
    resultados_completos = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        resultados = executor.map(obtener_datos_estacion_individual, lista_tareas)
        for region, comuna, sector, serie in resultados:
            if not serie.empty:
                if region not in resultados_completos: resultados_completos[region] = {}
                if comuna not in resultados_completos[region]: resultados_completos[region][comuna] = {}
                resultados_completos[region][comuna][sector] = serie
    return resultados_completos

datos_totales = descargar_todos_los_datos(var_api)
ahora = pd.Timestamp.now(tz='America/Santiago').tz_localize(None)

# ==========================================
# 4. PREPARACIÓN DE DATOS MAPA
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
                
                estado = "Critico" if valor_actual > limite_actual else "Normal"
                color = "red" if valor_actual > limite_actual else "green"
                tamanio = 15 if valor_actual > limite_actual else 8
                
                datos_mapa.append({
                    "Region": region, "Comuna": comuna, "Estacion": sector,
                    "Latitud": DICCIONARIO_ZONAS[region][comuna][sector][0], "Longitud": DICCIONARIO_ZONAS[region][comuna][sector][1],
                    "Concentracion": round(valor_actual, 1), "Estado": estado, "Color": color, "Tamaño": tamanio
                })

df_mapa = pd.DataFrame(datos_mapa)

# ==========================================
# 5. INTERFAZ PRINCIPAL B2B
# ==========================================
if cliente_activo == "Demo Publica":
    st.title("Datair | Inteligencia Ambiental")
else:
    st.title(f"Datair | Portal Corporativo: {cliente_activo}")

st.markdown(f"**Monitoreo Nacional de {contaminante_elegido}** | Limite Normativo: `{limite_actual} µg/m³`")

promedio_nacional = df_mapa["Concentracion"].mean() if not df_mapa.empty else 0
estaciones_criticas = len(df_mapa[df_mapa["Estado"] == "Critico"]) if not df_mapa.empty else 0
estado_pais = "Alerta Nacional" if estaciones_criticas > (len(df_mapa) * 0.2) else "Estable"

col1, col2, col3 = st.columns(3)
with col1: st.metric(label="Promedio Pais Actual", value=f"{promedio_nacional:.1f} µg/m³")
with col2: st.metric(label="Estado General", value=estado_pais)
with col3: st.metric(label="Estaciones Operativas", value=f"{len(df_mapa)} / {TOTAL_ESTACIONES}")

st.divider()

# --- SECCIÓN 1: MAPA NACIONAL ---
st.subheader("Monitoreo Geografico en Tiempo Real")
if not df_mapa.empty:
    fig_mapa = px.scatter_mapbox(
        df_mapa, lat="Latitud", lon="Longitud", hover_name="Estacion", 
        hover_data={"Latitud": False, "Longitud": False, "Region": True, "Comuna": True, "Concentracion": True, "Estado": True, "Tamaño": False, "Color": False},
        color="Estado", color_discrete_map={"Critico": "red", "Normal": "green"},
        size="Tamaño", zoom=4, center={"lat": -35.0, "lon": -71.0}, height=550
    )
    fig_mapa.update_layout(mapbox_style="carto-darkmatter", margin={"r":0,"t":0,"l":0,"b":0})
    st.plotly_chart(fig_mapa, use_container_width=True)

st.divider()

# --- SECCIÓN 2: ANÁLISIS REGIONAL Y COMUNAL ---
st.subheader("Analisis Zonal y Pronostico")

col_filtro1, col_filtro2 = st.columns(2)
with col_filtro1: region_elegida = st.selectbox("Selecciona la Region", list(DICCIONARIO_ZONAS.keys()), index=3)
with col_filtro2: comuna_elegida = st.selectbox("Selecciona la Comuna", list(DICCIONARIO_ZONAS[region_elegida].keys()))

datos_sectores_comuna = {}
lista_promedios_comunas = {}

if region_elegida in datos_totales:
    for comuna, sectores in datos_totales[region_elegida].items():
        series_comuna = []
        for sector, serie in sectores.items():
            series_comuna.append(serie)
            if comuna == comuna_elegida:
                datos_sectores_comuna[sector] = serie
        if series_comuna:
            lista_promedios_comunas[comuna] = pd.concat(series_comuna, axis=1).mean(axis=1)

df_region = pd.DataFrame(lista_promedios_comunas).reset_index().rename(columns={'index': 'Fecha y Hora'})
df_comuna_completo = pd.DataFrame(datos_sectores_comuna).reset_index().rename(columns={'index': 'Fecha y Hora'})

st.markdown(f"**Comparativa Regional ({region_elegida}) - Incluye Pronostico**")
fig_region = px.line(df_region, x='Fecha y Hora', y=df_region.columns[1:], labels={'value': 'Concentracion (µg/m³)', 'variable': 'Comuna'})
fig_region.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Limite Legal")
fig_region.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white", annotation_text="AHORA")
fig_region.add_vrect(x0=ahora, x1=df_region['Fecha y Hora'].max(), fillcolor="blue", opacity=0.1, layer="below", line_width=0)
st.plotly_chart(fig_region, use_container_width=True)

st.markdown(f"**Detalle Estaciones ({comuna_elegida}) - Incluye Pronostico**")
fig_comuna = px.line(df_comuna_completo, x='Fecha y Hora', y=df_comuna_completo.columns[1:], labels={'value': 'Concentracion (µg/m³)', 'variable': 'Estacion'})
fig_comuna.add_hline(y=limite_actual, line_dash="dot", line_color="red", annotation_text="Limite Legal")
fig_comuna.add_vline(x=ahora, line_width=2, line_dash="dash", line_color="white", annotation_text="AHORA")
fig_comuna.add_vrect(x0=ahora, x1=df_comuna_completo['Fecha y Hora'].max(), fillcolor="blue", opacity=0.1, layer="below", line_width=0)
st.plotly_chart(fig_comuna, use_container_width=True)

st.divider()

# --- SECCIÓN 3: EXPORTACIÓN ---
st.subheader("Generacion de Reportes de Cumplimiento")
st.write(f"Descarga informes gerenciales auditables para **{cliente_activo}**.")

df_region_historico = df_region[df_region['Fecha y Hora'] <= ahora]
df_comuna_historico = df_comuna_completo[df_comuna_completo['Fecha y Hora'] <= ahora]

def generar_excel_universal(df_datos, contaminante, limite, nombre_zona, tipo_zona, cliente):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_datos.to_excel(writer, sheet_name='Base_Datos', index=False, startrow=3)
        ws_resumen = writer.book.create_sheet('Dashboard_Ejecutivo', 0)
        ws_datos = writer.sheets['Base_Datos']
        
        ws_datos['A1'] = f"REGISTRO CONTINUO - {contaminante} ({nombre_zona}) - {cliente}"
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
        ws_datos.auto_filter.ref = f"A4:{ultima_letra}{len(df_datos)+4}"
        ws_datos.freeze_panes = 'A5'
        
        for row in range(1, 30):
            for col in range(1, 15):
                ws_resumen.cell(row=row, column=col).fill = PatternFill(start_color="F2F2F2", fill_type="solid")
        
        ws_resumen['B2'] = f"REPORTE DE AUDITORIA - {tipo_zona.upper()}"
        ws_resumen['B2'].font = Font(size=18, bold=True, color="FFFFFF")
        ws_resumen['B2'].fill = PatternFill(start_color="1F4E78", fill_type="solid")
        ws_resumen.merge_cells('B2:H3')
        ws_resumen['B2'].alignment = Alignment(horizontal="center", vertical="center")
        
        ws_resumen['B4'] = f"Cliente:"
        ws_resumen['C4'] = cliente
        ws_resumen['B5'], ws_resumen['C5'] = f"{tipo_zona} Evaluada:", nombre_zona
        ws_resumen['B6'], ws_resumen['C6'] = "Parametro Evaluado:", contaminante
        ws_resumen['B7'], ws_resumen['C7'] = "Limite Legal:", f"{limite} µg/m³"
        ws_resumen['C7'].font = Font(bold=True, color="C00000")
        
        kpi_headers = ['Sub-zona', 'Promedio', 'Peak Maximo', 'Horas Infraccion', 'Estado']
        for i, header in enumerate(kpi_headers, start=2):
            cell = ws_resumen.cell(row=10, column=i)
            cell.value, cell.font, cell.fill, cell.alignment = header, header_font, header_fill, Alignment(horizontal="center")
            ws_resumen.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 20
            
        for idx, subzona in enumerate(columnas_datos):
            row = 11 + idx
            horas_infraccion = (df_datos[subzona] > limite).sum()
            ws_resumen.cell(row=row, column=2).value = subzona
            ws_resumen.cell(row=row, column=3).value = round(df_datos[subzona].mean(), 2)
            ws_resumen.cell(row=row, column=4).value = round(df_datos[subzona].max(), 2)
            ws_resumen.cell(row=row, column=5).value = horas_infraccion
            estado_cell = ws_resumen.cell(row=row, column=6)
            if horas_infraccion > 0:
                estado_cell.value, estado_cell.font, estado_cell.fill = "CRITICO", red_font, red_fill
            else:
                estado_cell.value, estado_cell.font, estado_cell.fill = "CUMPLE", green_font, green_fill
                
            for col in range(2, 7):
                ws_resumen.cell(row=row, column=col).alignment = Alignment(horizontal="center")
                ws_resumen.cell(row=row, column=col).border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'), top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))
                
        chart = LineChart()
        chart.title = f"Monitoreo Continuo - {contaminante}"
        chart.style, chart.width, chart.height = 13, 25, 13
        data = Reference(ws_datos, min_col=2, min_row=4, max_col=len(df_datos.columns), max_row=len(df_datos)+4)
        cats = Reference(ws_datos, min_col=1, min_row=5, max_row=len(df_datos)+4)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws_resumen.add_chart(chart, "B16")
        
    return output.getvalue()

col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    excel_region = generar_excel_universal(df_region_historico, contaminante_elegido, limite_actual, region_elegida, "Region", cliente_activo)
    st.download_button(
        label=f"Descargar Auditoria Regional ({region_elegida})",
        data=excel_region,
        file_name=f"Auditoria_{region_elegida}_{contaminante_elegido}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with col_btn2:
    excel_comuna = generar_excel_universal(df_comuna_historico, contaminante_elegido, limite_actual, comuna_elegida, "Comuna", cliente_activo)
    st.download_button(
        label=f"Descargar Auditoria Comunal ({comuna_elegida})",
        data=excel_comuna,
        file_name=f"Auditoria_{comuna_elegida}_{contaminante_elegido}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )