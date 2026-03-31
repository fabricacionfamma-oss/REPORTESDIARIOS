import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import tempfile
import os
import calendar
from fpdf import FPDF
from datetime import timedelta

# ==========================================
# 0. DICCIONARIO DE MÁQUINAS Y GRUPOS FAMMA
# ==========================================
MAQUINAS_MAP = {
    # ESTAMPADO
    "LINEA 1.5": "LÍNEAS ESTAMPADO", 
    "LINEA 2": "LÍNEAS ESTAMPADO", 
    "LINEA 3": "LÍNEAS ESTAMPADO", 
    "LINEA 4": "LÍNEAS ESTAMPADO",
    
    # SOLDADURA
    "CELL 13 FAMMA": "CELDAS SOLDADURA",
    "CELL 14 FAMMA": "CELDAS SOLDADURA",
    "CELL 15A FAMMA": "CELDAS SOLDADURA",
    "CELL 15B FAMMA": "CELDAS SOLDADURA",
    "CELL 16 FAMMA": "CELDAS SOLDADURA",
    "CELL 17 FAMMA": "CELDAS SOLDADURA",
    "CELL 3 FAMMA": "CELDAS SOLDADURA",
    "PRP 1": "EQUIPOS PRP",
    "PRP 2": "EQUIPOS PRP",
    "PRP 3": "EQUIPOS PRP",
    
    "GENERAL": "LÍNEAS ESTAMPADO" # Fallback por si viene vacío en Estampado
}

GRUPOS_ESTAMPADO = ['LÍNEAS ESTAMPADO']
GRUPOS_SOLDADURA = ['CELDAS SOLDADURA', 'EQUIPOS PRP']

# ==========================================
# 1. CONFIGURACIÓN Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Generador de Reportes PDF - FAMMA", 
    layout="wide", 
    page_icon="📄"
)

st.markdown("""
<style>
    hr { margin-top: 1.5rem; margin-bottom: 1.5rem; }
    .stButton>button { height: 3rem; font-size: 16px; font-weight: bold; }
    .header-style { font-size: 26px; font-weight: bold; margin-bottom: 5px; color: #1F2937; }
</style>
""", unsafe_allow_html=True)

col_title, col_btn = st.columns([4, 1])
with col_title:
    st.markdown('<div class="header-style">📄 Reportes PDF - FAMMA</div>', unsafe_allow_html=True)
    st.write("Seleccione los parámetros para generar y descargar los reportes consolidados.")
with col_btn:
    st.write("") 
    if st.button("Limpiar Caché / Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# 2. CARGA DE DATOS ROBUSTA (GSHEETS - 8 HOJAS)
# ==========================================
@st.cache_data(ttl=300)
def load_data():
    try:
        try:
            url_base = st.secrets["connections"]["gsheets"]["spreadsheet"].strip()
        except Exception:
            st.error("Atención: No se encontró la configuración de secretos (.streamlit/secrets.toml).")
            return [pd.DataFrame()] * 8

        gid_datos = "0"
        gid_oee_diario = "1767654796"
        gid_prod = "315437448"
        gid_op_diario = "354131379"
        gid_oee_sem = "2079886194"
        gid_oee_men = "1696631148"
        gid_op_sem = "2038636509"
        gid_op_men = "1171574188"
        
        base_export = url_base.split("/edit")[0] + "/export?format=csv&gid="
        
        def process_df(url, is_daily=False):
            try:
                df = pd.read_csv(url)
            except Exception: return pd.DataFrame()
            
            # Normalización de la columna Máquina/Línea
            col_maq = next((c for c in df.columns if c.lower() in ['máquina', 'maquina', 'línea', 'linea', 'celda', 'equipo']), None)
            if col_maq and col_maq != 'Máquina':
                df.rename(columns={col_maq: 'Máquina'}, inplace=True)
            if 'Máquina' not in df.columns and len(df.columns) > 0:
                df['Máquina'] = 'General'
            
            cols_num = ['Tiempo (Min)', 'Buenas', 'Retrabajo', 'Observadas', 'OEE', 'Disponibilidad', 'Performance', 'Calidad', 'Eficiencia']
            for c in cols_num:
                matches = [col for col in df.columns if c.lower() in col.lower()]
                for match in matches:
                    df[match] = df[match].astype(str).str.replace(',', '.')
                    df[match] = df[match].str.replace('%', '')
                    df[match] = pd.to_numeric(df[match], errors='coerce').fillna(0.0)
            
            col_fecha = next((c for c in df.columns if 'fecha' in c.lower() and 'inicio' not in c.lower() and 'fin' not in c.lower()), None)
            if col_fecha:
                df['Fecha_DT'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
                df['Fecha_Filtro'] = df['Fecha_DT'].dt.normalize()
                if is_daily:
                    df = df.dropna(subset=['Fecha_Filtro'])
            
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].fillna('').astype(str).str.strip()
                    
            if 'Nivel Evento 3' in df.columns:
                def categorizar_estado(row):
                    texto = str(row.get('Nivel Evento 3', '')).upper() + " " + str(row.get('Nivel Evento 4', '')).upper() + " " + str(row.get('Evento', '')).upper()
                    if 'PRODUCCION' in texto or 'PRODUCCIÓN' in texto: return 'Producción'
                    if 'PROYECTO' in texto: return 'Proyecto'
                    if 'BAÑO' in texto or 'BANO' in texto or 'REFRIGERIO' in texto: return 'Descanso'
                    if 'PARADA PROGRAMADA' in texto: return 'Parada Programada'
                    return 'Falla/Gestión'
                
                df['Estado_Global'] = df.apply(categorizar_estado, axis=1)
                
                def obtener_detalle(row):
                    n3 = str(row.get('Nivel Evento 3', '')).strip()
                    n4 = str(row.get('Nivel Evento 4', '')).strip()
                    n5 = str(row.get('Nivel Evento 5', '')).strip()
                    n6 = str(row.get('Nivel Evento 6', '')).strip()
                    validos = [n for n in [n3, n4, n5, n6] if n and n.lower() not in ['none', 'nan', '']]
                    return validos[-1] if validos else str(row.get('Evento', 'Sin detalle'))
                    
                df['Detalle_Final'] = df.apply(obtener_detalle, axis=1)

                # CORRECCIÓN DE "FALLA ABIERTA" AL GRÁFICO DE TORTA
                def clasificar_macro(row):
                    n3 = str(row.get('Nivel Evento 3', '')).strip().upper()
                    n4 = str(row.get('Nivel Evento 4', '')).strip().upper()
                    n5 = str(row.get('Nivel Evento 5', '')).strip().upper()
                    
                    if 'GESTION' in n3 or 'GESTIÓN' in n3: return 'Gestión'
                    
                    if 'FALLA' in n3:
                        # Si en Nivel 4 dice "FALLA ABIERTA", saltamos a Nivel 5
                        if 'ABIERT' in n4 or n4 == 'FALLA' or n4 in ['NAN', 'NONE', '']:
                            if n5 not in ['NAN', 'NONE', '']:
                                return n5.title()
                        if n4 not in ['NAN', 'NONE', '']:
                            return n4.title()
                        return 'Falla General'
                        
                    return n3.title() if n3 not in ['NAN', 'NONE', ''] else 'Otros'
                
                df['Categoria_Macro'] = df.apply(clasificar_macro, axis=1)

            col_inicio = next((c for c in df.columns if 'inicio' in c.lower() or 'desde' in c.lower()), None)
            col_fin = next((c for c in df.columns if 'fin' in c.lower() or 'hasta' in c.lower()), None)
            if col_inicio: df['Inicio_Str'] = df[col_inicio].astype(str).str[:5]
            if col_fin: df['Fin_Str'] = df[col_fin].astype(str).str[:5]

            return df

        return (
            process_df(base_export + gid_datos, is_daily=True), 
            process_df(base_export + gid_oee_diario, is_daily=True), 
            process_df(base_export + gid_prod, is_daily=True), 
            process_df(base_export + gid_op_diario, is_daily=True),
            process_df(base_export + gid_oee_sem, is_daily=False),
            process_df(base_export + gid_oee_men, is_daily=False),
            process_df(base_export + gid_op_sem, is_daily=False),
            process_df(base_export + gid_op_men, is_daily=False)
        )
    except Exception as e:
        st.error(f"Error cargando datos: {e}")
        return [pd.DataFrame()] * 8

df_raw, df_oee_diario, df_prod_raw, df_op_diario_raw, df_oee_sem, df_oee_men, df_op_sem_raw, df_op_men_raw = load_data()

if df_raw.empty:
    st.warning("No hay datos cargados en la base principal.")
    st.stop()

# ==========================================
# 3. INTERFAZ: CONFIGURACIÓN PDF
# ==========================================
col_p1, col_p2, col_p3 = st.columns([1, 1.2, 1.5])

with col_p1:
    st.write("**1. Tipo de Reporte:**")
    pdf_tipo = st.radio("Período:", ["Diario", "Semanal", "Mensual"], horizontal=True, label_visibility="collapsed")

pdf_ini, pdf_fin = None, None
pdf_df_oee_target = pd.DataFrame()
pdf_df_op_target = pd.DataFrame()
pdf_label, file_label = "", ""

with col_p2:
    st.write("**2. Seleccione el Período:**")
    if pdf_tipo == "Diario":
        min_d = df_raw['Fecha_Filtro'].min().date() if not df_raw.empty else pd.to_datetime("today").date()
        max_d = df_raw['Fecha_Filtro'].max().date() if not df_raw.empty else pd.to_datetime("today").date()
        pdf_fecha = st.date_input("Día para PDF:", value=max_d, min_value=min_d, max_value=max_d, label_visibility="collapsed")
        
        pdf_ini, pdf_fin = pd.to_datetime(pdf_fecha), pd.to_datetime(pdf_fecha)
        pdf_df_oee_target = df_oee_diario[df_oee_diario['Fecha_Filtro'] == pdf_ini]
        pdf_df_op_target = df_op_diario_raw[df_op_diario_raw['Fecha_Filtro'] == pdf_ini]
        pdf_label = f"Día {pdf_fecha.strftime('%d-%m-%Y')}"
        file_label = pdf_label
        
    elif pdf_tipo == "Semanal":
        if not df_oee_sem.empty:
            col_sem = df_oee_sem.columns[0]
            opciones_sem = [s for s in df_oee_sem[col_sem].unique() if str(s).strip() != "" and str(s).lower() != "nan"]
            pdf_sem = st.selectbox("Semana para PDF:", opciones_sem, label_visibility="collapsed")
            
            pdf_df_oee_target = df_oee_sem[df_oee_sem[col_sem].astype(str) == str(pdf_sem)]
            col_sem_op = df_op_sem_raw.columns[0] if not df_op_sem_raw.empty else None
            if col_sem_op:
                pdf_df_op_target = df_op_sem_raw[df_op_sem_raw[col_sem_op].astype(str) == str(pdf_sem)]
            pdf_label = f"Semana {pdf_sem}"
            file_label = f"Semana_{pdf_sem}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
        else:
            st.warning("No hay datos semanales.")
                
    elif pdf_tipo == "Mensual":
        if not df_oee_men.empty:
            col_mes = df_oee_men.columns[0]
            opciones_mes = [m for m in df_oee_men[col_mes].unique() if str(m).strip() != "" and str(m).lower() != "nan"]
            pdf_mes = st.selectbox("Mes para PDF:", opciones_mes, label_visibility="collapsed")
            
            pdf_df_oee_target = df_oee_men[df_oee_men[col_mes].astype(str) == str(pdf_mes)]
            col_mes_op = df_op_men_raw.columns[0] if not df_op_men_raw.empty else None
            if col_mes_op:
                pdf_df_op_target = df_op_men_raw[df_op_men_raw[col_mes_op].astype(str) == str(pdf_mes)]
            pdf_label = f"Mes {pdf_mes}"
            file_label = f"Mes_{pdf_mes}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
        else:
            st.warning("No hay datos mensuales.")

# ==========================================
# 4. FUNCIONES HELPER PARA TIEMPO Y PDF
# ==========================================
def parse_time_to_mins(t_str):
    try:
        if pd.isna(t_str) or t_str in ['nan', 'None', '', '-']: return None
        parts = str(t_str).split(':'); return int(parts[0]) * 60 + int(parts[1])
    except: return None

def mins_to_time_str(m):
    if pd.isna(m) or m is None: return "-"
    m = int(m) % 1440; return f"{m//60:02d}:{m%60:02d}"

def mins_to_duration_str(m):
    if pd.isna(m) or m is None: return "00:00 hs"
    m = int(m); return f"{m//60:02d}:{m%60:02d} hs"

class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area; self.fecha_str = fecha_str; self.theme_color = theme_color

    def header(self):
        if os.path.exists("logo.png"): self.image("logo.png", 10, 8, 30)
        self.set_font("Times", 'B', 16); self.set_text_color(*self.theme_color)
        self.cell(0, 10, clean_text(f"REPORTE GERENCIAL - {self.area.upper()}"), ln=True, align='R')
        self.set_font("Arial", 'B', 10); self.set_text_color(100, 100, 100)
        self.cell(0, 6, clean_text(f"Periodo: {self.fecha_str}"), ln=True, align='R'); self.ln(5)

    def footer(self):
        self.set_y(-15); self.set_font("Arial", "I", 8); self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Pagina {self.page_no()}", 0, 0, "C")

def clean_text(text):
    if pd.isna(text): return "-"
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

def check_space(pdf, required_height):
    if pdf.get_y() + required_height > 275 and pdf.get_y() > 40:
        pdf.add_page(); return True
    return False

# CORRECCIÓN DE ESPACIADO VERTICAL EN SECCIONES
def print_section_title(pdf, title, theme_color):
    pdf.ln(5); pdf.set_font("Times", 'B', 14); pdf.set_text_color(*theme_color)
    pdf.cell(0, 6, clean_text(title), ln=True)
    x, y = pdf.get_x(), pdf.get_y()
    pdf.set_draw_color(*theme_color); pdf.set_line_width(0.5); pdf.line(x, y, x + 190, y)
    pdf.set_draw_color(0, 0, 0); pdf.set_line_width(0.2); pdf.set_text_color(0, 0, 0); pdf.ln(5)

def setup_table_header(pdf, theme_color):
    pdf.set_fill_color(*theme_color); pdf.set_text_color(255, 255, 255); pdf.set_draw_color(*theme_color)

def setup_table_row(pdf):
    pdf.set_fill_color(255, 255, 255); pdf.set_text_color(50, 50, 50); pdf.set_draw_color(200, 200, 200)

def set_pdf_color(pdf, val):
    if val < 0.85: pdf.set_text_color(220, 20, 20)
    elif val <= 0.95: pdf.set_text_color(200, 150, 0)
    else: pdf.set_text_color(33, 195, 84)

def get_metrics_direct(name_filter, target_df):
    m = {'OEE': 0.0, 'DISP': 0.0, 'PERF': 0.0, 'CAL': 0.0}
    if target_df.empty: return m
    mask = target_df.apply(lambda row: row.astype(str).str.upper().str.contains(name_filter.upper()), axis=1)
    datos = target_df[mask.any(axis=1)]
    if not datos.empty:
        fila = datos.iloc[0] 
        for key, col_search in {'OEE':['OEE'], 'DISP':['DISPONIBILIDAD', 'DISP'], 'PERF':['PERFORMANCE', 'PERFO'], 'CAL':['CALIDAD', 'CAL']}.items():
            actual_col = next((c for c in datos.columns if any(x in c.upper() for x in col_search)), None)
            if actual_col:
                val_str = str(fila[actual_col]).replace('%', '').replace(',', '.').strip()
                v = pd.to_numeric(val_str, errors='coerce')
                if pd.notna(v): m[key] = float(v/100 if v > 1.1 else v)
    return m

def print_pdf_metric_row(pdf, prefix, m):
    pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0, 0, 0)
    pdf.write(7, clean_text(f"{prefix} | OEE: "))
    set_pdf_color(pdf, m['OEE']); pdf.write(7, f"{m['OEE']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(7, clean_text("  |  Disp: "))
    set_pdf_color(pdf, m['DISP']); pdf.write(7, f"{m['DISP']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(7, clean_text("  |  Perf: "))
    set_pdf_color(pdf, m['PERF']); pdf.write(7, f"{m['PERF']:.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(7, clean_text("  |  Cal: "))
    set_pdf_color(pdf, m['CAL']); pdf.write(7, f"{m['CAL']:.1%}")
    pdf.set_text_color(0, 0, 0); pdf.ln(7)

def add_image_safe(pdf, img_path, w_mm, h_mm, center=True):
    if pdf.get_y() + h_mm > 275:
        pdf.add_page()
    x = (210 - w_mm) / 2 if center else pdf.get_x()
    y = pdf.get_y()
    pdf.image(img_path, x=x, y=y, w=w_mm)
    pdf.set_y(y + h_mm + 5)


# ==========================================
# 5. MOTOR GENERADOR DEL PDF 
# ==========================================
def crear_pdf(area, label_reporte, oee_target_df, op_target_df, ini_date, fin_date, p_tipo):
    if area.upper() == "ESTAMPADO":
        theme_color = (15, 76, 129); comp_color = (52, 152, 219)  
        chart_bars = ['#003366', '#3498DB', '#AED6F1']; pie_colors = px.colors.sequential.Blues_r
        grupos_area = GRUPOS_ESTAMPADO
    else:
        theme_color = (211, 84, 0); comp_color = (230, 126, 34) 
        chart_bars = ['#993300', '#E67E22', '#FAD7A1']; pie_colors = px.colors.sequential.Oranges_r
        grupos_area = GRUPOS_SOLDADURA
        
    hex_theme = '#%02x%02x%02x' % theme_color; hex_comp = '#%02x%02x%02x' % comp_color  
    
    # 1. Filtrado General
    if ini_date is not None and fin_date is not None:
        df_pdf_raw = df_raw[(df_raw['Fecha_Filtro'] >= ini_date) & (df_raw['Fecha_Filtro'] <= fin_date)]
        df_prod_pdf_raw = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= ini_date) & (df_prod_raw['Fecha_Filtro'] <= fin_date)] if not df_prod_raw.empty else pd.DataFrame()
    else:
        df_pdf_raw = pd.DataFrame(columns=df_raw.columns)
        df_prod_pdf_raw = pd.DataFrame(columns=df_prod_raw.columns)

    df_pdf = df_pdf_raw[df_pdf_raw['Fábrica'].str.contains(area, case=False, na=False)].copy()
    
    # 2. Tolerancia a falta de información de línea (Lógica inteligente para FAMMA)
    if 'Máquina' in df_pdf.columns:
        df_pdf['Máquina'] = df_pdf['Máquina'].replace({'': 'General', 'nan': 'General', 'None': 'General', 'NaN': 'General'})
    
    mapa_limpio = {str(k).strip().upper(): v for k, v in MAQUINAS_MAP.items()}
    
    def asignar_grupo(maq, area_name):
        maq_u = str(maq).strip().upper()
        if maq_u in mapa_limpio: return mapa_limpio[maq_u]
        
        # Lógica de rescate por palabras clave
        if 'LINEA' in maq_u: return 'LÍNEAS ESTAMPADO'
        if 'CELL' in maq_u or 'CELDA' in maq_u: return 'CELDAS SOLDADURA'
        if 'PRP' in maq_u: return 'EQUIPOS PRP'
        
        # Si no coincide nada, se asigna al grupo general del área
        return GRUPOS_ESTAMPADO[0] if area_name.upper() == "ESTAMPADO" else GRUPOS_SOLDADURA[0]

    if not df_pdf.empty:
        df_pdf['Grupo_Máquina'] = df_pdf['Máquina'].apply(lambda x: asignar_grupo(x, area))
    
    df_prod_pdf = pd.DataFrame()
    if not df_prod_pdf_raw.empty:
        df_prod_pdf = df_prod_pdf_raw[(df_prod_pdf_raw['Máquina'].str.contains(area, case=False, na=False)) | 
                                      (df_prod_pdf_raw['Máquina'].isin(df_pdf['Máquina'].unique()))].copy()
        
        if 'Máquina' in df_prod_pdf.columns:
            df_prod_pdf['Máquina'] = df_prod_pdf['Máquina'].replace({'': 'General', 'nan': 'General', 'None': 'General', 'NaN': 'General'})
            df_prod_pdf['Grupo_Máquina'] = df_prod_pdf['Máquina'].apply(lambda x: asignar_grupo(x, area))

    # Iniciar PDF e Índices
    pdf = ReportePDF(area, label_reporte, theme_color)
    pdf.set_auto_page_break(auto=True, margin=15); pdf.add_page()
    
    links_grupos = {g: pdf.add_link() for g in grupos_area}
    link_perfo = pdf.add_link(); link_tiempos = pdf.add_link()

    pdf.ln(10); pdf.set_font("Times", 'B', 18); pdf.set_text_color(*theme_color)
    pdf.cell(0, 10, clean_text("ÍNDICE DEL REPORTE"), ln=True, align='C')
    
    pdf.ln(10); pdf.set_font("Arial", 'U', 12); pdf.set_text_color(*comp_color)
    for g in grupos_area:
        pdf.cell(0, 8, clean_text(f"> Reporte detallado de Grupo: {g}"), ln=True, link=links_grupos[g])
    pdf.ln(5)
    pdf.cell(0, 8, clean_text("> Performance General de Operarios"), ln=True, link=link_perfo)
    pdf.cell(0, 8, clean_text("> Tablas de Tiempos Acumulados de Baño y Refrigerio"), ln=True, link=link_tiempos)

    if df_pdf.empty and oee_target_df.empty and df_prod_pdf.empty:
        pdf.add_page(); pdf.set_font("Arial", 'I', 12); pdf.set_text_color(100)
        pdf.cell(0, 10, f"No hay ningún tipo de dato registrado para el área {area} en este periodo.", ln=True)
        return pdf.output(dest='S').encode('latin-1')

    def dibujar_tabla_eventos_detallada(df_subset, col_detalle, titulo, color_t):
        if not df_subset.empty:
            check_space(pdf, 30); pdf.set_font("Arial", 'B', 9); pdf.set_text_color(*color_t)
            pdf.cell(0, 6, clean_text(f">> {titulo}:"), ln=True); pdf.ln(1)
            def dibujar_cabeceras():
                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                pdf.cell(15, 6, "Fecha", 1, 0, 'C', True)
                pdf.cell(12, 6, "Ini.", 1, 0, 'C', True)
                pdf.cell(12, 6, "Fin", 1, 0, 'C', True)
                pdf.cell(90, 6, "Detalle Registrado en Sistema", 1, 0, 'L', True)
                pdf.cell(12, 6, "Min", 1, 0, 'C', True)
                pdf.cell(49, 6, "Operador", 1, 1, 'L', True)
            
            dibujar_cabeceras(); setup_table_row(pdf); pdf.set_font("Arial", '', 8)
            df_subset['_sort_time'] = df_subset['Inicio_Str'].apply(lambda x: parse_time_to_mins(x) if pd.notna(x) else 9999)
            df_subset = df_subset.sort_values(['Fecha_Filtro', '_sort_time'], ascending=[True, True])
            for _, row in df_subset.iterrows():
                # CORRECCIÓN DE ESPACIOS: Altura aumentada a 6 y salto chequeado en 260
                if pdf.get_y() > 260:
                    pdf.add_page(); dibujar_cabeceras(); setup_table_row(pdf); pdf.set_font("Arial", '', 8)
                val_fecha = pd.to_datetime(row['Fecha_Filtro']).strftime('%d/%m') if pd.notna(row['Fecha_Filtro']) else "-"
                val_inicio = str(row['Inicio_Str'])[:5] if pd.notna(row['Inicio_Str']) and str(row['Inicio_Str']) != 'nan' else "-"
                val_fin = str(row['Fin_Str'])[:5] if pd.notna(row['Fin_Str']) and str(row['Fin_Str']) != 'nan' else "-"
                minutos = f"{row['Tiempo (Min)']:.0f}"
                operador = " " + str(row.get('Operador', '-'))[:35]
                detalle_str = " " + str(row[col_detalle]) if col_detalle in row and pd.notna(row[col_detalle]) else " Sin detalle"
                
                pdf.cell(15, 6, val_fecha, 'B', 0, 'C')
                pdf.cell(12, 6, val_inicio, 'B', 0, 'C')
                pdf.cell(12, 6, val_fin, 'B', 0, 'C')
                pdf.cell(90, 6, clean_text(detalle_str[:60]), 'B', 0, 'L')
                pdf.cell(12, 6, minutos, 'B', 0, 'C')
                pdf.cell(49, 6, clean_text(operador), 'B', 1, 'L')
            pdf.ln(5)

    # =========================================================
    # RECORRIDO POR CADA GRUPO 
    # =========================================================
    for g in grupos_area:
        df_pdf_g = df_pdf[df_pdf['Grupo_Máquina'] == g] if not df_pdf.empty else pd.DataFrame()
        df_prod_pdf_g = df_prod_pdf[df_prod_pdf['Grupo_Máquina'] == g] if not df_prod_pdf.empty else pd.DataFrame()
        
        maq_presentes = list(df_pdf_g['Máquina'].unique()) if not df_pdf_g.empty else []
        if not df_prod_pdf_g.empty:
            maq_presentes = list(set(maq_presentes + list(df_prod_pdf_g['Máquina'].unique())))
            
        m_general = get_metrics_direct(area, oee_target_df)
        
        maq_del_grupo_estaticas = [m for m, grp in MAQUINAS_MAP.items() if grp == g]
        todas_las_maqs = list(set(maq_presentes + maq_del_grupo_estaticas))
        
        maq_keys = []
        for maq in todas_las_maqs:
            m_maq = get_metrics_direct(maq, oee_target_df)
            if m_maq['OEE'] > 0:
                if (maq, m_maq) not in maq_keys:
                    maq_keys.append((maq, m_maq))
                
        if df_pdf_g.empty and df_prod_pdf_g.empty and m_general['OEE'] == 0 and not maq_keys:
            continue 
            
        pdf.add_page(); pdf.set_link(links_grupos[g]) 
        pdf.set_font("Times", 'B', 16); pdf.set_text_color(*theme_color)
        pdf.cell(0, 10, clean_text(f"SECCIÓN GRUPO: {g}"), ln=True, align='L', border='B'); pdf.ln(5)

        # ---------------------------------
        # 1. RESUMEN OEE 
        # ---------------------------------
        check_space(pdf, 35); print_section_title(pdf, "1. Resumen OEE del Grupo", theme_color)
        
        if m_general['OEE'] > 0:
            print_pdf_metric_row(pdf, f"General {area}", m_general)
            
        for maq, m_maq in maq_keys:
            print_pdf_metric_row(pdf, f"    > {maq}", m_maq)
        pdf.ln(3)

        # ---------------------------------
        # 2. HORARIOS O GRÁFICOS OEE
        # ---------------------------------
        if p_tipo == "Mensual":
            check_space(pdf, 25); print_section_title(pdf, "2. Análisis Visual de OEE por Máquina", theme_color)
            if not maq_keys and m_general['OEE'] > 0:
                maq_keys = [("General " + area, m_general)]
                
            if not maq_keys:
                pdf.set_font("Arial", 'I', 9); pdf.cell(0, 6, clean_text("No hay datos de OEE para graficar."), ln=True)
            else:
                for i in range(0, len(maq_keys), 2):
                    maq1, m_data1 = maq_keys[i]
                    maq2, m_data2 = maq_keys[i+1] if i+1 < len(maq_keys) else (None, None)
                    if pdf.get_y() + 65 > 260: pdf.add_page()
                    y_base = pdf.get_y()
                    
                    def build_oee_fig(m_name, m_data):
                        x_labels = ['OEE', 'Disp.', 'Perf.', 'Calidad']
                        y_vals = [m_data['OEE']*100, m_data['DISP']*100, m_data['PERF']*100, m_data['CAL']*100]
                        f = go.Figure(data=[go.Bar(x=x_labels, y=y_vals, text=[f"{v:.1f}%" for v in y_vals], textposition='auto', marker_color=['#2C3E50', '#F1C40F', '#3498DB', '#2ECC71'])])
                        max_y = max(y_vals) if y_vals else 100
                        y_max = max_y * 1.2 if max_y > 0 else 100
                        if y_max < 100: y_max = 110
                        f.update_layout(title=dict(text=f"{m_name}", font=dict(size=14)), height=250, width=380, margin=dict(t=40, b=20, l=20, r=20), plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, y_max], visible=False))
                        return f

                    fig1 = build_oee_fig(maq1, m_data1)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp1:
                        fig1.write_image(tmp1.name, engine="kaleido"); pdf.image(tmp1.name, x=10, y=y_base, w=95); os.remove(tmp1.name)
                        
                    if maq2:
                        fig2 = build_oee_fig(maq2, m_data2)
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp2:
                            fig2.write_image(tmp2.name, engine="kaleido"); pdf.image(tmp2.name, x=105, y=y_base, w=95); os.remove(tmp2.name)
                            
                    pdf.set_y(y_base + 65); pdf.ln(2)
        else:
            check_space(pdf, 35); print_section_title(pdf, "2. Horarios y Tiempo de Apertura", theme_color)
            df_pdf_g_horarios = df_pdf_g.copy()
            if not df_pdf_g_horarios.empty and 'Inicio_Str' in df_pdf_g_horarios.columns:
                col_turno = next((c for c in df_pdf_g_horarios.columns if 'turno' in c.lower()), None)
                if not col_turno:
                    df_pdf_g_horarios['Turno'] = 'A'
                else:
                    df_pdf_g_horarios['Turno'] = df_pdf_g_horarios[col_turno]
                    
                if p_tipo == "Diario":
                    tiempos_list = []
                    for (maq, turno), grp in df_pdf_g_horarios.groupby(['Máquina', 'Turno']):
                        intervals = []
                        for _, r in grp.iterrows():
                            ini = parse_time_to_mins(r['Inicio_Str']); fin = parse_time_to_mins(r['Fin_Str'])
                            if ini is not None and fin is not None:
                                if fin < ini and (ini - fin) > 720: fin += 1440
                                intervals.append([ini, fin])
                        if not intervals: continue
                        intervals.sort(key=lambda x: x[0]); merged = [intervals[0]]
                        for current in intervals[1:]:
                            last = merged[-1]
                            if current[0] <= last[1]: last[1] = max(last[1], current[1])
                            else: merged.append(current)
                        total_active = sum(iv[1] - iv[0] for iv in merged)
                        min_i, max_f = merged[0][0], merged[-1][1]; tiempo_bruto = max_f - min_i
                        unregistered_time = max(0, tiempo_bruto - total_active)
                        tiempos_list.append({'Máquina': maq, 'Turno': turno, 'Inicio': min_i, 'Fin': max_f, 'Total': total_active, 'NoReg': unregistered_time})
                    
                    df_horarios = pd.DataFrame(tiempos_list)
                    if not df_horarios.empty:
                        df_horarios = df_horarios.sort_values(['Máquina', 'Turno'])
                        def dibujar_cabeza_hora():
                            setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                            pdf.cell(35, 6, "Maquina", 1, 0, 'C', True); pdf.cell(15, 6, "Turno", 1, 0, 'C', True)
                            pdf.cell(25, 6, "Hora Inicio", 1, 0, 'C', True); pdf.cell(25, 6, "Hora Cierre", 1, 0, 'C', True)
                            pdf.cell(45, 6, "Apertura Neta", 1, 0, 'C', True); pdf.cell(45, 6, "No Registrado", 1, 1, 'C', True); pdf.ln()
                        
                        dibujar_cabeza_hora()
                        setup_table_row(pdf); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 8)
                        for _, r in df_horarios.iterrows():
                            if pdf.get_y() > 260: 
                                pdf.add_page(); dibujar_cabeza_hora(); setup_table_row(pdf); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 8)
                            pdf.cell(35, 6, " " + clean_text(str(r['Máquina'])[:15]), 1, 0, 'L'); pdf.cell(15, 6, clean_text(str(r['Turno'])), 1, 0, 'C')
                            pdf.cell(25, 6, clean_text(mins_to_time_str(r['Inicio'])), 1, 0, 'C'); pdf.cell(25, 6, clean_text(mins_to_time_str(r['Fin'])), 1, 0, 'C')
                            pdf.cell(45, 6, clean_text(mins_to_duration_str(r['Total'])), 1, 0, 'C'); pdf.cell(45, 6, clean_text(mins_to_duration_str(r['NoReg'])), 1, 1, 'C')
                        pdf.ln(5)
                else:
                    df_pdf_g_horarios['Fecha_DT'] = pd.to_datetime(df_pdf_g_horarios['Fecha_Filtro'])
                    df_pdf_g_horarios['Dia_Semana'] = df_pdf_g_horarios['Fecha_DT'].dt.dayofweek
                    horarios_list = []
                    for (maq, turno, dia), grp in df_pdf_g_horarios.groupby(['Máquina', 'Turno', 'Dia_Semana']):
                        if dia > 4: continue 
                        ini = grp['Inicio_Str'].apply(parse_time_to_mins).min(); fin = grp['Fin_Str'].apply(parse_time_to_mins).max()
                        if pd.notna(ini) and pd.notna(fin):
                            horarios_list.append({'Máquina': maq, 'Turno': turno, 'Dia': dia, 'Rango': f"{mins_to_time_str(ini)} - {mins_to_time_str(fin)}"})
                    if horarios_list:
                        df_h = pd.DataFrame(horarios_list).pivot_table(index=['Máquina', 'Turno'], columns='Dia', values='Rango', aggfunc='first').reset_index()
                        for i in range(5):
                            if i not in df_h.columns: df_h[i] = "-"
                        df_h = df_h.fillna("-").sort_values(['Máquina', 'Turno'])
                        def dibujar_cabeza_semana():
                            setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                            pdf.cell(25, 6, "Maquina", 1, 0, 'C', True); pdf.cell(12, 6, "Turno", 1, 0, 'C', True)
                            for d in ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"]: pdf.cell(30, 6, d, 1, 0, 'C', True)
                            pdf.ln()

                        dibujar_cabeza_semana()
                        setup_table_row(pdf); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 7)
                        for _, row in df_h.iterrows():
                            if pdf.get_y() > 260: 
                                pdf.add_page(); dibujar_cabeza_semana(); setup_table_row(pdf); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 7)
                            pdf.cell(25, 6, clean_text(row['Máquina']), 1); pdf.cell(12, 6, clean_text(str(row.get('Turno', '-'))), 1, 0, 'C')
                            for i in range(5): pdf.cell(30, 6, str(row.get(i, "-")), 1, 0, 'C')
                            pdf.ln()
                        pdf.ln(5)
            else:
                pdf.set_font("Arial", 'I', 9); pdf.cell(0, 6, clean_text("No hay horarios registrados para generar las estadísticas."), ln=True)

        # ---------------------------------
        # 3. DESGLOSE POR MÁQUINA
        # ---------------------------------
        maquinas_con_tiempo = []
        if not df_pdf_g.empty:
            for maq in sorted(df_pdf_g['Máquina'].unique()):
                df_maq = df_pdf_g[df_pdf_g['Máquina'] == maq]
                t_total = df_maq[df_maq['Estado_Global'].isin(['Producción', 'Falla/Gestión', 'Parada Programada', 'Proyecto', 'Descanso'])]['Tiempo (Min)'].sum()
                if t_total > 0: maquinas_con_tiempo.append(maq)
        
        if maquinas_con_tiempo:
            check_space(pdf, 25); print_section_title(pdf, "3. Analisis de Tiempos y Fallas por Máquina", theme_color)
            impresas = 0
            for maq in maquinas_con_tiempo:
                df_maq = df_pdf_g[df_pdf_g['Máquina'] == maq]
                t_prod = df_maq[df_maq['Estado_Global'] == 'Producción']['Tiempo (Min)'].sum()
                t_falla = df_maq[df_maq['Estado_Global'] == 'Falla/Gestión']['Tiempo (Min)'].sum()
                t_parada = df_maq[df_maq['Estado_Global'] == 'Parada Programada']['Tiempo (Min)'].sum()
                t_proy = df_maq[df_maq['Estado_Global'] == 'Proyecto']['Tiempo (Min)'].sum()
                t_desc = df_maq[df_maq['Estado_Global'] == 'Descanso']['Tiempo (Min)'].sum()
                
                if impresas > 0:
                    salto_ejecutado = check_space(pdf, 35) 
                    if not salto_ejecutado:
                        pdf.ln(6); pdf.set_draw_color(200, 200, 200); pdf.set_line_width(0.8)
                        pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 190, pdf.get_y())
                        pdf.set_draw_color(0, 0, 0); pdf.set_line_width(0.2); pdf.ln(6)
                
                impresas += 1
                pdf.set_font("Arial", 'B', 12); pdf.set_text_color(255, 255, 255); pdf.set_fill_color(*comp_color)
                pdf.cell(0, 8, clean_text(f"  MÁQUINA: {maq}"), border=0, ln=True, fill=True)
                pdf.set_font("Arial", 'I', 8); pdf.set_text_color(120, 120, 120); pdf.cell(0, 5, clean_text(f"  Grupo: {g}"), border=0, ln=True); pdf.ln(2)
                
                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                for col_name in ["Produccion", "Fallas/Gestion", "Paradas Prog.", "Proyecto", "Descansos"]: pdf.cell(38, 6, col_name, border=1, align='C', fill=True)
                pdf.ln(); setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                pdf.cell(38, 6, clean_text(mins_to_duration_str(t_prod)), border=1, align='C')
                pdf.cell(38, 6, clean_text(mins_to_duration_str(t_falla)), border=1, align='C')
                pdf.cell(38, 6, clean_text(mins_to_duration_str(t_parada)), border=1, align='C')
                pdf.cell(38, 6, clean_text(mins_to_duration_str(t_proy)), border=1, align='C')
                pdf.cell(38, 6, clean_text(mins_to_duration_str(t_desc)), border=1, align='C', ln=True); pdf.ln(4)
                
                df_maq_fallas = df_maq[df_maq['Estado_Global'] == 'Falla/Gestión']
                
                if p_tipo == "Mensual":
                    if not df_maq_fallas.empty:
                        agg_f15 = df_maq_fallas.groupby('Detalle_Final')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(15)
                        agg_f15 = agg_f15.sort_values('Tiempo (Min)', ascending=True) 
                        agg_f15['Label'] = agg_f15.apply(lambda r: f" {str(r['Detalle_Final'])[:60]} — {r['Tiempo (Min)']:.0f}m", axis=1)
                        max_x_val = agg_f15['Tiempo (Min)'].max() if not agg_f15.empty else 1
                        
                        trend_df = df_maq_fallas.groupby('Fecha_Filtro')['Tiempo (Min)'].sum().reset_index().sort_values('Fecha_Filtro')
                        trend_df['Fecha_Str'] = pd.to_datetime(trend_df['Fecha_Filtro']).dt.strftime('%d/%m')
                        
                        if pdf.get_y() + 65 > 260: pdf.add_page()
                        
                        pdf.set_font("Arial", 'B', 10); pdf.set_text_color(*comp_color)
                        pdf.cell(95, 6, clean_text("> Top 15 Fallas (por tiempo):"), 0, 0, 'L')
                        pdf.cell(95, 6, clean_text("> Tendencia Diaria de Fallas (Minutos):"), 0, 1, 'L')
                        
                        y_base_graficos = pdf.get_y()
                        
                        fig_top15 = px.bar(agg_f15, x='Tiempo (Min)', y='Detalle_Final', orientation='h', text='Label')
                        fig_top15.update_traces(marker_color=hex_comp, textposition='outside', textfont=dict(size=11, color='black'), cliponaxis=False)
                        fig_top15.update_layout(height=280, width=450, margin=dict(t=5, b=5, l=10, r=220), plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(visible=False, range=[0, max_x_val * 1.5]), yaxis=dict(title='', showticklabels=False))
                        
                        fig_trend = px.line(trend_df, x='Fecha_Str', y='Tiempo (Min)', markers=True)
                        fig_trend.update_traces(line_color=hex_comp, marker=dict(size=8, color=hex_theme))
                        fig_trend.update_layout(height=250, width=400, margin=dict(t=10, b=30, l=40, r=20), plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Minutos")
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_chart:
                            fig_top15.write_image(tmp_chart.name, engine="kaleido")
                            pdf.image(tmp_chart.name, x=5, y=y_base_graficos, w=105)
                            os.remove(tmp_chart.name)
                            
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_trend:
                            fig_trend.write_image(tmp_trend.name, engine="kaleido")
                            pdf.image(tmp_trend.name, x=110, y=y_base_graficos, w=90)
                            os.remove(tmp_trend.name)
                            
                        pdf.set_y(y_base_graficos + 65); pdf.ln(2)
                else: 
                    if not df_maq_fallas.empty:
                        h_mm_top3 = 40
                        if pdf.get_y() + 10 + h_mm_top3 > 260: pdf.add_page()
                        pdf.set_font("Arial", 'B', 10); pdf.set_text_color(*comp_color)
                        pdf.cell(0, 6, clean_text("> Top 3 Fallas (por tiempo):"), ln=True)
                        agg_f = df_maq_fallas.groupby('Detalle_Final')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(3)
                        agg_f['Label'] = agg_f.apply(lambda r: f" {str(r['Detalle_Final'])[:60]} — {r['Tiempo (Min)']:.0f} min ({(r['Tiempo (Min)']/max(t_falla,1))*100:.1f}%)", axis=1)
                        max_x_val = agg_f['Tiempo (Min)'].max() if not agg_f.empty else 1
                        fig_top3 = px.bar(agg_f, x='Tiempo (Min)', y='Detalle_Final', orientation='h', text='Label')
                        fig_top3.update_traces(marker_color=hex_comp, textposition='outside', textfont=dict(size=13, color='black'), cliponaxis=False)
                        fig_top3.update_layout(height=160, width=700, margin=dict(t=5, b=5, l=10, r=220), plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(visible=False, range=[0, max_x_val * 2.5]), yaxis=dict(title='', autorange="reversed", showticklabels=False))
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_chart:
                            fig_top3.write_image(tmp_chart.name, engine="kaleido")
                            add_image_safe(pdf, tmp_chart.name, w_mm=150, h_mm=h_mm_top3, center=False)
                            os.remove(tmp_chart.name)
                        
                        dibujar_tabla_eventos_detallada(df_maq_fallas, 'Detalle_Final', "Detalle de Tiempos Perdidos", comp_color)
                    
                    df_maq_paradas = df_maq[df_maq['Estado_Global'] == 'Parada Programada']
                    if not df_maq_paradas.empty:
                        dibujar_tabla_eventos_detallada(df_maq_paradas, 'Detalle_Final', "Paradas Programadas", theme_color)
        else:
            check_space(pdf, 25); print_section_title(pdf, "3. Analisis de Tiempos y Fallas por Máquina", theme_color)
            pdf.set_font("Arial", 'I', 9); pdf.set_text_color(100, 100, 100); pdf.cell(0, 6, clean_text("No hay desglose de tiempos registrado para las máquinas de este grupo."), ln=True); pdf.ln(5)

        # ---------------------------------
        # 4. RESUMEN VISUAL 
        # ---------------------------------
        resumen_global = df_pdf_g.groupby('Estado_Global')['Tiempo (Min)'].sum().reset_index() if not df_pdf_g.empty else pd.DataFrame()
        total_global = resumen_global['Tiempo (Min)'].sum() if not resumen_global.empty else 0

        if total_global > 0:
            check_space(pdf, 80); print_section_title(pdf, "4. Resumen Visual de Tiempos", theme_color); y_base = pdf.get_y()
            fig_g = px.pie(resumen_global, values='Tiempo (Min)', names='Estado_Global', hole=0.4, title="Global (Hs)", color_discrete_sequence=pie_colors)
            fig_g.update_traces(textinfo='percent+label', textposition='outside', textfont_size=11)
            fig_g.update_layout(width=420, height=300, margin=dict(t=40, b=50, l=40, r=40), showlegend=False, plot_bgcolor='rgba(0,0,0,0)')
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp1:
                fig_g.write_image(tmp1.name, engine="kaleido")
            
            df_fallas_grupo = df_pdf_g[df_pdf_g['Estado_Global'] == 'Falla/Gestión'].copy()
            if not df_fallas_grupo.empty and df_fallas_grupo['Tiempo (Min)'].sum() > 0:
                resumen_fallas = df_fallas_grupo.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
                fig_p = px.pie(resumen_fallas, values='Tiempo (Min)', names='Categoria_Macro', hole=0.4, title="Fallas por Tipo (Hs)", color_discrete_sequence=pie_colors)
                fig_p.update_traces(textinfo='percent+label', textposition='outside', textfont_size=11)
                fig_p.update_layout(width=420, height=300, margin=dict(t=40, b=50, l=40, r=40), showlegend=False, plot_bgcolor='rgba(0,0,0,0)')
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp2:
                    fig_p.write_image(tmp2.name, engine="kaleido")
                    
                if pdf.get_y() + 72 > 260: pdf.add_page(); y_base = pdf.get_y()
                pdf.image(tmp1.name, x=5, y=y_base, w=100)
                pdf.image(tmp2.name, x=105, y=y_base, w=100)
                os.remove(tmp2.name)
            else:
                if pdf.get_y() + 72 > 260: pdf.add_page(); y_base = pdf.get_y()
                pdf.image(tmp1.name, x=55, y=y_base, w=100)
                
            os.remove(tmp1.name)
            pdf.set_y(y_base + 75); pdf.ln(2)
        else:
            check_space(pdf, 25); print_section_title(pdf, "4. Resumen Visual de Tiempos", theme_color)
            pdf.set_font("Arial", 'I', 9); pdf.set_text_color(100, 100, 100); pdf.cell(0, 6, clean_text("No hay tiempos suficientes para generar los gráficos visuales."), ln=True); pdf.ln(5)

        # ---------------------------------
        # 5. PRODUCCIÓN POR MÁQUINA
        # ---------------------------------
        if not df_prod_pdf_g.empty and 'Buenas' in df_prod_pdf_g.columns:
            check_space(pdf, 75); print_section_title(pdf, "5. Produccion por Maquina", theme_color)
            
            prod_maq = df_prod_pdf_g.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
            fig_prod = px.bar(prod_maq, x='Máquina', y=['Buenas', 'Retrabajo', 'Observadas'], barmode='stack', color_discrete_sequence=chart_bars, text_auto=True)
            fig_prod.update_layout(width=800, height=300, margin=dict(t=20, b=40, l=20, r=20))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile3:
                fig_prod.write_image(tmpfile3.name, engine="kaleido")
                add_image_safe(pdf, tmpfile3.name, w_mm=155, h_mm=58)
                os.remove(tmpfile3.name)
            
            pdf.ln(3)
            
            def dibujar_cabeza_prod():
                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 9)
                pdf.cell(70, 6, "Codigo", 1, 0, 'C', True)
                pdf.cell(30, 6, "Buenas", 1, 0, 'C', True); pdf.cell(30, 6, "Retrab.", 1, 0, 'C', True); pdf.cell(30, 6, "Observ.", 1, 1, 'C', True)
            
            maquinas_prod = sorted(df_prod_pdf_g['Máquina'].unique())
            c_cod = next((c for c in df_prod_pdf_g.columns if 'código' in c.lower() or 'codigo' in c.lower()), 'Código')
            
            for maq_p in maquinas_prod:
                df_m_prod = df_prod_pdf_g[df_prod_pdf_g['Máquina'] == maq_p].groupby(c_cod)[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
                total_buenas = df_m_prod['Buenas'].sum(); total_retrabajo = df_m_prod['Retrabajo'].sum(); total_obs = df_m_prod['Observadas'].sum()
                total_piezas = total_buenas + total_retrabajo + total_obs
                
                check_space(pdf, 35)
                pdf.set_font("Arial", 'B', 10); pdf.set_text_color(*theme_color)
                pdf.cell(0, 6, clean_text(f"Top 5 Producción - {maq_p} (Total: {int(total_piezas)} piezas)"), ln=True)
                
                dibujar_cabeza_prod()
                setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                
                top5_prod = df_m_prod.sort_values('Buenas', ascending=False).head(5)
                for _, row in top5_prod.iterrows():
                    if pdf.get_y() > 260:
                        pdf.add_page(); dibujar_cabeza_prod(); setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                    pdf.cell(70, 6, " " + clean_text(str(row[c_cod])[:40]), 'B') 
                    pdf.cell(30, 6, str(int(row['Buenas'])), 'B', 0, 'C'); pdf.cell(30, 6, str(int(row['Retrabajo'])), 'B', 0, 'C'); pdf.cell(30, 6, str(int(row['Observadas'])), 'B', 1, 'C')
                    pdf.ln()
                pdf.ln(5)
        else:
            check_space(pdf, 25); print_section_title(pdf, "5. Produccion por Maquina", theme_color)
            pdf.set_font("Arial", 'I', 9); pdf.set_text_color(100, 100, 100); pdf.cell(0, 6, clean_text("No hay producción registrada para las máquinas de este grupo en el período."), ln=True); pdf.ln(5)

    # =========================================================================
    # SECCIÓN FINAL OPERARIOS 
    # =========================================================================
    check_space(pdf, 35)
    if pdf.get_y() > 40:
        pdf.ln(10); pdf.set_draw_color(*theme_color); pdf.set_line_width(1); pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.set_draw_color(0, 0, 0); pdf.set_line_width(0.2); pdf.ln(10)

    pdf.set_link(link_perfo); pdf.set_font("Times", 'B', 16); pdf.set_text_color(*theme_color)
    pdf.cell(0, 10, clean_text(f"SECCIÓN FINAL: PERFORMANCE Y TIEMPOS"), ln=True, align='L', border='B'); pdf.ln(5)
    print_section_title(pdf, "Performance de Operarios General", theme_color)
    
    if not op_target_df.empty:
        col_op = next((c for c in op_target_df.columns if 'operador' in c.lower() or 'nombre' in c.lower()), op_target_df.columns[1] if len(op_target_df.columns)>1 else op_target_df.columns[0])
        col_perf = next((c for c in op_target_df.columns if 'perf' in c.lower()), None)
        col_area = next((c for c in op_target_df.columns if 'fabrica' in c.lower() or 'fábrica' in c.lower()), None)
        
        if col_perf:
            df_filt = op_target_df.copy()
            if col_area:
                df_filt = df_filt[df_filt[col_area].astype(str).str.contains(area, case=False, na=False)]
            
            if not df_filt.empty:
                df_filt['Perf_Clean'] = pd.to_numeric(df_filt[col_perf].astype(str).str.replace('%', '').str.replace(',', '.'), errors='coerce').fillna(0)
                if df_filt['Perf_Clean'].mean() <= 1.5 and df_filt['Perf_Clean'].mean() > 0:
                    df_filt['Perf_Clean'] = df_filt['Perf_Clean'] * 100
                df_filt = df_filt.sort_values('Perf_Clean', ascending=False)
                
                w_op = 100 if col_area else 160
                
                def dibujar_cabeza_oper():
                    setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 9)
                    pdf.cell(w_op, 6, "Operador", 1, 0, 'C', True)
                    if col_area: pdf.cell(60, 6, "Fabrica", 1, 0, 'C', True)
                    pdf.cell(30, 6, "Perf.", 1, 1, 'C', True)

                dibujar_cabeza_oper()
                setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                for _, row in df_filt.iterrows():
                    if pdf.get_y() > 260: 
                        pdf.add_page(); dibujar_cabeza_oper(); setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                    perf_val = int(round(row['Perf_Clean']))
                    
                    pdf.cell(w_op, 6, " " + clean_text(str(row[col_op])[:60]), 'B')
                    if col_area: pdf.cell(60, 6, " " + clean_text(str(row[col_area])[:30]), 'B')
                    if perf_val >= 90: pdf.set_text_color(33, 195, 84)
                    elif perf_val >= 80: pdf.set_text_color(200, 150, 0)
                    else: pdf.set_text_color(220, 20, 20)
                    pdf.cell(30, 6, f"{perf_val}%", 'B', 1, 'C'); pdf.set_text_color(50, 50, 50)
                pdf.ln(5)
            else:
                pdf.set_font("Arial", 'I', 10); pdf.cell(0, 10, clean_text("No hay datos de performance registrados para esta área en este período."), ln=True)
        else:
            pdf.set_font("Arial", 'I', 10); pdf.cell(0, 10, clean_text("Faltan las columnas necesarias para calcular la performance en GSheets."), ln=True)

    def agregar_tabla_tiempos(titulo, palabras_clave):
        check_space(pdf, 30); print_section_title(pdf, titulo, theme_color)
        resumen_eventos = {}
        if not df_pdf.empty:
            col_t = 'Nivel Evento 4' if 'Nivel Evento 4' in df_pdf.columns else 'Nivel Evento 3'
            mask = df_pdf[col_t].apply(lambda val: isinstance(val, str) and any(kw in val.upper() for kw in palabras_clave))
            df_ev = df_pdf[mask]
            for _, r in df_ev.iterrows():
                t = float(r['Tiempo (Min)'])
                for op in str(r.get('Operador', '')).split('/'):
                    op = op.strip()
                    if op and op != '-':
                        if op not in resumen_eventos: resumen_eventos[op] = {'tiempo': 0.0, 'cantidad': 0}
                        resumen_eventos[op]['tiempo'] += t; resumen_eventos[op]['cantidad'] += 1

        if resumen_eventos:
            df_res = pd.DataFrame([{'Operador': k, 'Minutos': v['tiempo'], 'Cantidad': v['cantidad']} for k, v in resumen_eventos.items()]).sort_values('Minutos', ascending=False)
            df_res['Promedio'] = df_res['Minutos'] / df_res['Cantidad']
            
            def dibujar_cabeza_t():
                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 9)
                pdf.cell(85, 6, "Operador", 1, 0, 'C', True)
                pdf.cell(35, 6, "Total Min", 1, 0, 'C', True)
                pdf.cell(35, 6, "Cant. Veces", 1, 0, 'C', True)
                pdf.cell(35, 6, "Promedio Min", 1, 1, 'C', True)

            dibujar_cabeza_t()
            setup_table_row(pdf); pdf.set_font("Arial", '', 9)
            for _, r in df_res.iterrows():
                if pdf.get_y() > 260: 
                    pdf.add_page(); dibujar_cabeza_t(); setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                
                pdf.cell(85, 6, " " + clean_text(r['Operador'])[:45], 'B')
                pdf.cell(35, 6, f"{r['Minutos']:.1f}", 'B', 0, 'C')
                pdf.cell(35, 6, str(int(r['Cantidad'])), 'B', 0, 'C')
                pdf.cell(35, 6, f"{r['Promedio']:.1f}", 'B', 1, 'C')
            pdf.ln(5)
        else:
            pdf.set_font("Arial", 'I', 10); pdf.cell(0, 10, clean_text("No hay registros de tiempo acumulado para este ítem en el período."), ln=True)

    pdf.set_link(link_tiempos)
    agregar_tabla_tiempos("Tiempo de Baño Acumulado", ["BAÑO", "BANO"])
    agregar_tabla_tiempos("Tiempo de Refrigerio Acumulado", ["REFRIGERIO"])

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_pdf.name)
    with open(temp_pdf.name, "rb") as f: pdf_bytes = f.read()
    os.remove(temp_pdf.name)
    return pdf_bytes

# ==========================================
# 6. BOTONES DE EXPORTACIÓN EN PANTALLA
# ==========================================
with col_p3:
    st.write("**3. Generar y Descargar:**")
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("Preparar Reporte ESTAMPADO", use_container_width=True):
            with st.spinner("Construyendo documento PDF..."):
                try:
                    pdf_data = crear_pdf("Estampado", pdf_label, pdf_df_oee_target, pdf_df_op_target, pdf_ini, pdf_fin, pdf_tipo)
                    st.download_button("Descargar PDF Estampado", data=pdf_data, file_name=f"Estampado_{file_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
                    
    with col_btn2:
        if st.button("Preparar Reporte SOLDADURA", use_container_width=True):
            with st.spinner("Construyendo documento PDF..."):
                try:
                    pdf_data = crear_pdf("Soldadura", pdf_label, pdf_df_oee_target, pdf_df_op_target, pdf_ini, pdf_fin, pdf_tipo)
                    st.download_button("Descargar PDF Soldadura", data=pdf_data, file_name=f"Soldadura_{file_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
