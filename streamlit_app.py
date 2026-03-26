import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import tempfile
import os
from fpdf import FPDF

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
    if st.button("Actualizar Datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# 2. CARGA DE DATOS ROBUSTA
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
pdf_label = ""

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
        
    elif pdf_tipo == "Semanal":
        if not df_oee_sem.empty:
            col_sem = df_oee_sem.columns[0]
            opciones_sem = [s for s in df_oee_sem[col_sem].unique() if s.strip() != "" and str(s).lower() != "nan"]
            pdf_sem = st.selectbox("Semana para PDF:", opciones_sem, label_visibility="collapsed")
            
            pdf_df_oee_target = df_oee_sem[df_oee_sem[col_sem].astype(str) == str(pdf_sem)]
            col_sem_op = df_op_sem_raw.columns[0] if not df_op_sem_raw.empty else None
            if col_sem_op:
                pdf_df_op_target = df_op_sem_raw[df_op_sem_raw[col_sem_op].astype(str) == str(pdf_sem)]
            pdf_label = f"Semana {pdf_sem}"
            
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
            opciones_mes = [m for m in df_oee_men[col_mes].unique() if m.strip() != "" and str(m).lower() != "nan"]
            pdf_mes = st.selectbox("Mes para PDF:", opciones_mes, label_visibility="collapsed")
            
            pdf_df_oee_target = df_oee_men[df_oee_men[col_mes].astype(str) == str(pdf_mes)]
            col_mes_op = df_op_men_raw.columns[0] if not df_op_men_raw.empty else None
            if col_mes_op:
                pdf_df_op_target = df_op_men_raw[df_op_men_raw[col_mes_op].astype(str) == str(pdf_mes)]
            pdf_label = f"Mes {pdf_mes}"
            
            col_ini_p = next((c for c in pdf_df_oee_target.columns if 'inicio' in c.lower()), None)
            col_fin_p = next((c for c in pdf_df_oee_target.columns if 'fin' in c.lower()), None)
            if col_ini_p and col_fin_p and not pdf_df_oee_target.empty:
                pdf_ini = pd.to_datetime(pdf_df_oee_target.iloc[0][col_ini_p], dayfirst=True, errors='coerce')
                pdf_fin = pd.to_datetime(pdf_df_oee_target.iloc[0][col_fin_p], dayfirst=True, errors='coerce')
        else:
            st.warning("No hay datos mensuales.")

# ==========================================
# 4. FUNCIONES HELPER PARA TIEMPO
# ==========================================
def parse_time_to_mins(t_str):
    try:
        t = str(t_str).strip()
        if t in ['nan', 'None', '', '-']: return None
        parts = t.split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except:
        return None

def mins_to_time_str(m):
    if pd.isna(m) or m is None: return "-"
    m = int(m) % 1440
    return f"{m//60:02d}:{m%60:02d}"

def mins_to_duration_str(m):
    if pd.isna(m) or m is None: return "-"
    m = int(m)
    return f"{m//60:02d}:{m%60:02d} hs"

# ==========================================
# 5. CLASE PDF Y FUNCIONES DE ESTILO
# ==========================================
class ReportePDF(FPDF):
    def __init__(self, area, fecha_str, theme_color):
        super().__init__()
        self.area = area
        self.fecha_str = fecha_str
        self.theme_color = theme_color

    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 8, 30)
        
        self.set_font("Times", 'B', 16)
        self.set_text_color(*self.theme_color)
        self.cell(0, 10, clean_text(f"REPORTE GERENCIAL - {self.area.upper()}"), ln=True, align='R')
        
        self.set_font("Arial", 'I', 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 6, clean_text(f"Periodo: {self.fecha_str}"), ln=True, align='R')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Pagina {self.page_no()}", 0, 0, "C")

def clean_text(text):
    if pd.isna(text): return "-"
    text = str(text).replace('•', '-').replace('➤', '>')
    return text.encode('latin-1', 'replace').decode('latin-1')

def check_space(pdf, required_height):
    if pdf.get_y() + required_height > (pdf.h - 15):
        pdf.add_page()

def print_section_title(pdf, title, color):
    pdf.ln(3)
    pdf.set_font("Times", 'B', 14)
    pdf.set_text_color(*color)
    pdf.cell(0, 6, clean_text(title), ln=True)
    
    x, y = pdf.get_x(), pdf.get_y()
    pdf.set_draw_color(*color)
    pdf.set_line_width(0.5)
    pdf.line(x, y, x + 190, y)
    
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(3)

def setup_table_header(pdf, theme_color):
    pdf.set_fill_color(*theme_color)
    pdf.set_text_color(255, 255, 255)
    pdf.set_draw_color(*theme_color)

def setup_table_row(pdf):
    pdf.set_fill_color(255, 255, 255)
    pdf.set_text_color(50, 50, 50)
    pdf.set_draw_color(200, 200, 200)

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

def set_pdf_color(pdf, val):
    if val < 0.85: pdf.set_text_color(220, 20, 20)
    elif val <= 0.95: pdf.set_text_color(200, 150, 0)
    else: pdf.set_text_color(33, 195, 84)

def print_pdf_metric_row(pdf, prefix, m):
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.write(7, clean_text(f"{prefix} | OEE: "))
    set_pdf_color(pdf, m['OEE'])
    pdf.write(7, f"{m['OEE']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.write(7, clean_text("  |  Disp: "))
    set_pdf_color(pdf, m['DISP'])
    pdf.write(7, f"{m['DISP']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.write(7, clean_text("  |  Perf: "))
    set_pdf_color(pdf, m['PERF'])
    pdf.write(7, f"{m['PERF']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.write(7, clean_text("  |  Cal: "))
    set_pdf_color(pdf, m['CAL'])
    pdf.write(7, f"{m['CAL']:.1%}")
    
    pdf.set_text_color(0, 0, 0)
    pdf.ln(7)

def redactar_resumen_ejecutivo(pdf, area, df_pdf, df_oee_target):
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    if df_pdf.empty and df_oee_target.empty:
        pdf.multi_cell(0, 6, clean_text("No hay suficientes datos registrados en este periodo para generar un resumen ejecutivo."))
        return

    lineas = ['L1', 'L2', 'L3', 'L4'] if area.upper() == 'ESTAMPADO' else ['CELDA', 'PRP']
    mejores_oee = {}
    for l in lineas:
        m = get_metrics_direct(l, df_oee_target)
        if m['OEE'] > 0: mejores_oee[l] = m['OEE']
    
    texto_oee = ""
    if mejores_oee:
        mejor_maq = max(mejores_oee, key=mejores_oee.get)
        texto_oee = f"Durante este periodo, la linea/celda con mejor rendimiento general fue {mejor_maq} con un OEE de {mejores_oee[mejor_maq]:.1%}. "

    df_fallas = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.contains('FALLA', case=False)]
    texto_fallas = "No se registraron tiempos muertos por fallas significativos. "
    if not df_fallas.empty:
        total_falla_min = df_fallas['Tiempo (Min)'].sum()
        falla_top = df_fallas.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().idxmax()
        texto_fallas = f"Se registro un total de {total_falla_min:.1f} minutos de parada por fallas. La causa principal de impacto fue '{falla_top}'. "

    resumen = f"Resumen Ejecutivo: {texto_oee}{texto_fallas}Este informe abarca todos los registros documentados y consolidados en el periodo seleccionado."
    
    pdf.set_fill_color(245, 245, 245)
    pdf.multi_cell(0, 6, clean_text(resumen), border=0, fill=True)
    pdf.ln(4)

# ==========================================
# 6. MOTOR GENERADOR DEL PDF
# ==========================================
def crear_pdf(area, label_reporte, oee_target_df, op_target_df, ini_date, fin_date, p_tipo):
    # ASIGNACIÓN DE COLORES DINÁMICA
    if area.upper() == "ESTAMPADO":
        theme_color = (41, 128, 185)    # Azul institucional (Base de tablas)
        subtitle_color = (41, 128, 185) # Azul (Subtítulos)
        chart_bars = ['#1F77B4', '#AEC7E8', '#FF7F0E']
        hex_subtitle = '#%02x%02x%02x' % subtitle_color
    else:
        theme_color = (211, 84, 0)      # Naranja institucional (Base de tablas)
        subtitle_color = (41, 128, 185) # Azul (Exclusivo para subtítulos y gráficos de fallas)
        chart_bars = ['#E67E22', '#FAD7A1', '#1F77B4']
        hex_subtitle = '#%02x%02x%02x' % subtitle_color
        
    hex_theme = '#%02x%02x%02x' % theme_color

    if ini_date is not None and fin_date is not None:
        df_pdf_raw = df_raw[(df_raw['Fecha_Filtro'] >= ini_date) & (df_raw['Fecha_Filtro'] <= fin_date)]
        df_prod_pdf_raw = df_prod_raw[(df_prod_raw['Fecha_Filtro'] >= ini_date) & (df_prod_raw['Fecha_Filtro'] <= fin_date)] if not df_prod_raw.empty else pd.DataFrame()
    else:
        df_pdf_raw = pd.DataFrame(columns=df_raw.columns)
        df_prod_pdf_raw = pd.DataFrame(columns=df_prod_raw.columns)

    df_pdf = df_pdf_raw[df_pdf_raw['Fábrica'].str.contains(area, case=False, na=False)].copy()
    
    df_prod_pdf = pd.DataFrame()
    if not df_prod_pdf_raw.empty:
        df_prod_pdf = df_prod_pdf_raw[(df_prod_pdf_raw['Máquina'].str.contains(area, case=False, na=False)) | 
                                      (df_prod_pdf_raw['Máquina'].isin(df_pdf['Máquina'].unique()))].copy()

    # Iniciar PDF
    pdf = ReportePDF(area, label_reporte, theme_color)
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # 0. RESUMEN EJECUTIVO
    redactar_resumen_ejecutivo(pdf, area, df_pdf, oee_target_df)

    # 1. OEE
    check_space(pdf, 65)
    print_section_title(pdf, "1. Resumen General y OEE", subtitle_color)
    
    metrics_area = get_metrics_direct(area, oee_target_df)
    print_pdf_metric_row(pdf, f"General {area.upper()}", metrics_area)
    
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, clean_text("Tiempos Promedio (por registro en el periodo):"), ln=True)
    pdf.set_font("Arial", '', 10)
    
    if not df_pdf.empty and 'Nivel Evento 3' in df_pdf.columns:
        eventos_completos = df_pdf['Nivel Evento 3'].astype(str) + " " + df_pdf.get('Nivel Evento 4', '').astype(str)
        df_pdf['_evento_completo'] = eventos_completos
        
        avg_bano = df_pdf[df_pdf['_evento_completo'].str.contains('BAÑO|BANO', case=False, na=False)]['Tiempo (Min)'].mean()
        avg_refr = df_pdf[df_pdf['_evento_completo'].str.contains('REFRIGERIO', case=False, na=False)]['Tiempo (Min)'].mean()
        
        str_bano = f"   - Promedio Baño: {avg_bano:.1f} min" if pd.notna(avg_bano) else "   - Promedio Baño: Sin registros"
        str_refr = f"   - Promedio Refrigerio: {avg_refr:.1f} min" if pd.notna(avg_refr) else "   - Promedio Refrigerio: Sin registros"
        
        pdf.cell(0, 6, clean_text(str_bano), ln=True)
        pdf.cell(0, 6, clean_text(str_refr), ln=True)
    else:
        pdf.cell(0, 6, clean_text("   - Sin datos de tiempos para el area y periodo."), ln=True)
    
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, clean_text("Detalle OEE por Maquina/Linea:"), ln=True)
    lineas = ['L1', 'L2', 'L3', 'L4'] if area.upper() == 'ESTAMPADO' else ['CELDA', 'PRP']
    for l in lineas:
        m_l = get_metrics_direct(l, oee_target_df)
        print_pdf_metric_row(pdf, f"   > {l} ", m_l)
    pdf.ln(3) 

    # =========================================================
    # 2. HORARIOS Y TIEMPO DE APERTURA
    # =========================================================
    check_space(pdf, 50)
    print_section_title(pdf, "2. Horarios y Tiempo de Apertura", subtitle_color)
    
    col_inicio = next((c for c in df_pdf.columns if 'inicio' in c.lower() or 'desde' in c.lower()), None)
    col_fin = next((c for c in df_pdf.columns if 'fin' in c.lower() or 'hasta' in c.lower()), None)
    
    tiempo_teorico_area = 0
    
    def dibujar_tabla_eventos_detallada(df_subset, col_detalle, mostrar_categoria=False):
        setup_table_header(pdf, theme_color)
        pdf.set_font("Arial", 'B', 8)
        
        if mostrar_categoria:
            w_f, w_i, w_f2, w_c, w_d, w_m, w_o = 15, 11, 11, 25, 75, 11, 42
            pdf.cell(w_f, 7, "Fecha", border=1, align='C', fill=True)
            pdf.cell(w_i, 7, "Ini.", border=1, align='C', fill=True)
            pdf.cell(w_f2, 7, "Fin", border=1, align='C', fill=True)
            pdf.cell(w_c, 7, "Categoria", border=1, align='L', fill=True)
            pdf.cell(w_d, 7, "Detalle", border=1, align='C', fill=True)
            pdf.cell(w_m, 7, "Min", border=1, align='C', fill=True)
            pdf.cell(w_o, 7, "Operador", border=1, align='C', ln=True, fill=True)
        else:
            w_f, w_i, w_f2, w_d, w_m, w_o = 18, 14, 14, 86, 13, 45
            pdf.cell(w_f, 7, "Fecha", border=1, align='C', fill=True)
            pdf.cell(w_i, 7, "Ini.", border=1, align='C', fill=True)
            pdf.cell(w_f2, 7, "Fin", border=1, align='C', fill=True)
            pdf.cell(w_d, 7, "Detalle Evento", border=1, align='C', fill=True)
            pdf.cell(w_m, 7, "Min", border=1, align='C', fill=True)
            pdf.cell(w_o, 7, "Operador", border=1, align='C', ln=True, fill=True)
        
        setup_table_row(pdf)
        pdf.set_font("Arial", '', 8)
        
        if col_inicio and col_inicio in df_subset.columns:
            df_subset['_sort_time'] = df_subset[col_inicio].apply(lambda x: parse_time_to_mins(x) if pd.notna(x) else 9999)
            df_subset = df_subset.sort_values(['Fecha_Filtro', '_sort_time'], ascending=[True, True])
        else:
            df_subset = df_subset.sort_values(['Fecha_Filtro', 'Tiempo (Min)'], ascending=[True, False])
        
        for _, row in df_subset.iterrows():
            val_fecha = pd.to_datetime(row['Fecha_Filtro']).strftime('%d/%m') if pd.notna(row['Fecha_Filtro']) else "-"
            val_inicio = str(row[col_inicio])[:5] if col_inicio and str(row[col_inicio]) != 'nan' else "-"
            val_fin = str(row[col_fin])[:5] if col_fin and str(row[col_fin]) != 'nan' else "-"
            minutos = f"{row['Tiempo (Min)']:.0f}"
            operador = str(row['Operador'])[:22]
            detalle_str = str(row[col_detalle]) if col_detalle in row and pd.notna(row[col_detalle]) else str(row.get('Evento', '-'))
            
            if mostrar_categoria:
                categoria_str = " " + str(row.get('Nivel Evento 5', '-'))[:15]
                pdf.cell(w_f, 6, val_fecha, border='B', align='C')
                pdf.cell(w_i, 6, val_inicio, border='B', align='C')
                pdf.cell(w_f2, 6, val_fin, border='B', align='C')
                pdf.cell(w_c, 6, clean_text(categoria_str), border='B', align='L') 
                pdf.cell(w_d, 6, clean_text(detalle_str[:60]), border='B')
                pdf.cell(w_m, 6, minutos, border='B', align='C')
                pdf.cell(w_o, 6, clean_text(operador), border='B', ln=True)
            else:
                pdf.cell(w_f, 6, val_fecha, border='B', align='C')
                pdf.cell(w_i, 6, val_inicio, border='B', align='C')
                pdf.cell(w_f2, 6, val_fin, border='B', align='C')
                pdf.cell(w_d, 6, clean_text(detalle_str[:60]), border='B')
                pdf.cell(w_m, 6, minutos, border='B', align='C')
                pdf.cell(w_o, 6, clean_text(operador), border='B', ln=True)

    if col_inicio and col_fin and not df_pdf.empty:
        tiempos_list = []
        for (maq, fecha), g in df_pdf.groupby(['Máquina', 'Fecha_Filtro']):
            intervals = []
            for _, r in g.iterrows():
                ini = parse_time_to_mins(r[col_inicio])
                fin = parse_time_to_mins(r[col_fin])
                if ini is not None and fin is not None:
                    if fin < ini and (ini - fin) > 720: fin += 1440
                    intervals.append([ini, fin])
            
            if not intervals: continue
            
            intervals.sort(key=lambda x: x[0])
            merged = [intervals[0]]
            for current in intervals[1:]:
                last = merged[-1]
                if current[0] <= last[1]: 
                    last[1] = max(last[1], current[1])
                else:
                    merged.append(current)
            
            total_active = sum(iv[1] - iv[0] for iv in merged)
            min_i = merged[0][0]
            max_f = merged[-1][1]
            
            tiempo_bruto = max_f - min_i
            unregistered_time = max(0, tiempo_bruto - total_active)
            
            tiempos_list.append({'Máquina': maq, 'Inicio': min_i, 'Fin': max_f, 'Total': total_active, 'NoReg': unregistered_time})
            
        tiempo_teorico_area = sum((t['Fin'] - t['Inicio']) for t in tiempos_list) if tiempos_list else 0
        
        df_horarios = pd.DataFrame(tiempos_list)
        
        if not df_horarios.empty:
            if p_tipo == "Diario":
                df_res = df_horarios.sort_values('Máquina')
                col_h1, col_h2, col_h3, col_h4 = "Hora Inicio", "Hora Cierre", "Apertura Neta", "No Registrado"
            else:
                df_res = df_horarios.groupby('Máquina').mean().reset_index().sort_values('Máquina')
                pdf.set_font("Arial", 'I', 9)
                pdf.set_text_color(100, 100, 100)
                pdf.cell(0, 5, clean_text("Valores promedio calculados para el periodo seleccionado."), ln=True)
                pdf.ln(2)
                col_h1, col_h2, col_h3, col_h4 = "Inicio Prom.", "Cierre Prom.", "Apertura Neta Prom.", "No Reg. Prom."
            
            setup_table_header(pdf, theme_color)
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(46, 7, clean_text("Maquina"), border=1, fill=True)
            pdf.cell(28, 7, clean_text(col_h1), border=1, align='C', fill=True)
            pdf.cell(28, 7, clean_text(col_h2), border=1, align='C', fill=True)
            pdf.cell(44, 7, clean_text(col_h3), border=1, align='C', fill=True)
            pdf.cell(44, 7, clean_text(col_h4), border=1, align='C', ln=True, fill=True)
            
            setup_table_row(pdf)
            pdf.set_font("Arial", '', 9)
            for _, r in df_res.iterrows():
                pdf.cell(46, 7, clean_text(str(r['Máquina'])[:22]), border=1)
                pdf.cell(28, 7, clean_text(mins_to_time_str(r['Inicio'])), border=1, align='C')
                pdf.cell(28, 7, clean_text(mins_to_time_str(r['Fin'])), border=1, align='C')
                pdf.cell(44, 7, clean_text(mins_to_duration_str(r['Total'])), border=1, align='C')
                pdf.cell(44, 7, clean_text(mins_to_duration_str(r['NoReg'])), border=1, align='C', ln=True)
            pdf.ln(5)
        else:
            pdf.set_font("Arial", '', 10)
            pdf.cell(0, 8, clean_text("No hay datos de horarios validos para calcular."), ln=True)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, clean_text("No se encontraron datos de horarios en el reporte."), ln=True)

    # =========================================================
    # 3. ANÁLISIS DE FALLAS Y PARADAS POR MÁQUINA
    # =========================================================
    check_space(pdf, 40)
    print_section_title(pdf, "3. Analisis de Fallas y Paradas por Maquina", subtitle_color)
    
    df_fallas_area = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.upper().str.contains('FALLA', na=False)]
    df_paradas_area = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.upper().str.contains('PARADA PROGRAMADA', na=False)]
    df_proyectos_area = df_pdf[df_pdf['Nivel Evento 3'].astype(str).str.upper().str.contains('PROYECTO', na=False)]
    df_produccion_area = df_pdf[df_pdf['Evento'].astype(str).str.upper().str.contains('PRODUCCION|PRODUCCIÓN', na=False)]
    
    col_desc_parada = 'Nivel Evento 4'
    
    maquinas_con_eventos = sorted(set(df_pdf['Máquina'].unique()))
    hubo_eventos_en_grupo = False

    for maq in maquinas_con_eventos:
        df_maq_fallas = df_fallas_area[df_fallas_area['Máquina'] == maq]
        df_maq_paradas = df_paradas_area[df_paradas_area['Máquina'] == maq]
        df_maq_proyectos = df_proyectos_area[df_proyectos_area['Máquina'] == maq]
        
        t_prod = df_produccion_area[df_produccion_area['Máquina'] == maq]['Tiempo (Min)'].sum()
        t_falla = df_maq_fallas['Tiempo (Min)'].sum()
        t_pp = df_maq_paradas['Tiempo (Min)'].sum()
        t_proy = df_maq_proyectos['Tiempo (Min)'].sum() if not df_maq_proyectos.empty else 0
        
        if t_prod == 0 and t_falla == 0 and t_pp == 0 and t_proy == 0:
            continue
            
        hubo_eventos_en_grupo = True
        check_space(pdf, 60)
        
        # Título de la máquina DESTACADO EN AZUL
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(*subtitle_color) 
        pdf.cell(0, 9, clean_text(f"  MÁQUINA: {maq}"), border=0, ln=True, fill=True)
        pdf.ln(2)
        
        # Fila Resumen de Tiempos (Mantiene el Naranja institucional)
        setup_table_header(pdf, theme_color)
        pdf.set_font("Arial", 'B', 8)
        
        if t_proy > 0:
            pdf.cell(47, 6, clean_text("Tiempo de produccion"), border=1, align='C', fill=True)
            pdf.cell(47, 6, clean_text("Tiempo de falla"), border=1, align='C', fill=True)
            pdf.cell(48, 6, clean_text("Parada programada"), border=1, align='C', fill=True)
            pdf.cell(48, 6, clean_text("Tiempo de proyecto"), border=1, align='C', ln=True, fill=True)
            
            setup_table_row(pdf)
            pdf.set_font("Arial", '', 9)
            pdf.cell(47, 7, clean_text(mins_to_duration_str(t_prod)), border=1, align='C')
            pdf.cell(47, 7, clean_text(mins_to_duration_str(t_falla)), border=1, align='C')
            pdf.cell(48, 7, clean_text(mins_to_duration_str(t_pp)), border=1, align='C')
            pdf.cell(48, 7, clean_text(mins_to_duration_str(t_proy)), border=1, align='C', ln=True)
        else:
            pdf.cell(63, 6, clean_text("Tiempo de produccion"), border=1, align='C', fill=True)
            pdf.cell(63, 6, clean_text("Tiempo de falla"), border=1, align='C', fill=True)
            pdf.cell(64, 6, clean_text("Parada programada"), border=1, align='C', ln=True, fill=True)
            
            setup_table_row(pdf)
            pdf.set_font("Arial", '', 9)
            pdf.cell(63, 7, clean_text(mins_to_duration_str(t_prod)), border=1, align='C')
            pdf.cell(63, 7, clean_text(mins_to_duration_str(t_falla)), border=1, align='C')
            pdf.cell(64, 7, clean_text(mins_to_duration_str(t_pp)), border=1, align='C', ln=True)

        pdf.ln(2)
        
        # --- TOP 3 Fallas (Gráfico de Barras AZUL) ---
        if not df_maq_fallas.empty and 'Nivel Evento 6' in df_maq_fallas.columns:
            check_space(pdf, 60)
            pdf.set_font("Arial", 'B', 10)
            pdf.set_text_color(220, 20, 20)
            pdf.cell(0, 6, clean_text("Top 3 Fallas (por tiempo):"), ln=True)

            agg_f = df_maq_fallas.groupby('Nivel Evento 6')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(3)
            total_falla_maq = t_falla if t_falla > 0 else 1
            max_val = agg_f['Tiempo (Min)'].max()
            
            agg_f['Nivel Evento 6'] = agg_f['Nivel Evento 6'].apply(lambda x: str(x)[:30] + '...' if len(str(x)) > 30 else str(x))
            agg_f['Porcentaje'] = (agg_f['Tiempo (Min)'] / total_falla_maq) * 100
            agg_f['Label'] = agg_f.apply(lambda r: f"{r['Tiempo (Min)']:.0f} min ({r['Porcentaje']:.1f}%)", axis=1)
            
            fig_top3 = px.bar(agg_f, x='Tiempo (Min)', y='Nivel Evento 6', orientation='h', text='Label')
            fig_top3.update_traces(marker_color=hex_subtitle, textposition='outside', cliponaxis=False)
            fig_top3.update_layout(
                height=140, width=700,
                margin=dict(t=5, b=5, l=260, r=120), 
                plot_bgcolor='rgba(0,0,0,0)', 
                xaxis=dict(visible=False, range=[0, max_val * 1.35]),
                yaxis=dict(title='', autorange="reversed", tickfont=dict(size=12, color='black'))
            )
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_chart:
                fig_top3.write_image(tmp_chart.name, engine="kaleido")
                pdf.image(tmp_chart.name, w=150)
                os.remove(tmp_chart.name)
        
        # --- Tablas de Detalles Condicionales ---
        if not df_maq_fallas.empty:
            check_space(pdf, 25)
            pdf.set_font("Arial", 'B', 9)
            pdf.set_text_color(220, 20, 20)
            pdf.cell(0, 6, clean_text("> Detalle de fallas registradas:"), ln=True)
            dibujar_tabla_eventos_detallada(df_maq_fallas, 'Nivel Evento 6', mostrar_categoria=True)
            pdf.ln(2)
            
        if not df_maq_paradas.empty:
            check_space(pdf, 25)
            pdf.set_font("Arial", 'B', 9)
            pdf.set_text_color(200, 150, 0)
            pdf.cell(0, 6, clean_text("> Detalle de paradas programadas:"), ln=True)
            dibujar_tabla_eventos_detallada(df_maq_paradas, col_desc_parada, mostrar_categoria=False)
            pdf.ln(2)
            
        if not df_maq_proyectos.empty:
            check_space(pdf, 25)
            pdf.set_font("Arial", 'B', 9)
            pdf.set_text_color(33, 195, 84)
            pdf.cell(0, 6, clean_text("> Detalle de paradas por proyecto:"), ln=True)
            dibujar_tabla_eventos_detallada(df_maq_proyectos, col_desc_parada, mostrar_categoria=False)
            
        pdf.ln(6)

    if not hubo_eventos_en_grupo:
        pdf.set_font("Arial", 'I', 9)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 7, clean_text("No se registraron actividades en el area en el periodo."), ln=True)

    # =========================================================
    # 4. RESUMEN GENERAL DE TIEMPOS
    # =========================================================
    if not df_pdf.empty:
        check_space(pdf, 75)
        print_section_title(pdf, "4. Resumen General de Tiempos del Area", subtitle_color)
        
        def categorizar_evento(row):
            evento = str(row.get('Evento', '')).upper()
            if 'PRODUCCION' in evento or 'PRODUCCIÓN' in evento:
                return 'PRODUCCIÓN'
            ne3 = str(row.get('Nivel Evento 3', '')).upper()
            return ne3 if ne3 and ne3 != 'NAN' else 'S/D'

        df_pdf['Categoria_Resumen'] = df_pdf.apply(categorizar_evento, axis=1)
        
        resumen_tiempos = df_pdf.groupby('Categoria_Resumen')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False)
        tiempo_registrado = resumen_tiempos['Tiempo (Min)'].sum()
        tiempo_no_registrado = max(0, tiempo_teorico_area - tiempo_registrado)
        
        fig_pie = px.pie(resumen_tiempos, values='Tiempo (Min)', names='Categoria_Resumen', hole=0.4, 
                         color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_pie.update_layout(width=420, height=270, margin=dict(t=10, b=10, l=10, r=10), plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", y=-0.1))
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile2:
            fig_pie.write_image(tmpfile2.name, engine="kaleido")
            y_before_img = pdf.get_y()
            pdf.image(tmpfile2.name, x=10, y=y_before_img, w=110)
            
            pdf.set_xy(125, y_before_img + 12)
            pdf.set_font("Arial", 'B', 11)
            pdf.set_text_color(*theme_color)
            pdf.cell(70, 7, clean_text("Distribucion (Min):"), ln=True)
            
            for _, row in resumen_tiempos.iterrows():
                lbl = str(row['Categoria_Resumen']).upper()
                val = row['Tiempo (Min)']
                if val <= 0: continue
                
                if 'PROYECTO' in lbl:
                    lbl_print = lbl + " (no afecta disp.)"
                else:
                    lbl_print = lbl
                    
                pdf.set_x(125)
                pdf.set_font("Arial", 'B', 8)
                pdf.set_text_color(50, 50, 50)
                pdf.cell(45, 6, clean_text(lbl_print[:30] + ":"), border=0)
                pdf.set_font("Arial", '', 9)
                pdf.cell(25, 6, clean_text(mins_to_duration_str(val)), border=0, ln=True)
            
            if tiempo_no_registrado > 0:
                pdf.set_x(125)
                pdf.set_font("Arial", 'B', 8)
                pdf.set_text_color(200, 100, 0)
                pdf.cell(45, 6, clean_text("TIEMPO NO REGISTRADO:"), border=0)
                pdf.set_font("Arial", '', 9)
                pdf.cell(25, 6, clean_text(mins_to_duration_str(tiempo_no_registrado)), border=0, ln=True)
            
            pdf.set_y(y_before_img + 75)
            os.remove(tmpfile2.name)
        pdf.ln(3)
    
    # =========================================================
    # 5. PRODUCCIÓN POR MÁQUINA
    # =========================================================
    if not df_prod_pdf.empty and 'Buenas' in df_prod_pdf.columns:
        check_space(pdf, 80)
        print_section_title(pdf, "5. Produccion por Maquina", subtitle_color)
        prod_maq = df_prod_pdf.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
        fig_prod = px.bar(prod_maq, x='Máquina', y=['Buenas', 'Retrabajo', 'Observadas'], barmode='stack', color_discrete_sequence=chart_bars, text_auto=True)
        fig_prod.update_layout(width=800, height=350, margin=dict(t=40, b=100, l=40, r=40), plot_bgcolor='rgba(0,0,0,0)')
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile3:
            fig_prod.write_image(tmpfile3.name, engine="kaleido")
            pdf.image(tmpfile3.name, w=155)
            os.remove(tmpfile3.name)
            
        pdf.ln(3)
        check_space(pdf, 30)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(*theme_color)
        pdf.cell(0, 7, clean_text("Desglose por Codigo de Producto:"), ln=True)
        
        setup_table_header(pdf, theme_color)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(40, 7, clean_text("Maquina"), border=1, fill=True)
        pdf.cell(60, 7, clean_text("Codigo de Producto"), border=1, fill=True)
        pdf.cell(25, 7, clean_text("Buenas"), border=1, align='C', fill=True)
        pdf.cell(25, 7, clean_text("Retrabajo"), border=1, align='C', fill=True)
        pdf.cell(30, 7, clean_text("Observadas"), border=1, align='C', ln=True, fill=True)
        
        setup_table_row(pdf)
        pdf.set_font("Arial", '', 9)
        c_cod = next((c for c in df_prod_pdf.columns if 'código' in c.lower() or 'codigo' in c.lower()), 'Código')
        
        df_prod_group = df_prod_pdf.groupby(['Máquina', c_cod])[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index().sort_values('Máquina')
        for _, row in df_prod_group.iterrows():
            pdf.cell(40, 7, clean_text(str(row['Máquina'])[:25]), border='B')
            pdf.cell(60, 7, clean_text(str(row[c_cod])[:40]), border='B') 
            pdf.cell(25, 7, clean_text(str(int(row['Buenas']))), border='B', align='C')
            pdf.cell(25, 7, clean_text(str(int(row['Retrabajo']))), border='B', align='C')
            pdf.cell(30, 7, clean_text(str(int(row['Observadas']))), border='B', align='C', ln=True)
        pdf.ln(5)

    # =========================================================
    # 6. PERFORMANCE DE OPERARIOS 
    # =========================================================
    check_space(pdf, 60)
    print_section_title(pdf, "6. Performance de Operarios y Maquinas", subtitle_color)
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6, clean_text("Cuadro de desempeno y maquinas operadas en el sector."), ln=True)
    pdf.ln(3)
    
    if not op_target_df.empty:
        col_op = next((c for c in op_target_df.columns if 'operador' in c.lower() or 'nombre' in c.lower()), op_target_df.columns[1] if len(op_target_df.columns)>1 else op_target_df.columns[0])
        
        if p_tipo == "Diario":
            col_perf = op_target_df.columns[5] if len(op_target_df.columns) > 5 else None
            col_area = op_target_df.columns[14] if len(op_target_df.columns) > 14 else None
        else:
            col_perf = op_target_df.columns[7] if len(op_target_df.columns) > 7 else None
            col_area = op_target_df.columns[1] if len(op_target_df.columns) > 1 else None
        
        if col_perf and col_area:
            op_maq_map = {}
            if not df_prod_pdf.empty:
                col_maq_prod = next((c for c in df_prod_pdf.columns if 'máquina' in c.lower() or 'maquina' in c.lower()), None)
                limite_col = min(20, len(df_prod_pdf.columns))
                cols_ops = df_prod_pdf.columns[14:limite_col] if len(df_prod_pdf.columns) > 14 else []
                
                if col_maq_prod and len(cols_ops) > 0:
                    for _, r in df_prod_pdf.iterrows():
                        maq = str(r.get(col_maq_prod, '')).strip()
                        if maq and maq.lower() != 'nan':
                            for c in cols_ops:
                                op = str(r.get(c, '')).strip().upper()
                                if op and op != 'NAN' and op != 'NONE':
                                    if op not in op_maq_map: op_maq_map[op] = set()
                                    op_maq_map[op].add(maq)

            op_target_df['Perf_Clean'] = pd.to_numeric(op_target_df[col_perf].astype(str).str.replace('%', '').str.replace(',', '.'), errors='coerce').fillna(0)
            if op_target_df['Perf_Clean'].mean() <= 1.5 and op_target_df['Perf_Clean'].mean() > 0:
                op_target_df['Perf_Clean'] = op_target_df['Perf_Clean'] * 100
            
            df_grouped = op_target_df.copy()
            df_grouped['Perf_Int'] = df_grouped['Perf_Clean'].round().astype(int)
            df_grouped['Op_Upper'] = df_grouped[col_op].astype(str).str.strip().str.upper()
            
            df_grouped['Maquinas'] = df_grouped['Op_Upper'].apply(
                lambda x: ', '.join(sorted(op_maq_map.get(x, []))) if op_maq_map.get(x) else '-'
            )

            if p_tipo == "Diario":
                df_grouped = df_grouped.groupby(['Op_Upper', col_op, col_area, 'Maquinas']).agg(Perf_Int=('Perf_Int', 'mean')).reset_index()
                df_grouped['Perf_Int'] = df_grouped['Perf_Int'].round().astype(int)

            df_est = df_grouped[df_grouped[col_area].astype(str).str.contains('ESTAMPADO', case=False, na=False)].sort_values('Perf_Int', ascending=False)
            df_sol = df_grouped[df_grouped[col_area].astype(str).str.contains('SOLDADURA', case=False, na=False)].sort_values('Perf_Int', ascending=False)
            
            def imprimir_cuadro_perfo(titulo, df_seccion, t_color):
                check_space(pdf, 30)
                pdf.set_font("Arial", 'B', 12)
                pdf.set_text_color(*t_color)
                pdf.cell(0, 8, clean_text(titulo), ln=True)
                
                setup_table_header(pdf, t_color)
                pdf.set_font("Arial", 'B', 9)
                
                w_op = 60
                w_maq = 100
                w_perf = 30
                
                pdf.cell(w_op, 7, clean_text("Operador"), border=1, fill=True)
                pdf.cell(w_maq, 7, clean_text("Maquina(s) Asignada(s)"), border=1, fill=True)
                pdf.cell(w_perf, 7, clean_text("Performance"), border=1, align='C', ln=True, fill=True)
                
                setup_table_row(pdf)
                pdf.set_font("Arial", '', 9)
                
                if df_seccion.empty:
                    pdf.cell(w_op + w_maq + w_perf, 7, clean_text("Sin registros para esta area."), border='B', align='C', ln=True)
                else:
                    for _, row in df_seccion.iterrows():
                        perf_val = row['Perf_Int']
                        
                        pdf.cell(w_op, 7, clean_text(str(row[col_op])[:35]), border='B')
                        
                        m_str = str(row.get('Maquinas', '-'))
                        if len(m_str) > 60: m_str = m_str[:57] + "..."
                        pdf.cell(w_maq, 7, clean_text(m_str), border='B')
                        
                        if perf_val >= 90: pdf.set_text_color(33, 195, 84)
                        elif perf_val >= 80: pdf.set_text_color(200, 150, 0)
                        else: pdf.set_text_color(220, 20, 20)
                        
                        pdf.set_font("Arial", 'B', 9)
                        pdf.cell(w_perf, 7, clean_text(str(perf_val) + "%"), border='B', align='C', ln=True)
                        
                        pdf.set_text_color(50, 50, 50)
                        pdf.set_font("Arial", '', 9)
                pdf.ln(5)
                
            if area.upper() == "ESTAMPADO":
                imprimir_cuadro_perfo("Operarios ESTAMPADO", df_est, (41, 128, 185)) 
            elif area.upper() == "SOLDADURA":
                imprimir_cuadro_perfo("Operarios SOLDADURA", df_sol, theme_color) 
            
        else:
            pdf.set_font("Arial", '', 10)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 8, clean_text("Faltan columnas de base de datos para generar este cuadro."), ln=True)
    else:
        pdf.set_font("Arial", '', 10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 8, clean_text("No hay registros de performance de operarios para el periodo seleccionado."), ln=True)

    # =========================================================
    # 7 y 8. TABLAS DE PROMEDIO: BAÑO Y REFRIGERIO
    # =========================================================
    def agregar_tabla_tiempos_operarios(titulo, regex_keyword, numero_seccion):
        check_space(pdf, 30)
        print_section_title(pdf, f"{numero_seccion}. {titulo}", subtitle_color)

        try:
            if all(c in df_pdf.columns for c in ['Operador', 'Tiempo (Min)', 'Nivel Evento 3']):
                s_operario = df_pdf['Operador']
                s_tiempo = pd.to_numeric(df_pdf['Tiempo (Min)'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                s_evento = df_pdf['Nivel Evento 3'].astype(str)

                df_temp = pd.DataFrame({
                    'Operario': s_operario,
                    'Tiempo': s_tiempo,
                    'Evento': s_evento
                })

                mask = df_temp['Evento'].str.contains(regex_keyword, case=False, na=False)
                df_filtrado = df_temp[mask]

                if not df_filtrado.empty:
                    resumen = df_filtrado.groupby('Operario').agg(
                        Total_Min=('Tiempo', 'sum'),
                        Eventos=('Tiempo', 'count'),
                        Promedio=('Tiempo', 'mean')
                    ).reset_index().sort_values('Promedio', ascending=False)

                    setup_table_header(pdf, theme_color)
                    pdf.set_font("Arial", 'B', 9)
                    pdf.cell(70, 7, clean_text("Operador"), border=1, align='C', fill=True)
                    pdf.cell(30, 7, clean_text("Cant. Eventos"), border=1, align='C', fill=True)
                    pdf.cell(30, 7, clean_text("Total (Min)"), border=1, align='C', fill=True)
                    pdf.cell(30, 7, clean_text("Promedio (Min)"), border=1, align='C', ln=True, fill=True)

                    setup_table_row(pdf)
                    pdf.set_font("Arial", '', 9)
                    for _, r in resumen.iterrows():
                        op = clean_text(r['Operario']) if str(r['Operario']).strip() else "Desconocido"
                        pdf.cell(70, 7, op, border=1)
                        pdf.cell(30, 7, str(int(r['Eventos'])), border=1, align='C')
                        pdf.cell(30, 7, f"{r['Total_Min']:.1f}", border=1, align='C')
                        pdf.cell(30, 7, f"{r['Promedio']:.1f}", border=1, align='C', ln=True)
                    pdf.ln(5)
                else:
                    pdf.set_font("Arial", '', 10)
                    pdf.set_text_color(100, 100, 100)
                    pdf.cell(0, 8, clean_text("No se registraron tiempos para este evento en esta planta en el periodo."), ln=True)
                    pdf.ln(2)
            else:
                pdf.set_font("Arial", '', 10)
                pdf.set_text_color(100, 100, 100)
                pdf.cell(0, 8, clean_text("Error: Faltan las columnas de Operador, Tiempo o Evento para calcular."), ln=True)
        except Exception as e:
            pdf.set_font("Arial", '', 10)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 8, clean_text(f"Error procesando los datos: {str(e)}"), ln=True)

    agregar_tabla_tiempos_operarios("Tiempo Promedio de Bano por Operario", "BAÑO|BANO", "7")
    agregar_tabla_tiempos_operarios("Tiempo Promedio de Refrigerio por Operario", "REFRIGERIO", "8")

    # FINALIZAR
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp_pdf.name)
    with open(temp_pdf.name, "rb") as f: pdf_bytes = f.read()
    os.remove(temp_pdf.name)
    return pdf_bytes

# ==========================================
# 7. BOTONES DE EXPORTACIÓN EN PANTALLA
# ==========================================
with col_p3:
    st.write("**3. Generar y Descargar:**")
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("Preparar Reporte ESTAMPADO", use_container_width=True):
            with st.spinner("Construyendo documento PDF..."):
                try:
                    pdf_data = crear_pdf("Estampado", pdf_label, pdf_df_oee_target, pdf_df_op_target, pdf_ini, pdf_fin, pdf_tipo)
                    st.download_button("Descargar PDF Estampado", data=pdf_data, file_name=f"Estampado_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
                    
    with col_btn2:
        if st.button("Preparar Reporte SOLDADURA", use_container_width=True):
            with st.spinner("Construyendo documento PDF..."):
                try:
                    pdf_data = crear_pdf("Soldadura", pdf_label, pdf_df_oee_target, pdf_df_op_target, pdf_ini, pdf_fin, pdf_tipo)
                    st.download_button("Descargar PDF Soldadura", data=pdf_data, file_name=f"Soldadura_{pdf_label.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
