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
    "GENERAL": "LÍNEAS ESTAMPADO" 
}

GRUPOS_ESTAMPADO = ['LÍNEAS ESTAMPADO']
GRUPOS_SOLDADURA = ['CELDAS SOLDADURA', 'EQUIPOS PRP']

# ==========================================
# 1. CONFIGURACIÓN Y ESTILOS
# ==========================================
st.set_page_config(page_title="Generador de Reportes PDF - FAMMA", layout="wide", page_icon="📄")

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
    st.write("Seleccione los parámetros para generar y descargar los reportes consolidados directamente de la base de datos SQL.")
with col_btn:
    if st.button("Limpiar Caché", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# 2. CARGA Y LIMPIEZA DE DATOS DESDE SQL SERVER
# ==========================================
@st.cache_data(ttl=300)
def fetch_data_from_db(fecha_ini, fecha_fin, tipo_periodo, mes=None, anio=None):
    try:
        conn = st.connection("wii_bi", type="sql")
        ini_str = fecha_ini.strftime('%Y-%m-%d')
        fin_str = fecha_fin.strftime('%Y-%m-%d')

        df_trend = pd.DataFrame()
        df_horarios = pd.DataFrame()

        if tipo_periodo == "Mensual":
            q_prod = f"""
                SELECT c.Name as Máquina, pr.Code as Código, 
                       SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas
                FROM PROD_M_01 p JOIN CELL c ON p.CellId = c.CellId JOIN PRODUCT pr ON p.ProductId = pr.ProductId 
                WHERE p.Month = {mes} AND p.Year = {anio} GROUP BY c.Name, pr.Code
            """
            
            q_metrics = f"""
                SELECT c.Name as Máquina, 
                       SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas,
                       SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada,
                       (SUM(p.Performance * p.ProductiveTime) / NULLIF(SUM(p.ProductiveTime), 0)) as PERFORMANCE,
                       (SUM(p.Availability * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as DISPONIBILIDAD,
                       (SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) / NULLIF(SUM(p.Good + p.Rework + p.Scrap), 0)) as CALIDAD,
                       (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE
                FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId
                WHERE p.Month = {mes} AND p.Year = {anio}
                GROUP BY c.Name
            """

            q_op = f"""
                SELECT DISTINCT op.Name as Operador, p.Factory as Fábrica, 
                       (SUM(p.Performance * p.ProductiveTime) OVER(PARTITION BY p.OperatorId) / NULLIF(SUM(p.ProductiveTime) OVER(PARTITION BY p.OperatorId), 0)) as PERFORMANCE, 
                       SUM(p.BathTime) OVER(PARTITION BY p.OperatorId) as BathTime, 
                       SUM(p.BreakTime) OVER(PARTITION BY p.OperatorId) as BreakTime, 
                       SUM(p.FeedingTime) OVER(PARTITION BY p.OperatorId) as FeedingTime 
                FROM OPER_M_01 p JOIN OPERATOR op ON p.OperatorId = op.OperatorId 
                WHERE p.Month = {mes} AND p.Year = {anio}
            """
            df_op_target = conn.query(q_op)
            
            q_trend = f"""
                SELECT p.Month, c.Name as Máquina,
                       SUM(p.Oee * (p.ProductiveTime + p.DownTime)) as OEE_Num,
                       SUM(p.ProductiveTime + p.DownTime) as OEE_Den,
                       (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE,
                       SUM(p.Availability * (p.ProductiveTime + p.DownTime)) as Disp_Num,
                       SUM(p.Performance * p.ProductiveTime) as Perf_Num,
                       SUM(p.ProductiveTime) as T_Operativo,
                       SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) as Cal_Num,
                       SUM(p.Good + p.Rework + p.Scrap) as Piezas_Totales
                FROM PROD_M_03 p JOIN CELL c ON p.CellId = c.CellId
                WHERE p.Year = {anio} AND p.Month <= {mes}
                GROUP BY p.Month, c.Name
            """
            df_trend = conn.query(q_trend)
            
        else:
            q_prod = f"""
                SELECT c.Name as Máquina, pr.Code as Código, 
                       SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas
                FROM PROD_D_01 p JOIN CELL c ON p.CellId = c.CellId JOIN PRODUCT pr ON p.ProductId = pr.ProductId 
                WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}' GROUP BY c.Name, pr.Code
            """
            
            q_metrics = f"""
                SELECT c.Name as Máquina, 
                       SUM(p.Good) as Buenas, SUM(p.Rework) as Retrabajo, SUM(p.Scrap) as Observadas,
                       SUM(p.ProductiveTime) as T_Operativo, SUM(p.DownTime) as T_Parada,
                       (SUM(p.Performance * p.ProductiveTime) / NULLIF(SUM(p.ProductiveTime), 0)) as PERFORMANCE,
                       (SUM(p.Availability * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as DISPONIBILIDAD,
                       (SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) / NULLIF(SUM(p.Good + p.Rework + p.Scrap), 0)) as CALIDAD,
                       (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE
                FROM PROD_D_03 p JOIN CELL c ON p.CellId = c.CellId
                WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}'
                GROUP BY c.Name
            """
            
            q_op = f"""
                SELECT op.Name as Operador, p.Factory as Fábrica,
                       p.Performance, p.ProductiveTime
                FROM OPER_D_01 p 
                JOIN OPERATOR op ON p.OperatorId = op.OperatorId 
                WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}' 
            """
            df_op_raw = conn.query(q_op)
            
            if not df_op_raw.empty:
                df_op_raw['Performance'] = pd.to_numeric(df_op_raw['Performance'], errors='coerce').fillna(0)
                df_op_raw['ProductiveTime'] = pd.to_numeric(df_op_raw['ProductiveTime'], errors='coerce').fillna(0)
                df_op_raw['Perf_Num'] = df_op_raw['Performance'] * df_op_raw['ProductiveTime']
                
                df_op_target = df_op_raw.groupby(['Operador', 'Fábrica']).agg(
                    Perf_Num=('Perf_Num', 'sum'),
                    ProductiveTime=('ProductiveTime', 'sum')
                ).reset_index()
                
                df_op_target['PERFORMANCE'] = df_op_target['Perf_Num'] / df_op_target['ProductiveTime'].replace(0, 1)
            else:
                df_op_target = pd.DataFrame()

            # --- CONSULTA PARA HORARIOS Y TURNOS ADAPTADA A PERÍODOS DIARIOS/SEMANALES ---
            q_horarios = f"""
                WITH Tiempos_Turno AS (
                    SELECT CellId, TurnId, Date as Dia,
                           MIN(Started) as MinInicio,
                           MAX(Finish) as MaxFin
                    FROM EVENT_01
                    WHERE Date BETWEEN '{ini_str}' AND '{fin_str}'
                    GROUP BY CellId, TurnId, Date
                )
                SELECT c.Name as Máquina, tu.Name as Turno, t.Dia,
                       FORMAT(MIN(t.MinInicio), 'HH:mm') as Hora_Inicio,
                       FORMAT(MAX(t.MaxFin), 'HH:mm') as Hora_Cierre,
                       SUM(ISNULL(p.ProductiveTime, 0) + ISNULL(p.DownTime, 0)) as Apertura_Neta_Min,
                       CASE 
                           WHEN ISNULL(DATEDIFF(MINUTE, MIN(t.MinInicio), MAX(t.MaxFin)), 0) - SUM(ISNULL(p.ProductiveTime, 0) + ISNULL(p.DownTime, 0)) > 0 
                           THEN ISNULL(DATEDIFF(MINUTE, MIN(t.MinInicio), MAX(t.MaxFin)), 0) - SUM(ISNULL(p.ProductiveTime, 0) + ISNULL(p.DownTime, 0))
                           ELSE 0 
                       END as No_Registrado_Min
                FROM Tiempos_Turno t
                JOIN CELL c ON t.CellId = c.CellId
                JOIN TURN tu ON t.TurnId = tu.TurnId
                LEFT JOIN PROD_D_02 p ON t.CellId = p.CellId AND t.TurnId = p.TurnId AND t.Dia = p.Date
                GROUP BY c.Name, tu.Name, t.Dia
            """
            df_horarios = conn.query(q_horarios)

            if tipo_periodo == "Semanal":
                q_trend_semanal = f"""
                    SELECT p.Date as Fecha_Filtro, c.Name as Máquina,
                           SUM(p.Oee * (p.ProductiveTime + p.DownTime)) as OEE_Num,
                           SUM(p.ProductiveTime + p.DownTime) as OEE_Den,
                           (SUM(p.Oee * (p.ProductiveTime + p.DownTime)) / NULLIF(SUM(p.ProductiveTime + p.DownTime), 0)) as OEE,
                           SUM(p.Availability * (p.ProductiveTime + p.DownTime)) as Disp_Num,
                           SUM(p.Performance * p.ProductiveTime) as Perf_Num,
                           SUM(p.ProductiveTime) as T_Operativo,
                           SUM(p.Quality * (p.Good + p.Rework + p.Scrap)) as Cal_Num,
                           SUM(p.Good + p.Rework + p.Scrap) as Piezas_Totales
                    FROM PROD_D_03 p JOIN CELL c ON p.CellId = c.CellId
                    WHERE p.Date BETWEEN '{ini_str}' AND '{fin_str}'
                    GROUP BY p.Date, c.Name
                """
                df_trend = conn.query(q_trend_semanal)
            else:
                df_trend = pd.DataFrame()

        df_prod_target = conn.query(q_prod)
        df_metrics = conn.query(q_metrics)

        if not df_op_target.empty:
            df_op_target = df_op_target[~df_op_target['Operador'].str.lower().str.contains('usuario', na=False)]

        q_event = f"""
            SELECT e.Id as Evento_Id, c.Name as Máquina, e.Started as Inicio, e.Finish as Fin, 
                   e.Interval as [Tiempo (Min)], 
                   t1.Name as [Nivel Evento 1], t2.Name as [Nivel Evento 2], 
                   t3.Name as [Nivel Evento 3], t4.Name as [Nivel Evento 4], 
                   t5.Name as [Nivel Evento 5], t6.Name as [Nivel Evento 6],
                   t7.Name as [Nivel Evento 7], t8.Name as [Nivel Evento 8],
                   t9.Name as [Nivel Evento 9],
                   op_celda.Name as Operador_Celda,
                   op_req.Name as Operador_Req,
                   op_resp.Name as Operador_Resp,
                   e.Date as Fecha_Filtro, f.Name as Fábrica, tu.Name as Turno
            FROM EVENT_01 e
            LEFT JOIN CELL c ON e.CellId = c.CellId
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            LEFT JOIN EVENTTYPE t5 ON e.EventTypeLevel5 = t5.EventTypeId
            LEFT JOIN EVENTTYPE t6 ON e.EventTypeLevel6 = t6.EventTypeId
            LEFT JOIN EVENTTYPE t7 ON e.EventTypeLevel7 = t7.EventTypeId
            LEFT JOIN EVENTTYPE t8 ON e.EventTypeLevel8 = t8.EventTypeId
            LEFT JOIN EVENTTYPE t9 ON e.EventTypeLevel9 = t9.EventTypeId
            LEFT JOIN FACTORY f ON e.FactoryId = f.FactoryId
            LEFT JOIN TURN tu ON e.TurnId = tu.TurnId
            LEFT JOIN EVENT_OPERATOR_01 eo ON e.Id = eo.EventId
            LEFT JOIN OPERATOR op_celda ON eo.OperatorId = op_celda.OperatorId
            LEFT JOIN ANDON_01 a ON e.CellId = a.CellId AND e.Started = a.Started
            LEFT JOIN OPERATOR op_req ON a.RequesterOperatorId = op_req.OperatorId
            LEFT JOIN OPERATOR op_resp ON a.ResponserOperatorId = op_resp.OperatorId
            WHERE e.Date BETWEEN '{ini_str}' AND '{fin_str}'
        """
        df_raw = conn.query(q_event)

        if not df_raw.empty:
            df_raw['Fecha_Filtro'] = pd.to_datetime(df_raw['Fecha_Filtro']).dt.date
            df_raw['Inicio_Str'] = pd.to_datetime(df_raw['Inicio']).dt.strftime('%H:%M')
            df_raw['Fin_Str'] = pd.to_datetime(df_raw['Fin']).dt.strftime('%H:%M')
            df_raw['Tiempo (Min)'] = pd.to_numeric(df_raw['Tiempo (Min)'], errors='coerce').fillna(0)
            
            df_raw['Operador_Celda'] = df_raw['Operador_Celda'].fillna('').astype(str)
            df_raw['Operador_Req'] = df_raw['Operador_Req'].fillna('').astype(str)
            df_raw['Operador_Resp'] = df_raw['Operador_Resp'].fillna('').astype(str)

            cols_grupo = [c for c in df_raw.columns if c not in ['Operador_Celda', 'Operador_Req', 'Operador_Resp']]

            def agrupar_nombres(ops):
                n = [str(x).strip() for x in ops.unique() if pd.notna(x) and str(x).strip() != '']
                return ' / '.join(n)

            df_raw = df_raw.groupby(cols_grupo, dropna=False).agg({
                'Operador_Celda': agrupar_nombres,
                'Operador_Req': agrupar_nombres,
                'Operador_Resp': agrupar_nombres
            }).reset_index()

            def determinar_operador_final(row):
                resp = row['Operador_Resp']
                req = row['Operador_Req']
                celda = row['Operador_Celda']
                
                if resp:
                    reales = [n.strip() for n in resp.split('/') if 'usuario' not in n.lower() and 'admin' not in n.lower()]
                    if reales: return ' / '.join(reales)
                if req:
                    reales = [n.strip() for n in req.split('/') if 'usuario' not in n.lower() and 'admin' not in n.lower()]
                    if reales: return ' / '.join(reales)
                if celda:
                    reales = [n.strip() for n in celda.split('/') if 'usuario' not in n.lower() and 'admin' not in n.lower()]
                    if reales: return ' / '.join(reales)
                return '-'

            df_raw['Operador'] = df_raw.apply(determinar_operador_final, axis=1)

            cols_niveles = [c for c in df_raw.columns if 'Nivel Evento' in c]

            def categorizar_estado(row):
                texto_completo = " ".join([str(row.get(c, '')) for c in cols_niveles]).upper()
                if 'PRODUCCION' in texto_completo or 'PRODUCCIÓN' in texto_completo: return 'Producción'
                if 'PROYECTO' in texto_completo: return 'Proyecto'
                if 'BAÑO' in texto_completo or 'BANO' in texto_completo or 'REFRIGERIO' in texto_completo: return 'Descanso'
                if 'PARADA PROGRAMADA' in texto_completo: return 'Parada Programada'
                return 'Falla/Gestión'

            def clasificar_macro(row):
                texto_completo = " ".join([str(row.get(c, '')) for c in cols_niveles]).upper()
                categorias_clave = ["MANTENIMIENTO", "MATRICERIA", "DISPOSITIVOS", "TECNOLOGIA", "GESTION", "LOGISTICA", "CALIDAD"]
                for cat in categorias_clave:
                    if cat in texto_completo:
                        return cat.capitalize()
                return 'Otra Falla/Gestión'

            def obtener_detalle_final(row):
                niveles = [str(row.get(c, '')) for c in cols_niveles]
                validos = [n.strip() for n in niveles if n.strip() and n.strip().lower() not in ['none', 'nan', 'null']]
                if not validos: return "Sin detalle en sistema"
                ultimo_dato = validos[-1].upper()
                estado = row.get('Estado_Global', '')
                categoria = row.get('Categoria_Macro', '')
                if estado == 'Falla/Gestión':
                    if categoria != 'Otra Falla/Gestión':
                        return f"[{categoria.upper()}] {ultimo_dato}"
                    return ultimo_dato
                return validos[-1]

            df_raw['Estado_Global'] = df_raw.apply(categorizar_estado, axis=1)
            df_raw['Categoria_Macro'] = df_raw.apply(clasificar_macro, axis=1)
            df_raw['Detalle_Final'] = df_raw.apply(obtener_detalle_final, axis=1)

        return df_raw, df_prod_target, df_op_target, df_trend, df_metrics, df_horarios

    except Exception as e:
        st.error(f"Error ejecutando consulta a base de datos wii_bi: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==========================================
# 3. INTERFAZ: CONFIGURACIÓN PERIODO
# ==========================================
col_p1, col_p2, col_p3 = st.columns([1, 1.2, 2.0])

with col_p1:
    st.write("**1. Tipo de Reporte:**")
    pdf_tipo = st.radio("Período:", ["Diario", "Semanal", "Mensual"], horizontal=True, label_visibility="collapsed")

with col_p2:
    st.write("**2. Seleccione el Período:**")
    today = pd.to_datetime("today").date()
    pdf_ini, pdf_fin, pdf_mes, pdf_anio = None, None, None, None
    pdf_label, file_label = "", ""

    if pdf_tipo == "Diario":
        pdf_fecha = st.date_input("Día para PDF:", value=today)
        pdf_ini = pdf_fin = pd.to_datetime(pdf_fecha)
        pdf_label = f"Dia {pdf_fecha.strftime('%d-%m-%Y')}"
        file_label = pdf_label
        
    elif pdf_tipo == "Semanal":
        fecha_ref = st.date_input("Seleccione un día de la semana deseada:", value=today)
        dt_ref = pd.to_datetime(fecha_ref)
        pdf_ini = dt_ref - timedelta(days=dt_ref.weekday()); pdf_fin = pdf_ini + timedelta(days=6) 
        semana_num = pdf_ini.isocalendar().week
        pdf_label = f"Semana {semana_num} ({pdf_ini.strftime('%d/%m/%Y')} al {pdf_fin.strftime('%d/%m/%Y')})"
        file_label = f"Semana_{semana_num}_{pdf_ini.strftime('%d-%m-%Y')}_al_{pdf_fin.strftime('%d-%m-%Y')}"
        
    elif pdf_tipo == "Mensual":
        c_m, c_y = st.columns(2)
        mes_list = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        with c_m: mes_sel = st.selectbox("Mes", mes_list, index=today.month-1)
        with c_y: anio_sel = st.selectbox("Año", range(2023, today.year + 2), index=today.year-2023)
        pdf_mes = mes_list.index(mes_sel) + 1; pdf_anio = anio_sel
        pdf_ini = pd.to_datetime(f"{pdf_anio}-{pdf_mes}-01")
        last_day = calendar.monthrange(pdf_anio, pdf_mes)[1]
        pdf_fin = pd.to_datetime(f"{pdf_anio}-{pdf_mes}-{last_day}")
        pdf_label = f"{mes_sel} {pdf_anio}"; file_label = f"{mes_sel}_{pdf_anio}"

df_raw, pdf_df_prod_target, pdf_df_op_target, df_trend, df_metrics, df_horarios = fetch_data_from_db(pdf_ini, pdf_fin, pdf_tipo, mes=pdf_mes, anio=pdf_anio)

# --- CREAR COPIAS DE SEGURIDAD PARA EL ANEXO DEL PDF ---
df_metrics_ORIGINAL = df_metrics.copy()
df_prod_ORIGINAL = pdf_df_prod_target.copy()

# ==========================================
# 4. FUNCIONES HELPER PDF
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
        if os.path.exists("logo.jpg"): self.image("logo.jpg", 10, 8, 30)
        elif os.path.exists("logo.png"): self.image("logo.png", 10, 8, 30)
        self.set_font("Times", 'B', 16); self.set_text_color(*self.theme_color)
        self.cell(0, 10, clean_text(f"REPORTE GERENCIAL - {self.area.upper()}"), ln=True, align='R')
        self.set_font("Arial", 'B', 10); self.set_text_color(100, 100, 100)
        self.cell(0, 6, clean_text(f"Periodo: {self.fecha_str}"), ln=True, align='R'); self.ln(5)

    def footer(self):
        self.set_y(-15); self.set_font("Arial", "I", 8); self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Pagina {self.page_no()}", 0, 0, "C")

def clean_text(text):
    if pd.isna(text): return "-"
    # Usamos caracteres ASCII seguros para FPDF
    return str(text).replace('•', '-').replace('➤', '>').encode('latin-1', 'replace').decode('latin-1')

def check_space(pdf, required_height):
    if pdf.get_y() + required_height > 275 and pdf.get_y() > 40:
        pdf.add_page(); return True
    return False

def print_section_title(pdf, title, theme_color):
    pdf.ln(3); pdf.set_font("Times", 'B', 14); pdf.set_text_color(*theme_color)
    pdf.cell(0, 6, clean_text(title), ln=True)
    x, y = pdf.get_x(), pdf.get_y()
    pdf.set_draw_color(*theme_color); pdf.set_line_width(0.5); pdf.line(x, y, x + 190, y)
    pdf.set_draw_color(0, 0, 0); pdf.set_line_width(0.2); pdf.set_text_color(0, 0, 0); pdf.ln(3)

def setup_table_header(pdf, theme_color):
    pdf.set_fill_color(*theme_color); pdf.set_text_color(255, 255, 255); pdf.set_draw_color(*theme_color)

def setup_table_row(pdf):
    pdf.set_fill_color(255, 255, 255); pdf.set_text_color(50, 50, 50); pdf.set_draw_color(200, 200, 200)

def set_pdf_color_metric(pdf, val, metric_name):
    targets = {
        'OEE': 0.75,
        'DISPONIBILIDAD': 0.88,
        'PERFORMANCE': 0.90,
        'CALIDAD': 0.95
    }
    target = targets.get(metric_name.upper(), 0.85)
    
    if val >= target:
        pdf.set_text_color(33, 195, 84) # Verde
    else:
        pdf.set_text_color(220, 20, 20) # Rojo

def print_pdf_metric_row(pdf, prefix, m):
    pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0, 0, 0)
    pdf.write(7, clean_text(f"{prefix} | OEE: "))
    set_pdf_color_metric(pdf, m.get('OEE', 0), 'OEE'); pdf.write(7, f"{m.get('OEE', 0):.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(7, clean_text("  |  Disp: "))
    set_pdf_color_metric(pdf, m.get('DISPONIBILIDAD', 0), 'DISPONIBILIDAD'); pdf.write(7, f"{m.get('DISPONIBILIDAD', 0):.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(7, clean_text("  |  Perf: "))
    set_pdf_color_metric(pdf, m.get('PERFORMANCE', 0), 'PERFORMANCE'); pdf.write(7, f"{m.get('PERFORMANCE', 0):.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.write(7, clean_text("  |  Cal: "))
    set_pdf_color_metric(pdf, m.get('CALIDAD', 0), 'CALIDAD'); pdf.write(7, f"{m.get('CALIDAD', 0):.1%}")
    
    pdf.set_text_color(0, 0, 0); pdf.ln(7)

def add_image_safe(pdf, img_path, w_mm, h_mm, center=True):
    if pdf.get_y() + h_mm > 275:
        pdf.add_page()
    x = (210 - w_mm) / 2 if center else pdf.get_x()
    y = pdf.get_y()
    pdf.image(img_path, x=x, y=y, w=w_mm)
    pdf.set_y(y + h_mm + 5)

# ==========================================
# 5.A. MOTOR PARA RESUMEN EJECUTIVO (SOLO MENSUAL)
# ==========================================
def crear_pdf_resumen_ejecutivo(fecha_str, df_trend, df_metrics_pdf):
    theme_color = (44, 62, 80) 
    pdf = ReportePDF("GLOBAL PLANTA", fecha_str, theme_color)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    print_section_title(pdf, "RESUMEN EJECUTIVO: KPI POR PLANTA", theme_color)

    def get_planta(maq_name):
        maq_upper = str(maq_name).strip().upper()
        if 'CELL' in maq_upper or 'CELDA' in maq_upper or 'PRP' in maq_upper or 'SOLD' in maq_upper:
            return 'SOLDADURA'
        if 'LINEA' in maq_upper or 'LÍNEA' in maq_upper:
            return 'ESTAMPADO'
        return 'OTRO'

    df_met_all = df_metrics_pdf.copy()
    if not df_met_all.empty and df_met_all['OEE'].max() > 1.5:
        df_met_all['OEE'] = df_met_all['OEE'] / 100.0
        df_met_all['DISPONIBILIDAD'] = df_met_all['DISPONIBILIDAD'] / 100.0
        df_met_all['PERFORMANCE'] = df_met_all['PERFORMANCE'] / 100.0
        df_met_all['CALIDAD'] = df_met_all['CALIDAD'] / 100.0

    df_met_all['Planta'] = df_met_all['Máquina'].apply(get_planta)
    
    df_met_all['T_Planificado'] = df_met_all['T_Operativo'].fillna(0) + df_met_all['T_Parada'].fillna(0)
    df_met_all['Piezas_Totales'] = df_met_all['Buenas'].fillna(0) + df_met_all['Retrabajo'].fillna(0) + df_met_all['Observadas'].fillna(0)

    df_met_all['OEE_Num'] = df_met_all['OEE'].fillna(0) * df_met_all['T_Planificado']
    df_met_all['Disp_Num'] = df_met_all['DISPONIBILIDAD'].fillna(0) * df_met_all['T_Planificado']
    df_met_all['Perf_Num'] = df_met_all['PERFORMANCE'].fillna(0) * df_met_all['T_Operativo']
    df_met_all['Cal_Num'] = df_met_all['CALIDAD'].fillna(0) * df_met_all['Piezas_Totales']

    met_planta = df_met_all.groupby('Planta')[['OEE_Num', 'Disp_Num', 'Perf_Num', 'Cal_Num', 'T_Planificado', 'T_Operativo', 'Piezas_Totales']].sum()

    def calc_metrics(p_name):
        if p_name in met_planta.index:
            row = met_planta.loc[p_name]
            oee = row['OEE_Num'] / row['T_Planificado'] if row['T_Planificado'] > 0 else 0
            disp = row['Disp_Num'] / row['T_Planificado'] if row['T_Planificado'] > 0 else 0
            perf = row['Perf_Num'] / row['T_Operativo'] if row['T_Operativo'] > 0 else 0
            cal = row['Cal_Num'] / row['Piezas_Totales'] if row['Piezas_Totales'] > 0 else 0
            return oee, disp, perf, cal
        return 0, 0, 0, 0

    oee_est, disp_est, perf_est, cal_est = calc_metrics('ESTAMPADO')
    oee_sol, disp_sol, perf_sol, cal_sol = calc_metrics('SOLDADURA')

    def draw_kpi_row(y, title, oee, disp, perf, cal):
        pdf.set_xy(10, y)
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(*theme_color)
        pdf.cell(0, 6, clean_text(title), ln=1)
        y_boxes = pdf.get_y() + 2
        w = 42; spacing = 5; x_start = 13.5
        
        def draw_box(x, title_box, val):
            pdf.set_xy(x, y_boxes)
            pdf.set_font("Arial", 'B', 9)
            pdf.set_fill_color(*theme_color)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(w, 8, clean_text(title_box), border=1, align='C', fill=True, ln=2)
            
            pdf.set_fill_color(245, 245, 245)
            set_pdf_color_metric(pdf, val, title_box)
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(w, 12, f"{val*100:.1f}%", border=1, align='C', fill=True)
        
        draw_box(x_start, "OEE", oee)
        draw_box(x_start + w + spacing, "DISPONIBILIDAD", disp)
        draw_box(x_start + 2*(w + spacing), "PERFORMANCE", perf)
        draw_box(x_start + 3*(w + spacing), "CALIDAD", cal)
        
        return y_boxes + 25

    y_curr = pdf.get_y() + 5
    y_curr = draw_kpi_row(y_curr, "INDICADORES: ESTAMPADO", oee_est, disp_est, perf_est, cal_est)
    y_curr += 8
    y_curr = draw_kpi_row(y_curr, "INDICADORES: SOLDADURA", oee_sol, disp_sol, perf_sol, cal_sol)

    if not df_trend.empty:
        pdf.set_y(y_curr + 10)
        pdf.set_font("Arial", 'B', 12); pdf.set_text_color(*theme_color)
        pdf.cell(0, 6, clean_text("Evolución Mensual Histórica (4 Indicadores por Planta)"), ln=True)

        df_trend_all = df_trend.copy()
        df_trend_all['Planta'] = df_trend_all['Máquina'].apply(get_planta)

        trend_planta = df_trend_all[df_trend_all['Planta'] != 'OTRO'].groupby(['Month', 'Planta'])[['OEE_Num', 'OEE_Den', 'Disp_Num', 'Perf_Num', 'Cal_Num', 'T_Operativo', 'Piezas_Totales']].sum().reset_index()
        
        trend_planta['OEE'] = (trend_planta['OEE_Num'] / trend_planta['OEE_Den']).fillna(0)
        trend_planta['DISP'] = (trend_planta['Disp_Num'] / trend_planta['OEE_Den']).fillna(0)
        trend_planta['PERF'] = (trend_planta['Perf_Num'] / trend_planta['T_Operativo']).fillna(0)
        trend_planta['CAL'] = (trend_planta['Cal_Num'] / trend_planta['Piezas_Totales']).fillna(0)

        trend_melt = trend_planta.melt(id_vars=['Month', 'Planta'], value_vars=['OEE', 'DISP', 'PERF', 'CAL'], var_name='Indicador', value_name='Valor')
        
        if trend_melt['Valor'].max() <= 10.0 and trend_melt['Valor'].max() > 0:
            trend_melt['Valor'] = trend_melt['Valor'] * 100

        meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
        trend_melt['Mes_Nombre'] = trend_melt['Month'].map(meses_map)

        fig_glob = px.bar(
            trend_melt, x='Mes_Nombre', y='Valor', color='Indicador', facet_row='Planta',
            barmode='group', text_auto='.0f',
            color_discrete_map={'OEE': '#2C3E50', 'DISP': '#2980B9', 'PERF': '#F39C12', 'CAL': '#27AE60'}
        )
        fig_glob.update_layout(
            height=450, width=800, margin=dict(t=30, b=20, l=20, r=20),
            yaxis_title='Porcentaje (%)', xaxis_title='',
            plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        fig_glob.update_yaxes(range=[0, 110])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_glob:
            fig_glob.write_image(tmp_glob.name, engine="kaleido")
            add_image_safe(pdf, tmp_glob.name, w_mm=190, h_mm=115, center=True)
            os.remove(tmp_glob.name)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 5.B. MOTOR GENERADOR DEL PDF PRINCIPAL
# ==========================================
def crear_pdf(area, label_reporte, op_target_df, prod_target_df, df_pdf_raw, p_tipo, df_trend, df_metrics_pdf, df_horarios, override_estampado=False, df_metrics_ORIG=None, df_prod_ORIG=None):
    if area.upper() == "ESTAMPADO":
        theme_color = (15, 76, 129); comp_color = (52, 152, 219)  
        chart_bars = ['#003366', '#3498DB', '#AED6F1']; pie_colors = px.colors.sequential.Blues_r
        grupos_area = GRUPOS_ESTAMPADO
    else:
        theme_color = (211, 84, 0); comp_color = (230, 126, 34) 
        chart_bars = ['#993300', '#E67E22', '#FAD7A1']; pie_colors = px.colors.sequential.Oranges_r
        grupos_area = GRUPOS_SOLDADURA
        
    hex_theme = '#%02x%02x%02x' % theme_color; hex_comp = '#%02x%02x%02x' % comp_color  

    if not df_metrics_pdf.empty and df_metrics_pdf['OEE'].max() > 1.5:
        df_metrics_pdf['OEE'] = df_metrics_pdf['OEE'] / 100.0
        df_metrics_pdf['DISPONIBILIDAD'] = df_metrics_pdf['DISPONIBILIDAD'] / 100.0
        df_metrics_pdf['PERFORMANCE'] = df_metrics_pdf['PERFORMANCE'] / 100.0
        df_metrics_pdf['CALIDAD'] = df_metrics_pdf['CALIDAD'] / 100.0

    def asignar_grupo_dinamico(maq):
        maq_u = str(maq).strip().upper()
        if maq_u in MAQUINAS_MAP: return MAQUINAS_MAP[maq_u]
        if 'CELL' in maq_u or 'CELDA' in maq_u: return 'CELDAS SOLDADURA'
        if 'LINEA' in maq_u or 'LÍNEA' in maq_u: return 'LÍNEAS ESTAMPADO'
        if 'PRP' in maq_u or 'SOLD' in maq_u: return 'EQUIPOS PRP'
        return 'Otro'

    df_pdf = pd.DataFrame(columns=['Máquina', 'Fábrica', 'Estado_Global', 'Tiempo (Min)', 'Operador', 'Nivel Evento 2'])
    if not df_pdf_raw.empty:
        mask_area = (df_pdf_raw['Fábrica'].astype(str).str.contains(area, case=False, na=False)) | \
                    (df_pdf_raw.get('Nivel Evento 2', pd.Series(dtype=str)).astype(str).str.contains(area, case=False, na=False))
        df_pdf = df_pdf_raw[mask_area].copy()
    
    df_pdf['Máquina_Match'] = df_pdf['Máquina'].astype(str).str.strip().str.upper()
    df_pdf['Grupo_Máquina'] = df_pdf['Máquina_Match'].apply(asignar_grupo_dinamico)

    df_prod_pdf = pd.DataFrame(columns=['Máquina', 'Buenas', 'Retrabajo', 'Observadas'])
    if not prod_target_df.empty: df_prod_pdf = prod_target_df.copy()
    df_prod_pdf['Máquina_Match'] = df_prod_pdf['Máquina'].astype(str).str.strip().str.upper()
    df_prod_pdf['Grupo_Máquina'] = df_prod_pdf['Máquina_Match'].apply(asignar_grupo_dinamico)

    pdf = ReportePDF(area, label_reporte, theme_color)
    pdf.set_auto_page_break(auto=True, margin=15); pdf.add_page()
    
    links_resumen_grupo = {g: pdf.add_link() for g in grupos_area}
    links_detalle_grupo = {g: pdf.add_link() for g in grupos_area}
    link_perfo = pdf.add_link(); link_tiempos = pdf.add_link()

    # --- ÍNDICE ---
    pdf.ln(10); pdf.set_font("Times", 'B', 18); pdf.set_text_color(*theme_color)
    pdf.cell(0, 10, clean_text("ÍNDICE DEL REPORTE"), ln=True, align='C')
    pdf.ln(10); pdf.set_font("Arial", 'U', 11); pdf.set_text_color(*comp_color)
    for g in grupos_area:
        pdf.cell(0, 7, clean_text(f">> Grupo {g} - Resumen General del Área"), ln=True, link=links_resumen_grupo[g])
        pdf.cell(0, 7, clean_text(f"      -> Ir al Detalle Individual Máquina a Máquina"), ln=True, link=links_detalle_grupo[g])
        pdf.ln(1)
    pdf.ln(4)
    pdf.cell(0, 8, clean_text(">> Performance General de Operarios"), ln=True, link=link_perfo)
    pdf.cell(0, 8, clean_text(">> Tablas de Tiempos Acumulados de Descanso"), ln=True, link=link_tiempos)

    def obtener_metricas_maquina(maq_name):
        maq_row = df_metrics_pdf[df_metrics_pdf['Máquina'] == maq_name]
        if maq_row.empty: return None
        r = maq_row.iloc[0]
        return {
            'OEE': r['OEE'], 'DISPONIBILIDAD': r['DISPONIBILIDAD'], 
            'PERFORMANCE': r['PERFORMANCE'], 'CALIDAD': r['CALIDAD'], 
            'T_Planificado': (r['T_Operativo'] + r['T_Parada']) if pd.notna(r['T_Operativo']) else 0,
            'T_Operativo': r['T_Operativo'] if pd.notna(r['T_Operativo']) else 0, 
            'Buenas': r['Buenas'] if pd.notna(r['Buenas']) else 0, 
            'Totales': (r['Buenas'] + r['Retrabajo'] + r['Observadas']) if pd.notna(r['Buenas']) else 0
        }

    def dibujar_tabla_eventos_detallada(df_subset, col_detalle, titulo, color_t):
        if not df_subset.empty:
            check_space(pdf, 25); pdf.set_font("Arial", 'B', 9); pdf.set_text_color(*color_t)
            pdf.cell(0, 6, clean_text(f">> {titulo}:"), ln=True); pdf.ln(1)
            def dibujar_cabeceras():
                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                pdf.cell(15, 6, "Fecha", 1, 0, 'C', True)
                pdf.cell(12, 6, "Ini.", 1, 0, 'C', True)
                pdf.cell(12, 6, "Fin", 1, 0, 'C', True)
                pdf.cell(106, 6, "Detalle Registrado en Sistema", 1, 0, 'L', True)
                pdf.cell(12, 6, "Min", 1, 0, 'C', True)
                pdf.cell(33, 6, "Operador", 1, 1, 'L', True)
            
            dibujar_cabeceras(); setup_table_row(pdf); pdf.set_font("Arial", '', 8)
            df_subset['_sort_time'] = df_subset['Inicio_Str'].apply(lambda x: parse_time_to_mins(x) if pd.notna(x) else 9999)
            df_subset = df_subset.sort_values(['Fecha_Filtro', '_sort_time'], ascending=[True, True])
            for _, row in df_subset.iterrows():
                if pdf.get_y() > 265:
                    pdf.add_page(); dibujar_cabeceras(); setup_table_row(pdf); pdf.set_font("Arial", '', 8)
                val_fecha = pd.to_datetime(row['Fecha_Filtro']).strftime('%d/%m') if pd.notna(row['Fecha_Filtro']) else "-"
                val_inicio = str(row['Inicio_Str'])[:5] if pd.notna(row['Inicio_Str']) else "-"
                val_fin = str(row['Fin_Str'])[:5] if pd.notna(row['Fin_Str']) else "-"
                minutos = f"{row['Tiempo (Min)']:.0f}"
                operador = " " + str(row['Operador'])[:15]
                detalle_str = " " + str(row[col_detalle]) if col_detalle in row and pd.notna(row[col_detalle]) else " Sin detalle"
                
                pdf.cell(15, 5, val_fecha, 'B', 0, 'C')
                pdf.cell(12, 5, val_inicio, 'B', 0, 'C')
                pdf.cell(12, 5, val_fin, 'B', 0, 'C')
                pdf.cell(106, 5, clean_text(detalle_str[:70]), 'B', 0, 'L')
                pdf.cell(12, 5, minutos, 'B', 0, 'C')
                pdf.cell(33, 5, clean_text(operador), 'B', 1, 'L')
            pdf.ln(4)

    # --- RECORRIDO POR CADA GRUPO ---
    for g in grupos_area:
        maq_del_grupo = sorted(list(set(df_pdf[df_pdf['Grupo_Máquina'] == g]['Máquina'].tolist() + df_prod_pdf[df_prod_pdf['Grupo_Máquina'] == g]['Máquina'].tolist())))
        df_pdf_g = df_pdf[df_pdf['Máquina'].isin(maq_del_grupo)]
        if df_pdf_g.empty and not any(m in df_prod_pdf['Máquina'].values for m in maq_del_grupo): continue
            
        pdf.add_page(); pdf.set_link(links_resumen_grupo[g]) 
        pdf.set_font("Times", 'B', 16); pdf.set_text_color(*theme_color)
        pdf.cell(0, 10, clean_text(f"SECCIÓN GRUPO: {g}"), ln=True, align='L', border='B'); pdf.ln(5)

        # 1. RESUMEN OEE GRUPAL
        print_section_title(pdf, "1. Resumen OEE del Grupo", theme_color)
        g_plan = 0; g_op = 0; g_buenas = 0; g_totales = 0; g_disp_w = 0; g_perf_w = 0; g_oee_w = 0
        maquinas_metricas = {}
        for maq in maq_del_grupo:
            m = obtener_metricas_maquina(maq)
            if m:
                maquinas_metricas[maq] = m
                tp = m['T_Planificado'] if m['T_Planificado'] > 0 or not override_estampado else 1
                to = m['T_Operativo'] if m['T_Operativo'] > 0 or not override_estampado else 1
                g_plan += tp; g_op += to; g_buenas += m['Buenas']; g_totales += m['Totales']
                g_disp_w += m['DISPONIBILIDAD'] * tp; g_perf_w += m['PERFORMANCE'] * to; g_oee_w += m['OEE'] * tp
        
        g_oee = g_oee_w / g_plan if g_plan > 0 else 0
        print_pdf_metric_row(pdf, f"Total {g}", {'OEE': g_oee, 'DISPONIBILIDAD': g_disp_w/g_plan if g_plan>0 else 0, 'PERFORMANCE': g_perf_w/g_op if g_op>0 else 0, 'CALIDAD': g_buenas/g_totales if g_totales>0 else 0})
        for maq, metrics in maquinas_metricas.items(): print_pdf_metric_row(pdf, f"    > {maq}", metrics)
        pdf.ln(5)

        # ==========================================
        # NUEVO: 2. Gráficos Evolución o KPIs
        # ==========================================
        if p_tipo == "Mensual":
            check_space(pdf, 80)
            print_section_title(pdf, "2. Evolución Histórica OEE por Máquina", theme_color)
            
            if not df_trend.empty:
                df_trend_g = df_trend[df_trend['Máquina'].isin(maq_del_grupo)].copy()
                if not df_trend_g.empty:
                    if df_trend_g['OEE'].max() <= 10.0:
                        df_trend_g['OEE'] = df_trend_g['OEE'] * 100
                    
                    meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
                    df_trend_g['Mes_Nombre'] = df_trend_g['Month'].map(meses_map)
                    
                    fig_trend_oee = px.bar(
                        df_trend_g, x='Mes_Nombre', y='OEE', color='Máquina', 
                        barmode='group', text_auto='.1f', color_discrete_sequence=px.colors.qualitative.Prism
                    )
                    fig_trend_oee.update_layout(
                        height=350, width=800, margin=dict(t=20, b=20, l=20, r=20),
                        yaxis_title='OEE (%)', xaxis_title='', legend_title='Máquinas', 
                        plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, 110])
                    )
                    
                    y_base = pdf.get_y()
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_oee:
                        fig_trend_oee.write_image(tmp_oee.name, engine="kaleido")
                        pdf.image(tmp_oee.name, x=10, y=y_base, w=190)
                        os.remove(tmp_oee.name)
                        
                    pdf.set_y(y_base + 90); pdf.ln(2)
                else:
                    pdf.set_font("Arial", 'I', 9); pdf.cell(0, 6, clean_text("No hay datos históricos para graficar en este grupo."), ln=True)
            else:
                pdf.set_font("Arial", 'I', 9); pdf.cell(0, 6, clean_text("No hay datos de evolución para graficar en este periodo."), ln=True)
        
        else: # Semanal o Diario
            check_space(pdf, 80)
            print_section_title(pdf, f"2. Comparativa de KPIs entre Máquinas ({p_tipo})", theme_color)
            
            df_m_g = df_metrics_pdf[df_metrics_pdf['Máquina'].isin(maq_del_grupo)].copy()
            if not df_m_g.empty:
                df_m_g_melt = df_m_g.melt(id_vars=['Máquina'], value_vars=['OEE', 'DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD'], var_name='Indicador', value_name='Valor')
                
                df_m_g_melt['Valor'] = df_m_g_melt['Valor'] * 100
                
                fig_kpis = px.bar(
                    df_m_g_melt, x='Indicador', y='Valor', color='Máquina', 
                    barmode='group', text_auto='.1f',
                    color_discrete_sequence=px.colors.qualitative.Prism
                )
                fig_kpis.update_layout(
                    height=350, width=800, margin=dict(t=20, b=20, l=20, r=20),
                    yaxis_title='Porcentaje (%)', xaxis_title='', 
                    plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[0, 110]),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title="")
                )
                
                y_base = pdf.get_y()
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_kpi:
                    fig_kpis.write_image(tmp_kpi.name, engine="kaleido")
                    pdf.image(tmp_kpi.name, x=10, y=y_base, w=190)
                    os.remove(tmp_kpi.name)
                    
                pdf.set_y(y_base + 90); pdf.ln(2)
            else:
                pdf.set_font("Arial", 'I', 9); pdf.cell(0, 6, clean_text("No hay datos de KPIs para graficar en este periodo."), ln=True)

        # 3. HORARIOS (SOLO DIARIO/SEMANAL)
        if p_tipo in ["Diario", "Semanal"]:
            check_space(pdf, 25); print_section_title(pdf, "3. Horarios y Tiempo de Apertura", theme_color)
            setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)

            if not df_horarios.empty:
                df_horarios_grupo = df_horarios[df_horarios['Máquina'].isin(maq_del_grupo)].copy()

                if not df_horarios_grupo.empty:
                    if p_tipo == "Semanal":
                        w_maq = 35; w_tur = 15; w_day = 27
                        pdf.cell(w_maq, 6, "Maquina", 1, 0, 'C', True)
                        pdf.cell(w_tur, 6, "Turno", 1, 0, 'C', True)
                        dias = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"]
                        for d in dias:
                            pdf.cell(w_day, 6, d, 1, 0 if d != "Viernes" else 1, 'C', True)

                        setup_table_row(pdf); pdf.set_font("Arial", '', 8)

                        df_horarios_grupo['Dia'] = pd.to_datetime(df_horarios_grupo['Dia'])
                        df_horarios_grupo['Weekday'] = df_horarios_grupo['Dia'].dt.weekday
                        df_horarios_grupo['Rango'] = df_horarios_grupo.apply(
                            lambda row: f"{row['Hora_Inicio']}-{row['Hora_Cierre']}" if pd.notna(row['Hora_Inicio']) else "", axis=1
                        )

                        grouped = df_horarios_grupo.groupby(['Máquina', 'Turno'])
                        for (maq_name, turno), group in grouped:
                            pdf.cell(w_maq, 5, " " + clean_text(maq_name), 1, 0, 'L')
                            pdf.cell(w_tur, 5, clean_text(turno), 1, 0, 'C')

                            for day_idx in range(5):
                                day_data = group[group['Weekday'] == day_idx]
                                if not day_data.empty:
                                    rango = day_data.iloc[0]['Rango']
                                    pdf.cell(w_day, 5, rango, 1, 0 if day_idx < 4 else 1, 'C')
                                else:
                                    pdf.cell(w_day, 5, "", 1, 0 if day_idx < 4 else 1, 'C')
                            pdf.ln()

                    else: # Diario
                        w_maq = 35; w_tur = 20; w_hor = 30; w_tie = 35
                        pdf.cell(w_maq, 6, "Maquina", 1, 0, 'C', True)
                        pdf.cell(w_tur, 6, "Turno", 1, 0, 'C', True)
                        pdf.cell(w_hor, 6, "Hora Inicio", 1, 0, 'C', True)
                        pdf.cell(w_hor, 6, "Hora Cierre", 1, 0, 'C', True)
                        pdf.cell(w_tie, 6, "Apertura Neta", 1, 0, 'C', True)
                        pdf.cell(w_tie, 6, "No Registrado", 1, 1, 'C', True)

                        setup_table_row(pdf); pdf.set_font("Arial", '', 8)
                        for _, r_hor in df_horarios_grupo.sort_values(['Máquina', 'Turno']).iterrows():
                            pdf.cell(w_maq, 5, " " + clean_text(r_hor['Máquina']), 1, 0, 'L')
                            pdf.cell(w_tur, 5, clean_text(r_hor['Turno']), 1, 0, 'C')
                            
                            hora_ini = str(r_hor['Hora_Inicio'])
                            try:
                                h, m = map(int, hora_ini.split(':'))
                                total_min = h * 60 + m
                                if 360 <= total_min < 375: pdf.set_text_color(33, 195, 84)
                                elif 375 <= total_min < 380: pdf.set_text_color(218, 165, 32)
                                elif 380 <= total_min < 420: pdf.set_text_color(220, 20, 20)
                                elif total_min >= 420: pdf.set_text_color(128, 0, 128)
                                else: pdf.set_text_color(50, 50, 50)
                            except:
                                pdf.set_text_color(50, 50, 50)
                            
                            pdf.cell(w_hor, 5, hora_ini, 1, 0, 'C')
                            pdf.set_text_color(50, 50, 50) 
                            pdf.cell(w_hor, 5, clean_text(r_hor['Hora_Cierre']), 1, 0, 'C')
                            
                            apertura_str = mins_to_duration_str(r_hor.get('Apertura_Neta_Min', 0))
                            no_reg_str = mins_to_duration_str(r_hor.get('No_Registrado_Min', 0))
                            
                            pdf.cell(w_tie, 5, apertura_str, 1, 0, 'C')
                            pdf.cell(w_tie, 5, no_reg_str, 1, 1, 'C')
                else:
                    pdf.cell(185, 5, "No hay registros de turnos para este grupo.", 1, 1, 'C')
            else:
                pdf.cell(185, 5, "No hay registros de turnos para este periodo.", 1, 1, 'C')
                
            pdf.ln(5)

        # 4. RESUMEN GENERAL DE TIEMPOS DEL GRUPO
        check_space(pdf, 30)
        num_section_tiempos = "3." if p_tipo == "Mensual" else "4."
        print_section_title(pdf, f"{num_section_tiempos} Resumen General del Grupo (Tiempos Consolidados)", theme_color)
        
        t_prod_g = df_pdf_g[df_pdf_g['Estado_Global'] == 'Producción']['Tiempo (Min)'].sum()
        t_falla_g = df_pdf_g[df_pdf_g['Estado_Global'] == 'Falla/Gestión']['Tiempo (Min)'].sum()
        t_parada_g = df_pdf_g[df_pdf_g['Estado_Global'] == 'Parada Programada']['Tiempo (Min)'].sum()
        t_proy_g = df_pdf_g[df_pdf_g['Estado_Global'] == 'Proyecto']['Tiempo (Min)'].sum()
        t_desc_g = df_pdf_g[df_pdf_g['Estado_Global'] == 'Descanso']['Tiempo (Min)'].sum()
        
        pdf.set_font("Arial", 'B', 12); pdf.set_text_color(255, 255, 255); pdf.set_fill_color(*comp_color)
        pdf.cell(0, 8, clean_text(f"  RESUMEN TOTAL ACUMULADO GRUPO: {g}"), border=0, ln=True, fill=True)
        pdf.set_font("Arial", 'I', 8); pdf.set_text_color(120, 120, 120); pdf.cell(0, 5, clean_text("  Acumulado total de todas las máquinas de la familia"), border=0, ln=True); pdf.ln(2)
        
        setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
        for col_name in ["Produccion", "Fallas/Gestion", "Paradas Prog.", "Proyecto", "Descansos"]: pdf.cell(38, 6, col_name, border=1, align='C', fill=True)
        pdf.ln(); setup_table_row(pdf); pdf.set_font("Arial", '', 9)
        pdf.cell(38, 5, clean_text(mins_to_duration_str(t_prod_g)), border=1, align='C')
        pdf.cell(38, 5, clean_text(mins_to_duration_str(t_falla_g)), border=1, align='C')
        pdf.cell(38, 5, clean_text(mins_to_duration_str(t_parada_g)), border=1, align='C')
        pdf.cell(38, 5, clean_text(mins_to_duration_str(t_proy_g)), border=1, align='C')
        pdf.cell(38, 5, clean_text(mins_to_duration_str(t_desc_g)), border=1, align='C', ln=True); pdf.ln(4)
        
        # --- Desglose de Fallas del Grupo ---
        df_g_fallas = df_pdf_g[df_pdf_g['Estado_Global'] == 'Falla/Gestión'].copy()
        if not df_g_fallas.empty:
            agg_f15 = df_g_fallas.groupby('Detalle_Final')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(15)
            agg_f15 = agg_f15.sort_values('Tiempo (Min)', ascending=True) 
            agg_f15['Label'] = agg_f15.apply(lambda r: f" {str(r['Detalle_Final'])[:60]} — {r['Tiempo (Min)']:.0f}m", axis=1)
            max_x_val = agg_f15['Tiempo (Min)'].max() if not agg_f15.empty else 1
            
            if p_tipo == "Diario":
                df_g_fallas['Eje_Temp'] = pd.to_datetime(df_g_fallas['Inicio']).dt.strftime('%H:00')
            else:
                df_g_fallas['Eje_Temp'] = pd.to_datetime(df_g_fallas['Fecha_Filtro']).dt.strftime('%d/%m')
                
            trend_df = df_g_fallas.groupby(['Eje_Temp', 'Máquina'])['Tiempo (Min)'].sum().reset_index()
            trend_df = trend_df.sort_values('Eje_Temp')
            
            if pdf.get_y() + 65 > 275: pdf.add_page()
            pdf.set_font("Arial", 'B', 10); pdf.set_text_color(*comp_color)
            pdf.cell(95, 6, clean_text("> Top 15 Fallas del Grupo (por tiempo):"), 0, 0, 'L')
            pdf.cell(95, 6, clean_text("> Tendencia Temporal de Fallas (Minutos):"), 0, 1, 'L')
            
            y_base_graficos = pdf.get_y()
            
            fig_top15 = px.bar(agg_f15, x='Tiempo (Min)', y='Detalle_Final', orientation='h', text='Label')
            fig_top15.update_traces(marker_color=hex_comp, textposition='outside', textfont=dict(size=11, color='black'), cliponaxis=False)
            fig_top15.update_layout(height=250, width=450, margin=dict(t=5, b=5, l=10, r=220), plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(visible=False, range=[0, max_x_val * 1.5]), yaxis=dict(title='', showticklabels=False))
            
            fig_trend = px.line(trend_df, x='Eje_Temp', y='Tiempo (Min)', color='Máquina', markers=True, color_discrete_sequence=px.colors.qualitative.Set1)
            fig_trend.update_layout(height=250, width=400, margin=dict(t=10, b=30, l=40, r=20), plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Minutos", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=""))
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_chart:
                fig_top15.write_image(tmp_chart.name, engine="kaleido")
                pdf.image(tmp_chart.name, x=5, y=y_base_graficos, w=105)
                os.remove(tmp_chart.name)
                
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_trend:
                fig_trend.write_image(tmp_trend.name, engine="kaleido")
                pdf.image(tmp_trend.name, x=110, y=y_base_graficos, w=90)
                os.remove(tmp_trend.name)
                
            pdf.set_y(y_base_graficos + 60); pdf.ln(2)

        # --- Desglose de Paradas Programadas del Grupo ---
        df_g_paradas = df_pdf_g[df_pdf_g['Estado_Global'] == 'Parada Programada'].copy()
        if not df_g_paradas.empty:
            if pdf.get_y() > 180: pdf.add_page()
            pdf.set_font("Arial", 'B', 10); pdf.set_text_color(*theme_color) 
            pdf.cell(95, 6, clean_text("> SMED / Paradas Programadas Grupo:"), 0, 0, 'L')
            pdf.cell(95, 6, clean_text("> Tendencia Temporal (Minutos):"), 0, 1, 'L')
            
            y_base_p = pdf.get_y()
            
            resumen_p = df_g_paradas.groupby('Detalle_Final').agg(Cantidad=('Tiempo (Min)', 'count'), Total_Min=('Tiempo (Min)', 'sum')).reset_index()
            resumen_p['Promedio_Min'] = resumen_p['Total_Min'] / resumen_p['Cantidad']
            resumen_p = resumen_p.sort_values('Total_Min', ascending=False)
            
            if p_tipo == "Diario":
                df_g_paradas['Eje_Temp'] = pd.to_datetime(df_g_paradas['Inicio']).dt.strftime('%H:00')
            else:
                df_g_paradas['Eje_Temp'] = pd.to_datetime(df_g_paradas['Fecha_Filtro']).dt.strftime('%d/%m')
                
            trend_p = df_g_paradas.groupby(['Eje_Temp', 'Máquina'])['Tiempo (Min)'].sum().reset_index()
            trend_p = trend_p.sort_values('Eje_Temp')
            
            fig_trend_p = px.line(trend_p, x='Eje_Temp', y='Tiempo (Min)', color='Máquina', markers=True, color_discrete_sequence=px.colors.qualitative.Safe)
            fig_trend_p.update_layout(height=320, width=420, margin=dict(t=10, b=100, l=40, r=10), plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Total Minutos", legend=dict(orientation="h", yanchor="top", y=-0.3, xanchor="center", x=0.5, font=dict(size=8), title=""))
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_trend_p:
                fig_trend_p.write_image(tmp_trend_p.name, engine="kaleido")
                pdf.image(tmp_trend_p.name, x=105, y=y_base_p, w=100)
                os.remove(tmp_trend_p.name)
                
            pdf.set_y(y_base_p + 2)
            setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
            pdf.cell(50, 5, "Evento", 1, 0, 'C', True); pdf.cell(12, 5, "Cant.", 1, 0, 'C', True); pdf.cell(16, 5, "Total", 1, 0, 'C', True); pdf.cell(16, 5, "Prom.", 1, 1, 'C', True)
            setup_table_row(pdf); pdf.set_font("Arial", '', 7)
            max_y_tab = pdf.get_y()
            
            for _, rp in resumen_p.head(10).iterrows():
                pdf.cell(50, 4.5, " " + clean_text(rp['Detalle_Final'])[:33], 'B', 0, 'L')
                pdf.cell(12, 4.5, str(int(rp['Cantidad'])), 'B', 0, 'C')
                pdf.cell(16, 4.5, f"{rp['Total_Min']:.0f}m", 'B', 0, 'C')
                pdf.cell(16, 4.5, f"{rp['Promedio_Min']:.1f}m", 'B', 1, 'C')
                max_y_tab = pdf.get_y()
                
            pdf.set_y(max(max_y_tab, y_base_p + 75) + 5)

        # 5. RESUMEN VISUAL DE TIEMPOS DEL GRUPO
        resumen_global = df_pdf_g.groupby('Estado_Global')['Tiempo (Min)'].sum().reset_index() if not df_pdf_g.empty else pd.DataFrame()
        total_global = resumen_global['Tiempo (Min)'].sum() if not resumen_global.empty else 0

        pdf.add_page()
        num_section_visual = "4." if p_tipo == "Mensual" else "5."
        
        if total_global > 0:
            print_section_title(pdf, f"{num_section_visual} Resumen Visual de Tiempos del Grupo", theme_color); y_base = pdf.get_y()
            fig_g = px.pie(resumen_global, values='Tiempo (Min)', names='Estado_Global', hole=0.4, title="Estructura de Tiempos (Hs)", color_discrete_sequence=pie_colors)
            fig_g.update_traces(textinfo='percent+label', textposition='outside', textfont_size=11)
            fig_g.update_layout(width=420, height=300, margin=dict(t=40, b=50, l=80, r=80), showlegend=False, plot_bgcolor='rgba(0,0,0,0)')
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp1:
                fig_g.write_image(tmp1.name, engine="kaleido")
            
            df_fallas_grupo = df_pdf_g[df_pdf_g['Estado_Global'] == 'Falla/Gestión'].copy()
            if not df_fallas_grupo.empty and df_fallas_grupo['Tiempo (Min)'].sum() > 0:
                resumen_fallas = df_fallas_grupo.groupby('Categoria_Macro')['Tiempo (Min)'].sum().reset_index()
                fig_p = px.pie(resumen_fallas, values='Tiempo (Min)', names='Categoria_Macro', hole=0.4, title="Fallas Distribuidas por Área (Hs)", color_discrete_sequence=pie_colors)
                fig_p.update_traces(textinfo='percent+label', textposition='outside', textfont_size=11)
                fig_p.update_layout(width=420, height=300, margin=dict(t=40, b=50, l=80, r=80), showlegend=False, plot_bgcolor='rgba(0,0,0,0)')
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp2:
                    fig_p.write_image(tmp2.name, engine="kaleido")
                    
                pdf.image(tmp1.name, x=5, y=y_base, w=100)
                pdf.image(tmp2.name, x=105, y=y_base, w=100)
                os.remove(tmp2.name)
            else:
                pdf.image(tmp1.name, x=55, y=y_base, w=100)
                
            os.remove(tmp1.name)
            pdf.set_y(y_base + 75); pdf.ln(2)
        else:
            print_section_title(pdf, f"{num_section_visual} Resumen Visual de Tiempos del Grupo", theme_color)
            pdf.set_font("Arial", 'I', 9); pdf.set_text_color(100, 100, 100); pdf.cell(0, 6, clean_text("No hay datos de tiempo suficientes para generar gráficos de torta."), ln=True); pdf.ln(5)

        # 6. PRODUCCIÓN POR GRUPO
        df_prod_pdf_g = df_prod_pdf[df_prod_pdf['Grupo_Máquina'] == g] if not df_prod_pdf.empty else pd.DataFrame()
        if not df_prod_pdf_g.empty:
            pdf.add_page()
            print_section_title(pdf, "Desglose de Producción del Grupo", theme_color)
            
            prod_maq = df_prod_pdf_g.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
            fig_prod = px.bar(prod_maq, x='Máquina', y=['Buenas', 'Retrabajo', 'Observadas'], barmode='stack', color_discrete_sequence=chart_bars, text_auto=True)
            fig_prod.update_layout(width=800, height=300, margin=dict(t=20, b=40, l=20, r=20))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile3:
                fig_prod.write_image(tmpfile3.name, engine="kaleido")
                add_image_safe(pdf, tmpfile3.name, w_mm=155, h_mm=58)
                os.remove(tmpfile3.name)
            
            pdf.ln(2)
            
            def dibujar_cabeza_prod():
                setup_table_header(pdf, theme_color)
                pdf.set_font("Arial", 'B', 8)
                pdf.cell(70, 5, "Codigo Producto", 1, 0, 'C', True)
                pdf.cell(30, 5, "Buenas", 1, 0, 'C', True)
                pdf.cell(30, 5, "Retrab.", 1, 0, 'C', True)
                pdf.cell(30, 5, "Observ.", 1, 1, 'C', True)
            
            maquinas_prod = sorted(df_prod_pdf_g['Máquina'].unique())
            
            for maq_p in maquinas_prod:
                df_m_prod = df_prod_pdf_g[df_prod_pdf_g['Máquina'] == maq_p].groupby('Código')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
                total_buenas = df_m_prod['Buenas'].sum()
                total_retrabajo = df_m_prod['Retrabajo'].sum()
                total_obs = df_m_prod['Observadas'].sum()
                total_piezas = total_buenas + total_retrabajo + total_obs
                
                check_space(pdf, 25)
                pdf.set_font("Arial", 'B', 9); pdf.set_text_color(*theme_color)
                pdf.cell(0, 5, clean_text(f"Top 5 Códigos Producidos - {maq_p} (Total: {int(total_piezas)} pzs)"), ln=True)
                
                dibujar_cabeza_prod()
                setup_table_row(pdf); pdf.set_font("Arial", '', 8)
                
                top5_prod = df_m_prod.sort_values('Buenas', ascending=False).head(5)
                for _, row in top5_prod.iterrows():
                    if pdf.get_y() > 265:
                        pdf.add_page(); dibujar_cabeza_prod(); setup_table_row(pdf); pdf.set_font("Arial", '', 8)
                    pdf.cell(70, 4.5, " " + clean_text(str(row['Código'])[:45]), 'B') 
                    pdf.cell(30, 4.5, str(int(row['Buenas'])), 'B', 0, 'C')
                    pdf.cell(30, 4.5, str(int(row['Retrabajo'])), 'B', 0, 'C')
                    pdf.cell(30, 4.5, str(int(row['Observadas'])), 'B', 1, 'C')
                pdf.ln(3) 

        # ==========================================
        # 7. ANÁLISIS INDIVIDUAL POR MÁQUINA
        # ==========================================
        maquinas_con_tiempo = []
        if not df_pdf_g.empty:
            for maq in sorted(df_pdf_g['Máquina'].unique()):
                df_maq_temp = df_pdf_g[df_pdf_g['Máquina'] == maq]
                t_total = df_maq_temp[df_maq_temp['Estado_Global'].isin(['Producción', 'Falla/Gestión', 'Parada Programada', 'Proyecto', 'Descanso'])]['Tiempo (Min)'].sum()
                if t_total > 0: maquinas_con_tiempo.append(maq)

        if maquinas_con_tiempo:
            pdf.add_page()
            pdf.set_link(links_detalle_grupo[g]) 
            pdf.set_font("Times", 'B', 16)
            pdf.set_text_color(*theme_color)
            pdf.cell(0, 10, clean_text(f"DESGLOSE DETALLADO POR MÁQUINA - GRUPO {g}"), ln=True, border='B')
            pdf.ln(5)

            for i, maq in enumerate(maquinas_con_tiempo):
                if p_tipo in ["Semanal", "Mensual"]: 
                    if i > 0: pdf.add_page() 
                else: 
                    check_space(pdf, 60)

                df_maq = df_pdf_g[df_pdf_g['Máquina'] == maq]
                t_prod = df_maq[df_maq['Estado_Global'] == 'Producción']['Tiempo (Min)'].sum()
                t_falla = df_maq[df_maq['Estado_Global'] == 'Falla/Gestión']['Tiempo (Min)'].sum()
                t_parada = df_maq[df_maq['Estado_Global'] == 'Parada Programada']['Tiempo (Min)'].sum()
                t_proy = df_maq[df_maq['Estado_Global'] == 'Proyecto']['Tiempo (Min)'].sum()
                t_desc = df_maq[df_maq['Estado_Global'] == 'Descanso']['Tiempo (Min)'].sum()

                pdf.set_font("Arial", 'B', 12); pdf.set_text_color(255, 255, 255); pdf.set_fill_color(*comp_color)
                pdf.cell(0, 8, clean_text(f"  MÁQUINA: {maq}"), border=0, ln=True, fill=True)
                pdf.set_font("Arial", 'I', 8); pdf.set_text_color(120, 120, 120); pdf.cell(0, 5, clean_text(f"  Grupo: {g}"), border=0, ln=True); pdf.ln(2)

                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                for col in ["Produccion", "Fallas/Gestion", "Paradas Prog.", "Descansos"]: pdf.cell(47.5, 6, col, 1, 0, 'C', True)
                pdf.ln(); setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                pdf.cell(47.5, 6, clean_text(mins_to_duration_str(t_prod)), 1, 0, 'C')
                pdf.cell(47.5, 6, clean_text(mins_to_duration_str(t_falla)), 1, 0, 'C')
                pdf.cell(47.5, 6, clean_text(mins_to_duration_str(t_parada)), 1, 0, 'C')
                pdf.cell(47.5, 6, clean_text(mins_to_duration_str(t_desc)), 1, 1, 'C'); pdf.ln(5)

                df_maq_fallas = df_maq[df_maq['Estado_Global'] == 'Falla/Gestión']
                if not df_maq_fallas.empty:
                    pdf.set_font("Arial", 'B', 10)
                    pdf.set_text_color(*comp_color)
                    
                    if p_tipo in ["Semanal", "Mensual"]:
                        pdf.cell(100, 6, clean_text("> Top Fallas (Detalle Completo):"), 0, 0, 'L')
                        pdf.cell(90, 6, clean_text("> Tendencia Temporal de Fallas (Minutos):"), 0, 1, 'L')
                        ancho_texto = 65
                    else:
                        pdf.cell(0, 6, clean_text("> Top Fallas (Detalle Completo):"), 0, 1, 'L')
                        ancho_texto = 140
                    
                    y_inicio_seccion = pdf.get_y()
                    
                    agg_f = df_maq_fallas.groupby('Detalle_Final')['Tiempo (Min)'].sum().reset_index().sort_values('Tiempo (Min)', ascending=False).head(8)
                    max_time = agg_f['Tiempo (Min)'].max() if not agg_f.empty else 1
                    
                    pdf.set_font("Arial", '', 7)
                    pdf.set_text_color(50, 50, 50)
                    
                    y_tabla = y_inicio_seccion
                    ancho_barra_max = 25 if p_tipo in ["Semanal", "Mensual"] else 30
                    
                    for _, row in agg_f.iterrows():
                        pdf.set_xy(10, y_tabla)
                        nombre_falla = clean_text(row['Detalle_Final'])
                        mins = row['Tiempo (Min)']
                        
                        y_antes = pdf.get_y()
                        pdf.multi_cell(ancho_texto, 3.5, nombre_falla, border=0, align='L')
                        y_despues = pdf.get_y()
                        
                        valor_barra = (mins / max_time) * ancho_barra_max
                        pdf.set_fill_color(*comp_color)
                        pdf.rect(10 + ancho_texto + 2, y_antes + 1, valor_barra, 2, 'F')
                        
                        pdf.set_xy(10 + ancho_texto + 2 + valor_barra + 1, y_antes)
                        pdf.set_font("Arial", 'B', 6)
                        pdf.cell(10, 4, f"{mins:.0f}m", 0, 0, 'L')
                        pdf.set_font("Arial", '', 7)
                        
                        y_tabla = max(y_despues, y_antes + 4) + 1 
                        if p_tipo in ["Semanal", "Mensual"] and y_tabla > y_inicio_seccion + 55: break

                    if p_tipo in ["Semanal", "Mensual"]:
                        trend_f = df_maq_fallas.groupby('Fecha_Filtro')['Tiempo (Min)'].sum().reset_index()
                        trend_f['Fecha_Filtro'] = pd.to_datetime(trend_f['Fecha_Filtro']).sort_values()
                        fig_line = px.line(trend_f, x='Fecha_Filtro', y='Tiempo (Min)', markers=True)
                        fig_line.update_traces(line_color=hex_comp, marker=dict(size=6, color=hex_theme))
                        fig_line.update_xaxes(tickformat="%d/%m")
                        fig_line.update_layout(height=220, width=400, margin=dict(t=10, b=30, l=40, r=10), plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Minutos")

                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as t2:
                            fig_line.write_image(t2.name, engine="kaleido")
                            pdf.image(t2.name, x=110, y=y_inicio_seccion, w=95)
                            os.remove(t2.name)
                        
                        pdf.set_y(max(y_tabla, y_inicio_seccion + 60))
                        pdf.ln(5)
                    else:
                        pdf.set_y(y_tabla + 5)

                # SMED Individual Máquina
                df_maq_paradas = df_maq[df_maq['Estado_Global'] == 'Parada Programada']
                if not df_maq_paradas.empty:
                    pdf.set_font("Arial", 'B', 10); pdf.set_text_color(*theme_color)
                    if p_tipo in ["Semanal", "Mensual"]:
                        pdf.cell(95, 6, clean_text("> SMED / Paradas Programadas:"), 0, 0, 'L')
                        pdf.cell(95, 6, clean_text("> Tendencia Temporal (Promedio en Minutos):"), 0, 1, 'L')
                    else:
                        pdf.cell(0, 6, clean_text("> SMED / Paradas Programadas:"), 0, 1, 'L')

                    y_smed = pdf.get_y()
                    res_p = df_maq_paradas.groupby('Detalle_Final').agg(C=('Tiempo (Min)', 'count'), T=('Tiempo (Min)', 'sum')).reset_index().sort_values('T', ascending=False).head(5)

                    if p_tipo in ["Semanal", "Mensual"]:
                        trend_p = df_maq_paradas.groupby(['Fecha_Filtro', 'Detalle_Final'])['Tiempo (Min)'].mean().reset_index()
                        trend_p['Fecha_Filtro'] = pd.to_datetime(trend_p['Fecha_Filtro']).sort_values()
                        trend_p['Detalle_Corto'] = trend_p['Detalle_Final'].apply(lambda x: str(x)[:25] + "..." if len(str(x)) > 25 else str(x))
                        top_5_eventos = res_p.head(5)['Detalle_Final'].tolist()
                        trend_p_filtrado = trend_p[trend_p['Detalle_Final'].isin(top_5_eventos)]
                        
                        fig_trend_p = px.line(trend_p_filtrado, x='Fecha_Filtro', y='Tiempo (Min)', color='Detalle_Corto', markers=True, color_discrete_sequence=px.colors.qualitative.Safe)
                        fig_trend_p.update_xaxes(tickformat="%d/%m")
                        fig_trend_p.update_layout(height=200, width=420, margin=dict(t=10, b=20, l=40, r=10), plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Promedio Min", showlegend=False)
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_trend_p:
                            fig_trend_p.write_image(tmp_trend_p.name, engine="kaleido")
                            pdf.image(tmp_trend_p.name, x=105, y=y_smed, w=95)
                            os.remove(tmp_trend_p.name)

                    pdf.set_y(y_smed + 2)
                    setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 8)
                    ancho_ev = 50 if p_tipo in ["Semanal", "Mensual"] else 110
                    ancho_col = 15 if p_tipo in ["Semanal", "Mensual"] else 25
                    
                    pdf.cell(ancho_ev, 6, "Evento", 1, 0, 'C', True)
                    pdf.cell(ancho_col, 6, "Cant.", 1, 0, 'C', True)
                    pdf.cell(ancho_col, 6, "Total", 1, 0, 'C', True)
                    pdf.cell(ancho_col, 6, "Prom.", 1, 1, 'C', True)
                    
                    setup_table_row(pdf); pdf.set_font("Arial", '', 7)
                    for _, rp in res_p.iterrows():
                        caracteres_max = 32 if p_tipo in ["Semanal", "Mensual"] else 70
                        pdf.cell(ancho_ev, 4.5, " " + clean_text(rp['Detalle_Final'])[:caracteres_max], 'B')
                        pdf.cell(ancho_col, 4.5, str(int(rp['C'])), 'B', 0, 'C')
                        pdf.cell(ancho_col, 4.5, f"{rp['T']:.0f}m", 'B', 0, 'C')
                        pdf.cell(ancho_col, 4.5, f"{rp['T']/rp['C']:.1f}m", 'B', 1, 'C')
                    
                    if p_tipo in ["Semanal", "Mensual"]:
                        pdf.set_y(max(pdf.get_y(), y_smed + 55) + 5)
                    else:
                        pdf.ln(5)

    # --- SECCIÓN PERFORMANCE DE OPERARIOS ---
    pdf.add_page(); pdf.set_link(link_perfo)
    print_section_title(pdf, "Performance de Operarios General", theme_color)
    if area.upper() == "ESTAMPADO" and override_estampado:
        pdf.set_font("Arial", 'B', 9); pdf.set_text_color(220, 20, 20)
        pdf.multi_cell(0, 5, clean_text("AVISO: Indicadores modificados manualmente. La tabla de performance individual se oculta por inconsistencia con datos originales."))
    else:
        if not op_target_df.empty:
            df_filt = op_target_df[op_target_df['Fábrica'].astype(str).str.contains(area, case=False, na=False)].copy()
            if df_filt.empty and not df_pdf.empty:
                ops_activos = []
                for op_list in df_pdf['Operador'].unique():
                    if pd.notna(op_list) and op_list != '-': ops_activos.extend([o.strip() for o in op_list.split('/')])
                df_filt = op_target_df[op_target_df['Operador'].isin(ops_activos)].copy()
                
            if not df_filt.empty:
                df_filt = df_filt.drop_duplicates(subset=['Operador']).copy()
                df_filt['PERFORMANCE'] = pd.to_numeric(df_filt['PERFORMANCE'], errors='coerce').fillna(0)
                df_filt = df_filt.sort_values('PERFORMANCE', ascending=False)
                
                operador_maquinas = {}
                if not df_pdf.empty:
                    for _, r in df_pdf.iterrows():
                        maq = str(r['Máquina']).strip()
                        ops = str(r['Operador']).split('/')
                        for o in ops:
                            o = o.strip()
                            if o and o != '-':
                                if o not in operador_maquinas: operador_maquinas[o] = set()
                                operador_maquinas[o].add(maq)

                setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 9)
                pdf.cell(50, 6, "Operador", 1, 0, 'C', True); pdf.cell(35, 6, "Fabrica", 1, 0, 'C', True)
                pdf.cell(85, 6, "Maquinas Operadas", 1, 0, 'C', True); pdf.cell(20, 6, "Perf.", 1, 1, 'C', True)

                setup_table_row(pdf); pdf.set_font("Arial", '', 9)
                for _, row in df_filt.iterrows():
                    perf_val = int(round(row['PERFORMANCE']))
                    op_name = clean_text(str(row['Operador'])).strip()
                    if 'usuario' in op_name.lower() or 'admin' in op_name.lower(): continue
                    
                    if pdf.get_y() > 270: 
                        pdf.add_page(); setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 9)
                        pdf.cell(50, 6, "Operador", 1, 0, 'C', True); pdf.cell(35, 6, "Fabrica", 1, 0, 'C', True)
                        pdf.cell(85, 6, "Maquinas Operadas", 1, 0, 'C', True); pdf.cell(20, 6, "Perf.", 1, 1, 'C', True)
                        setup_table_row(pdf); pdf.set_font("Arial", '', 9)

                    maq_set = operador_maquinas.get(op_name, set())
                    maq_str = ", ".join(sorted(list(maq_set))) if maq_set else "-"

                    pdf.cell(50, 5, " " + op_name[:28], 'B'); pdf.cell(35, 5, " " + clean_text(str(row['Fábrica'])[:18]), 'B')
                    pdf.cell(85, 5, " " + clean_text(maq_str[:50]), 'B')
                        
                    if perf_val >= 90: pdf.set_text_color(33, 195, 84) # Verde estricto
                    else: pdf.set_text_color(220, 20, 20) # Rojo estricto
                    
                    pdf.cell(20, 5, f"{perf_val}%", 'B', 1, 'C'); pdf.set_text_color(50, 50, 50)
                pdf.ln(10)

    # =========================================================================
    # SECCIÓN: TIEMPOS DE DESCANSO (BAÑO Y REFRIGERIO)
    # =========================================================================
    pdf.add_page()
    pdf.set_link(link_tiempos)
    print_section_title(pdf, "Tablas de Tiempos Acumulados de Descanso", theme_color)

    df_descansos = df_pdf[df_pdf['Estado_Global'] == 'Descanso']

    if not df_descansos.empty:
        def generar_tabla_descanso(titulo, palabras_clave, limite_minutos):
            mask = df_descansos['Detalle_Final'].astype(str).str.contains('|'.join(palabras_clave), case=False, na=False) | \
                   df_descansos['Nivel Evento 2'].astype(str).str.contains('|'.join(palabras_clave), case=False, na=False)
            df_tipo = df_descansos[mask]

            if not df_tipo.empty:
                pdf.set_font("Arial", 'B', 12)
                pdf.set_text_color(*theme_color)
                pdf.cell(0, 8, clean_text(titulo), ln=True)

                agg_desc = df_tipo.groupby('Operador').agg(Total_Min=('Tiempo (Min)', 'sum'), Cant_Veces=('Tiempo (Min)', 'count')).reset_index().sort_values('Total_Min', ascending=False)

                setup_table_header(pdf, theme_color)
                pdf.set_font("Arial", 'B', 9)
                pdf.cell(70, 6, "Operador", 1, 0, 'C', True)
                pdf.cell(30, 6, "Total Min", 1, 0, 'C', True)
                pdf.cell(30, 6, "Cant. Veces", 1, 0, 'C', True)
                pdf.cell(30, 6, "Promedio Min", 1, 1, 'C', True)

                setup_table_row(pdf)
                pdf.set_font("Arial", '', 9)
                for _, r in agg_desc.iterrows():
                    op_name = clean_text(str(r['Operador'])).strip()
                    if 'usuario' in op_name.lower() or 'admin' in op_name.lower() or op_name == '-': 
                        continue
                    
                    if pdf.get_y() > 270:
                        pdf.add_page(); setup_table_header(pdf, theme_color); pdf.set_font("Arial", 'B', 9)
                        pdf.cell(70, 6, "Operador", 1, 0, 'C', True); pdf.cell(30, 6, "Total Min", 1, 0, 'C', True)
                        pdf.cell(30, 6, "Cant. Veces", 1, 0, 'C', True); pdf.cell(30, 6, "Promedio Min", 1, 1, 'C', True)
                        setup_table_row(pdf); pdf.set_font("Arial", '', 9)

                    # Evaluar color Rojo según límite
                    is_over = False
                    if p_tipo == "Diario":
                        if r['Total_Min'] > limite_minutos: is_over = True
                    else:
                        if (r['Total_Min'] / r['Cant_Veces']) > limite_minutos: is_over = True
                        
                    pdf.set_text_color(50, 50, 50)
                    pdf.cell(70, 5, " " + op_name[:35], 'B')
                    
                    if is_over: pdf.set_text_color(220, 20, 20)
                    pdf.cell(30, 5, f"{r['Total_Min']:.1f}", 'B', 0, 'C')
                    
                    pdf.set_text_color(50, 50, 50)
                    pdf.cell(30, 5, str(int(r['Cant_Veces'])), 'B', 0, 'C')
                    
                    if is_over: pdf.set_text_color(220, 20, 20)
                    pdf.cell(30, 5, f"{(r['Total_Min'] / r['Cant_Veces']):.1f}", 'B', 1, 'C')
                    
                    pdf.set_text_color(50, 50, 50)
                pdf.ln(8)

        generar_tabla_descanso("Tiempo de Baño Acumulado", ['baño', 'bano'], limite_minutos=8)
        generar_tabla_descanso("Tiempo de Refrigerio Acumulado", ['refrigerio'], limite_minutos=17)

    else:
        pdf.set_font("Arial", '', 10)
        pdf.set_text_color(50, 50, 50)
        pdf.cell(0, 6, "No hay datos de descansos (baño o refrigerio) registrados para esta área en este período.", ln=True)

    # =========================================================================
    # ANEXO: AUDITORÍA DE AJUSTES MANUALES
    # =========================================================================
    if area.upper() == "ESTAMPADO" and override_estampado and df_metrics_ORIG is not None:
        pdf.add_page()
        print_section_title(pdf, "ANEXO: AUDITORÍA DE AJUSTES MANUALES", (220, 20, 20))
        pdf.set_font("Arial", '', 9)
        pdf.multi_cell(0, 5, clean_text("Este anexo detalla las diferencias entre los valores capturados automáticamente por el sistema Wiidem y los valores corregidos para este reporte."))
        pdf.ln(5)

        # 1. Auditoría Indicadores
        pdf.set_font("Arial", 'B', 10); pdf.cell(0, 7, "1. Comparativa de OEE e Indicadores", ln=True)
        escala_orig = 100 if not df_metrics_ORIG.empty and df_metrics_ORIG['OEE'].max() > 1.5 else 1
        maqs_est = df_metrics_pdf[df_metrics_pdf['Máquina'].apply(asignar_grupo_dinamico) == 'LÍNEAS ESTAMPADO']['Máquina'].unique()

        for maq in maqs_est:
            r_orig = df_metrics_ORIG[df_metrics_ORIG['Máquina'] == maq].iloc[0] if not df_metrics_ORIG[df_metrics_ORIG['Máquina'] == maq].empty else None
            r_nuevo = df_metrics_pdf[df_metrics_pdf['Máquina'] == maq].iloc[0] if not df_metrics_pdf[df_metrics_pdf['Máquina'] == maq].empty else None
            
            if r_orig is not None and r_nuevo is not None:
                diffs = []
                for k, lbl in [('OEE','OEE'),('DISPONIBILIDAD','Disp'),('PERFORMANCE','Perf'),('CALIDAD','Cal')]:
                    v_o = r_orig[k] if escala_orig == 100 else r_orig[k]*100
                    v_n = r_nuevo[k]*100 # El nuevo ya está estandarizado a 0.9X
                    if abs(v_n - v_o) > 0.01: diffs.append((lbl, v_o, v_n))
                
                if diffs:
                    pdf.set_font("Arial", 'B', 9); pdf.set_text_color(*theme_color); pdf.cell(0, 6, clean_text(f">> Máquina: {maq}"), ln=True)
                    setup_table_header(pdf, (220, 20, 20)); pdf.set_font("Arial", 'B', 8)
                    pdf.cell(40, 6, "Métrica", 1, 0, 'C', True); pdf.cell(40, 6, "Valor Original", 1, 0, 'C', True); pdf.cell(40, 6, "Valor Corregido", 1, 0, 'C', True); pdf.cell(40, 6, "Diferencia", 1, 1, 'C', True)
                    setup_table_row(pdf); pdf.set_font("Arial", '', 8)
                    for lbl, vo, vn in diffs:
                        pdf.cell(40, 5, " "+lbl, 1)
                        pdf.cell(40, 5, f"{vo:.2f}%", 1, 0, 'C')
                        pdf.cell(40, 5, f"{vn:.2f}%", 1, 0, 'C')
                        pdf.set_text_color(128,0,128) if (vn-vo)>0 else pdf.set_text_color(220,20,20)
                        pdf.cell(40, 5, f"{vn-vo:+.2f}%", 1, 1, 'C'); pdf.set_text_color(50,50,50)
                    pdf.ln(5)

        # 2. Auditoría Producción
        pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0, 0, 0); pdf.cell(0, 7, "2. Comparativa de Produccion (Piezas Buenas)", ln=True)
        setup_table_header(pdf, (220, 20, 20)); pdf.set_font("Arial", 'B', 8)
        pdf.cell(45, 6, "Maquina", 1, 0, 'C', True); pdf.cell(55, 6, "Codigo", 1, 0, 'C', True); pdf.cell(30, 6, "Orig.", 1, 0, 'C', True); pdf.cell(30, 6, "Corr.", 1, 0, 'C', True); pdf.cell(30, 6, "Dif.", 1, 1, 'C', True)
        setup_table_row(pdf); pdf.set_font("Arial", '', 8)
        p_o = df_prod_ORIG.groupby(['Máquina','Código'])['Buenas'].sum().reset_index()
        p_n = prod_target_df.groupby(['Máquina','Código'])['Buenas'].sum().reset_index()
        merged = pd.merge(p_o, p_n, on=['Máquina','Código'], suffixes=('_O','_N'))
        for _, r in merged[merged['Buenas_O'] != merged['Buenas_N']].iterrows():
            vo, vn = int(round(r['Buenas_O'])), int(round(r['Buenas_N']))
            pdf.cell(45, 5, " "+clean_text(r['Máquina'][:20]), 1); pdf.cell(55, 5, " "+clean_text(r['Código'][:30]), 1)
            pdf.cell(30, 5, str(vo), 1, 0, 'C'); pdf.cell(30, 5, str(vn), 1, 0, 'C')
            pdf.set_text_color(128,0,128) if (vn-vo)>0 else pdf.set_text_color(220,20,20)
            pdf.cell(30, 5, f"{vn-vo:+d}", 1, 1, 'C'); pdf.set_text_color(50,50,50)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 5.5. CORRECCIÓN MANUAL DE DATOS (ESTAMPADO)
# ==========================================
st.divider()

habilitar_edicion = False  # Bandera por defecto

def es_maquina_estampado(maq):
    maq_u = str(maq).strip().upper()
    if maq_u in MAQUINAS_MAP and MAQUINAS_MAP[maq_u] == 'LÍNEAS ESTAMPADO': return True
    if 'LINEA' in maq_u or 'LÍNEA' in maq_u: return True
    return False

with st.expander("🛠️ Corrección Manual de Datos - LÍNEAS ESTAMPADO"):
    st.markdown("Utilice esta sección si la producción o los indicadores no se cerraron a tiempo en el sistema y necesita **forzar los valores** para el reporte.")
    habilitar_edicion = st.toggle("Habilitar sobreescritura manual", value=False)
    
    if habilitar_edicion:
        col_ed1, col_ed2 = st.columns(2)
        
        df_met_editado = pd.DataFrame()
        df_prod_editado = pd.DataFrame()
        df_met_est_orig = pd.DataFrame()
        df_prod_est_orig = pd.DataFrame()
        
        with col_ed1:
            st.write("**1. Indicadores (Performance, Disp, Calidad)**")
            st.caption("Edite los valores. El OEE se recalculará automáticamente.")
            
            if not df_metrics.empty and 'Máquina' in df_metrics.columns:
                mask_met_est = df_metrics['Máquina'].apply(es_maquina_estampado)
                df_met_est_orig = df_metrics[mask_met_est][['Máquina', 'DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD', 'OEE']].copy()
                df_met_est = df_metrics[mask_met_est][['Máquina', 'DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD']].copy()
                
                df_met_editado = st.data_editor(
                    df_met_est, 
                    column_config={"Máquina": st.column_config.TextColumn("Máquina", disabled=True)}, 
                    hide_index=True, use_container_width=True, key="ed_met"
                )
            else:
                st.info("No hay datos de indicadores cargados para este período.")
            
        with col_ed2:
            st.write("**2. Producción por Código (Cantidades)**")
            st.caption("Ajuste las cantidades de piezas producidas.")
            
            if not pdf_df_prod_target.empty and 'Máquina' in pdf_df_prod_target.columns:
                mask_prod_est = pdf_df_prod_target['Máquina'].apply(es_maquina_estampado)
                df_prod_est_orig = pdf_df_prod_target[mask_prod_est][['Máquina', 'Código', 'Buenas', 'Retrabajo', 'Observadas']].copy()
                df_prod_est = df_prod_est_orig.copy()
                
                df_prod_editado = st.data_editor(
                    df_prod_est,
                    column_config={
                        "Máquina": st.column_config.TextColumn("Máquina", disabled=True),
                        "Código": st.column_config.TextColumn("Código", disabled=True)
                    },
                    hide_index=True, use_container_width=True, key="ed_prod"
                )
            else:
                st.info("No hay datos de producción cargados para este período.")

        # --- LÓGICA DE ACTUALIZACIÓN EN MEMORIA Y CÁLCULO DE IMPACTO ---
        if not df_met_editado.empty:
            escala_100 = False
            if not df_metrics.empty and df_metrics['DISPONIBILIDAD'].max() > 1.5:
                escala_100 = True

            comparacion_oee = []

            for _, row in df_met_editado.iterrows():
                maq = row['Máquina']
                idx = df_metrics[df_metrics['Máquina'] == maq].index
                
                if not idx.empty:
                    d, p, c = row['DISPONIBILIDAD'], row['PERFORMANCE'], row['CALIDAD']
                    
                    if escala_100:
                        d_val = d * 100 if (d <= 1.5 and d > 0) else d
                        p_val = p * 100 if (p <= 1.5 and p > 0) else p
                        c_val = c * 100 if (c <= 1.5 and c > 0) else c
                        oee_val = (d_val / 100.0) * (p_val / 100.0) * (c_val / 100.0) * 100.0
                    else:
                        d_val = d / 100.0 if d > 1.5 else d
                        p_val = p / 100.0 if p > 1.5 else p
                        c_val = c / 100.0 if c > 1.5 else c
                        oee_val = d_val * p_val * c_val

                    df_metrics.loc[idx, ['DISPONIBILIDAD', 'PERFORMANCE', 'CALIDAD', 'OEE']] = [d_val, p_val, c_val, oee_val]
                    
                    oee_orig = df_met_est_orig[df_met_est_orig['Máquina'] == maq]['OEE'].values[0]
                    if not escala_100: oee_orig = oee_orig * 100
                    oee_val_disp = oee_val if escala_100 else oee_val * 100
                    
                    if round(oee_orig, 2) != round(oee_val_disp, 2):
                        comparacion_oee.append({
                            "Máquina": maq,
                            "OEE Original": f"{oee_orig:.2f}%",
                            "OEE Corregido": f"{oee_val_disp:.2f}%",
                            "Diferencia": f"{oee_val_disp - oee_orig:+.2f}%"
                        })

        if not df_prod_editado.empty:
            mask_prod_global = pdf_df_prod_target['Máquina'].apply(es_maquina_estampado)
            pdf_df_prod_target = pdf_df_prod_target[~mask_prod_global] 
            pdf_df_prod_target = pd.concat([pdf_df_prod_target, df_prod_editado], ignore_index=True)
            
            prod_agrup = df_prod_editado.groupby('Máquina')[['Buenas', 'Retrabajo', 'Observadas']].sum().reset_index()
            for _, row in prod_agrup.iterrows():
                if not df_metrics.empty:
                    idx = df_metrics[df_metrics['Máquina'] == row['Máquina']].index
                    if not idx.empty:
                        df_metrics.loc[idx, ['Buenas', 'Retrabajo', 'Observadas']] = [row['Buenas'], row['Retrabajo'], row['Observadas']]

            comparacion_prod = []
            df_prod_orig_agrup = df_prod_est_orig.groupby(['Máquina', 'Código'])['Buenas'].sum().reset_index()
            df_prod_nuevo_agrup = df_prod_editado.groupby(['Máquina', 'Código'])['Buenas'].sum().reset_index()
            merged_prod = pd.merge(df_prod_orig_agrup, df_prod_nuevo_agrup, on=['Máquina', 'Código'], suffixes=('_Orig', '_Nueva'))
            merged_prod['Dif'] = merged_prod['Buenas_Nueva'] - merged_prod['Buenas_Orig']
            
            for _, r in merged_prod[merged_prod['Dif'] != 0].iterrows():
                comparacion_prod.append({
                    "Máquina": r['Máquina'],
                    "Código": r['Código'],
                    "Buenas Orig": int(round(r['Buenas_Orig'])),
                    "Buenas Nuevas": int(round(r['Buenas_Nueva'])),
                    "Diferencia": f"{int(round(r['Dif'])):+d}"
                })

        # --- MOSTRAR IMPACTO EN PANTALLA ---
        if comparacion_oee or comparacion_prod:
            st.markdown("### 📊 Impacto de las Correcciones")
            c_imp1, c_imp2 = st.columns(2)
            with c_imp1:
                st.write("**Cambios en OEE**")
                if comparacion_oee: st.dataframe(pd.DataFrame(comparacion_oee), hide_index=True, use_container_width=True)
                else: st.success("No se registraron cambios en los indicadores OEE.")
            with c_imp2:
                st.write("**Cambios en Piezas Buenas**")
                if comparacion_prod: st.dataframe(pd.DataFrame(comparacion_prod), hide_index=True, use_container_width=True)
                else: st.success("No se registraron cambios en la producción.")

# ==========================================
# 6. BOTONES DE EXPORTACIÓN EN PANTALLA
# ==========================================
with col_p3:
    st.write("**3. Generar y Descargar:**")
    
    if pdf_tipo == "Mensual":
        col_btn1, col_btn2, col_btn3 = st.columns(3)
    else:
        col_btn1, col_btn2 = st.columns(2)
        
    with col_btn1:
        if st.button("Reporte ESTAMPADO", use_container_width=True):
            with st.spinner("Generando PDF Estampado..."):
                try:
                    pdf_data = crear_pdf(
                        "Estampado", 
                        pdf_label, 
                        pdf_df_op_target, 
                        pdf_df_prod_target, 
                        df_raw, 
                        pdf_tipo, 
                        df_trend, 
                        df_metrics, 
                        df_horarios,
                        override_estampado=habilitar_edicion,
                        df_metrics_ORIG=df_metrics_ORIGINAL,
                        df_prod_ORIG=df_prod_ORIGINAL
                    )
                    st.download_button("Descargar Estampado", data=pdf_data, file_name=f"FAMMA_Estampado_{file_label}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
                    
    with col_btn2:
        if st.button("Reporte SOLDADURA", use_container_width=True):
            with st.spinner("Generando PDF Soldadura..."):
                try:
                    pdf_data = crear_pdf(
                        "Soldadura", 
                        pdf_label, 
                        pdf_df_op_target, 
                        pdf_df_prod_target, 
                        df_raw, 
                        pdf_tipo, 
                        df_trend, 
                        df_metrics, 
                        df_horarios, 
                        override_estampado=False
                    )
                    st.download_button("Descargar Soldadura", data=pdf_data, file_name=f"FAMMA_Soldadura_{file_label}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as e:
                    st.error(f"Error generando PDF: {e}")
                    
    if pdf_tipo == "Mensual":
        with col_btn3:
            if st.button("Resumen Ejecutivo", use_container_width=True):
                with st.spinner("Generando Resumen Ejecutivo Global..."):
                    try:
                        pdf_resumen = crear_pdf_resumen_ejecutivo(pdf_label, df_trend, df_metrics)
                        st.download_button("Descargar Resumen", data=pdf_resumen, file_name=f"FAMMA_Resumen_Planta_{file_label}.pdf", mime="application/pdf", use_container_width=True)
                    except Exception as e:
                        st.error(f"Error generando PDF: {e}")
