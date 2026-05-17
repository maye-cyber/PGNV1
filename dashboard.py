import streamlit as st
import pandas as pd
import plotly.express as px
import os
from fpdf import FPDF
import io
import plotly.io as pio

# Configuracion para exportacion de imagenes (Kaleido)
pio.renderers.default = "browser"
pio.templates.default = "plotly_white"

# CONFIGURACION DE LA PAGINA
st.set_page_config(
    page_title="Monitor PGN - Dashboard",
    layout="wide"
)

# PALETA DE COLORES CORPORATIVA PGN
COLOR_AZUL = "#003366"
COLOR_AMARILLO = "#FFCC00"
COLOR_NEGRO = "#212529"
COLOR_FONDO = "#F8F9FA"

# CSS PARA ESTILO INSTITUCIONAL
st.markdown(f"""
    <style>
    .stApp {{ background-color: {COLOR_FONDO}; }}
    .main-header {{
        background-color: {COLOR_AZUL}; padding: 25px; border-radius: 10px;
        color: white; text-align: center; margin-bottom: 25px; border-bottom: 5px solid {COLOR_AMARILLO};
    }}
    .main-header h1 {{ color: white !important; font-weight: bold; margin: 0; padding: 0; }}
    .main-header p {{ color: {COLOR_AMARILLO} !important; font-size: 1.3rem; margin: 5px 0 0 0; font-weight: 500; }}
    [data-testid="stMetricValue"] {{ color: {COLOR_AZUL} !important; font-size: 1.8rem !important; }}
    [data-testid="stMetricLabel"] {{ color: {COLOR_NEGRO} !important; font-weight: bold; }}
    div[data-testid="stPopover"] button {{ border-radius: 8px; border: 1px solid {COLOR_AZUL}; color: {COLOR_AZUL}; width: 100%; }}
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1>Monitoreo y Alarmas</h1>
        <p>Dashboard</p>
    </div>
""", unsafe_allow_html=True)

# SISTEMAS OFICIALES RECONOCIDOS
SISTEMAS_OFICIALES = sorted(["SIM", "X-ROAD", "SIRI", "PORTAL WEB", "INTRANET", "HOMINIS", "INSAP", "APPS PORTAL", "APPS INTRANET", "STRATEGOS", "SIAF", "GESTOR DOKUS", "DOKUS", "ITA", "SIGDEA PORTAL EMPLEADO", "SIGDEA SEDE ELECTRONICA", "SIGDEA", "ALFA", "APP MOVIL", "APPS EXTERNAS", "NUEVA SEDE ELECTRONICA", "IGA", "REGLA DE NEGOCIO", "SIM HOMINIS"], key=len, reverse=True)

def extraer_alarmas(valor):
    val = str(valor).strip().upper()
    if val in ['OK', 'O.K.', 'NAN', 'NONE', '', 'NA', 'NINGUNO']: return []
    for sep in [',', ';', '\n', ' Y ', ' E ', ' - ']: val = val.replace(sep, '|')
    partes = [p.strip() for p in val.split('|') if p.strip()]
    encontrados = []
    for p in partes:
        match_oficial = False
        for oficial in SISTEMAS_OFICIALES:
            if oficial in p: encontrados.append(oficial); match_oficial = True; break
        if not match_oficial and p not in ['OK', 'O.K.', 'ALERTA', 'ALARMA']:
            p_l = p.replace("SERVICIO ", "").strip()
            if len(p_l) > 1: encontrados.append(p_l)
    return list(dict.fromkeys([e.upper() for e in encontrados]))

@st.cache_data
def cargar_y_procesar_todo(archivo_path_or_buf):
    try:
        dict_hojas = pd.read_excel(archivo_path_or_buf, sheet_name=None)
        lista_final = []
        for mes_hoja, df_hoja in dict_hojas.items():
            df_hoja.columns = [c.strip() for c in df_hoja.columns]
            mes_nombre = str(mes_hoja).strip().capitalize()
            for col in ['Monitoreo fecha', 'Horario control']:
                if col in df_hoja.columns: df_hoja[col] = df_hoja[col].ffill()
            for _, row in df_hoja.iterrows():
                app_val = row.get('Aplicativo', 'OK')
                apps = extraer_alarmas(app_val)
                if not apps:
                    lista_final.append({**row.to_dict(), 'Sistemas': 'OK', 'Es_Alarma': False, 'Mes': mes_nombre})
                else:
                    for app in apps:
                        lista_final.append({**row.to_dict(), 'Sistemas': app, 'Es_Alarma': True, 'Mes': mes_nombre})
        df_final = pd.DataFrame(lista_final)
        if 'Horario control' in df_final.columns:
            def limpiar_hora(h):
                h_s = str(h).lower()
                if any(x in h_s for x in ['08:', '8 am', '8:00']): return '8 am'
                if any(x in h_s for x in ['12:', '12 pm', '12:00']): return '12 pm'
                if any(x in h_s for x in ['16:', '4 pm', '16:00']): return '4 pm'
                return None
            df_final['Horario_Normalizado'] = df_final['Horario control'].apply(limpiar_hora)
        return df_final
    except Exception as e:
        return None

def generar_pdf_completo(total_a, s_top, m_top, fig_s, fig_m, fig_h):
    from fpdf.enums import XPos, YPos
    from datetime import datetime
    
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # PORTADA / HEADER
    pdf.add_page()
    pdf.set_fill_color(0, 51, 102)  # Azul PGN
    pdf.rect(0, 0, 210, 50, 'F')
    
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 22)
    pdf.ln(10)
    pdf.cell(190, 15, "INFORME DE MONITOREO Y ALARMAS", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.set_font("helvetica", "B", 14)
    pdf.set_text_color(255, 204, 0) # Amarillo PGN
    pdf.cell(190, 10, f"Procuraduría General de la Nación - {datetime.now().strftime('%Y')}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    
    pdf.set_text_color(0, 0, 0)
    pdf.ln(25)
    
    # RESUMEN EJECUTIVO
    pdf.set_font("helvetica", "B", 16)
    pdf.set_draw_color(0, 51, 102)
    pdf.set_line_width(0.5)
    pdf.cell(190, 10, "1. RESUMEN EJECUTIVO", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    pdf.set_font("helvetica", "", 12)
    def c(t): return str(t).encode('latin-1', 'replace').decode('latin-1')
    
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(190, 10, c(f"Total de Incidencias Registradas: {total_a}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
    pdf.cell(190, 10, c(f"Sistema con Mayor Afectación: {s_top}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(190, 10, c(f"Mes con Mayor Actividad: {m_top}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
    pdf.cell(190, 10, c(f"Fecha de Generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.ln(10)
    
    # GRÁFICO 1: SISTEMAS
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(190, 10, "2. DISTRIBUCIÓN POR SISTEMA", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    try:
        img_s = fig_s.to_image(format="png", width=1000, height=600)
        pdf.image(io.BytesIO(img_s), x=15, w=180)
    except Exception as e:
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(190, 10, f"(Error al cargar gráfico: {str(e)})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # GRÁFICO 2: TENDENCIA MENSUAL
    pdf.add_page()
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(190, 10, "3. TENDENCIA MENSUAL", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    try:
        img_m = fig_m.to_image(format="png", width=1000, height=500)
        pdf.image(io.BytesIO(img_m), x=15, w=180)
    except Exception as e:
        pdf.cell(190, 10, "(Error al cargar gráfico)", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
    pdf.ln(10)
    
    # GRÁFICO 3: HORARIOS
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(190, 10, "4. ANÁLISIS POR HORARIO DE CONTROL", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    try:
        img_h = fig_h.to_image(format="png", width=1000, height=500)
        pdf.image(io.BytesIO(img_h), x=15, w=180)
    except Exception as e:
        pdf.cell(190, 10, "(Error al cargar gráfico)", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # FOOTER (Simplificado en cada página usando alias_nb_pages)
    pdf.alias_nb_pages()
    
    return pdf.output()
    
@st.dialog("Detalle de Incidencias", width="large")
def mostrar_detalle_casos(mes, df_casos):
    st.markdown(f"### Casos registrados en **{mes}**")
    columnas_mostrar = ['Sistemas', 'inconvenientes.', 'Comentario admin']
    # Filtrar solo si las columnas existen en el dataframe
    cols_existentes = [c for c in columnas_mostrar if c in df_casos.columns]
    
    if not df_casos.empty:
        # Renombrar columnas para mejor presentacion si es necesario
        df_display = df_casos[cols_existentes].copy()
        df_display.columns = [c.capitalize().replace('.', '') for c in df_display.columns]
        
        st.dataframe(
            df_display, 
            width="stretch", 
            hide_index=True,
            column_config={
                "Inconvenientes": st.column_config.TextColumn("Descripción del Problema", width="large"),
                "Comentario admin": st.column_config.TextColumn("Comentario Admin", width="medium")
            }
        )
    else:
        st.info(f"No hay detalles específicos de alarmas para el mes de {mes}.")
    
    if st.button("Cerrar"):
        st.rerun()

# FLUJO
ruta_excel = "c:/PRATICAS/DSB/monitoreos 2025.xlsx"

# Cargador de archivos discreto
with st.expander("Cargar nuevo archivo Excel"):
    archivo_subido = st.file_uploader("Sube un archivo para actualizar los datos:", type=["xlsx"])

data_source = archivo_subido if archivo_subido else (ruta_excel if os.path.exists(ruta_excel) else None)

if data_source:
    df = cargar_y_procesar_todo(data_source)
    if df is not None:
        # PESTAÑAS PRINCIPALES
        tab_met, tab_list = st.tabs(["Metricas y Graficos", "Listado Detallado"])
        
        # DEFINICION DE FILTROS (DENTRO DE LA LOGICA PERO SIN SIDEBAR)
        ord_m = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
        
        # Inicializar variables de seleccion para evitar errores antes de los popovers
        s_alarmas = sorted(df[df['Es_Alarma']]['Sistemas'].unique().tolist())
        m_disp = sorted(df['Mes'].unique().tolist(), key=lambda x: ord_m.index(x) if x in ord_m else 99)
        
        # Los filtros se renderizaran dentro de las pestañas o en un area comun
        # Para que afecten a todo, calculamos la data filtrada primero usando placeholders o widgets invisibles?
        # En Streamlit, si pones el widget en un lugar, su valor es global.
        
        # FILTROS DISCRETOS (En el area principal)
        with st.container():
            col_f1, col_f2, col_f3 = st.columns([1, 1, 3])
            with col_f1:
                with st.popover("Filtrar Sistemas"):
                    s_sel = st.multiselect("Seleccionar:", s_alarmas, default=s_alarmas)
            with col_f2:
                with st.popover("Filtrar Meses"):
                    m_sel = st.multiselect("Seleccionar:", m_disp, default=m_disp)
        
        df_f = df[df['Mes'].isin(m_sel)].copy()
        df_a = df_f[(df_f['Es_Alarma']) & (df_f['Sistemas'].isin(s_sel))].copy()
        total_a = len(df_a)
        if total_a > 0:
            top_s_val = df_a.groupby('Sistemas').size().idxmax(); count_s = df_a.groupby('Sistemas').size().max(); p_s = (count_s / total_a) * 100
            top_m_val = df_a.groupby('Mes').size().idxmax(); count_m = df_a.groupby('Mes').size().max(); p_m = (count_m / total_a) * 100
            
            # FIGURAS
            fig_s = px.bar(df_a.groupby('Sistemas').size().reset_index(name='C').sort_values('C', ascending=False), x='Sistemas', y='C', text_auto=True, color='Sistemas', title="Ranking de Incidencias por Sistema", color_discrete_sequence=px.colors.qualitative.Bold)
            fig_s.update_layout(showlegend=False, xaxis_tickangle=-45, template="plotly_white")
            fig_m = px.bar(df_f.groupby('Mes')['Es_Alarma'].sum().reset_index(name='C').sort_values('Mes', key=lambda x: x.map({v: i for i, v in enumerate(ord_m)})), x='Mes', y='C', text_auto=True, color='Mes', title="Incidencias por Mes", color_discrete_sequence=px.colors.sequential.YlOrRd)
            fig_m.update_layout(showlegend=False, template="plotly_white")
            fig_h = px.bar(df_a.dropna(subset=['Horario_Normalizado']).groupby('Horario_Normalizado').size().reset_index(name='C').sort_values('Horario_Normalizado', key=lambda x: x.map({'8 am': 1, '12 pm': 2, '4 pm': 3})), x='Horario_Normalizado', y='C', text_auto=True, color='Horario_Normalizado', color_discrete_map={'8 am': '#003366', '12 pm': '#FFCC00', '4 pm': '#212529'}, title="Incidencias por Horario", template="plotly_white")
            fig_h.update_layout(showlegend=False)

            with tab_met:
                c1, c2, c_pdf = st.columns([1, 1, 0.6])
                with c1: st.metric("SISTEMA MAS AFECTADO", top_s_val, delta=f"{count_s} ({p_s:.1f}%)", delta_color="inverse")
                with c2: st.metric("MES MAS CRITICO", top_m_val, delta=f"{count_m} ({p_m:.1f}%)", delta_color="inverse")
                with c_pdf:
                    # OPTIMIZACION: Generar PDF solo si el usuario lo solicita
                    if st.button("Preparar Reporte PDF", key="btn_prep_pdf"):
                        try:
                            with st.spinner("Generando PDF..."):
                                pdf_data = generar_pdf_completo(total_a, top_s_val, top_m_val, fig_s, fig_m, fig_h)
                                st.download_button(
                                    label="📥 Descargar Reporte PDF",
                                    data=bytes(pdf_data),
                                    file_name=f"Reporte_Monitoreo_PGN_{top_m_val}.pdf",
                                    mime="application/pdf",
                                    key="btn_download_pdf"
                                )
                                st.success("¡Reporte listo para descargar!")
                        except Exception as e:
                            st.error(f"Error al generar el PDF: {str(e)}")
                            st.info("Asegúrate de que no haya filtros vacíos.")
                
                st.markdown("---")
                cl, cr = st.columns(2)
                with cl: st.plotly_chart(fig_s, use_container_width=True)
                with cr: 
                    event_m = st.plotly_chart(fig_m, use_container_width=True, on_select="rerun")
                    # Logica de Drill-down
                    if event_m and "selection" in event_m and "points" in event_m["selection"] and len(event_m["selection"]["points"]) > 0:
                        sel_mes = event_m["selection"]["points"][0]["x"]
                        df_det = df_a[df_a['Mes'] == sel_mes]
                        mostrar_detalle_casos(sel_mes, df_det)
                
                st.markdown("---")
                st.plotly_chart(fig_h, use_container_width=True)

            with tab_list:
                st.markdown("### Listado Completo de Alarmas")
                st.info("Este listado muestra todos los casos que coinciden con los filtros seleccionados en la barra lateral.")
                
                columnas_vista = ['Mes', 'Sistemas', 'inconvenientes.', 'Comentario admin']
                df_view = df_a[[c for c in columnas_vista if c in df_a.columns]].copy()
                df_view.columns = [c.capitalize().replace('.', '') for c in df_view.columns]
                
                st.dataframe(
                    df_view, 
                    width="stretch", 
                    hide_index=True,
                    column_config={
                        "Inconvenientes": st.column_config.TextColumn("Descripción", width="large"),
                        "Comentario admin": st.column_config.TextColumn("Admin Info", width="medium")
                    }
                )
        else:
            st.warning("No hay alarmas registradas para los filtros seleccionados.")
else:
    st.info("Sube un archivo Excel para comenzar.")
