# ═══════════════════════════════════════════════════════════════════════
# COFEPRIS — Registros sanitarios de medicamentos (otorgados)
# Interfaz Streamlit sobre la lógica del notebook COFEPRIS_final_1.ipynb
# ═══════════════════════════════════════════════════════════════════════
import io
import re
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title='COFEPRIS — Registros de medicamentos | Bayer',
    layout='wide',
)

# ─────────────────── design system (tokens HSL de Cora) ─────────────────
# Tipografía: DM Sans (UI) + JetBrains Mono (datos). Sidebar negro puro.
# Cards rounded-xl con borde sutil. Labels uppercase con tracking.
DS = {
    'primary': 'hsl(220 90% 50%)',
    'foreground': 'hsl(220 15% 10%)',
    'muted': 'hsl(220 15% 96%)',
    'muted_fg': 'hsl(220 10% 40%)',
    'border': 'hsl(220 15% 88%)',
}
# Series del chart (hex requerido por Vega; par validado para CVD/contraste)
COLOR_FILTRO = '#0D59F2'   # primary
COLOR_TOTAL = '#0D9488'    # teal-600

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&family=JetBrains+Mono:wght@400;500;600&display=swap');

html, body, .stApp, p, span, div, input, textarea, button, label, h1, h2, h3, h4 {
    font-family: 'DM Sans', sans-serif;
}

/* ── encabezados ── */
h1 { font-size: 1.4rem !important; font-weight: 700 !important;
     letter-spacing: -0.01em; color: hsl(220 15% 10%); }
h3 { font-size: 0.72rem !important; text-transform: uppercase;
     letter-spacing: 0.08em; font-weight: 600 !important;
     color: hsl(220 10% 40%) !important;
     border-bottom: 1px solid hsl(220 15% 88% / .5);
     padding-bottom: 0.5rem !important; }

/* ── sidebar negro puro ── */
[data-testid="stSidebar"] { background: #000 !important; }
[data-testid="stSidebar"] * { color: #fff; }
[data-testid="stSidebar"] h2 {
    font-size: 0.72rem !important; text-transform: uppercase;
    letter-spacing: 0.08em; font-weight: 600 !important;
    color: rgba(255,255,255,.65) !important; }
[data-testid="stSidebar"] h3 {
    color: rgba(255,255,255,.65) !important;
    border-bottom: 1px solid rgba(255,255,255,.15); }
[data-testid="stSidebar"] label p { font-size: 12px; color: rgba(255,255,255,.85); }
[data-testid="stSidebar"] .stButton button {
    background: rgba(255,255,255,.10); border: 1px solid rgba(255,255,255,.20);
    border-radius: 0.75rem; color: #fff; font-size: 13px;
    transition: all .2s; }
[data-testid="stSidebar"] .stButton button:hover { background: rgba(255,255,255,.20); }
[data-testid="stSidebar"] input, [data-testid="stSidebar"] textarea {
    color: hsl(220 15% 10%) !important; }
[data-testid="stSidebar"] [data-baseweb="select"] span { color: hsl(220 15% 10%); }

/* ── métricas como cards ── */
[data-testid="stMetric"] {
    background: #fff; border: 1px solid hsl(220 15% 88% / .6);
    border-radius: 0.75rem; padding: 1rem;
    box-shadow: 0 1px 2px rgb(0 0 0 / .05); }
[data-testid="stMetricLabel"] p {
    text-transform: uppercase; font-size: 10px !important;
    letter-spacing: 0.08em; font-weight: 600;
    color: hsl(220 10% 40%) !important; }
[data-testid="stMetricValue"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 1.55rem !important; font-weight: 600;
    color: hsl(220 15% 10%); }

/* ── botones y descargas ── */
.stButton button, .stDownloadButton button {
    border-radius: 0.75rem !important; font-size: 13px;
    transition: all .2s; }
.stDownloadButton button {
    border: 1px solid hsl(220 15% 88%); background: #fff;
    color: hsl(220 15% 10%); }
.stDownloadButton button:hover {
    background: hsl(220 15% 96% / .6); border-color: hsl(220 90% 50%);
    color: hsl(220 90% 50%); }

/* ── expanders y captions ── */
[data-testid="stExpander"] {
    border: 1px solid hsl(220 15% 88% / .6) !important;
    border-radius: 0.75rem !important; background: #fff; }
[data-testid="stCaptionContainer"] p { font-size: 12px; color: hsl(220 10% 40%); }
</style>
""", unsafe_allow_html=True)

# ── logo de Bayer ──
# Usa el archivo local del repo si existe (assets/logo_bayer.png o .svg);
# si no, cae al logo oficial publicado en Wikimedia Commons.
LOGO = next((str(p) for p in (Path('assets/logo_bayer.png'),
                              Path('assets/logo_bayer.svg'),
                              Path('logo_bayer.png'),
                              Path('logo_bayer.svg')) if p.exists()), None)
LOGO_URL = 'https://upload.wikimedia.org/wikipedia/commons/f/f7/Logo_Bayer.svg'
try:
    st.logo(LOGO or LOGO_URL, size='large')
except Exception:
    pass

URL_XLSX = ('https://registros.cofepris.gob.mx/BRSDM/public/completo/'
            'Visor_Registros_Medicamentos.xlsx')
NAVEGADOR = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/122.0 Safari/537.36'),
    'Accept-Language': 'es-MX,es;q=0.9,en;q=0.8',
}

# Preset de oncología (el mismo del notebook)
PRESET_ONCO = {
    'atc_grupos': ['L01', 'L02'],
    'terminos': (r'CANCER|CARCINOMA|NEOPLAS|TUMOR|ONCOLOG|QUIMIOTERAP|LEUCEMIA|'
                 r'LINFOMA|MIELOMA|SARCOMA|MELANOMA|METASTAS|ANTINEOPLAS|CITOTOXIC'),
    'rescate_sin_atc': True,
}

# ─────────────────────── utilidades del notebook ───────────────────────
COLUMNAS = {
    'numero_registro': ['numero de registro'],
    'denominacion_distintiva': ['denominacion distintiva'],
    'fecha_expedicion_vigencia': ['fecha expedicion vigencia'],
    'fecha_expedicion_prorroga': ['fecha expedicion vigencia prorroga'],
    'estado': ['estado'],
    'forma_farmaceutica': ['forma farmaceutica'],
    'indicaciones_terapeuticas': ['indicaciones terapeuticas'],
    'contraindicaciones': ['contra indicaciones', 'contraindicaciones'],
    'vida_util': ['vida util'],
    'fraccion': ['fraccion'],
    'denominacion_generica': ['denominacion generica'],
    'via_administracion': ['vista administracion', 'via administracion'],
    'tipo_medicamento': ['tipo medicamento'],
    'presentacion': ['presentacion'],
    'cantidad': ['cantidad'],
    'sistema_organico': ['sistema organico'],
    'grupo_farmacologico': ['grupo farmacologico'],
    'subgrupo_farmacologico': ['subgrupo farmacologico'],
    'subgrupo_quimico': ['subgrupo quimico'],
    'sustancia_quimica': ['sustancia quimica'],
    'titular': ['titular'],
    'domicilio': ['domicilio'],
    'fabricantes_medicamentos': ['fabricantes medicamentos'],
    'fabricantes_farmacos': ['fabricantes farmacos'],
    'acondicionado_por': ['acondicionado por'],
    'acondicionado_extranjero': ['acondicionado extranjero'],
    'distribuidores': ['distribuidores'],
    'unidad_farmacovigilancia': ['unidad farmaco vigilancia'],
    'fecha_emision': ['fecha emision'],
}
MESES = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
         'julio': 7, 'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10,
         'noviembre': 11, 'diciembre': 12}
_RE_ES = re.compile(r'(\d{1,2})\s+DE\s+([A-ZÁÉÍÓÚÑ]+)\s+DE\s+(\d{4})', re.IGNORECASE)


def norm(s):
    s = unicodedata.normalize('NFKD', str(s))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip().lower()


def SAM(s):
    """Sin Acentos, Mayúsculas."""
    s = unicodedata.normalize('NFKD', str(s or ''))
    return ''.join(c for c in s if not unicodedata.combining(c)).upper()


def _fecha_es(t):
    m = _RE_ES.search(t or '')
    if not m:
        return pd.NaT
    d, mt, a = m.groups()
    mes = MESES.get(norm(mt))
    if not mes:
        return pd.NaT
    try:
        return pd.Timestamp(int(a), mes, int(d))
    except ValueError:
        return pd.NaT


def partir_vigencia(serie):
    """'29 DE NOVIEMBRE DE 2018 / 29 DE NOVIEMBRE DE 2023' -> (expedición, vigencia)."""
    iz, de = [], []
    for v in serie.fillna(''):
        p = str(v).split('/')
        iz.append(_fecha_es(p[0]))
        de.append(_fecha_es(p[1]) if len(p) > 1 else pd.NaT)
    return pd.DataFrame({'desde': iz, 'hasta': de}, index=serie.index)


def limpiar(otorg):
    """Renombra columnas con tolerancia a erratas y deriva las fechas."""
    disp, ren = {norm(c): c for c in otorg.columns}, {}
    for k, vs in COLUMNAS.items():
        o = next((disp[v] for v in vs if v in disp), None)
        if o:
            ren[o] = k
    otorg = otorg.rename(columns=ren)
    for c in otorg.columns:
        if otorg[c].dtype == object:
            otorg[c] = (otorg[c].astype(str)
                        .replace({'nan': None, 'None': None, '': None}).str.strip())
    v = partir_vigencia(otorg['fecha_expedicion_vigencia'])
    otorg['fecha_expedicion'] = v['desde']
    otorg['fecha_vigencia_hasta'] = v['hasta']
    p = partir_vigencia(otorg['fecha_expedicion_prorroga'])
    otorg['fecha_prorroga_hasta'] = p['hasta']
    otorg['vigencia_efectiva_hasta'] = (otorg['fecha_prorroga_hasta']
                                        .fillna(otorg['fecha_vigencia_hasta']))
    return otorg


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def descargar_base():
    """Descarga el XLSX oficial con reintentos. Regresa (df, error)."""
    ultimo_error = None
    for i in range(1, 4):
        try:
            r = requests.get(URL_XLSX, headers=NAVEGADOR, timeout=120)
            if r.status_code == 200 and len(r.content) > 1_000_000:
                df = pd.read_excel(io.BytesIO(r.content), dtype=str, engine='openpyxl')
                return limpiar(df), None
            ultimo_error = f'HTTP {r.status_code} ({len(r.content):,} bytes)'
        except Exception as e:
            ultimo_error = f'{type(e).__name__}: {e}'
        time.sleep(3 * i)
    return None, ultimo_error


@st.cache_data(show_spinner=False)
def cargar_subido(contenido: bytes):
    df = pd.read_excel(io.BytesIO(contenido), dtype=str, engine='openpyxl')
    return limpiar(df)


def aplicar_filtro(df, atc_grupos=(), atc_sistemas=(), terminos=None,
                   rescate_sin_atc=False, estados=(), tipos=(),
                   titular_contiene=None, fabricante_contiene=None,
                   desde=None, hasta=None, con_traza=True):
    """Subconjunto de df según los criterios dados. Vacío = no filtrar por eso."""
    d = df
    grupo = d['grupo_farmacologico'].fillna('').map(SAM).str.strip()
    sist = d['sistema_organico'].fillna('').map(SAM).str.strip()
    sin_atc = sist.isin(['', 'N/A']) | grupo.isin(['', 'N/A'])

    hay_terapeutico = bool(atc_grupos or atc_sistemas or terminos)
    if hay_terapeutico:
        sel = pd.Series(False, index=d.index)
        traza = pd.Series(None, index=d.index, dtype=object)

        if atc_grupos:
            m = grupo.str.startswith(tuple(atc_grupos))
            sel |= m
            traza = traza.mask(m & traza.isna(), 'ATC ' + grupo.str[:3])

        if atc_sistemas:
            m = sist.str.startswith(tuple(s + ' ' for s in atc_sistemas))
            sel |= m
            traza = traza.mask(m & traza.isna(), 'sistema ' + sist.str[:1])

        if terminos:
            txt = (d['denominacion_generica'].fillna('').map(SAM) + ' ' +
                   d['indicaciones_terapeuticas'].fillna('').map(SAM) + ' ' +
                   d['subgrupo_farmacologico'].fillna('').map(SAM))
            menciona = txt.str.contains(terminos, regex=True, na=False)

            if rescate_sin_atc:
                m = menciona & sin_atc & ~sel
                sel |= m
                traza = traza.mask(m & traza.isna(), 'texto (sin ATC)')
            elif not (atc_grupos or atc_sistemas):
                sel |= menciona
                traza = traza.mask(menciona & traza.isna(), 'texto')

        d = d[sel].copy()
        if con_traza:
            d['deteccion'] = traza[sel]
    else:
        d = d.copy()
        if con_traza:
            d['deteccion'] = 'sin filtro terapéutico'

    if estados:
        d = d[d['estado'].fillna('').str.upper().isin([e.upper() for e in estados])]
    if tipos:
        objetivo = [SAM(t) for t in tipos]
        d = d[d['tipo_medicamento'].fillna('').map(SAM).isin(objetivo)]
    if titular_contiene:
        d = d[d['titular'].fillna('').map(SAM)
              .str.contains(SAM(titular_contiene), regex=False)]
    if fabricante_contiene:
        d = d[d['fabricantes_medicamentos'].fillna('').map(SAM)
              .str.contains(SAM(fabricante_contiene), regex=False)]

    if desde:
        d = d[d['fecha_expedicion'] >= pd.Timestamp(desde)]
    if hasta:
        d = d[d['fecha_expedicion'] <= pd.Timestamp(hasta)]

    return d


def serie_mensual(d, etiqueta):
    return (d.dropna(subset=['fecha_expedicion'])
            .assign(mes=lambda x: x['fecha_expedicion'].dt.to_period('M').astype(str))
            .groupby('mes').size().rename(etiqueta))


# ═══════════════════════════ carga de datos ════════════════════════════
c_logo, c_tit = st.columns([1, 10], vertical_alignment='center')
with c_logo:
    st.image(LOGO or LOGO_URL, width=72)
with c_tit:
    st.title('COFEPRIS — Registros sanitarios de medicamentos')
    st.caption('Serie mensual de registros otorgados, con filtros configurables. '
               'Fuente: XLSX oficial del visor de COFEPRIS (actualizado a diario).')
st.markdown('<div style="border-bottom:1px solid hsl(220 15% 88% / .5);'
            'margin:0 0 1rem 0"></div>', unsafe_allow_html=True)

with st.spinner('Descargando la base oficial de COFEPRIS…'):
    otorg, error = descargar_base()

if otorg is None:
    st.error(f'No se pudo descargar el XLSX de COFEPRIS ({error}). '
             'Puedes subirlo manualmente:')
    st.markdown(f'Descárgalo de [aquí]({URL_XLSX}) y súbelo abajo.')
    archivo = st.file_uploader('Visor_Registros_Medicamentos.xlsx', type=['xlsx'])
    if archivo is None:
        st.stop()
    otorg = cargar_subido(archivo.getvalue())

# ═══════════════════════════ sidebar: filtros ══════════════════════════
DESC_ATC = {
    'L01': 'antineoplásicos', 'L02': 'terapia endocrina',
    'L03': 'inmunoestimulantes', 'L04': 'inmunosupresores',
    'C07': 'betabloqueadores', 'J01': 'antibacterianos',
}

grupos_disp = sorted({g[:3] for g in otorg['grupo_farmacologico'].dropna().map(SAM)
                      if re.match(r'^[A-Z]\d\d', g)})
sistemas_disp = sorted({s[0] for s in otorg['sistema_organico'].dropna().map(SAM)
                        if re.match(r'^[A-Z]\s', s)})
estados_disp = sorted(otorg['estado'].dropna().unique())
tipos_disp = sorted(otorg['tipo_medicamento'].dropna().unique())
anios_disp = sorted(otorg['fecha_expedicion'].dropna().dt.year.unique(), reverse=True)


def _aplicar_preset_onco():
    st.session_state['w_grupos'] = PRESET_ONCO['atc_grupos']
    st.session_state['w_terminos'] = PRESET_ONCO['terminos']
    st.session_state['w_rescate'] = PRESET_ONCO['rescate_sin_atc']
    st.session_state['w_sistemas'] = []


def _limpiar_filtros():
    st.session_state['w_grupos'] = []
    st.session_state['w_sistemas'] = []
    st.session_state['w_terminos'] = ''
    st.session_state['w_rescate'] = False
    st.session_state['w_estados'] = []
    st.session_state['w_tipos'] = []
    st.session_state['w_titular'] = ''
    st.session_state['w_fabricante'] = ''


with st.sidebar:
    st.header('Filtros')
    c1, c2 = st.columns(2)
    c1.button('Preset oncología', on_click=_aplicar_preset_onco,
              width='stretch',
              help='L01 + L02 + términos de cáncer + rescate sin ATC')
    c2.button('Limpiar', on_click=_limpiar_filtros, width='stretch')

    st.subheader('Terapéutico')
    atc_grupos = st.multiselect(
        'Grupos ATC (prefijo del Grupo Farmacológico)', grupos_disp, key='w_grupos',
        format_func=lambda g: f'{g} — {DESC_ATC[g]}' if g in DESC_ATC else g)
    atc_sistemas = st.multiselect(
        'Sistemas ATC (letra del Sistema Orgánico)', sistemas_disp, key='w_sistemas')
    terminos = st.text_area(
        'Búsqueda por texto (regex, SIN acentos, MAYÚSCULAS)', key='w_terminos',
        height=80,
        help='Busca en nombre genérico + indicaciones + subgrupo. '
             'Ej: CANCER|TUMOR|LEUCEMIA')
    rescate = st.checkbox(
        'Rescatar por texto los registros SIN clasificación ATC', key='w_rescate',
        help='Sin esto pierdes Tamoxifeno, Enzalutamida y ~109 más.')

    st.subheader('Administrativo')
    estados = st.multiselect('Estado', estados_disp, key='w_estados')
    tipos = st.multiselect('Tipo de medicamento', tipos_disp, key='w_tipos')
    titular = st.text_input('Titular contiene', key='w_titular',
                            placeholder='PFIZER')
    fabricante = st.text_input('Fabricante contiene', key='w_fabricante',
                               placeholder='ROCHE')

    st.subheader('Temporal (fecha de expedición)')
    anio = st.selectbox('Año', ['Toda la historia'] + [str(a) for a in anios_disp],
                        index=0, key='w_anio')
    cd, ch = st.columns(2)
    desde_txt = cd.text_input('Desde', placeholder='2020-01-01', key='w_desde')
    hasta_txt = ch.text_input('Hasta', placeholder='2026-12-31', key='w_hasta')

    st.subheader('Opciones')
    incluir_total = st.checkbox('Incluir serie de TODA la base (comparación)',
                                value=True, key='w_total')

# el año manda sobre desde/hasta si estos vienen vacíos
desde = desde_txt.strip() or None
hasta = hasta_txt.strip() or None
if anio != 'Toda la historia':
    desde = desde or f'{anio}-01-01'
    hasta = hasta or f'{anio}-12-31'

try:
    filtrado = aplicar_filtro(
        otorg,
        atc_grupos=tuple(atc_grupos),
        atc_sistemas=tuple(atc_sistemas),
        terminos=terminos.strip() or None,
        rescate_sin_atc=rescate,
        estados=tuple(estados),
        tipos=tuple(tipos),
        titular_contiene=titular.strip() or None,
        fabricante_contiene=fabricante.strip() or None,
        desde=desde, hasta=hasta,
    )
except re.error as e:
    st.error(f'La expresión regular de búsqueda no es válida: {e}')
    st.stop()
except ValueError as e:
    st.error(f'Fecha no válida (usa formato AAAA-MM-DD): {e}')
    st.stop()

# ═══════════════════════════ panel principal ═══════════════════════════
m1, m2, m3, m4 = st.columns(4)
m1.metric('Registros en la base', f'{len(otorg):,}')
m2.metric('Registros filtrados', f'{len(filtrado):,}')
vig = (filtrado['estado'].fillna('').str.upper() == 'VIGENTE').sum()
m3.metric('Vigentes (del filtro)', f'{vig:,}')
m4.metric('Titulares distintos', f"{filtrado['titular'].nunique():,}")

# ── serie mensual ──
series = [serie_mensual(filtrado, 'otorgados_filtro')]
if incluir_total:
    series.append(serie_mensual(otorg, 'otorgados_total'))
otorg_mes = pd.concat(series, axis=1).fillna(0).astype(int).sort_index()
otorg_mes.index.name = 'mes'
if desde or hasta:  # acotar la serie total al mismo rango para comparar
    idx = otorg_mes.index.to_series()
    if desde:
        otorg_mes = otorg_mes[idx >= pd.Timestamp(desde).strftime('%Y-%m')]
        idx = otorg_mes.index.to_series()
    if hasta:
        otorg_mes = otorg_mes[idx <= pd.Timestamp(hasta).strftime('%Y-%m')]

st.subheader('Serie mensual de otorgados')
if len(otorg_mes):
    NOMBRES = {'otorgados_filtro': 'Filtro', 'otorgados_total': 'Toda la base'}
    largo = otorg_mes.reset_index().melt('mes', var_name='serie',
                                         value_name='registros')
    largo['serie'] = largo['serie'].map(NOMBRES)
    largo['fecha'] = pd.to_datetime(largo['mes'])

    una_serie = largo['serie'].nunique() == 1
    escala = alt.Scale(domain=['Filtro', 'Toda la base'],
                       range=[COLOR_FILTRO, COLOR_TOTAL])
    color = (alt.value(COLOR_FILTRO) if una_serie
             else alt.Color('serie:N', scale=escala,
                            legend=alt.Legend(title=None, orient='top')))

    cerca = alt.selection_point(nearest=True, on='pointermove',
                                fields=['fecha'], empty=False)
    base = alt.Chart(largo).encode(
        x=alt.X('fecha:T', title=None,
                axis=alt.Axis(format='%b %Y', grid=False)),
        y=alt.Y('registros:Q', title='registros / mes',
                axis=alt.Axis(gridColor='#EEF0F3', tickCount=5)),
    )
    lineas = base.mark_line(strokeWidth=2).encode(color=color)
    puntos = base.mark_point(size=70, filled=True).encode(
        color=color,
        opacity=alt.condition(cerca, alt.value(1), alt.value(0)),
        tooltip=[alt.Tooltip('mes:N', title='Mes'),
                 alt.Tooltip('serie:N', title='Serie'),
                 alt.Tooltip('registros:Q', title='Registros')],
    ).add_params(cerca)
    regla = base.mark_rule(color='#DCDFE5').encode(
        opacity=alt.condition(cerca, alt.value(1), alt.value(0)))

    chart = (
        (lineas + regla + puntos)
        .properties(height=320)
        .configure(font='DM Sans')
        .configure_axis(domainOpacity=0, tickOpacity=0,
                        labelColor='#5C6370', titleColor='#5C6370',
                        labelFont='JetBrains Mono', labelFontSize=10,
                        titleFontSize=10, titleFontWeight=600)
        .configure_legend(labelColor='#16181D', labelFontSize=12)
        .configure_view(strokeOpacity=0)
    )
    st.altair_chart(chart, width='stretch')
    with st.expander('Ver tabla de la serie mensual'):
        st.dataframe(otorg_mes, width='stretch')
else:
    st.info('No hay registros con fecha de expedición en el rango elegido.')

# ── por qué entró cada registro (badges) ──
if 'deteccion' in filtrado and filtrado['deteccion'].notna().any():
    COLOR_PILL = {'ATC': '#0D59F2', 'sistema': '#0D9488',
                  'texto': '#7C3AED', 'sin filtro': '#5C6370'}

    def _pill(texto, n):
        c = next((v for k, v in COLOR_PILL.items() if texto.startswith(k)),
                 '#5C6370')
        return (f'<span style="display:inline-flex;align-items:center;gap:4px;'
                f'padding:2px 8px;border-radius:6px;font-size:11px;'
                f'font-weight:500;border:1px solid {c}40;color:{c};'
                f'background:{c}12;font-family:\'JetBrains Mono\',monospace">'
                f'{texto} · {n:,}</span>')

    conteo = filtrado['deteccion'].value_counts()
    st.markdown(
        '<div style="display:flex;gap:8px;flex-wrap:wrap;margin:0 0 1rem 0">'
        + ''.join(_pill(k, v) for k, v in conteo.items()) + '</div>',
        unsafe_allow_html=True)

# ── tabla de registros ──
st.subheader(f'Registros filtrados ({len(filtrado):,})')
COLS = ['numero_registro', 'denominacion_distintiva', 'denominacion_generica',
        'estado', 'deteccion', 'grupo_farmacologico', 'subgrupo_farmacologico',
        'sustancia_quimica', 'tipo_medicamento', 'forma_farmaceutica', 'titular',
        'fabricantes_medicamentos', 'fecha_expedicion', 'vigencia_efectiva_hasta',
        'indicaciones_terapeuticas']
vista = filtrado[[c for c in COLS if c in filtrado.columns]]
st.dataframe(vista, width='stretch', height=420)

# ═══════════════════════════ descargas ═════════════════════════════════
st.subheader('Descargar')
LIMITE = 32_767  # máximo de caracteres por celda en Excel


def para_excel(d):
    d = d.copy()
    for c in d.columns:
        if d[c].dtype == object:
            largo = d[c].astype(str).str.len() > LIMITE
            if largo.any():
                d.loc[largo, c] = d.loc[largo, c].astype(str).str[:LIMITE - 3] + '...'
    for c in d.select_dtypes(include=['datetimetz']).columns:
        d[c] = d[c].dt.tz_localize(None)
    return d


stamp = datetime.now().strftime('%Y%m%d')
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as xw:
    otorg_mes.reset_index().to_excel(xw, sheet_name='Serie_mensual', index=False)
    para_excel(vista).to_excel(xw, sheet_name='Otorgados_filtrados', index=False)

d1, d2 = st.columns(2)
d1.download_button(
    f'Excel — cofepris_{stamp}.xlsx', buf.getvalue(),
    file_name=f'cofepris_{stamp}.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    width='stretch')
d2.download_button(
    f'CSV completo (sin truncar) — cofepris_otorgados_{stamp}.csv',
    filtrado.to_csv(index=False, encoding='utf-8-sig'),
    file_name=f'cofepris_otorgados_{stamp}.csv',
    mime='text/csv', width='stretch')

st.caption('Los datos se descargan del visor oficial de COFEPRIS y se cachean '
           '6 horas. La parte de solicitudes (PDFs de gob.mx) no está incluida: '
           'la mayoría de los PDFs están cifrados; para esa serie, tramita el '
           'dato por Plataforma Nacional de Transparencia.')
