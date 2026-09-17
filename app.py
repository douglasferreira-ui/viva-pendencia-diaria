from __future__ import annotations

import hashlib
import html
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core import (
    build_excel_export,
    calculate_metrics,
    dashboard_table,
    filter_scope,
    filter_view,
    now_sp,
    priority_sort,
    read_excel_bytes,
    read_excel_path,
    read_history,
    today_sp,
    upsert_history,
)

from inadimplencia import (
    build_receivables_export,
    read_receivables_bytes,
    read_receivables_path,
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
LATEST_FILE = DATA_DIR / "latest.xlsx"
HISTORY_FILE = DATA_DIR / "history.csv"
INAD_FILE = DATA_DIR / "inadimplencia_latest.xlsx"
BASE_DISPLAY_DATETIME_FILE = DATA_DIR / "base_display_datetime.txt"
LOGO_FILE = APP_DIR / "assets" / "logo_viva.png"

BRAND_BLUE = "#2F4EA2"
BRAND_CYAN = "#34C5C9"
BRAND_NAVY = "#17336F"
BRAND_BG = "#F5F8FC"
RED = "#D93A3A"
ORANGE = "#F28C28"
YELLOW = "#E6B93B"
PURPLE = "#7B61FF"
GREEN = "#20A875"
GRAY = "#6B7280"

# Usuários e permissões. As senhas ficam apenas como hash SHA-256.
# USERANONI é o administrador: vê todas as unidades e é o único que pode publicar a base diária.
USER_ACCOUNTS = {
    "USERANONI": {
        "password_hash": "348c5a9ca2fbcf4acc8de1b53d7f24c8aee370f28de2c0bf83c68ab8119a1417",
        "label": "Acesso geral",
        "admin": True,
        "unit_aliases": [],
    },
    "VCVIVA": {
        "password_hash": "ccc9cd59afe7ed42c50807c22933464d2e6de021ed2d6378d9166aae51dd78cd",
        "label": "Vila Clementino",
        "admin": False,
        "unit_aliases": ["VILA CLEMENTINO"],
    },
    "GESTORA": {
        "password_hash": "bad0486ea85aab426009343df34b300472f498f40563851b1f32605e04297e65",
        "label": "Cuidar Bem",
        "admin": False,
        "unit_aliases": ["CUIDAR BEM"],
    },
    "VIVAINDEPENDENCIA": {
        "password_hash": "3ca37cb255ac637fa831668857bd7cd5e42da2ed900ea791211a2611f317b5b2",
        "label": "Independência",
        "admin": False,
        "unit_aliases": ["INDEPENDENCIA", "VIVA INDEPENDENCIA"],
    },
    "LIMEIRAVIVA": {
        "password_hash": "7032f39d8b30e56be4bd2cac6c3506f09da74a1dad18df3c31541ca7d5bbc149",
        "label": "Limeira Nova",
        "admin": False,
        "unit_aliases": ["LIMEIRA NOVA", "VIVA LIMEIRA NOVA"],
    },
    "RIO2VIVA": {
        "password_hash": "b28677b11a18046732eb3018da8649eeac2cbeb30b2240dd2727c2f4aa68db62",
        "label": "Rio Claro 2",
        "admin": False,
        "unit_aliases": ["RIO CLARO 2", "VIVA RIO CLARO 2", "RC2", "RC 2"],
    },
    "BARAOVIVA": {
        "password_hash": "bab41044abba49069b3839e8bea5f91353c19a545b3f6db029f2cd31566ccca7",
        "label": "Viva Barão",
        "admin": False,
        "unit_aliases": ["VIVA BARAO", "BARAO"],
    },
    "MATRIZVIVA": {
        "password_hash": "9afc87ff94a0ce1536edf429940996b6345c49d8e3396a63a7c4e9a42503b149",
        "label": "Viva Rio Claro Matriz",
        "admin": False,
        "unit_aliases": ["VIVA RIO CLARO MATRIZ", "RIO CLARO MATRIZ", "RC MATRIZ", "VIVA RIO CLARO", "RIO CLARO"],
    },
    "SCVIVA": {
        "password_hash": "5b5166f67bae34a0b7fc53522a34a38b605ef8308355b02b68ad1de706ab70c9",
        "label": "Viva São Carlos",
        "admin": False,
        "unit_aliases": ["VIVA SAO CARLOS", "SAO CARLOS"],
    },
}

st.set_page_config(
    page_title="Viva Cuidar | Pendência Diária",
    page_icon="💙",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------
# Estilo
# -----------------------------
st.markdown(
    f"""
    <style>
      :root {{ --blue:{BRAND_BLUE}; --cyan:{BRAND_CYAN}; --navy:{BRAND_NAVY}; }}
      .stApp {{ background: {BRAND_BG}; }}
      [data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid #E7ECF4; }}
      .block-container {{ padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1600px; }}
      h1,h2,h3 {{ color: {BRAND_NAVY}; letter-spacing: -0.02em; }}
      .app-header {{
        background: linear-gradient(135deg, #FFFFFF 0%, #F7FBFF 65%, #EDF9FA 100%);
        border: 1px solid #E4EAF3;
        border-radius: 22px;
        padding: 18px 22px;
        box-shadow: 0 10px 30px rgba(23,51,111,.07);
        margin-bottom: 16px;
      }}
      .header-title {{ font-size: 1.7rem; font-weight: 800; color:{BRAND_NAVY}; margin:0; }}
      .header-sub {{ color:#667085; margin-top:5px; font-size:.96rem; }}
      .status-badge {{
        display:inline-block; padding:6px 10px; border-radius:999px;
        background:#EAF8F8; color:#11767A; font-weight:700; font-size:.78rem;
      }}
      .kpi-card {{
        background:#FFFFFF; border:1px solid #E7ECF4; border-radius:18px; padding:16px 17px;
        box-shadow:0 7px 20px rgba(23,51,111,.055); min-height:118px;
      }}
      .kpi-label {{ color:#667085; font-size:.78rem; font-weight:700; text-transform:uppercase; letter-spacing:.04em; }}
      .kpi-value {{ color:{BRAND_NAVY}; font-size:1.85rem; line-height:1.1; font-weight:850; margin-top:8px; }}
      .kpi-foot {{ color:#8490A3; font-size:.78rem; margin-top:7px; }}
      .radar-card {{
        border-radius:18px; padding:16px 18px; color:#FFFFFF; min-height:120px;
        box-shadow:0 8px 24px rgba(15,23,42,.12);
      }}
      .radar-label {{ font-size:.82rem; font-weight:800; text-transform:uppercase; opacity:.92; }}
      .radar-value {{ font-size:2.05rem; font-weight:900; margin-top:8px; line-height:1; }}
      .radar-foot {{ font-size:.78rem; opacity:.9; margin-top:9px; }}
      .section-card {{
        background:#FFFFFF; border:1px solid #E7ECF4; border-radius:20px;
        padding:16px 16px 8px 16px; box-shadow:0 7px 20px rgba(23,51,111,.045);
      }}
      .login-wrap {{ max-width:560px; margin:4vh auto 0 auto; background:#FFFFFF; border:1px solid #E7ECF4;
        border-radius:24px; padding:28px; box-shadow:0 18px 48px rgba(23,51,111,.12); }}
      .login-logo-box {{ background:linear-gradient(135deg, #FFFFFF 0%, #F8FBFF 55%, #ECFBFC 100%); border:1px solid #E7ECF4; border-radius:22px; padding:18px 14px; margin-bottom:10px; }}
      .login-title {{ text-align:center; color:{BRAND_NAVY}; font-size:1.65rem; font-weight:850; margin:8px 0 2px; }}
      .login-sub {{ text-align:center; color:#718096; margin-bottom:18px; }}
      div[data-testid="stMetric"] {{ background:#FFFFFF; border:1px solid #E7ECF4; border-radius:16px; padding:12px; }}
      .stButton > button {{ border-radius:12px; font-weight:750; }}
      .stDownloadButton > button {{ border-radius:12px; font-weight:750; }}
      [data-testid="stDataFrame"] {{ border-radius:14px; overflow:hidden; }}
      .small-note {{ color:#7A8699; font-size:.78rem; }}
      .danger-note {{ background:#FFF1F1; border:1px solid #FFD9D9; color:#9B1C1C; border-radius:14px; padding:10px 12px; }}
      .money-card {{ background:#FFFFFF; border:1px solid #E7ECF4; border-radius:18px; padding:16px 17px; box-shadow:0 7px 20px rgba(23,51,111,.055); min-height:118px; }}
      .money-card .label {{ color:#667085; font-size:.76rem; font-weight:800; text-transform:uppercase; letter-spacing:.04em; }}
      .money-card .value {{ color:{BRAND_NAVY}; font-size:1.55rem; line-height:1.1; font-weight:850; margin-top:8px; }}
      .money-card .foot {{ color:#8490A3; font-size:.78rem; margin-top:7px; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def _secret(name: str, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def _normalize_key(value: str) -> str:
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.split())


def get_account(username: str):
    return USER_ACCOUNTS.get(_normalize_key(username))


def password_ok(account: dict, password: str) -> bool:
    digest = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return digest == account.get("password_hash", "")


def current_account() -> dict:
    username = st.session_state.get("login_user", "")
    return get_account(username) or {}


def is_admin() -> bool:
    return bool(current_account().get("admin", False))


def restrict_dataframe_for_user(df: pd.DataFrame) -> pd.DataFrame:
    account = current_account()
    if account.get("admin", False):
        return df.copy()

    aliases = {_normalize_key(x) for x in account.get("unit_aliases", [])}
    if not aliases:
        return df.iloc[0:0].copy()

    unit_keys = df["LOCAL"].astype(str).map(_normalize_key)
    return df[unit_keys.isin(aliases)].copy()


def login_screen():
    st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
    if LOGO_FILE.exists():
        st.markdown('<div class="login-logo-box">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([0.55, 2.4, 0.55])
        with c2:
            st.image(str(LOGO_FILE), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="login-title">Pendência Diária</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="login-sub">Painel operacional Viva Cuidar • acesso restrito</div>',
        unsafe_allow_html=True,
    )
    with st.form("login_form"):
        user = st.text_input("Usuário", placeholder="Digite seu usuário")
        password = st.text_input("Senha", type="password", placeholder="Digite sua senha")
        submitted = st.form_submit_button("Entrar", type="primary", use_container_width=True)
    if submitted:
        account = get_account(user)
        if account and password_ok(account, password):
            st.session_state["authenticated"] = True
            st.session_state["login_user"] = _normalize_key(user)
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")
    st.markdown('</div>', unsafe_allow_html=True)


def require_login():
    if not st.session_state.get("authenticated", False):
        login_screen()
        st.stop()


def load_latest_data():
    if not LATEST_FILE.exists():
        return None, None
    try:
        df = read_excel_path(LATEST_FILE)
        updated = pd.Timestamp(LATEST_FILE.stat().st_mtime, unit="s", tz="UTC").tz_convert("America/Sao_Paulo")
        return df, updated
    except Exception as exc:
        st.error(f"Não foi possível abrir a base salva: {exc}")
        return None, None


def publish_upload(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    df = read_excel_bytes(file_bytes)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LATEST_FILE.write_bytes(file_bytes)
    upsert_history(df, HISTORY_FILE)
    return df


def load_manual_base_datetime(fallback=None):
    if BASE_DISPLAY_DATETIME_FILE.exists():
        try:
            raw = BASE_DISPLAY_DATETIME_FILE.read_text(encoding="utf-8").strip()
            if raw:
                return pd.Timestamp(raw)
        except Exception:
            pass
    return fallback


def save_manual_base_datetime(date_value, time_value):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    combined = datetime.combine(date_value, time_value)
    BASE_DISPLAY_DATETIME_FILE.write_text(combined.isoformat(), encoding="utf-8")
    return pd.Timestamp(combined)


def clear_manual_base_datetime():
    if BASE_DISPLAY_DATETIME_FILE.exists():
        BASE_DISPLAY_DATETIME_FILE.unlink()


def load_inadimplencia_data():
    if not INAD_FILE.exists():
        return None, None
    try:
        df = read_receivables_path(INAD_FILE)
        updated = pd.Timestamp(INAD_FILE.stat().st_mtime, unit="s", tz="UTC").tz_convert("America/Sao_Paulo")
        return df, updated
    except Exception as exc:
        st.error(f"Não foi possível abrir a base de inadimplência: {exc}")
        return None, None


def publish_inadimplencia_upload(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    df = read_receivables_bytes(file_bytes)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    INAD_FILE.write_bytes(file_bytes)
    return df


def brl(value: float) -> str:
    text = f"R$ {float(value):,.2f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def money_card(label: str, value: str, foot: str = ""):
    return f"""
      <div class="money-card">
        <div class="label">{html.escape(label)}</div>
        <div class="value">{html.escape(value)}</div>
        <div class="foot">{html.escape(foot)}</div>
      </div>
    """


def kpi_card(label: str, value: str, foot: str = ""):
    return f"""
      <div class="kpi-card">
        <div class="kpi-label">{html.escape(label)}</div>
        <div class="kpi-value">{html.escape(value)}</div>
        <div class="kpi-foot">{html.escape(foot)}</div>
      </div>
    """


def radar_card(label: str, value: int, foot: str, color: str):
    return f"""
      <div class="radar-card" style="background:{color};">
        <div class="radar-label">{html.escape(label)}</div>
        <div class="radar-value">{value}</div>
        <div class="radar-foot">{html.escape(foot)}</div>
      </div>
    """


def empty_chart(message: str):
    fig = go.Figure()
    fig.add_annotation(text=message, x=0.5, y=0.5, showarrow=False, font=dict(size=15, color="#7B8798"))
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=40, b=10), xaxis_visible=False, yaxis_visible=False)
    return fig


def chart_layout(fig, title: str, height: int = 340):
    fig.update_layout(
        title=dict(text=title, font=dict(size=16, color=BRAND_NAVY)),
        height=height,
        margin=dict(l=18, r=18, t=55, b=20),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Arial", color="#475467"),
        legend_title_text="",
    )
    return fig


def render_table(df: pd.DataFrame, key: str, height: int = 480):
    table = dashboard_table(df)
    if table.empty:
        st.info("Nenhum caso encontrado para este recorte.")
        return
    st.dataframe(
        table,
        use_container_width=True,
        hide_index=True,
        height=height,
        key=key,
        column_config={
            "DATA FINAL": st.column_config.DateColumn("Data final", format="DD/MM/YYYY"),
            "Dias Atraso": st.column_config.NumberColumn("Dias atraso", format="%d"),
            "Dias até Vencer": st.column_config.NumberColumn("Dias até vencer", format="%d"),
            "LOCAL": st.column_config.TextColumn("Unidade"),
        },
    )


require_login()

# -----------------------------
# Sidebar: atualização + filtros
# -----------------------------
with st.sidebar:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), use_container_width=True)
    st.markdown("### Central de Pendências")

    account = current_account()
    if account.get("admin", False):
        st.caption("Acesso geral • todas as unidades")

        with st.expander("📤 Atualizar base do dia", expanded=not LATEST_FILE.exists()):
            uploaded = st.file_uploader("Envie a planilha .xlsx", type=["xlsx"], key="daily_upload")
            if uploaded is not None:
                if st.button("Publicar atualização", type="primary", use_container_width=True):
                    try:
                        with st.spinner("Validando e publicando a base..."):
                            new_df = publish_upload(uploaded)
                        st.success(f"Base publicada com {len(new_df):,} linhas.".replace(",", "."))
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
            st.caption("A nova base substitui a anterior e atualiza o histórico do dia.")

        current_display_dt = load_manual_base_datetime()
        if current_display_dt is None:
            if LATEST_FILE.exists():
                current_display_dt = pd.Timestamp(LATEST_FILE.stat().st_mtime, unit="s", tz="UTC").tz_convert("America/Sao_Paulo").tz_localize(None)
            else:
                current_display_dt = pd.Timestamp.now()

        with st.expander("🗓️ Data exibida da base", expanded=False):
            st.caption("Somente o administrador altera a data/hora mostrada no cabeçalho para todos os usuários.")
            manual_date = st.date_input(
                "Data",
                value=current_display_dt.date(),
                format="DD/MM/YYYY",
                key="manual_base_date",
            )
            manual_time = st.time_input(
                "Hora",
                value=current_display_dt.time().replace(microsecond=0),
                step=60,
                key="manual_base_time",
            )
            cdt1, cdt2 = st.columns(2)
            with cdt1:
                if st.button("Salvar data/hora", type="primary", use_container_width=True, key="save_manual_base_dt"):
                    save_manual_base_datetime(manual_date, manual_time)
                    st.success("Data/hora atualizada.")
                    st.rerun()
            with cdt2:
                if st.button("Usar data automática", use_container_width=True, key="clear_manual_base_dt"):
                    clear_manual_base_datetime()
                    st.success("Data automática restaurada.")
                    st.rerun()

        with st.expander("💰 Atualizar inadimplência", expanded=False):
            inad_upload = st.file_uploader(
                "Envie o MonitorReceber .xlsx",
                type=["xlsx"],
                key="inadimplencia_upload",
            )
            if inad_upload is not None:
                if st.button("Publicar inadimplência", type="primary", use_container_width=True, key="publish_inad"): 
                    try:
                        with st.spinner("Validando e publicando a inadimplência..."):
                            inad_df_new = publish_inadimplencia_upload(inad_upload)
                        abertos = inad_df_new[inad_df_new["É inadimplente"]]
                        st.success(f"Inadimplência publicada: {len(abertos)} título(s) em aberto.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
            st.caption("Somente o administrador pode atualizar esta base.")
    else:
        st.caption(f"Unidade: {account.get('label', 'Acesso restrito')}")
        st.info("A atualização da base é exclusiva do usuário administrador.")

    st.divider()

    if st.button("Sair", use_container_width=True):
        st.session_state.clear()
        st.rerun()


df, updated_at = load_latest_data()
if df is None:
    st.markdown(
        f"""
        <div class="app-header">
          <div class="header-title">Pendência Diária</div>
          <div class="header-sub">Envie a planilha do dia no menu lateral para iniciar o painel.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if is_admin():
        st.info("Nenhuma base foi publicada ainda. Use **Atualizar base do dia** na barra lateral.")
    else:
        st.info("A base do dia ainda não foi publicada pelo administrador.")
    st.stop()

# Aplica a permissão de unidade antes de qualquer filtro ou KPI.
df = restrict_dataframe_for_user(df)

if df.empty and not is_admin():
    account = current_account()
    st.warning(
        f"A unidade **{account.get('label', '')}** não foi encontrada na base publicada. "
        "Peça ao administrador para conferir o nome da unidade na planilha do dia."
    )

# filtros estruturais
with st.sidebar:
    st.markdown("### Filtros")
    all_units = sorted(df["LOCAL"].dropna().unique().tolist())
    if is_admin():
        unidades = st.multiselect("Unidade", all_units, placeholder="Todas as unidades")
    else:
        unidades = []

    all_groups = sorted(df["Grupo"].dropna().unique().tolist())
    grupos = st.multiselect("Grupo", all_groups, placeholder="Todos os grupos")

    busca = st.text_input("Busca rápida", placeholder="Cliente, contrato, equipamento...")
    visao = st.selectbox(
        "Situação da lista",
        ["Todos", "Atrasados", "Vence hoje", "Vence em 5 dias", "Locados em dia", "Disponíveis", "Com OS"],
        index=0,
    )
    faixas = st.multiselect("Faixa de atraso", ["24h", "48h", "72h+"], placeholder="Todas")

scope = filter_scope(df, unidades=unidades, grupos=grupos, search=busca)
view_df = filter_view(scope, visao=visao, faixas=faixas)
metrics = calculate_metrics(scope)

display_updated_at = load_manual_base_datetime(updated_at)
updated_text = display_updated_at.strftime("%d/%m/%Y às %H:%M") if display_updated_at is not None else "—"

# -----------------------------
# Cabeçalho
# -----------------------------
st.markdown('<div class="app-header">', unsafe_allow_html=True)
h1, h2 = st.columns([1.1, 4.2])
with h1:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), use_container_width=True)
with h2:
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap; height:100%;">
          <div>
            <div class="header-title">Pendência Diária • Operação de Locações</div>
            <div class="header-sub">Olhe de manhã, identifique o crítico e priorize a ação do dia.</div>
          </div>
          <div style="text-align:right;">
            <div class="status-badge">● Base atualizada</div>
            <div class="header-sub">{updated_text}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# KPIs
# -----------------------------
cols = st.columns(8)
kpis = [
    ("Equipamentos", f"{metrics.total}", "No recorte selecionado"),
    ("Ocupação", f"{metrics.ocupacao_pct:.1f}%", f"{metrics.locados} locados"),
    ("Locados", f"{metrics.locados}", "Contratos / com contrato"),
    ("Atrasados", f"{metrics.atrasados}", "Data final já vencida"),
    ("Locados em atraso", f"{metrics.atraso_locados_pct:.1f}%", "Atrasados ÷ locados"),
    ("Vence hoje", f"{metrics.vence_hoje}", "Ação imediata"),
    ("Vence em 5 dias", f"{metrics.vence_5_dias}", "Próximos 5 dias"),
    ("Disponíveis", f"{metrics.disponiveis}", "Prontos para locação"),
]
for col, (label, value, foot) in zip(cols, kpis):
    with col:
        st.markdown(kpi_card(label, value, foot), unsafe_allow_html=True)

st.markdown("#### Radar de atraso")
r1, r2, r3, r4 = st.columns(4)
with r1:
    st.markdown(radar_card("24 horas", metrics.faixa_24h, f"Acumulado ≥24h: {metrics.acumulado_24h}", YELLOW), unsafe_allow_html=True)
with r2:
    st.markdown(radar_card("48 horas", metrics.faixa_48h, f"Acumulado ≥48h: {metrics.acumulado_48h}", ORANGE), unsafe_allow_html=True)
with r3:
    st.markdown(radar_card("72 horas +", metrics.faixa_72h, f"Críticos acumulados: {metrics.acumulado_72h}", RED), unsafe_allow_html=True)
with r4:
    st.markdown(radar_card("Atraso acumulado", metrics.atrasados, "Todos os contratos vencidos", BRAND_NAVY), unsafe_allow_html=True)

# -----------------------------
# Gráficos principais
# -----------------------------
st.markdown("#### Visão executiva")
g1, g2 = st.columns(2)

with g1:
    occ_df = pd.DataFrame(
        {
            "Situação": ["Locados", "Disponíveis", "Outros"],
            "Quantidade": [metrics.locados, metrics.disponiveis, metrics.outros],
        }
    )
    occ_df = occ_df[occ_df["Quantidade"] > 0]
    if occ_df.empty:
        fig = empty_chart("Sem dados para exibir")
    else:
        fig = px.pie(
            occ_df,
            names="Situação",
            values="Quantidade",
            hole=0.68,
            color="Situação",
            color_discrete_map={"Locados": BRAND_BLUE, "Disponíveis": BRAND_CYAN, "Outros": "#B8C2D1"},
        )
        fig.update_traces(textposition="inside", textinfo="percent+label")
        fig.add_annotation(text=f"<b>{metrics.ocupacao_pct:.1f}%</b><br>ocupação", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color=BRAND_NAVY))
    st.plotly_chart(chart_layout(fig, "Ocupação do parque"), use_container_width=True, config={"displayModeBar": False})

with g2:
    aging = pd.DataFrame(
        {
            "Faixa": ["24h", "48h", "72h+"],
            "Casos": [metrics.faixa_24h, metrics.faixa_48h, metrics.faixa_72h],
        }
    )
    fig = px.bar(
        aging,
        x="Faixa",
        y="Casos",
        text="Casos",
        color="Faixa",
        color_discrete_map={"24h": YELLOW, "48h": ORANGE, "72h+": RED},
    )
    fig.update_traces(textposition="outside")
    fig.update_yaxes(rangemode="tozero", gridcolor="#EEF2F7")
    fig.update_xaxes(title=None)
    fig.update_layout(showlegend=False)
    st.plotly_chart(chart_layout(fig, "Atrasos por criticidade"), use_container_width=True, config={"displayModeBar": False})

# Linha 2 de gráficos
c1, c2 = st.columns(2)
with c1:
    today = pd.Timestamp(today_sp())
    future = scope[scope["É Locado"] & scope["DATA FINAL"].notna()].copy()
    future = future[(future["DATA FINAL"] >= today) & (future["DATA FINAL"] <= today + pd.Timedelta(days=7))]
    if future.empty:
        fig = empty_chart("Sem vencimentos nos próximos 7 dias")
    else:
        due = future.groupby("DATA FINAL", as_index=False).size().rename(columns={"size": "Casos"})
        due["Data"] = due["DATA FINAL"].dt.strftime("%d/%m")
        fig = px.bar(due, x="Data", y="Casos", text="Casos")
        fig.update_traces(marker_color=BRAND_BLUE, textposition="outside")
        fig.update_yaxes(rangemode="tozero", gridcolor="#EEF2F7")
    st.plotly_chart(chart_layout(fig, "Vencimentos nos próximos 7 dias"), use_container_width=True, config={"displayModeBar": False})

with c2:
    overdue_groups = scope[scope["Atrasado"]].copy()
    if overdue_groups.empty:
        fig = empty_chart("Nenhum atraso no recorte")
    else:
        top = overdue_groups.groupby("Grupo", as_index=False).size().rename(columns={"size": "Casos"})
        top = top.sort_values("Casos", ascending=True).tail(10)
        fig = px.bar(top, x="Casos", y="Grupo", orientation="h", text="Casos")
        fig.update_traces(marker_color=RED, textposition="outside")
        fig.update_xaxes(rangemode="tozero", gridcolor="#EEF2F7")
    st.plotly_chart(chart_layout(fig, "Top grupos com atraso"), use_container_width=True, config={"displayModeBar": False})

# Unidade: útil quando a base tiver mais de uma loja
if scope["LOCAL"].nunique() > 1:
    unit_rows = []
    for unidade, part in scope.groupby("LOCAL"):
        m = calculate_metrics(part)
        unit_rows.append({"Unidade": unidade, "Ocupação %": m.ocupacao_pct, "Atraso %": m.atraso_locados_pct, "Atrasados": m.atrasados})
    unit_df = pd.DataFrame(unit_rows).sort_values("Ocupação %", ascending=False)
    fig = px.bar(
        unit_df,
        x="Unidade",
        y=["Ocupação %", "Atraso %"],
        barmode="group",
        labels={"value": "%", "variable": "Indicador"},
    )
    fig.update_yaxes(range=[0, max(100, float(unit_df[["Ocupação %", "Atraso %"]].max().max()) + 10)], gridcolor="#EEF2F7")
    st.plotly_chart(chart_layout(fig, "Ocupação e atraso por unidade", 380), use_container_width=True, config={"displayModeBar": False})

# -----------------------------
# Casos operacionais
# -----------------------------
st.markdown("#### Fila operacional")
filter_caption = f"Exibindo **{len(view_df)}** caso(s) após os filtros da lista. Os KPIs acima usam o recorte de unidade/grupo/busca."
st.caption(filter_caption)

critical = scope[scope["Atrasado"] & (scope["Dias Atraso"] >= 3)]
attention = scope[scope["Atrasado"] & scope["Dias Atraso"].isin([1, 2])]
next5 = scope[scope["Vence em 5 dias"] | scope["Vence Hoje"]]
locados = scope[scope["É Locado"]]

tab_names = [
    "📋 Lista filtrada",
    "🔴 Críticos 72h+",
    "🟠 24h / 48h",
    "🟣 Vence até 5 dias",
    "💙 Todos os locados",
    "📊 Resumo por grupo",
    "📈 Histórico",
]
if is_admin():
    tab_names.append("💰 Inadimplência")

_tabs = st.tabs(tab_names)
tab_all, tab_crit, tab_24_48, tab_next, tab_loc, tab_summary, tab_history = _tabs[:7]
tab_inad = _tabs[7] if is_admin() else None

with tab_all:
    render_table(view_df, "table_all")

with tab_crit:
    st.caption("Ordenado do maior atraso para o menor.")
    render_table(priority_sort(critical), "table_critical")

with tab_24_48:
    render_table(priority_sort(attention), "table_attention")

with tab_next:
    render_table(priority_sort(next5), "table_next5")

with tab_loc:
    render_table(priority_sort(locados), "table_locados")

with tab_summary:
    if scope.empty:
        st.info("Sem dados para resumir.")
    else:
        group_rows = []
        for group, part in scope.groupby("Grupo"):
            m = calculate_metrics(part)
            group_rows.append(
                {
                    "Grupo": group,
                    "Total": m.total,
                    "Locados": m.locados,
                    "Disponíveis": m.disponiveis,
                    "Atrasados": m.atrasados,
                    "Vence hoje": m.vence_hoje,
                    "Vence 5 dias": m.vence_5_dias,
                    "Ocupação %": round(m.ocupacao_pct, 1),
                    "Atraso locados %": round(m.atraso_locados_pct, 1),
                }
            )
        group_df = pd.DataFrame(group_rows).sort_values(["Atrasados", "Locados"], ascending=[False, False])
        st.dataframe(
            group_df,
            use_container_width=True,
            hide_index=True,
            height=520,
            column_config={
                "Ocupação %": st.column_config.ProgressColumn("Ocupação %", min_value=0, max_value=100, format="%.1f%%"),
                "Atraso locados %": st.column_config.ProgressColumn("Atraso locados %", min_value=0, max_value=100, format="%.1f%%"),
            },
        )

with tab_history:
    hist = read_history(HISTORY_FILE)
    allowed_history_units = sorted(df["LOCAL"].dropna().astype(str).unique().tolist())

    if not is_admin():
        allowed_keys = {_normalize_key(x) for x in allowed_history_units}
        if not hist.empty and "unidade" in hist.columns:
            hist = hist[hist["unidade"].astype(str).map(_normalize_key).isin(allowed_keys)]

    if hist.empty or hist["data"].nunique() < 2:
        st.info("O histórico começa a ganhar tendência depois de pelo menos 2 dias de atualização.")
        if not hist.empty:
            st.dataframe(hist.sort_values(["data", "unidade"], ascending=[False, True]), use_container_width=True, hide_index=True)
    else:
        hist_scope = hist.copy()
        if is_admin():
            if unidades:
                hist_scope = hist_scope[hist_scope["unidade"].isin(unidades)]
            else:
                hist_scope = hist_scope[hist_scope["unidade"] == "TODAS"]
        # Para usuários de loja, o histórico já foi limitado à unidade autorizada acima.

        if hist_scope.empty:
            st.info("Sem histórico para a unidade selecionada.")
        else:
            hist_scope["Data"] = hist_scope["data"].dt.strftime("%d/%m")
            color_field = "unidade" if (is_admin() and unidades) or (not is_admin() and hist_scope["unidade"].nunique() > 1) else None
            fig = px.line(
                hist_scope,
                x="Data",
                y="ocupacao_pct",
                color=color_field,
                markers=True,
                labels={"ocupacao_pct": "Ocupação %", "unidade": "Unidade"},
            )
            fig.update_yaxes(range=[0, 100], gridcolor="#EEF2F7")
            st.plotly_chart(chart_layout(fig, "Evolução da ocupação"), use_container_width=True, config={"displayModeBar": False})

            fig2 = px.line(
                hist_scope,
                x="Data",
                y="atrasados",
                color=color_field,
                markers=True,
                labels={"atrasados": "Atrasados", "unidade": "Unidade"},
            )
            fig2.update_yaxes(rangemode="tozero", gridcolor="#EEF2F7")
            st.plotly_chart(chart_layout(fig2, "Evolução dos atrasos"), use_container_width=True, config={"displayModeBar": False})


# -----------------------------
# Inadimplência — exclusivo administrador
# -----------------------------
if is_admin() and tab_inad is not None:
    with tab_inad:
        inad_df, inad_updated = load_inadimplencia_data()
        if inad_df is None:
            st.info(
                "Nenhuma base de inadimplência foi publicada. Use **Atualizar inadimplência** no menu lateral e envie o arquivo MonitorReceber."
            )
        else:
            inad_updated_text = inad_updated.strftime("%d/%m/%Y às %H:%M") if inad_updated is not None else "—"
            st.markdown("### Inadimplência")
            st.caption(f"Valores vencidos que continuam em aberto para baixa • Base atualizada em {inad_updated_text}")

            inad_open = inad_df[inad_df["É inadimplente"]].copy()

            if inad_open.empty:
                st.success("Não há títulos vencidos em aberto na base publicada.")
            else:
                # -------------------------
                # Filtros administrativos
                # -------------------------
                f1, f2, f3, f4 = st.columns([1.15, 1, 1, 1.55])
                with f1:
                    inad_units = sorted(inad_open["Local"].dropna().astype(str).unique().tolist())
                    inad_selected_units = st.multiselect(
                        "Local",
                        inad_units,
                        placeholder="Todas as unidades",
                        key="inad_units_v3",
                    )
                with f2:
                    origem_options = ["Todos"] + [x for x in ["Venda", "Locação", "Intermediação", "Não identificado"] if x in inad_open["Origem"].unique()]
                    inad_origem = st.selectbox(
                        "Origem",
                        origem_options,
                        index=0,
                        key="inad_origem_v3",
                    )
                with f3:
                    inad_faturamento = st.selectbox(
                        "Faturamento",
                        ["Todos", "Faturado", "Não faturado"],
                        index=0,
                        key="inad_faturamento_v3",
                    )
                with f4:
                    inad_search = st.text_input(
                        "Buscar",
                        placeholder="Cliente, contrato, pedido, fatura ou documento...",
                        key="inad_search_v3",
                    )

                f5, f6 = st.columns([1, 1])
                with f5:
                    inad_periodo = st.selectbox(
                        "Período do vencimento",
                        ["Todos", "Últimos 30 dias", "Últimos 60 dias", "Últimos 90 dias", "Últimos 120 dias"],
                        index=0,
                        key="inad_periodo_v3",
                    )
                with f6:
                    inad_faixa = st.selectbox(
                        "Faixa de atraso",
                        ["Todos", "24h", "48h", "72h+"],
                        index=0,
                        key="inad_faixa_v3",
                    )

                inad_filtered = inad_open.copy()

                if inad_selected_units:
                    inad_filtered = inad_filtered[inad_filtered["Local"].isin(inad_selected_units)]

                if inad_origem != "Todos":
                    inad_filtered = inad_filtered[inad_filtered["Origem"] == inad_origem]

                if inad_faturamento != "Todos":
                    inad_filtered = inad_filtered[inad_filtered["Status faturamento"] == inad_faturamento]

                # Atalhos de data: considera a data de vencimento até hoje.
                if inad_periodo != "Todos":
                    dias_periodo = int(inad_periodo.split()[1])
                    limite = pd.Timestamp(today_sp()) - pd.Timedelta(days=dias_periodo)
                    inad_filtered = inad_filtered[
                        inad_filtered["Vencimento"].notna()
                        & (inad_filtered["Vencimento"] >= limite)
                        & (inad_filtered["Vencimento"] <= pd.Timestamp(today_sp()))
                    ]

                if inad_faixa == "24h":
                    inad_filtered = inad_filtered[inad_filtered["Dias em atraso"] == 1]
                elif inad_faixa == "48h":
                    inad_filtered = inad_filtered[inad_filtered["Dias em atraso"] == 2]
                elif inad_faixa == "72h+":
                    inad_filtered = inad_filtered[inad_filtered["Dias em atraso"] >= 3]

                if inad_search.strip():
                    term = _normalize_key(inad_search)
                    h = pd.Series("", index=inad_filtered.index, dtype=str)
                    for col in ["Nome Fantasia", "Contrato", "Documento", "Local", "Origem", "Status faturamento"]:
                        if col in inad_filtered.columns:
                            h = h + " | " + inad_filtered[col].fillna("").astype(str)
                    for col in ["Fatura", "Pedido"]:
                        if col in inad_filtered.columns:
                            h = h + " | " + inad_filtered[col].fillna(0).astype(str)
                    inad_filtered = inad_filtered[h.map(_normalize_key).str.contains(term, na=False, regex=False)]

                # -------------------------
                # KPIs do recorte
                # -------------------------
                total_aberto = float(inad_filtered["Em aberto"].sum())
                titulos = int(len(inad_filtered))
                clientes = int(inad_filtered["Nome Fantasia"].replace("", pd.NA).dropna().nunique())
                ticket = (total_aberto / titulos) if titulos else 0.0
                maior_atraso = int(inad_filtered["Dias em atraso"].max()) if titulos else 0

                venda_df = inad_filtered[inad_filtered["Origem"] == "Venda"]
                locacao_df = inad_filtered[inad_filtered["Origem"] == "Locação"]
                interm_df = inad_filtered[inad_filtered["Origem"] == "Intermediação"]
                nao_fat_df = inad_filtered[inad_filtered["Status faturamento"] == "Não faturado"]

                m1, m2, m3, m4, m5 = st.columns(5)
                with m1:
                    st.markdown(money_card("Total em aberto", brl(total_aberto), f"{titulos} título(s) vencido(s)"), unsafe_allow_html=True)
                with m2:
                    st.markdown(money_card("Venda em aberto", brl(venda_df["Em aberto"].sum()), f"{len(venda_df)} título(s)"), unsafe_allow_html=True)
                with m3:
                    st.markdown(money_card("Locação em aberto", brl(locacao_df["Em aberto"].sum()), f"{len(locacao_df)} título(s)"), unsafe_allow_html=True)
                with m4:
                    st.markdown(money_card("Intermediação", brl(interm_df["Em aberto"].sum()), f"{len(interm_df)} título(s)"), unsafe_allow_html=True)
                with m5:
                    st.markdown(money_card("Não faturado", brl(nao_fat_df["Em aberto"].sum()), f"{len(nao_fat_df)} título(s)"), unsafe_allow_html=True)

                m6, m7, m8, m9 = st.columns(4)
                with m6:
                    st.markdown(money_card("Clientes inadimplentes", str(clientes), "Clientes distintos no recorte"), unsafe_allow_html=True)
                with m7:
                    st.markdown(money_card("Títulos em aberto", str(titulos), "Pendências para baixa"), unsafe_allow_html=True)
                with m8:
                    st.markdown(money_card("Ticket médio", brl(ticket), "Valor médio por título"), unsafe_allow_html=True)
                with m9:
                    st.markdown(money_card("Maior atraso", f"{maior_atraso} dias", "Maior vencimento em aberto"), unsafe_allow_html=True)

                # Radar sempre respeita TODOS os filtros aplicados.
                r24 = inad_filtered[inad_filtered["Dias em atraso"] == 1]
                r48 = inad_filtered[inad_filtered["Dias em atraso"] == 2]
                r72 = inad_filtered[inad_filtered["Dias em atraso"] >= 3]
                rc1, rc2, rc3, rc4 = st.columns(4)
                with rc1:
                    st.markdown(radar_card("24 horas", len(r24), brl(r24["Em aberto"].sum()), YELLOW), unsafe_allow_html=True)
                with rc2:
                    st.markdown(radar_card("48 horas", len(r48), brl(r48["Em aberto"].sum()), ORANGE), unsafe_allow_html=True)
                with rc3:
                    st.markdown(radar_card("72 horas +", len(r72), brl(r72["Em aberto"].sum()), RED), unsafe_allow_html=True)
                with rc4:
                    st.markdown(radar_card("Total vencido", len(inad_filtered), brl(inad_filtered["Em aberto"].sum()), BRAND_NAVY), unsafe_allow_html=True)

                # -------------------------
                # Gráficos
                # -------------------------
                g1, g2 = st.columns(2)
                with g1:
                    by_origin = (
                        inad_filtered.groupby("Origem", as_index=False)["Em aberto"].sum()
                        .sort_values("Em aberto", ascending=True)
                    )
                    if by_origin.empty:
                        fig = empty_chart("Sem valores para o filtro atual")
                    else:
                        fig = px.bar(by_origin, x="Em aberto", y="Origem", orientation="h", text="Em aberto")
                        fig.update_traces(marker_color=BRAND_BLUE, texttemplate="R$ %{text:,.2f}", textposition="outside")
                        fig.update_xaxes(gridcolor="#EEF2F7")
                    st.plotly_chart(chart_layout(fig, "Em aberto por origem", 350), use_container_width=True, config={"displayModeBar": False})

                with g2:
                    by_fat = (
                        inad_filtered.groupby("Status faturamento", as_index=False)["Em aberto"].sum()
                        .sort_values("Em aberto", ascending=True)
                    )
                    if by_fat.empty:
                        fig = empty_chart("Sem valores para o filtro atual")
                    else:
                        fig = px.bar(by_fat, x="Em aberto", y="Status faturamento", orientation="h", text="Em aberto")
                        fig.update_traces(marker_color=ORANGE, texttemplate="R$ %{text:,.2f}", textposition="outside")
                        fig.update_xaxes(gridcolor="#EEF2F7")
                    st.plotly_chart(chart_layout(fig, "Faturado x não faturado", 350), use_container_width=True, config={"displayModeBar": False})

                g3, g4 = st.columns(2)
                with g3:
                    by_unit = (
                        inad_filtered.groupby("Local", as_index=False)["Em aberto"].sum()
                        .sort_values("Em aberto", ascending=True)
                    )
                    if by_unit.empty:
                        fig = empty_chart("Sem valores para o filtro atual")
                    else:
                        fig = px.bar(by_unit, x="Em aberto", y="Local", orientation="h", text="Em aberto")
                        fig.update_traces(marker_color=BRAND_CYAN, texttemplate="R$ %{text:,.2f}", textposition="outside")
                        fig.update_xaxes(gridcolor="#EEF2F7")
                    st.plotly_chart(chart_layout(fig, "Inadimplência por unidade", 370), use_container_width=True, config={"displayModeBar": False})

                with g4:
                    aging_base = inad_filtered.copy()
                    def _aging_label(d):
                        if pd.isna(d):
                            return "Sem data"
                        d = int(d)
                        if d == 1:
                            return "24h"
                        if d == 2:
                            return "48h"
                        if d <= 30:
                            return "3 a 30 dias"
                        if d <= 60:
                            return "31 a 60 dias"
                        if d <= 90:
                            return "61 a 90 dias"
                        if d <= 120:
                            return "91 a 120 dias"
                        return "121 dias+"
                    aging_base["Faixa operacional"] = aging_base["Dias em atraso"].map(_aging_label)
                    order = ["24h", "48h", "3 a 30 dias", "31 a 60 dias", "61 a 90 dias", "91 a 120 dias", "121 dias+"]
                    aging = aging_base.groupby("Faixa operacional", as_index=False)["Em aberto"].sum()
                    aging["Faixa operacional"] = pd.Categorical(aging["Faixa operacional"], categories=order, ordered=True)
                    aging = aging.sort_values("Faixa operacional")
                    if aging.empty:
                        fig = empty_chart("Sem valores para o filtro atual")
                    else:
                        fig = px.bar(aging, x="Faixa operacional", y="Em aberto", text="Em aberto")
                        fig.update_traces(marker_color=RED, texttemplate="R$ %{text:,.2f}", textposition="outside")
                        fig.update_yaxes(gridcolor="#EEF2F7")
                        fig.update_xaxes(title=None)
                    st.plotly_chart(chart_layout(fig, "Faixa de atraso: 30 / 60 / 90 / 120 dias", 370), use_container_width=True, config={"displayModeBar": False})

                g5, g6 = st.columns(2)
                with g5:
                    top_clients = (
                        inad_filtered.groupby("Nome Fantasia", as_index=False)["Em aberto"].sum()
                        .sort_values("Em aberto", ascending=False)
                        .head(10)
                        .sort_values("Em aberto", ascending=True)
                    )
                    if top_clients.empty:
                        fig = empty_chart("Sem clientes para o filtro atual")
                    else:
                        fig = px.bar(top_clients, x="Em aberto", y="Nome Fantasia", orientation="h", text="Em aberto")
                        fig.update_traces(marker_color=BRAND_BLUE, texttemplate="R$ %{text:,.2f}", textposition="outside")
                        fig.update_xaxes(gridcolor="#EEF2F7")
                    st.plotly_chart(chart_layout(fig, "Maiores valores em aberto por cliente", 390), use_container_width=True, config={"displayModeBar": False})

                with g6:
                    daily = (
                        inad_filtered.groupby("Vencimento", as_index=False)["Em aberto"].sum()
                        .sort_values("Vencimento")
                    )
                    if daily.empty:
                        fig = empty_chart("Sem vencimentos para o filtro atual")
                    else:
                        daily["Data"] = daily["Vencimento"].dt.strftime("%d/%m")
                        fig = px.bar(daily, x="Data", y="Em aberto", text="Em aberto")
                        fig.update_traces(marker_color=PURPLE, texttemplate="R$ %{text:,.2f}", textposition="outside")
                        fig.update_yaxes(gridcolor="#EEF2F7")
                    st.plotly_chart(chart_layout(fig, "Valores em aberto por vencimento", 390), use_container_width=True, config={"displayModeBar": False})

                # -------------------------
                # Tabela operacional
                # -------------------------
                st.markdown("#### Títulos em aberto para baixa")
                table_cols = [
                    "Local", "Nome Fantasia", "Origem", "Status faturamento", "Contrato", "Pedido", "Fatura",
                    "Documento", "Tipo Documento", "Emissão", "Vencimento", "Dias em atraso", "Faixa atraso",
                    "Valor Título", "Valor Recebido", "Em aberto", "Vencido com Juros", "Telefone", "Celular",
                    "Representante", "Observação",
                ]
                inad_table = inad_filtered[[c for c in table_cols if c in inad_filtered.columns]].copy()
                inad_table = inad_table.sort_values(["Dias em atraso", "Em aberto"], ascending=[False, False])
                st.dataframe(
                    inad_table,
                    use_container_width=True,
                    hide_index=True,
                    height=560,
                    column_config={
                        "Emissão": st.column_config.DateColumn("Emissão", format="DD/MM/YYYY"),
                        "Vencimento": st.column_config.DateColumn("Vencimento", format="DD/MM/YYYY"),
                        "Dias em atraso": st.column_config.NumberColumn("Dias atraso", format="%d"),
                        "Pedido": st.column_config.NumberColumn("Pedido", format="%.0f"),
                        "Fatura": st.column_config.NumberColumn("Fatura", format="%.0f"),
                        "Valor Título": st.column_config.NumberColumn("Valor título", format="R$ %.2f"),
                        "Valor Recebido": st.column_config.NumberColumn("Recebido", format="R$ %.2f"),
                        "Em aberto": st.column_config.NumberColumn("Em aberto", format="R$ %.2f"),
                        "Vencido com Juros": st.column_config.NumberColumn("Com juros", format="R$ %.2f"),
                    },
                    key="inad_table_v3",
                )

                d1, d2 = st.columns([1.3, 3.7])
                with d1:
                    inad_export = build_receivables_export(inad_filtered)
                    st.download_button(
                        "⬇️ Baixar inadimplência filtrada",
                        data=inad_export,
                        file_name=f"inadimplencia_viva_{today_sp().isoformat()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="download_inad_v3",
                    )
                with d2:
                    st.caption(
                        "Origem: Pedido preenchido = Venda; contrato INTERM = Intermediação; demais contratos = Locação. "
                        "Não faturado = sem número de Fatura e sem DATA FAT na Observação. O período de 30/60/90/120 dias usa a data de vencimento."
                    )

# -----------------------------
# Exportação
# -----------------------------
st.divider()
e1, e2, e3 = st.columns([1.3, 1, 2.7])
with e1:
    export_bytes = build_excel_export(view_df)
    st.download_button(
        "⬇️ Baixar lista filtrada",
        data=export_bytes,
        file_name=f"pendencias_viva_{today_sp().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with e2:
    if is_admin() and LATEST_FILE.exists():
        st.download_button(
            "⬇️ Baixar base atual",
            data=LATEST_FILE.read_bytes(),
            file_name="base_atual_viva.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
with e3:
    st.caption(
        "Regra do painel: 24h = 1 dia vencido; 48h = 2 dias; 72h+ = 3 dias ou mais. "
        "Ocupação = equipamentos locados ÷ total do recorte. Atraso % = atrasados ÷ locados."
    )

st.markdown(
    f'<div class="small-note" style="text-align:center;margin-top:18px;">Viva Cuidar • Pendência Diária • {now_sp().strftime("%d/%m/%Y")}</div>',
    unsafe_allow_html=True,
)
