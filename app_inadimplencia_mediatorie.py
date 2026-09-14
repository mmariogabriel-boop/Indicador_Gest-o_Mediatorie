"""Painel Streamlit de inadimplência da Mediatorie.

Utiliza planilha por faixa, resumo_mensal e titulos_antes_d30 de bases.py.
"""

from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

from bases import planilha, titulos_antes_d30, resumo_mensal

# Compatibilidade durante o deploy: o app novo usa a competência do mês atual
# como referência. Se o Streamlit ainda estiver com um bases.py antigo em cache,
# evita ImportError e monta a referência a partir da data disponível.
try:
    from bases import DATA_REFERENCIA, COMPETENCIA_REFERENCIA
except ImportError:
    import bases as _bases

    DATA_REFERENCIA = getattr(
        _bases,
        "DATA_REFERENCIA",
        pd.Timestamp.today().normalize(),
    )
    COMPETENCIA_REFERENCIA = pd.Timestamp(DATA_REFERENCIA).strftime("%m/%Y")


# =========================================================
# CONFIGURAÇÃO GERAL
# =========================================================

st.set_page_config(
    page_title="Inadimplência | Mediatorie",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

PASTA_PROJETO = Path(__file__).resolve().parent
CAMINHOS_LOGO = [
    PASTA_PROJETO / "assets" / "logo_mediatorie.png",
    PASTA_PROJETO / "logo_mediatorie.png",
]
CAMINHO_LOGO = next((caminho for caminho in CAMINHOS_LOGO if caminho.exists()), None)

VERDE = "#86BC25"
VERDE_ESCURO = "#5F8E16"
VERDE_CLARO = "#DCEBC4"
VERDE_BARRA = "rgba(134, 188, 37, 0.58)"
VERMELHO = "#D64545"
VERMELHO_ESCURO = "#A92F2F"
AZUL = "#1E3A5F"
GRAFITE = "#4B4F4D"
FUNDO = "#F5F7F2"

METAS_CLASSIFICACAO = {
    "Atual": 2.00,
    "D30": 2.00,
    "D60": 1.50,
    "D90": 1.00,
    "D+": 1.00,
}

COLUNAS_OBRIGATORIAS = [
    "Comp",
    "Em aberto",
    "Faturado",
    "Percentual de Inadimplência",
    "Classificação",
]


# =========================================================
# IDENTIDADE VISUAL
# =========================================================

st.markdown(
    f"""
    <style>
        :root {{ color-scheme: light !important; }}

        html, body, .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"] {{
            background-color: {FUNDO} !important;
            color: {GRAFITE} !important;
        }}

        [data-testid="stHeader"] {{
            background-color: #353A37 !important;
        }}

        [data-testid="stToolbar"] *,
        [data-testid="stDecoration"] * {{
            color: #FFFFFF !important;
        }}

        .block-container {{
            max-width: 1500px;
            padding-top: 3.25rem !important;
            padding-bottom: 2rem;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #FFFFFF 0%, #F0F4EA 100%);
            border-right: 1px solid #DDE6D3;
        }}

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3,
        [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
        [data-testid="stSidebar"] small {{
            color: {GRAFITE} !important;
        }}

        [data-baseweb="select"] > div {{
            background-color: #FFFFFF !important;
            border-color: #CFDAC5 !important;
            color: {GRAFITE} !important;
        }}

        [data-baseweb="select"] input,
        [data-baseweb="select"] svg {{
            color: {GRAFITE} !important;
            fill: {GRAFITE} !important;
        }}

        [data-baseweb="tag"] {{
            background-color: {VERDE_ESCURO} !important;
            border-radius: 6px !important;
        }}

        [data-baseweb="tag"] span,
        [data-baseweb="tag"] svg {{
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
        }}

        [data-baseweb="popover"],
        [data-baseweb="menu"] {{
            background-color: #FFFFFF !important;
            color: {GRAFITE} !important;
        }}

        [data-baseweb="menu"] li {{ color: {GRAFITE} !important; }}

        [data-testid="stMetric"] {{
            background-color: #FFFFFF;
            border: 1px solid #E2E8DC;
            border-left: 5px solid {VERDE};
            border-radius: 12px;
            padding: 16px 18px;
            box-shadow: 0 4px 14px rgba(75, 79, 77, 0.06);
        }}

        [data-testid="stMetricLabel"] p {{ color: #667064 !important; }}
        [data-testid="stMetricValue"] {{ color: {GRAFITE} !important; }}
        [data-testid="stMetricDelta"] * {{ font-weight: 650 !important; }}

        div[data-baseweb="tab-list"] {{ gap: 8px; }}

        button[data-baseweb="tab"] {{
            background-color: #FFFFFF;
            color: #657066 !important;
            border: 1px solid #E3E9DF;
            border-radius: 9px 9px 0 0;
            padding: 10px 18px;
        }}

        button[data-baseweb="tab"] p {{ color: #657066 !important; }}

        button[data-baseweb="tab"][aria-selected="true"] {{
            color: {VERDE_ESCURO};
            border-bottom-color: {VERDE};
            background-color: #F3F8EC;
        }}

        button[data-baseweb="tab"][aria-selected="true"] p {{
            color: {VERDE_ESCURO} !important;
            font-weight: 700 !important;
        }}

        .mediatorie-header {{
            padding: 10px 0 18px 0;
            margin-top: 0.35rem;
            border-bottom: 3px solid {VERDE};
        }}

        .mediatorie-kicker {{
            color: {VERDE_ESCURO};
            font-size: 0.84rem;
            font-weight: 700;
            letter-spacing: 0.12em;
            text-transform: uppercase;
        }}

        .mediatorie-title {{
            color: {GRAFITE};
            font-size: 2.05rem;
            font-weight: 750;
            line-height: 1.15;
            margin: 4px 0;
        }}

        .mediatorie-subtitle {{ color: #69736A; font-size: 1rem; }}

        .sidebar-kicker {{
            background: linear-gradient(135deg, {VERDE_ESCURO}, {VERDE});
            color: #FFFFFF !important;
            border-radius: 10px;
            padding: 13px 15px;
            margin: 5px 0 20px 0;
            font-size: 0.83rem;
            font-weight: 750;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }}

        .status-ok {{
            background: #EFF8E5;
            color: {VERDE_ESCURO};
            border: 1px solid #CDE3AC;
            border-radius: 10px;
            padding: 12px 16px;
            margin: 8px 0 18px 0;
            font-weight: 600;
        }}

        .status-alerta {{
            background: #FFF0F0;
            color: {VERMELHO_ESCURO};
            border: 1px solid #F0C7C7;
            border-radius: 10px;
            padding: 12px 16px;
            margin: 8px 0 18px 0;
            font-weight: 600;
        }}

        [data-testid="stCaptionContainer"] p,
        [data-testid="stExpander"] p,
        [data-testid="stDataFrame"] {{
            color: {GRAFITE} !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNÇÕES DE APOIO
# =========================================================

def moeda_brasileira(valor: float) -> str:
    if pd.isna(valor):
        return "R$ 0,00"

    return (
        f"R$ {valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def moeda_resumida(valor: float) -> str:
    if pd.isna(valor):
        return "R$ 0,00"

    if abs(valor) >= 1_000_000:
        return f"R$ {valor / 1_000_000:.2f} mi".replace(".", ",")

    if abs(valor) >= 1_000:
        return f"R$ {valor / 1_000:.0f} mil".replace(".", ",")

    return moeda_brasileira(valor)


def formatar_percentual(valor: float) -> str:
    if pd.isna(valor):
        return "0,00%"
    return f"{valor:.2f}%".replace(".", ",")


def preparar_dados(base: pd.DataFrame) -> pd.DataFrame:
    colunas_ausentes = [
        coluna for coluna in COLUNAS_OBRIGATORIAS if coluna not in base.columns
    ]

    if colunas_ausentes:
        raise KeyError(
            "As seguintes colunas não foram encontradas na base: "
            + ", ".join(colunas_ausentes)
        )

    dados = base.copy()

    for coluna in ["Em aberto", "Faturado", "Percentual de Inadimplência"]:
        dados[coluna] = pd.to_numeric(dados[coluna], errors="coerce")

    dados["Classificação"] = (
        dados["Classificação"].fillna("").astype(str).str.strip()
    )
    dados["Ordem"] = pd.to_datetime(dados["Comp"], format="%m/%Y", errors="coerce")

    dados = (
        dados.loc[dados["Classificação"].ne("")]
        .dropna(
            subset=[
                "Ordem",
                "Em aberto",
                "Faturado",
                "Percentual de Inadimplência",
            ]
        )
        .sort_values("Ordem")
        .reset_index(drop=True)
    )

    dados["Meta"] = dados["Classificação"].map(METAS_CLASSIFICACAO)
    dados["Desvio"] = dados["Percentual de Inadimplência"] - dados["Meta"]
    dados["Status"] = dados["Desvio"].le(0).map(
        {True: "Dentro da meta", False: "Acima da meta"}
    )
    dados["Comp_exibicao"] = (
        dados["Comp"].astype(str)
        + "<br><b>"
        + dados["Classificação"]
        + "</b>"
    )

    dados["Hover aberto"] = dados["Em aberto"].map(moeda_brasileira)
    dados["Hover faturado"] = dados["Faturado"].map(moeda_brasileira)
    dados["Rótulo aberto"] = dados["Em aberto"].map(moeda_resumida)
    dados["Rótulo percentual"] = dados["Percentual de Inadimplência"].map(
        formatar_percentual
    )
    dados["Rótulo meta"] = dados["Meta"].map(formatar_percentual)

    return dados


def preparar_dados_financeiros(base: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por competência; nunca soma faturamento repetido por faixa."""
    obrigatorias = ["Comp", "Em aberto", "Faturado"]
    ausentes = [coluna for coluna in obrigatorias if coluna not in base.columns]
    if ausentes:
        raise KeyError(f"Colunas ausentes no resumo_mensal: {ausentes}")
    dados = base.copy()
    dados["Ordem"] = pd.to_datetime(dados["Comp"], format="%m/%Y", errors="coerce")
    for coluna in ["Em aberto", "Faturado"]:
        dados[coluna] = pd.to_numeric(dados[coluna], errors="coerce").replace(
            [float("inf"), float("-inf")], float("nan")
        )
    dados = dados.dropna(subset=["Ordem", "Em aberto"]).sort_values("Ordem")
    if dados["Ordem"].duplicated().any():
        raise ValueError("resumo_mensal deve conter apenas uma linha por competência.")
    dados["Comp"] = dados["Ordem"].dt.strftime("%m/%Y")
    dados["Percentual de Inadimplência"] = (
        dados["Em aberto"]
        .div(dados["Faturado"].where(dados["Faturado"].gt(0)))
        .mul(100).round(2)
    )
    dados["Comp_exibicao"] = dados["Comp"]
    dados["Hover aberto"] = dados["Em aberto"].map(moeda_brasileira)
    dados["Hover faturado"] = dados["Faturado"].map(
        lambda valor: "Não informado" if pd.isna(valor) else moeda_brasileira(valor)
    )
    dados["Rótulo aberto"] = dados["Em aberto"].map(moeda_resumida)
    dados["Rótulo percentual"] = dados["Percentual de Inadimplência"].map(
        lambda valor: "" if pd.isna(valor) else formatar_percentual(valor)
    )
    return dados.reset_index(drop=True)


def criar_cabecalho() -> None:
    if CAMINHO_LOGO is not None:
        coluna_logo, coluna_texto = st.columns([0.85, 5.15], gap="large")
        with coluna_logo:
            st.image(str(CAMINHO_LOGO), width=175)
    else:
        coluna_texto = st.container()

    with coluna_texto:
        st.markdown(
            """
            <div class="mediatorie-header">
                <div class="mediatorie-kicker">Gestão financeira</div>
                <div class="mediatorie-title">Análise de Inadimplência</div>
                <div class="mediatorie-subtitle">
                    Indicadores por competência, considerando no realizado apenas valores já vencidos.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def criar_grafico_meta_realizado(dados: pd.DataFrame) -> go.Figure:
    valores = pd.concat(
        [dados["Percentual de Inadimplência"], dados["Meta"]]
    ).dropna()
    maior_valor = float(valores.max()) if not valores.empty else 0.0
    limite = max(4, math.ceil(maior_valor * 1.30))

    fig = go.Figure()

    # O realizado é a informação principal e fica nas barras.
    fig.add_trace(
        go.Bar(
            x=dados["Comp_exibicao"],
            y=dados["Percentual de Inadimplência"],
            name="Realizado",
            marker=dict(
                color=VERDE_BARRA,
                line=dict(color=VERDE_ESCURO, width=1.2),
            ),
            text=dados["Rótulo percentual"],
            textposition="outside",
            textfont=dict(size=12, color="#000000"),
            cliponaxis=False,
            customdata=dados[
                [
                    "Classificação",
                    "Status",
                    "Hover aberto",
                    "Hover faturado",
                ]
            ].to_numpy(),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Realizado: <b>%{y:.2f}%</b><br>"
                "Classificação: %{customdata[0]}<br>"
                "Status: %{customdata[1]}<br>"
                "Em aberto: %{customdata[2]}<br>"
                "Faturado: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    # A meta é a referência e fica em uma linha pontilhada.
    fig.add_trace(
        go.Scatter(
            x=dados["Comp_exibicao"],
            y=dados["Meta"],
            name="Meta",
            mode="lines+markers+text",
            text=dados["Rótulo meta"],
            textposition="bottom center",
            textfont=dict(size=11, color="#000000"),
            line=dict(color=VERDE_ESCURO, width=3, dash="dot"),
            marker=dict(
                size=9,
                color="#FFFFFF",
                line=dict(color=VERDE_ESCURO, width=2),
            ),
            hovertemplate="Meta: <b>%{y:.2f}%</b><extra></extra>",
        )
    )

    fig.update_layout(
        title=dict(
            text=(
                "Meta x inadimplência realizada"
                "<br><sup>Barras = realizado | Linha pontilhada = meta</sup>"
            ),
            x=0.01,
            font=dict(color="#000000"),
        ),
        template="plotly_white",
        height=570,
        margin=dict(l=45, r=30, t=95, b=80),
        legend=dict(
            orientation="h",
            y=1.10,
            x=1,
            xanchor="right",
            font=dict(color="#000000"),
        ),
        hovermode="x unified",
        bargap=0.42,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Arial", size=13, color="#000000"),
        hoverlabel=dict(
            bgcolor="#FFFFFF",
            bordercolor="#CBD8C0",
            font=dict(color=GRAFITE, size=12),
        ),
        uniformtext_minsize=9,
        uniformtext_mode="hide",
    )
    fig.update_yaxes(
        title_text="Percentual de inadimplência",
        range=[0, limite],
        ticksuffix="%",
        tickformat=".1f",
        tickfont=dict(color="#000000"),
        title_font=dict(color="#000000"),
        gridcolor="rgba(148,163,184,0.25)",
        zeroline=False,
    )
    fig.update_xaxes(
        title_text="Competência e faixa de atraso",
        showgrid=False,
        automargin=True,
        tickfont=dict(color="#000000"),
        title_font=dict(color="#000000"),
    )

    return fig


def criar_grafico_financeiro(dados: pd.DataFrame) -> go.Figure:
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=dados["Comp_exibicao"],
            y=dados["Em aberto"],
            name="Em aberto",
            marker=dict(
                color=VERDE_BARRA,
                line=dict(color=VERDE_ESCURO, width=1.2),
            ),
            text=dados["Rótulo aberto"],
            textposition="inside",
            insidetextanchor="middle",
            textangle=0,
            textfont=dict(size=11, color="#000000"),
            cliponaxis=False,
            customdata=dados[["Hover aberto", "Hover faturado"]].to_numpy(),
            hovertemplate=(
                "Em aberto: <b>%{customdata[0]}</b><br>"
                "Faturado: %{customdata[1]}<extra></extra>"
            ),
        ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(
            x=dados["Comp_exibicao"],
            y=dados["Percentual de Inadimplência"],
            name="Percentual em aberto",
            mode="lines+markers+text",
            text=dados["Rótulo percentual"],
            textposition="top center",
            connectgaps=False,
            cliponaxis=False,
            textfont=dict(size=11, color="#000000"),
            line=dict(color=AZUL, width=4),
            marker=dict(
                size=10,
                color="#FFFFFF",
                line=dict(color=AZUL, width=3),
            ),
            hovertemplate="Saldo em aberto / faturado: <b>%{y:.2f}%</b><extra></extra>",
        ),
        secondary_y=True,
    )

    maior_aberto = float(dados["Em aberto"].max())
    limite_aberto = max(maior_aberto * 1.30, 1)

    percentuais_validos = dados["Percentual de Inadimplência"].dropna()
    maior_percentual = float(percentuais_validos.max()) if not percentuais_validos.empty else 0.0
    limite_percentual = max(4, math.ceil(maior_percentual * 1.30))

    fig.update_yaxes(
        title_text="Valor em aberto",
        range=[0, limite_aberto],
        tickprefix="R$ ",
        tickformat="~s",
        tickfont=dict(color="#000000"),
        title_font=dict(color="#000000"),
        gridcolor="rgba(148,163,184,0.25)",
        zeroline=False,
        secondary_y=False,
    )
    fig.update_yaxes(
        title_text="Percentual em aberto",
        range=[0, limite_percentual],
        ticksuffix="%",
        tickformat=".1f",
        tickfont=dict(color="#000000"),
        title_font=dict(color="#000000"),
        showgrid=False,
        zeroline=False,
        secondary_y=True,
    )
    fig.update_layout(
        title=dict(
            text=(
                "Evolução financeira"
                "<br><sup>Valor em aberto e percentual sobre o faturamento</sup>"
            ),
            x=0.01,
            font=dict(color="#000000"),
        ),
        template="plotly_white",
        height=570,
        margin=dict(l=45, r=45, t=95, b=80),
        legend=dict(
            orientation="h",
            y=1.10,
            x=1,
            xanchor="right",
            font=dict(color="#000000"),
        ),
        hovermode="x unified",
        bargap=0.44,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Arial", size=13, color="#000000"),
        hoverlabel=dict(
            bgcolor="#FFFFFF",
            bordercolor="#CBD8C0",
            font=dict(color=GRAFITE, size=12),
        ),
        uniformtext_minsize=9,
        uniformtext_mode="hide",
    )
    fig.update_xaxes(
        title_text="Competência de vencimento",
        type="category",
        categoryorder="array",
        categoryarray=dados["Comp"].tolist(),
        showgrid=False,
        automargin=True,
        tickfont=dict(color="#000000"),
        title_font=dict(color="#000000"),
    )

    return fig


def exibir_kpis(dados: pd.DataFrame) -> None:
    linha_atual = dados.loc[dados["Classificação"].eq("Atual")]
    linha_referencia = linha_atual.iloc[-1] if not linha_atual.empty else dados.iloc[-1]

    competencia = linha_referencia["Comp"]
    classificacao = linha_referencia["Classificação"]
    desvio = float(linha_referencia["Desvio"])

    coluna1, coluna2 = st.columns(2, gap="large")
    coluna1.metric(
        label=f"Em aberto — {competencia}",
        value=moeda_resumida(linha_referencia["Em aberto"]),
        help=f"Competência classificada como {classificacao}.",
    )
    coluna2.metric(
        label=f"Inadimplência — {competencia}",
        value=formatar_percentual(linha_referencia["Percentual de Inadimplência"]),
        delta=(
            f"{desvio:+.2f} p.p. em relação à meta"
            .replace(".", ",")
        ),
        delta_color="inverse",
        help=f"Meta da faixa {classificacao}: {formatar_percentual(linha_referencia['Meta'])}.",
    )

    if desvio <= 0:
        classe = "status-ok"
        mensagem = (
            f"✓ {competencia} está dentro da meta definida para {classificacao}."
        )
    else:
        classe = "status-alerta"
        mensagem = (
            f"⚠ {competencia} está acima da meta definida para {classificacao}."
        )

    st.markdown(
        f'<div class="{classe}">{mensagem}</div>',
        unsafe_allow_html=True,
    )


def tabela_formatada(dados: pd.DataFrame):
    tabela = dados[
        [
            "Comp",
            "Classificação",
            "Em aberto",
            "Faturado",
            "Percentual de Inadimplência",
            "Meta",
            "Desvio",
            "Status",
        ]
    ].copy()

    return tabela.style.format(
        {
            "Em aberto": moeda_brasileira,
            "Faturado": moeda_brasileira,
            "Percentual de Inadimplência": formatar_percentual,
            "Meta": formatar_percentual,
            "Desvio": lambda valor: f"{valor:+.2f} p.p.".replace(".", ","),
        }
    )


def preparar_titulos_antes_d30(base: pd.DataFrame) -> pd.DataFrame:
    colunas_obrigatorias = [
        "Comp",
        "Em aberto",
        "Dias de atraso",
        "Dias para D30",
        "Data de entrada no D30",
        "Classificação",
    ]

    colunas_ausentes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in base.columns
    ]

    if colunas_ausentes:
        raise KeyError(
            "A base de títulos antes do D30 não possui as colunas: "
            + ", ".join(colunas_ausentes)
        )

    dados = base.copy()
    dados["Em aberto"] = pd.to_numeric(
        dados["Em aberto"],
        errors="coerce"
    )
    dados["Dias de atraso"] = pd.to_numeric(
        dados["Dias de atraso"],
        errors="coerce"
    ).astype("Int64")
    dados["Dias para D30"] = pd.to_numeric(
        dados["Dias para D30"],
        errors="coerce"
    ).astype("Int64")
    dados["Data de entrada no D30"] = pd.to_datetime(
        dados["Data de entrada no D30"],
        errors="coerce"
    )

    return (
        dados.loc[dados["Classificação"].eq("Atual")]
        .sort_values(
            ["Data de entrada no D30", "Dias para D30"],
            ascending=[True, True]
        )
        .reset_index(drop=True)
    )


def tabela_titulos_formatada(dados: pd.DataFrame) -> pd.DataFrame:
    """Mantém os tipos dos dados e todas as linhas, sem usar Pandas Styler."""
    tabela = dados.copy()

    colunas_data = [
        "Dt.lançamento",
        "Data vencimento líq.",
        "Data de vencimento",
        "Data de entrada no D30",
        "data_hoje",
    ]

    for coluna in colunas_data:
        if coluna in tabela.columns:
            tabela[coluna] = pd.to_datetime(
                tabela[coluna],
                errors="coerce",
                format="mixed",
                dayfirst=True,
            )
    return tabela


def configurar_colunas_tabela(dados: pd.DataFrame) -> dict:
    """Formatação nativa: preserva ordenação numérica e não cria estilos por célula."""
    configuracao = {}
    for coluna in ["Em aberto", "Faturado", "Valor a vencer", "Valor vence hoje", "Valor vencido"]:
        if coluna in dados.columns:
            configuracao[coluna] = st.column_config.NumberColumn(coluna, format="R$ %.2f")
    for coluna in ["Dias de atraso", "Dias para vencimento", "Dias para D30", "Contagem_dias"]:
        if coluna in dados.columns:
            configuracao[coluna] = st.column_config.NumberColumn(coluna, format="%d")
    for coluna in ["Dt.lançamento", "Data vencimento líq.", "Data de vencimento",
                   "Data de entrada no D30", "data_hoje"]:
        if coluna in dados.columns:
            configuracao[coluna] = st.column_config.DateColumn(coluna, format="DD/MM/YYYY")
    if "Percentual de Inadimplência" in dados.columns:
        configuracao["Percentual de Inadimplência"] = st.column_config.NumberColumn(
            "Percentual em aberto", format="%.2f%%"
        )
    return configuracao


# =========================================================
# PREPARAÇÃO DA BASE
# =========================================================

try:
    dados_completos = preparar_dados(planilha)
    dados_mensais = preparar_dados_financeiros(resumo_mensal)
    dados_titulos_antes_d30 = preparar_titulos_antes_d30(
        titulos_antes_d30
    )
except (KeyError, ValueError) as erro:
    st.error(str(erro))
    st.stop()


# =========================================================
# CABEÇALHO E FILTROS
# =========================================================

criar_cabecalho()

with st.sidebar:
    st.markdown(
        '<div class="sidebar-kicker">Painel financeiro</div>',
        unsafe_allow_html=True,
    )
    st.markdown("### Filtros")

    ordem_classificacoes = ["Atual", "D30", "D60", "D90", "D+"]
    classificacoes_existentes = set(
        dados_completos["Classificação"].dropna().astype(str)
    )
    classificacoes_disponiveis = [
        faixa for faixa in ordem_classificacoes
        if faixa in classificacoes_existentes
    ]
    classificacoes = st.multiselect(
        "Faixa de atraso",
        options=classificacoes_disponiveis,
        default=classificacoes_disponiveis,
    )

    # Inclui também meses sem percentual e títulos futuros da quarta aba.
    competencias_base = pd.concat(
        [dados_completos["Comp"], dados_mensais["Comp"], dados_titulos_antes_d30["Comp"]],
        ignore_index=True,
    ).dropna().drop_duplicates()
    datas_competencias = pd.to_datetime(competencias_base, format="%m/%Y", errors="coerce")
    competencias_disponiveis = (
        datas_competencias.dropna().drop_duplicates().sort_values().dt.strftime("%m/%Y").tolist()
    )
    competencias = st.multiselect(
        "Competência",
        options=competencias_disponiveis,
        default=competencias_disponiveis,
    )

    st.divider()
    st.caption(
        f"Competência Atual: {COMPETENCIA_REFERENCIA} "
        f"(data de referência: {DATA_REFERENCIA:%d/%m/%Y})."
    )
    st.caption(
        "Faixas por competência: Atual = mês corrente | "
        "D30 = 1 competência anterior | D60 = 2 | "
        "D90 = 3 | D+ = 4 ou mais."
    )
    st.caption(
        "Metas configuradas: Atual e D30 = 2,00% | D60 = 1,50% | "
        "D90 e D+ = 1,00%."
    )


dados_filtrados = dados_completos.loc[
    dados_completos["Classificação"].isin(classificacoes)
    & dados_completos["Comp"].isin(competencias)
].copy()

dados_financeiros = dados_mensais.loc[dados_mensais["Comp"].isin(competencias)].copy()


# =========================================================
# ABAS DO PAINEL
# =========================================================

aba_executiva, aba_financeira, aba_tabela, aba_antes_d30 = st.tabs(
    [
        "Visão executiva",
        "Evolução financeira",
        "Tabela detalhada",
        "Títulos antes do D30",
    ]
)

with aba_executiva:
    if dados_filtrados.empty:
        st.info("Nenhum percentual válido para as competências e faixas selecionadas.")
    else:
        exibir_kpis(dados_filtrados)
        st.plotly_chart(
            criar_grafico_meta_realizado(dados_filtrados),
            use_container_width=True,
            theme=None,
            config={"displaylogo": False, "locale": "pt-BR", "responsive": True},
        )

with aba_financeira:
    st.caption(
        "Totais por competência: uma barra e um percentual por mês. "
        f"Competência Atual: {COMPETENCIA_REFERENCIA}. "
        "Esta aba considera todas as faixas e respeita apenas o filtro de competência."
    )
    st.caption(
        "Percentual realizado = valor já vencido ÷ faturamento do mês × 100. "
        "A competência define a faixa (Atual/D30/D60/D90/D+), mas o valor é atualizado por dia: "
        "títulos a vencer e os que vencem hoje não entram no realizado."
    )
    if dados_financeiros.empty:
        st.info("Nenhum dado mensal para as competências selecionadas.")
    else:
        if dados_financeiros["Percentual de Inadimplência"].isna().any():
            st.caption(
                "Meses com faturamento ausente, zero ou negativo mantêm a barra de valor "
                "em aberto, mas não apresentam percentual."
            )
        st.plotly_chart(
            criar_grafico_financeiro(dados_financeiros),
            use_container_width=True,
            theme=None,
            config={"displaylogo": False, "locale": "pt-BR", "responsive": True},
        )
        with st.expander("Conferir os valores do gráfico"):
            colunas_conferencia = [
                coluna for coluna in ["Comp", "Em aberto", "Valor a vencer",
                                      "Valor vence hoje", "Valor vencido", "Faturado",
                                      "Percentual de Inadimplência", "Situação faturamento"]
                if coluna in dados_financeiros.columns
            ]
            conferencia = dados_financeiros[colunas_conferencia]
            st.dataframe(
                conferencia, column_config=configurar_colunas_tabela(conferencia),
                use_container_width=True, hide_index=True,
            )

with aba_tabela:
    st.subheader("Detalhamento por competência")
    st.caption(
        f"Classificação referente à competência {COMPETENCIA_REFERENCIA}. "
        "A tabela respeita os filtros aplicados na barra lateral."
    )
    st.dataframe(
        tabela_formatada(dados_filtrados),
        use_container_width=True,
        hide_index=True,
    )

with aba_antes_d30:
    st.subheader("Títulos da competência Atual — antes do D30")
    st.caption(
        f"Esta lista mostra os títulos da competência {COMPETENCIA_REFERENCIA}. "
        "A faixa permanece Atual até a virada da competência, mas o indicador só soma cada título "
        "depois que sua data de vencimento perde a validade. "
        "No primeiro dia do mês seguinte, a competência anterior passa para D30. "
        "Os dias de atraso são apenas informativos e não alteram a faixa."
    )

    titulos_filtrados = dados_titulos_antes_d30.loc[
        dados_titulos_antes_d30["Comp"].isin(competencias)
    ].copy()

    if titulos_filtrados.empty:
        st.success(
            "Não existem títulos antes do D30 para as competências "
            "selecionadas."
        )
    else:
        coluna_quantidade, coluna_valor = st.columns(2, gap="large")
        coluna_quantidade.metric(
            label="Quantidade de títulos",
            value=f"{len(titulos_filtrados):,}".replace(",", "."),
        )
        coluna_valor.metric(
            label="Valor nominal da competência Atual",
            value=moeda_resumida(titulos_filtrados["Em aberto"].sum()),
        )

        st.dataframe(
            tabela_titulos_formatada(titulos_filtrados),
            column_config=configurar_colunas_tabela(titulos_filtrados),
            use_container_width=True,
            hide_index=True,
            height=520,
        )

        st.download_button(
            label="Baixar lista em CSV",
            data=titulos_filtrados.to_csv(
                index=False,
                sep=";",
                decimal=",",
                encoding="utf-8-sig"
            ).encode("utf-8-sig"),
            file_name="titulos_antes_d30.csv",
            mime="text/csv",
            use_container_width=False,
        )

st.caption(
    "Regra: a competência define a faixa e o dia define o valor vencido. "
    "Atual = competência corrente; D30 = 1 competência anterior; "
    "D60 = 2; D90 = 3; D+ = 4 ou mais. "
    "A quantidade de dias de atraso não define a faixa. "
    "Exemplo em 09/2026: Atual = 09/2026, D30 = 08/2026, "
    "D60 = 07/2026, D90 = 06/2026 e D+ = 05/2026 ou anterior."
)
