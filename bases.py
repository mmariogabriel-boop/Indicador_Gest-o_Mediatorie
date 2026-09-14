"""Bases de inadimplência da Mediatorie.

Regra oficial dos indicadores por competência:
    competência atual        -> Atual
    1 competência anterior   -> D30
    2 competências anteriores-> D60
    3 competências anteriores-> D90
    4 ou mais anteriores     -> D+

Exemplo em 09/2026:
    09/2026 -> Atual
    08/2026 -> D30
    07/2026 -> D60
    06/2026 -> D90
    05/2026 ou anterior -> D+

A FAIXA (Atual/D30/D60/D90/D+) é definida pela competência.
O VALOR do indicador, porém, é calculado diariamente: só entra no numerador
o título cuja Data vencimento líq. já perdeu a validade (vencimento < data de referência).
Títulos a vencer e títulos que vencem hoje NÃO entram em "Em aberto" do indicador.

Saídas disponíveis para importação pelo app:
    planilha             : uma linha por competência e classificação.
    titulos_antes_d30     : títulos da competência Atual.
    resumo_mensal        : uma linha por competência, sem duplicar faturamento.
    base_titulos         : todos os registros elegíveis, classificados por competência.
    faturado_por_mes     : faturamento consolidado por competência.
    avisos               : inconsistências encontradas na preparação.

O percentual mantém a fórmula do projeto:
    Valor vencido / Faturado * 100

A competência define a faixa; o dia define quais títulos daquela competência
já venceram e, portanto, entram no valor do indicador.
"""

from pathlib import Path
import re
import warnings

import numpy as np
import pandas as pd


# =========================================================
# CAMINHOS
# =========================================================

PASTA_PROJETO = Path(__file__).resolve().parent
ARQUIVO_TITULOS = PASTA_PROJETO / "CustomerLineItems (25).xlsx"
ARQUIVO_FATURADO = PASTA_PROJETO / "teste.xlsx"
ABA_FATURADO = "Planilha1"


# =========================================================
# CONFIGURAÇÕES
# =========================================================

EMPRESA_ALVO = 1000
CODIGOS_EXCLUIR = ["1000", "1800", "1600", "10000"]

COLUNA_EMPRESA = "Empresa"
COLUNA_LANCAMENTO = "Lançamento contábil"
COLUNA_MONTANTE = "Montante (ME)"
COLUNA_DATA = "Data vencimento líq."

# Competência atual + 11 competências anteriores.
MESES_ANALISE = 12

ORDEM_CLASSIFICACAO = {
    "Atual": 0,
    "D30": 1,
    "D60": 2,
    "D90": 3,
    "D+": 4,
    "Futura": 5,
}

COLUNAS_DETALHE_PREFERENCIAIS = [
    COLUNA_EMPRESA,
    "Cliente",
    "Nome do cliente",
    "Conta",
    "Nº documento",
    "Documento",
    "Atribuição",
    "Referência",
    COLUNA_LANCAMENTO,
    COLUNA_DATA,
    COLUNA_MONTANTE,
    "Dt.lançamento",
    "Data de vencimento",
]


# =========================================================
# VALIDAÇÕES E CONVERSÕES
# =========================================================


def validar_arquivo(caminho):
    caminho = Path(caminho)
    if not caminho.is_file():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {caminho}\n"
            "Coloque as planilhas na pasta deste bases.py ou ajuste "
            "ARQUIVO_TITULOS e ARQUIVO_FATURADO no início do código."
        )


def validar_colunas(dataframe, colunas, nome_arquivo):
    ausentes = [coluna for coluna in colunas if coluna not in dataframe.columns]
    if ausentes:
        raise KeyError(f"Colunas ausentes em {nome_arquivo}: {ausentes}")


def normalizar_cabecalhos(dataframe):
    dados = dataframe.copy()
    dados.columns = dados.columns.astype(str).str.strip()

    if dados.columns.duplicated().any():
        raise ValueError("A planilha contém nomes de colunas repetidos.")

    return dados


def converter_datas(valores):
    """Converte datas Excel/texto brasileiro/ISO. Datas inválidas viram NaT."""
    return pd.to_datetime(
        valores,
        format="mixed",
        dayfirst=True,
        errors="coerce",
    ).dt.normalize()


def normalizar_referencia(data_referencia=None):
    """Data usada para saber qual é a competência Atual."""
    if data_referencia is None:
        return pd.Timestamp.today().normalize()

    referencia = pd.to_datetime(
        data_referencia,
        format="mixed",
        dayfirst=True,
        errors="coerce",
    )

    if pd.isna(referencia):
        raise ValueError("A data de referência é inválida.")

    return referencia.normalize()


def obter_periodo_referencia(data_referencia=None):
    """Retorna a competência correspondente à data de execução/simulação."""
    return normalizar_referencia(data_referencia).to_period("M")


# =========================================================
# REGRAS DE PRAZO E CLASSIFICAÇÃO
# =========================================================


def calcular_dias_atraso(data_vencimento, data_referencia=None):
    """Dias vencidos para consulta operacional; não define a faixa do indicador."""
    vencimento = converter_datas(pd.Series([data_vencimento])).iloc[0]

    if pd.isna(vencimento):
        return pd.NA

    return max((normalizar_referencia(data_referencia) - vencimento).days, 0)


def adicionar_prazos(titulos, data_referencia=None):
    """Calcula campos operacionais por título, sem definir a faixa do indicador."""
    referencia = normalizar_referencia(data_referencia)

    dados = titulos.copy()
    dados[COLUNA_DATA] = converter_datas(dados[COLUNA_DATA])
    dados["data_hoje"] = referencia
    dados["Periodo"] = dados[COLUNA_DATA].dt.to_period("M")
    dados["Comp"] = dados["Periodo"].dt.strftime("%m/%Y")

    # Negativo = ainda vai vencer; zero = vence hoje; positivo = vencido.
    dias = (referencia - dados[COLUNA_DATA]).dt.days.astype("Int64")
    dados["Contagem_dias"] = dias

    dados["Contagem_Prazo"] = np.select(
        [
            dias.lt(0).to_numpy(dtype=bool, na_value=False),
            dias.eq(0).to_numpy(dtype=bool, na_value=False),
            dias.gt(0).to_numpy(dtype=bool, na_value=False),
        ],
        ["A vencer", "Vence hoje", "Vencido"],
        default="Data inválida",
    )

    dados["Dias de atraso"] = dias.clip(lower=0)
    dados["Dias para vencimento"] = (-dias).clip(lower=0)

    # Pela regra por competência, todos os títulos do mês entram em D30
    # juntos no primeiro dia da competência seguinte.
    dados["Data de entrada no D30"] = (
        dados["Periodo"] + 1
    ).dt.to_timestamp()

    dados["Dias para D30"] = (
        (dados["Data de entrada no D30"] - referencia)
        .dt.days
        .clip(lower=0)
        .astype("Int64")
    )

    return dados


def classificar_por_competencia(titulos, data_referencia=None):
    """Classifica Atual/D30/D60/D90/D+ exclusivamente pela competência.

    Exemplo se a data de referência estiver em 09/2026:
        09/2026 -> Atual
        08/2026 -> D30
        07/2026 -> D60
        06/2026 -> D90
        <= 05/2026 -> D+
        >= 10/2026 -> Futura
    """
    referencia = normalizar_referencia(data_referencia)
    dados = titulos.copy()

    if "Periodo" not in dados.columns:
        dados["Periodo"] = converter_datas(dados[COLUNA_DATA]).dt.to_period("M")

    meses_atraso = (
        (referencia.year - dados["Periodo"].dt.year) * 12
        + (referencia.month - dados["Periodo"].dt.month)
    ).astype("Int64")

    dados["Meses de atraso"] = meses_atraso

    dados["Classificação"] = np.select(
        [
            meses_atraso.eq(0).to_numpy(dtype=bool, na_value=False),
            meses_atraso.eq(1).to_numpy(dtype=bool, na_value=False),
            meses_atraso.eq(2).to_numpy(dtype=bool, na_value=False),
            meses_atraso.eq(3).to_numpy(dtype=bool, na_value=False),
            meses_atraso.ge(4).to_numpy(dtype=bool, na_value=False),
            meses_atraso.lt(0).to_numpy(dtype=bool, na_value=False),
        ],
        ["Atual", "D30", "D60", "D90", "D+", "Futura"],
        default="Data inválida",
    )

    return dados


# =========================================================
# PREPARAÇÃO DOS TÍTULOS
# =========================================================


def preparar_titulos(titulos, data_referencia, avisos_saida):
    dados = normalizar_cabecalhos(titulos)

    validar_colunas(
        dados,
        [COLUNA_EMPRESA, COLUNA_LANCAMENTO, COLUNA_MONTANTE, COLUNA_DATA],
        ARQUIVO_TITULOS.name,
    )

    dados[COLUNA_EMPRESA] = pd.to_numeric(
        dados[COLUNA_EMPRESA],
        errors="coerce",
    )

    dados[COLUNA_LANCAMENTO] = (
        dados[COLUNA_LANCAMENTO]
        .astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    dados[COLUNA_MONTANTE] = pd.to_numeric(
        dados[COLUNA_MONTANTE],
        errors="coerce",
    ).replace([np.inf, -np.inf], np.nan)

    dados[COLUNA_DATA] = converter_datas(dados[COLUNA_DATA])

    # Mantém a regra do projeto: excluir códigos por trecho contido no texto.
    padrao = "|".join(re.escape(codigo) for codigo in CODIGOS_EXCLUIR)
    excluidos = dados[COLUNA_LANCAMENTO].str.contains(padrao, na=False)

    empresa_e_codigo = (
        dados[COLUNA_EMPRESA].eq(EMPRESA_ALVO)
        & ~excluidos
    )

    invalidos = empresa_e_codigo & (
        dados[COLUNA_DATA].isna()
        | dados[COLUNA_MONTANTE].isna()
    )

    if invalidos.any():
        avisos_saida.append(
            f"{invalidos.sum()} registro(s) descartado(s) por vencimento "
            "ou montante ausente/inválido."
        )

    dados = dados.loc[
        empresa_e_codigo
        & dados[COLUNA_MONTANTE].ge(0)
        & dados[COLUNA_DATA].notna()
    ].copy()

    if dados.empty:
        raise ValueError(
            "Não há títulos válidos após os filtros da empresa e dos lançamentos."
        )

    # Não deduplica: cada linha elegível do Excel é preservada.
    dados = adicionar_prazos(dados, data_referencia)
    dados = classificar_por_competencia(dados, data_referencia)

    return dados.reset_index(drop=True)


def selecionar_titulos_antes_d30(titulos_classificados):
    """Retorna a competência Atual, que migrará inteira para D30 no próximo mês."""
    detalhe = titulos_classificados.loc[
        titulos_classificados["Classificação"].eq("Atual")
    ].copy()

    calculadas = [
        "Contagem_Prazo",
        "Contagem_dias",
        "Dias de atraso",
        "Dias para vencimento",
        "Dias para D30",
        "Data de entrada no D30",
        "Meses de atraso",
        "Classificação",
        "data_hoje",
    ]

    preferenciais = [
        coluna
        for coluna in COLUNAS_DETALHE_PREFERENCIAIS
        if coluna in detalhe.columns
    ]

    colunas = list(dict.fromkeys(["Comp"] + preferenciais + calculadas))

    # Mantém identificadores adicionais existentes no arquivo original.
    colunas += [
        coluna
        for coluna in detalhe.columns
        if coluna not in colunas and coluna != "Periodo"
    ]

    return (
        detalhe
        .sort_values(
            ["Data de entrada no D30", COLUNA_DATA],
            kind="stable",
        )[colunas]
        .rename(columns={COLUNA_MONTANTE: "Em aberto"})
        .reset_index(drop=True)
    )


def gerar_titulos_antes_d30(titulos, data_referencia=None):
    """Compatibilidade: recebe títulos e aplica a regra mensal por competência."""
    referencia = normalizar_referencia(data_referencia)
    dados = adicionar_prazos(titulos, referencia)
    dados = classificar_por_competencia(dados, referencia)
    return selecionar_titulos_antes_d30(dados)


# =========================================================
# FATURAMENTO
# =========================================================


def preparar_faturamento(faturado, avisos_saida):
    dados = normalizar_cabecalhos(faturado)

    validar_colunas(
        dados,
        ["mês", "faturado"],
        ARQUIVO_FATURADO.name,
    )

    # Aceita competência como data Excel, texto de data ou MM/AAAA.
    datas = dados["mês"].copy()
    texto = datas.astype("string").str.strip()
    so_mes = texto.str.fullmatch(r"\d{1,2}/\d{4}", na=False)

    datas = datas.astype(object)
    datas.loc[so_mes] = "01/" + texto.loc[so_mes]

    dados["Periodo"] = converter_datas(datas).dt.to_period("M")
    dados["faturado"] = pd.to_numeric(
        dados["faturado"],
        errors="coerce",
    ).replace([np.inf, -np.inf], np.nan)

    invalidos = dados["Periodo"].isna() | dados["faturado"].isna()

    if invalidos.any():
        avisos_saida.append(
            f"{invalidos.sum()} linha(s) de faturamento ignorada(s) "
            "por data ou valor inválido."
        )

    return (
        dados.loc[~invalidos]
        .groupby("Periodo", as_index=False)["faturado"]
        .sum()
        .rename(columns={"faturado": "Faturado"})
    )


# =========================================================
# RESUMOS
# =========================================================


def resumir_aberto(titulos, chaves):
    """Resume os títulos preservando competência e corte diário.

    Regra do indicador:
    - competência define Atual/D30/D60/D90/D+;
    - somente vencimento ANTERIOR à data de referência entra em "Em aberto";
    - título que vence hoje ainda não perdeu a validade e não entra no numerador;
    - título a vencer também não entra no numerador.
    """
    dados = titulos.copy()

    vencido = dados["Contagem_dias"].gt(0)
    vence_hoje = dados["Contagem_dias"].eq(0)
    a_vencer = dados["Contagem_dias"].lt(0)

    dados["Valor a vencer"] = dados[COLUNA_MONTANTE].where(a_vencer, 0)
    dados["Valor vence hoje"] = dados[COLUNA_MONTANTE].where(vence_hoje, 0)
    dados["Valor vencido"] = dados[COLUNA_MONTANTE].where(vencido, 0)
    dados["Registro vencido"] = vencido.astype("int64")

    resumo = dados.groupby(chaves, as_index=False).agg(**{
        # IMPORTANTE: o valor do indicador é somente o que já perdeu a validade.
        "Em aberto": ("Valor vencido", "sum"),
        "Quantidade vencida": ("Registro vencido", "sum"),
        "Quantidade de registros": (COLUNA_MONTANTE, "size"),
        "Menor atraso": ("Dias de atraso", "min"),
        "Dias de atraso": ("Dias de atraso", "max"),
        "Valor a vencer": ("Valor a vencer", "sum"),
        "Valor vence hoje": ("Valor vence hoje", "sum"),
        "Valor vencido": ("Valor vencido", "sum"),
    })

    return resumo


def juntar_faturamento(resumo, faturado, validacao):
    resultado = resumo.merge(
        faturado,
        on="Periodo",
        how="left",
        validate=validacao,
    )

    denominador = resultado["Faturado"].where(
        resultado["Faturado"].gt(0)
    )

    resultado["Percentual de Inadimplência"] = (
        resultado["Em aberto"]
        .div(denominador)
        .mul(100)
        .round(2)
    )

    resultado["Situação faturamento"] = np.select(
        [
            resultado["Faturado"].isna(),
            resultado["Faturado"].eq(0),
            resultado["Faturado"].lt(0),
        ],
        [
            "Sem faturamento informado",
            "Faturamento zero",
            "Faturamento negativo",
        ],
        default="Válido",
    )

    resultado["Comp"] = resultado["Periodo"].dt.strftime("%m/%Y")

    if "Classificação" in resultado.columns:
        resultado["_ordem_faixa"] = (
            resultado["Classificação"]
            .map(ORDEM_CLASSIFICACAO)
            .fillna(999)
        )

        resultado = resultado.sort_values(
            ["Periodo", "_ordem_faixa"]
        ).drop(columns="_ordem_faixa")
    else:
        resultado = resultado.sort_values("Periodo")

    principais = [
        "Comp",
        "Em aberto",
        "Faturado",
        "Percentual de Inadimplência",
        "Dias de atraso",
    ]

    if "Classificação" in resultado.columns:
        principais.append("Classificação")

    restantes = [
        coluna
        for coluna in resultado.columns
        if coluna not in principais and coluna != "Periodo"
    ]

    return resultado[principais + restantes].reset_index(drop=True)


# =========================================================
# MONTAGEM DOS RELATÓRIOS
# =========================================================


def montar_relatorios(titulos, faturado, data_referencia=None):
    """Prepara os relatórios usando competência + corte diário.

    A competência define a faixa do indicador. A data de referência também
    define quais títulos já perderam a validade e entram no valor realizado.
    Assim, a competência Atual cresce ao longo do mês conforme os vencimentos
    efetivamente passam.
    """
    referencia = normalizar_referencia(data_referencia)
    periodo_referencia = referencia.to_period("M")
    periodo_inicio = periodo_referencia - (MESES_ANALISE - 1)

    mensagens = []

    base = preparar_titulos(
        titulos,
        referencia,
        mensagens,
    )

    faturamento = preparar_faturamento(
        faturado,
        mensagens,
    )

    # Indicadores: competência Atual + 11 anteriores.
    # Competências futuras ficam fora do painel principal.
    recorte = base.loc[
        base["Periodo"].between(periodo_inicio, periodo_referencia)
    ].copy()

    if recorte.empty:
        raise ValueError(
            "Não há títulos nas competências do período de análise "
            f"({periodo_inicio.strftime('%m/%Y')} a "
            f"{periodo_referencia.strftime('%m/%Y')})."
        )

    # Como a classificação é por competência, cada competência terá somente
    # uma faixa (Atual, D30, D60, D90 ou D+).
    por_faixa = juntar_faturamento(
        resumir_aberto(
            recorte,
            ["Periodo", "Classificação"],
        ),
        faturamento,
        "many_to_one",
    )

    mensal = juntar_faturamento(
        resumir_aberto(
            recorte,
            ["Periodo"],
        ),
        faturamento,
        "one_to_one",
    )

    sem_percentual = mensal.loc[
        mensal["Situação faturamento"].ne("Válido"),
        "Comp",
    ]

    if not sem_percentual.empty:
        mensagens.append(
            "Percentual não calculado nas competências "
            + ", ".join(sem_percentual)
            + ": faturamento ausente, zero ou negativo. "
            "Os valores em aberto foram mantidos."
        )

    faturamento_recorte = faturamento.loc[
        faturamento["Periodo"].between(
            periodo_inicio,
            periodo_referencia,
        )
    ].copy()

    faturamento_recorte["Comp"] = (
        faturamento_recorte["Periodo"].dt.strftime("%m/%Y")
    )

    return {
        "planilha": por_faixa,
        "titulos_antes_d30": selecionar_titulos_antes_d30(base),
        "resumo_mensal": mensal,
        "base_titulos": base,
        "faturado_por_mes": faturamento_recorte[
            ["Comp", "Faturado"]
        ].reset_index(drop=True),
        "avisos": mensagens,
        "data_referencia": referencia,
        "periodo_referencia": periodo_referencia,
        "competencia_referencia": periodo_referencia.strftime("%m/%Y"),
    }


def gerar_relatorios(
    data_referencia=None,
    arquivo_titulos=None,
    arquivo_faturado=None,
):
    caminho_titulos = (
        Path(arquivo_titulos)
        if arquivo_titulos is not None
        else ARQUIVO_TITULOS
    )

    caminho_faturado = (
        Path(arquivo_faturado)
        if arquivo_faturado is not None
        else ARQUIVO_FATURADO
    )

    validar_arquivo(caminho_titulos)
    validar_arquivo(caminho_faturado)

    return montar_relatorios(
        pd.read_excel(caminho_titulos),
        pd.read_excel(
            caminho_faturado,
            sheet_name=ABA_FATURADO,
        ),
        data_referencia,
    )


def gerar_bases(data_referencia=None):
    """Mantém o retorno de dois DataFrames usado nos códigos anteriores."""
    relatorios = gerar_relatorios(data_referencia)
    return (
        relatorios["planilha"],
        relatorios["titulos_antes_d30"],
    )


def gerar_planilha(data_referencia=None):
    return gerar_relatorios(data_referencia)["planilha"]


# =========================================================
# IMPORTAÇÕES UTILIZADAS PELO APP
# =========================================================

_relatorios = gerar_relatorios()

planilha = _relatorios["planilha"]
titulos_antes_d30 = _relatorios["titulos_antes_d30"]
resumo_mensal = _relatorios["resumo_mensal"]
base_titulos = _relatorios["base_titulos"]
faturado_por_mes = _relatorios["faturado_por_mes"]
avisos = _relatorios["avisos"]

DATA_REFERENCIA = _relatorios["data_referencia"]
PERIODO_REFERENCIA = _relatorios["periodo_referencia"]
COMPETENCIA_REFERENCIA = _relatorios["competencia_referencia"]

# Aliases mantidos apenas para não quebrar versões antigas do app durante deploy.
# Na regra nova, eles representam a competência de referência atual, e não o
# último mês encerrado.
DATA_FECHAMENTO = DATA_REFERENCIA
COMPETENCIA_FECHADA = COMPETENCIA_REFERENCIA

for aviso in avisos:
    warnings.warn(
        aviso,
        UserWarning,
        stacklevel=1,
    )


if __name__ == "__main__":
    print(f"Títulos: {ARQUIVO_TITULOS}")
    print(
        f"Empresa: {EMPRESA_ALVO} | "
        f"Data de referência: {DATA_REFERENCIA:%d/%m/%Y} | "
        f"Competência Atual: {COMPETENCIA_REFERENCIA}"
    )

    print("\nResumo por competência e classificação:")
    print(planilha.to_string(index=False))

    print("\nResumo mensal:")
    print(resumo_mensal.to_string(index=False))

    print("\nTítulos da competência Atual (antes do D30):")
    print(titulos_antes_d30.to_string(index=False))
"""Bases de inadimplência da Mediatorie.

Regra oficial dos indicadores por competência:
    competência atual        -> Atual
    1 competência anterior   -> D30
    2 competências anteriores-> D60
    3 competências anteriores-> D90
    4 ou mais anteriores     -> D+

Exemplo em 09/2026:
    09/2026 -> Atual
    08/2026 -> D30
    07/2026 -> D60
    06/2026 -> D90
    05/2026 ou anterior -> D+

A FAIXA (Atual/D30/D60/D90/D+) é definida pela competência.
O VALOR do indicador, porém, é calculado diariamente: só entra no numerador
o título cuja Data vencimento líq. já perdeu a validade (vencimento < data de referência).
Títulos a vencer e títulos que vencem hoje NÃO entram em "Em aberto" do indicador.

Saídas disponíveis para importação pelo app:
    planilha             : uma linha por competência e classificação.
    titulos_antes_d30     : títulos da competência Atual.
    resumo_mensal        : uma linha por competência, sem duplicar faturamento.
    base_titulos         : todos os registros elegíveis, classificados por competência.
    faturado_por_mes     : faturamento consolidado por competência.
    avisos               : inconsistências encontradas na preparação.

O percentual mantém a fórmula do projeto:
    Valor vencido / Faturado * 100

A competência define a faixa; o dia define quais títulos daquela competência
já venceram e, portanto, entram no valor do indicador.
"""

from pathlib import Path
import re
import warnings

import numpy as np
import pandas as pd


# =========================================================
# CAMINHOS
# =========================================================

PASTA_PROJETO = Path(__file__).resolve().parent
ARQUIVO_TITULOS = PASTA_PROJETO / "CustomerLineItems (25).xlsx"
ARQUIVO_FATURADO = PASTA_PROJETO / "teste.xlsx"
ABA_FATURADO = "Planilha1"


# =========================================================
# CONFIGURAÇÕES
# =========================================================

EMPRESA_ALVO = 1000
CODIGOS_EXCLUIR = ["1000", "1800", "1600", "10000"]

COLUNA_EMPRESA = "Empresa"
COLUNA_LANCAMENTO = "Lançamento contábil"
COLUNA_MONTANTE = "Montante (ME)"
COLUNA_DATA = "Data vencimento líq."

# Competência atual + 11 competências anteriores.
MESES_ANALISE = 12

ORDEM_CLASSIFICACAO = {
    "Atual": 0,
    "D30": 1,
    "D60": 2,
    "D90": 3,
    "D+": 4,
    "Futura": 5,
}

COLUNAS_DETALHE_PREFERENCIAIS = [
    COLUNA_EMPRESA,
    "Cliente",
    "Nome do cliente",
    "Conta",
    "Nº documento",
    "Documento",
    "Atribuição",
    "Referência",
    COLUNA_LANCAMENTO,
    COLUNA_DATA,
    COLUNA_MONTANTE,
    "Dt.lançamento",
    "Data de vencimento",
]


# =========================================================
# VALIDAÇÕES E CONVERSÕES
# =========================================================


def validar_arquivo(caminho):
    caminho = Path(caminho)
    if not caminho.is_file():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {caminho}\n"
            "Coloque as planilhas na pasta deste bases.py ou ajuste "
            "ARQUIVO_TITULOS e ARQUIVO_FATURADO no início do código."
        )


def validar_colunas(dataframe, colunas, nome_arquivo):
    ausentes = [coluna for coluna in colunas if coluna not in dataframe.columns]
    if ausentes:
        raise KeyError(f"Colunas ausentes em {nome_arquivo}: {ausentes}")


def normalizar_cabecalhos(dataframe):
    dados = dataframe.copy()
    dados.columns = dados.columns.astype(str).str.strip()

    if dados.columns.duplicated().any():
        raise ValueError("A planilha contém nomes de colunas repetidos.")

    return dados


def converter_datas(valores):
    """Converte datas Excel/texto brasileiro/ISO. Datas inválidas viram NaT."""
    return pd.to_datetime(
        valores,
        format="mixed",
        dayfirst=True,
        errors="coerce",
    ).dt.normalize()


def normalizar_referencia(data_referencia=None):
    """Data usada para saber qual é a competência Atual."""
    if data_referencia is None:
        return pd.Timestamp.today().normalize()

    referencia = pd.to_datetime(
        data_referencia,
        format="mixed",
        dayfirst=True,
        errors="coerce",
    )

    if pd.isna(referencia):
        raise ValueError("A data de referência é inválida.")

    return referencia.normalize()


def obter_periodo_referencia(data_referencia=None):
    """Retorna a competência correspondente à data de execução/simulação."""
    return normalizar_referencia(data_referencia).to_period("M")


# =========================================================
# REGRAS DE PRAZO E CLASSIFICAÇÃO
# =========================================================


def calcular_dias_atraso(data_vencimento, data_referencia=None):
    """Dias vencidos para consulta operacional; não define a faixa do indicador."""
    vencimento = converter_datas(pd.Series([data_vencimento])).iloc[0]

    if pd.isna(vencimento):
        return pd.NA

    return max((normalizar_referencia(data_referencia) - vencimento).days, 0)


def adicionar_prazos(titulos, data_referencia=None):
    """Calcula campos operacionais por título, sem definir a faixa do indicador."""
    referencia = normalizar_referencia(data_referencia)

    dados = titulos.copy()
    dados[COLUNA_DATA] = converter_datas(dados[COLUNA_DATA])
    dados["data_hoje"] = referencia
    dados["Periodo"] = dados[COLUNA_DATA].dt.to_period("M")
    dados["Comp"] = dados["Periodo"].dt.strftime("%m/%Y")

    # Negativo = ainda vai vencer; zero = vence hoje; positivo = vencido.
    dias = (referencia - dados[COLUNA_DATA]).dt.days.astype("Int64")
    dados["Contagem_dias"] = dias

    dados["Contagem_Prazo"] = np.select(
        [
            dias.lt(0).to_numpy(dtype=bool, na_value=False),
            dias.eq(0).to_numpy(dtype=bool, na_value=False),
            dias.gt(0).to_numpy(dtype=bool, na_value=False),
        ],
        ["A vencer", "Vence hoje", "Vencido"],
        default="Data inválida",
    )

    dados["Dias de atraso"] = dias.clip(lower=0)
    dados["Dias para vencimento"] = (-dias).clip(lower=0)

    # Pela regra por competência, todos os títulos do mês entram em D30
    # juntos no primeiro dia da competência seguinte.
    dados["Data de entrada no D30"] = (
        dados["Periodo"] + 1
    ).dt.to_timestamp()

    dados["Dias para D30"] = (
        (dados["Data de entrada no D30"] - referencia)
        .dt.days
        .clip(lower=0)
        .astype("Int64")
    )

    return dados


def classificar_por_competencia(titulos, data_referencia=None):
    """Classifica Atual/D30/D60/D90/D+ exclusivamente pela competência.

    Exemplo se a data de referência estiver em 09/2026:
        09/2026 -> Atual
        08/2026 -> D30
        07/2026 -> D60
        06/2026 -> D90
        <= 05/2026 -> D+
        >= 10/2026 -> Futura
    """
    referencia = normalizar_referencia(data_referencia)
    dados = titulos.copy()

    if "Periodo" not in dados.columns:
        dados["Periodo"] = converter_datas(dados[COLUNA_DATA]).dt.to_period("M")

    meses_atraso = (
        (referencia.year - dados["Periodo"].dt.year) * 12
        + (referencia.month - dados["Periodo"].dt.month)
    ).astype("Int64")

    dados["Meses de atraso"] = meses_atraso

    dados["Classificação"] = np.select(
        [
            meses_atraso.eq(0).to_numpy(dtype=bool, na_value=False),
            meses_atraso.eq(1).to_numpy(dtype=bool, na_value=False),
            meses_atraso.eq(2).to_numpy(dtype=bool, na_value=False),
            meses_atraso.eq(3).to_numpy(dtype=bool, na_value=False),
            meses_atraso.ge(4).to_numpy(dtype=bool, na_value=False),
            meses_atraso.lt(0).to_numpy(dtype=bool, na_value=False),
        ],
        ["Atual", "D30", "D60", "D90", "D+", "Futura"],
        default="Data inválida",
    )

    return dados


# =========================================================
# PREPARAÇÃO DOS TÍTULOS
# =========================================================


def preparar_titulos(titulos, data_referencia, avisos_saida):
    dados = normalizar_cabecalhos(titulos)

    validar_colunas(
        dados,
        [COLUNA_EMPRESA, COLUNA_LANCAMENTO, COLUNA_MONTANTE, COLUNA_DATA],
        ARQUIVO_TITULOS.name,
    )

    dados[COLUNA_EMPRESA] = pd.to_numeric(
        dados[COLUNA_EMPRESA],
        errors="coerce",
    )

    dados[COLUNA_LANCAMENTO] = (
        dados[COLUNA_LANCAMENTO]
        .astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    dados[COLUNA_MONTANTE] = pd.to_numeric(
        dados[COLUNA_MONTANTE],
        errors="coerce",
    ).replace([np.inf, -np.inf], np.nan)

    dados[COLUNA_DATA] = converter_datas(dados[COLUNA_DATA])

    # Mantém a regra do projeto: excluir códigos por trecho contido no texto.
    padrao = "|".join(re.escape(codigo) for codigo in CODIGOS_EXCLUIR)
    excluidos = dados[COLUNA_LANCAMENTO].str.contains(padrao, na=False)

    empresa_e_codigo = (
        dados[COLUNA_EMPRESA].eq(EMPRESA_ALVO)
        & ~excluidos
    )

    invalidos = empresa_e_codigo & (
        dados[COLUNA_DATA].isna()
        | dados[COLUNA_MONTANTE].isna()
    )

    if invalidos.any():
        avisos_saida.append(
            f"{invalidos.sum()} registro(s) descartado(s) por vencimento "
            "ou montante ausente/inválido."
        )

    dados = dados.loc[
        empresa_e_codigo
        & dados[COLUNA_MONTANTE].ge(0)
        & dados[COLUNA_DATA].notna()
    ].copy()

    if dados.empty:
        raise ValueError(
            "Não há títulos válidos após os filtros da empresa e dos lançamentos."
        )

    # Não deduplica: cada linha elegível do Excel é preservada.
    dados = adicionar_prazos(dados, data_referencia)
    dados = classificar_por_competencia(dados, data_referencia)

    return dados.reset_index(drop=True)


def selecionar_titulos_antes_d30(titulos_classificados):
    """Retorna a competência Atual, que migrará inteira para D30 no próximo mês."""
    detalhe = titulos_classificados.loc[
        titulos_classificados["Classificação"].eq("Atual")
    ].copy()

    calculadas = [
        "Contagem_Prazo",
        "Contagem_dias",
        "Dias de atraso",
        "Dias para vencimento",
        "Dias para D30",
        "Data de entrada no D30",
        "Meses de atraso",
        "Classificação",
        "data_hoje",
    ]

    preferenciais = [
        coluna
        for coluna in COLUNAS_DETALHE_PREFERENCIAIS
        if coluna in detalhe.columns
    ]

    colunas = list(dict.fromkeys(["Comp"] + preferenciais + calculadas))

    # Mantém identificadores adicionais existentes no arquivo original.
    colunas += [
        coluna
        for coluna in detalhe.columns
        if coluna not in colunas and coluna != "Periodo"
    ]

    return (
        detalhe
        .sort_values(
            ["Data de entrada no D30", COLUNA_DATA],
            kind="stable",
        )[colunas]
        .rename(columns={COLUNA_MONTANTE: "Em aberto"})
        .reset_index(drop=True)
    )


def gerar_titulos_antes_d30(titulos, data_referencia=None):
    """Compatibilidade: recebe títulos e aplica a regra mensal por competência."""
    referencia = normalizar_referencia(data_referencia)
    dados = adicionar_prazos(titulos, referencia)
    dados = classificar_por_competencia(dados, referencia)
    return selecionar_titulos_antes_d30(dados)


# =========================================================
# FATURAMENTO
# =========================================================


def preparar_faturamento(faturado, avisos_saida):
    dados = normalizar_cabecalhos(faturado)

    validar_colunas(
        dados,
        ["mês", "faturado"],
        ARQUIVO_FATURADO.name,
    )

    # Aceita competência como data Excel, texto de data ou MM/AAAA.
    datas = dados["mês"].copy()
    texto = datas.astype("string").str.strip()
    so_mes = texto.str.fullmatch(r"\d{1,2}/\d{4}", na=False)

    datas = datas.astype(object)
    datas.loc[so_mes] = "01/" + texto.loc[so_mes]

    dados["Periodo"] = converter_datas(datas).dt.to_period("M")
    dados["faturado"] = pd.to_numeric(
        dados["faturado"],
        errors="coerce",
    ).replace([np.inf, -np.inf], np.nan)

    invalidos = dados["Periodo"].isna() | dados["faturado"].isna()

    if invalidos.any():
        avisos_saida.append(
            f"{invalidos.sum()} linha(s) de faturamento ignorada(s) "
            "por data ou valor inválido."
        )

    return (
        dados.loc[~invalidos]
        .groupby("Periodo", as_index=False)["faturado"]
        .sum()
        .rename(columns={"faturado": "Faturado"})
    )


# =========================================================
# RESUMOS
# =========================================================


def resumir_aberto(titulos, chaves):
    """Resume os títulos preservando competência e corte diário.

    Regra do indicador:
    - competência define Atual/D30/D60/D90/D+;
    - somente vencimento ANTERIOR à data de referência entra em "Em aberto";
    - título que vence hoje ainda não perdeu a validade e não entra no numerador;
    - título a vencer também não entra no numerador.
    """
    dados = titulos.copy()

    vencido = dados["Contagem_dias"].gt(0)
    vence_hoje = dados["Contagem_dias"].eq(0)
    a_vencer = dados["Contagem_dias"].lt(0)

    dados["Valor a vencer"] = dados[COLUNA_MONTANTE].where(a_vencer, 0)
    dados["Valor vence hoje"] = dados[COLUNA_MONTANTE].where(vence_hoje, 0)
    dados["Valor vencido"] = dados[COLUNA_MONTANTE].where(vencido, 0)
    dados["Registro vencido"] = vencido.astype("int64")

    resumo = dados.groupby(chaves, as_index=False).agg(**{
        # IMPORTANTE: o valor do indicador é somente o que já perdeu a validade.
        "Em aberto": ("Valor vencido", "sum"),
        "Quantidade vencida": ("Registro vencido", "sum"),
        "Quantidade de registros": (COLUNA_MONTANTE, "size"),
        "Menor atraso": ("Dias de atraso", "min"),
        "Dias de atraso": ("Dias de atraso", "max"),
        "Valor a vencer": ("Valor a vencer", "sum"),
        "Valor vence hoje": ("Valor vence hoje", "sum"),
        "Valor vencido": ("Valor vencido", "sum"),
    })

    return resumo


def juntar_faturamento(resumo, faturado, validacao):
    resultado = resumo.merge(
        faturado,
        on="Periodo",
        how="left",
        validate=validacao,
    )

    denominador = resultado["Faturado"].where(
        resultado["Faturado"].gt(0)
    )

    resultado["Percentual de Inadimplência"] = (
        resultado["Em aberto"]
        .div(denominador)
        .mul(100)
        .round(2)
    )

    resultado["Situação faturamento"] = np.select(
        [
            resultado["Faturado"].isna(),
            resultado["Faturado"].eq(0),
            resultado["Faturado"].lt(0),
        ],
        [
            "Sem faturamento informado",
            "Faturamento zero",
            "Faturamento negativo",
        ],
        default="Válido",
    )

    resultado["Comp"] = resultado["Periodo"].dt.strftime("%m/%Y")

    if "Classificação" in resultado.columns:
        resultado["_ordem_faixa"] = (
            resultado["Classificação"]
            .map(ORDEM_CLASSIFICACAO)
            .fillna(999)
        )

        resultado = resultado.sort_values(
            ["Periodo", "_ordem_faixa"]
        ).drop(columns="_ordem_faixa")
    else:
        resultado = resultado.sort_values("Periodo")

    principais = [
        "Comp",
        "Em aberto",
        "Faturado",
        "Percentual de Inadimplência",
        "Dias de atraso",
    ]

    if "Classificação" in resultado.columns:
        principais.append("Classificação")

    restantes = [
        coluna
        for coluna in resultado.columns
        if coluna not in principais and coluna != "Periodo"
    ]

    return resultado[principais + restantes].reset_index(drop=True)


# =========================================================
# MONTAGEM DOS RELATÓRIOS
# =========================================================


def montar_relatorios(titulos, faturado, data_referencia=None):
    """Prepara os relatórios usando competência + corte diário.

    A competência define a faixa do indicador. A data de referência também
    define quais títulos já perderam a validade e entram no valor realizado.
    Assim, a competência Atual cresce ao longo do mês conforme os vencimentos
    efetivamente passam.
    """
    referencia = normalizar_referencia(data_referencia)
    periodo_referencia = referencia.to_period("M")
    periodo_inicio = periodo_referencia - (MESES_ANALISE - 1)

    mensagens = []

    base = preparar_titulos(
        titulos,
        referencia,
        mensagens,
    )

    faturamento = preparar_faturamento(
        faturado,
        mensagens,
    )

    # Indicadores: competência Atual + 11 anteriores.
    # Competências futuras ficam fora do painel principal.
    recorte = base.loc[
        base["Periodo"].between(periodo_inicio, periodo_referencia)
    ].copy()

    if recorte.empty:
        raise ValueError(
            "Não há títulos nas competências do período de análise "
            f"({periodo_inicio.strftime('%m/%Y')} a "
            f"{periodo_referencia.strftime('%m/%Y')})."
        )

    # Como a classificação é por competência, cada competência terá somente
    # uma faixa (Atual, D30, D60, D90 ou D+).
    por_faixa = juntar_faturamento(
        resumir_aberto(
            recorte,
            ["Periodo", "Classificação"],
        ),
        faturamento,
        "many_to_one",
    )

    mensal = juntar_faturamento(
        resumir_aberto(
            recorte,
            ["Periodo"],
        ),
        faturamento,
        "one_to_one",
    )

    sem_percentual = mensal.loc[
        mensal["Situação faturamento"].ne("Válido"),
        "Comp",
    ]

    if not sem_percentual.empty:
        mensagens.append(
            "Percentual não calculado nas competências "
            + ", ".join(sem_percentual)
            + ": faturamento ausente, zero ou negativo. "
            "Os valores em aberto foram mantidos."
        )

    faturamento_recorte = faturamento.loc[
        faturamento["Periodo"].between(
            periodo_inicio,
            periodo_referencia,
        )
    ].copy()

    faturamento_recorte["Comp"] = (
        faturamento_recorte["Periodo"].dt.strftime("%m/%Y")
    )

    return {
        "planilha": por_faixa,
        "titulos_antes_d30": selecionar_titulos_antes_d30(base),
        "resumo_mensal": mensal,
        "base_titulos": base,
        "faturado_por_mes": faturamento_recorte[
            ["Comp", "Faturado"]
        ].reset_index(drop=True),
        "avisos": mensagens,
        "data_referencia": referencia,
        "periodo_referencia": periodo_referencia,
        "competencia_referencia": periodo_referencia.strftime("%m/%Y"),
    }


def gerar_relatorios(
    data_referencia=None,
    arquivo_titulos=None,
    arquivo_faturado=None,
):
    caminho_titulos = (
        Path(arquivo_titulos)
        if arquivo_titulos is not None
        else ARQUIVO_TITULOS
    )

    caminho_faturado = (
        Path(arquivo_faturado)
        if arquivo_faturado is not None
        else ARQUIVO_FATURADO
    )

    validar_arquivo(caminho_titulos)
    validar_arquivo(caminho_faturado)

    return montar_relatorios(
        pd.read_excel(caminho_titulos),
        pd.read_excel(
            caminho_faturado,
            sheet_name=ABA_FATURADO,
        ),
        data_referencia,
    )


def gerar_bases(data_referencia=None):
    """Mantém o retorno de dois DataFrames usado nos códigos anteriores."""
    relatorios = gerar_relatorios(data_referencia)
    return (
        relatorios["planilha"],
        relatorios["titulos_antes_d30"],
    )


def gerar_planilha(data_referencia=None):
    return gerar_relatorios(data_referencia)["planilha"]


# =========================================================
# IMPORTAÇÕES UTILIZADAS PELO APP
# =========================================================

_relatorios = gerar_relatorios()

planilha = _relatorios["planilha"]
titulos_antes_d30 = _relatorios["titulos_antes_d30"]
resumo_mensal = _relatorios["resumo_mensal"]
base_titulos = _relatorios["base_titulos"]
faturado_por_mes = _relatorios["faturado_por_mes"]
avisos = _relatorios["avisos"]

DATA_REFERENCIA = _relatorios["data_referencia"]
PERIODO_REFERENCIA = _relatorios["periodo_referencia"]
COMPETENCIA_REFERENCIA = _relatorios["competencia_referencia"]

# Aliases mantidos apenas para não quebrar versões antigas do app durante deploy.
# Na regra nova, eles representam a competência de referência atual, e não o
# último mês encerrado.
DATA_FECHAMENTO = DATA_REFERENCIA
COMPETENCIA_FECHADA = COMPETENCIA_REFERENCIA

for aviso in avisos:
    warnings.warn(
        aviso,
        UserWarning,
        stacklevel=1,
    )


if __name__ == "__main__":
    print(f"Títulos: {ARQUIVO_TITULOS}")
    print(
        f"Empresa: {EMPRESA_ALVO} | "
        f"Data de referência: {DATA_REFERENCIA:%d/%m/%Y} | "
        f"Competência Atual: {COMPETENCIA_REFERENCIA}"
    )

    print("\nResumo por competência e classificação:")
    print(planilha.to_string(index=False))

    print("\nResumo mensal:")
    print(resumo_mensal.to_string(index=False))

    print("\nTítulos da competência Atual (antes do D30):")
    print(titulos_antes_d30.to_string(index=False))
