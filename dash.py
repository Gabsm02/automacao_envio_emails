"""
Dashboard do histórico de envios às detentoras.

Instalação: pip install -U streamlit plotly pandas openpyxl
Execução:   streamlit run dashboard.py

Deixe o arquivo historico_envios.xlsx na mesma pasta deste script.
"""

import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =====================================================
# CONFIGURAÇÕES
# =====================================================
ARQUIVO = Path(__file__).parent / "historico_envios.xlsx"
COLUNAS_OBRIGATORIAS = ["UFSIGLA", "DETENTORA", "DATA_ENVIO", "CORRIGIDO"]
LIMITE_ATRASO_DIAS = 30

CORRIGIDO = "Corrigido"
PENDENTE = "Pendente"
SEM_STATUS = "Sem status"
ORDEM_STATUS = [CORRIGIDO, PENDENTE, SEM_STATUS]
CORES_STATUS = {
    CORRIGIDO: "#0CA30C",
    PENDENTE: "#EB3C7D",
    SEM_STATUS: "#6250D6",
}

FAIXAS_ATRASO = {
    "0–7": (0, 7),
    "8–15": (8, 15),
    "16–30": (16, 30),
    "31–60": (31, 60),
    "60+": (61, None),
}

FORMATO_DATA = "DD/MM/YYYY HH:mm"
REGEX_EMAIL = re.compile(r"[^@\s;,]+@[^@\s;,]+\.[A-Za-z]{2,}")

st.set_page_config(
    page_title="Histórico de Envios",
    page_icon="📧",
    layout="wide",
)


# =====================================================
# CARREGAMENTO E TRATAMENTO
# =====================================================
def converter_data(valor):
    """Converte uma célula, que pode vir como data do Excel ou texto dd/mm/aaaa."""
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, pd.Timestamp) or hasattr(valor, "year"):
        return pd.Timestamp(valor)
    return pd.to_datetime(str(valor).strip(), dayfirst=True, errors="coerce")


def padronizar_status(valor):
    if pd.isna(valor):
        return SEM_STATUS
    texto = str(valor).strip().upper()
    if texto == "SIM":
        return CORRIGIDO
    if texto in ("NAO", "NÃO"):
        return PENDENTE
    return SEM_STATUS


@st.cache_data(show_spinner="Lendo a planilha...")
def carregar_dados(modificado_em):
    # "modificado_em" faz o cache ser refeito sempre que o arquivo é salvo
    df = pd.read_excel(ARQUIVO)
    df.columns = [str(coluna).strip().upper() for coluna in df.columns]

    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Coluna(s) não encontrada(s): {', '.join(faltando)}. "
            f"Colunas da planilha: {', '.join(df.columns)}"
        )

    for coluna in ["ID DETENTORA", "EMAIL", "DATA_COBRANCA"]:
        if coluna not in df.columns:
            df[coluna] = pd.NA

    # Colunas de texto: o ID pode misturar números e letras
    for coluna in ["UFSIGLA", "DETENTORA", "ID DETENTORA", "EMAIL"]:
        df[coluna] = df[coluna].map(
            lambda valor: pd.NA if pd.isna(valor) else str(valor).strip()
        ).astype("string")

    df["DETENTORA"] = df["DETENTORA"].str.upper().fillna("NÃO INFORMADA")
    df["UF"] = df["UFSIGLA"].str[:2].str.upper()

    # A planilha mistura datas do Excel com textos no formato dd/mm/aaaa
    for coluna in ["DATA_ENVIO", "DATA_COBRANCA"]:
        df[coluna] = pd.to_datetime(df[coluna].map(converter_data), errors="coerce")

    df["STATUS"] = df["CORRIGIDO"].map(padronizar_status)
    return df


def carregar_ou_parar():
    try:
        modificado_em = ARQUIVO.stat().st_mtime
        return carregar_dados(modificado_em), modificado_em
    except FileNotFoundError:
        st.error(f"Planilha não encontrada em: {ARQUIVO}")
    except PermissionError:
        st.error("Não foi possível ler a planilha. Feche o arquivo no Excel e atualize.")
    except KeyError as erro:
        st.error(erro.args[0])
    st.stop()


def calcular_campos_de_prazo(df):
    # Fica fora do cache para os dias continuarem certos de um dia para o outro
    df = df.copy()
    agora = pd.Timestamp.now()
    df["DIAS_DESDE_ENVIO"] = (agora - df["DATA_ENVIO"]).dt.days
    ultimo_contato = df[["DATA_ENVIO", "DATA_COBRANCA"]].max(axis=1)
    df["DIAS_SEM_CONTATO"] = (agora - ultimo_contato).dt.days
    return df


# =====================================================
# GRÁFICOS
# =====================================================
def cor_de_fundo():
    try:
        return "#171614" if st.context.theme.type == "dark" else "#FFFFFF"
    except AttributeError:
        return "#FFFFFF"


def estilizar(fig, altura):
    fig.update_layout(
        height=altura,
        margin=dict(l=8, r=16, t=76, b=8),
        barcornerradius=4,
        hoverlabel=dict(namelength=-1),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.01, x=0,
            title_text="", traceorder="normal",
        ),
    )
    return fig


def contar_por_status(df, coluna):
    tabela = (
        df.groupby([coluna, "STATUS"]).size().unstack(fill_value=0)
        .reindex(columns=ORDEM_STATUS, fill_value=0)
    )
    return tabela


def grafico_status_horizontal(df, coluna, titulo):
    tabela = contar_por_status(df, coluna).sort_values(
        [PENDENTE, CORRIGIDO], ascending=False
    )
    fig = go.Figure()
    for status in ORDEM_STATUS:
        if tabela[status].sum() == 0:
            continue
        valores = tabela[status]
        fig.add_bar(
            name=status,
            orientation="h",
            y=tabela.index,
            x=valores,
            marker=dict(
                color=CORES_STATUS[status],
                line=dict(color=cor_de_fundo(), width=2),
            ),
            text=[str(v) if v else "" for v in valores],
            textposition="inside",
            insidetextanchor="middle",
            textangle=0,
            textfont=dict(color="#FFFFFF"),
            hovertemplate=f"%{{y}}<br>{status}: %{{x}}<extra></extra>",
        )
    fig.update_layout(barmode="stack", title=titulo)
    fig.update_yaxes(autorange="reversed", title=None, type="category")
    fig.update_xaxes(title="Registros", dtick=5 if tabela.sum(axis=1).max() > 10 else 1)
    return estilizar(fig, max(260, 70 + 36 * len(tabela)))


def grafico_atraso(pendentes):
    contagem = {}
    for faixa, (minimo, maximo) in FAIXAS_ATRASO.items():
        filtro = pendentes["DIAS_DESDE_ENVIO"] >= minimo
        if maximo is not None:
            filtro &= pendentes["DIAS_DESDE_ENVIO"] <= maximo
        contagem[faixa] = int(filtro.sum())

    fig = go.Figure(
        go.Bar(
            x=list(contagem.keys()),
            y=list(contagem.values()),
            marker_color=CORES_STATUS[PENDENTE],
            text=list(contagem.values()),
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{x} dias<br>Pendentes: %{y}<extra></extra>",
        )
    )
    fig.update_layout(title="Há quanto tempo as pendências estão abertas", bargap=0.35)
    fig.update_xaxes(title="Dias desde o envio", type="category")
    fig.update_yaxes(title="Pendentes", rangemode="tozero")
    return estilizar(fig, 400)


def grafico_envios_semana(df):
    dados = df.dropna(subset=["DATA_ENVIO"]).copy()
    dados["SEMANA"] = dados["DATA_ENVIO"].dt.to_period("W-SUN").dt.start_time
    tabela = contar_por_status(dados, "SEMANA")
    semanas = pd.date_range(tabela.index.min(), tabela.index.max(), freq="7D")
    tabela = tabela.reindex(semanas, fill_value=0)
    rotulos = semanas.strftime("%d/%m")

    fig = go.Figure()
    for status in ORDEM_STATUS:
        if tabela[status].sum() == 0:
            continue
        fig.add_bar(
            name=status,
            x=rotulos,
            y=tabela[status],
            marker=dict(
                color=CORES_STATUS[status],
                line=dict(color=cor_de_fundo(), width=2),
            ),
            hovertemplate=f"Semana de %{{x}}<br>{status}: %{{y}}<extra></extra>",
        )
    fig.update_layout(barmode="stack", title="Envios por semana", bargap=0.3)
    fig.update_xaxes(title="Início da semana (segunda-feira)", type="category")
    fig.update_yaxes(title="Envios")
    return estilizar(fig, 380)


# =====================================================
# TABELAS
# =====================================================
def resumo_por_detentora(df):
    tabela = contar_por_status(df, "DETENTORA")
    tabela["Total"] = tabela.sum(axis=1)
    com_status = tabela[CORRIGIDO] + tabela[PENDENTE]
    tabela["Taxa de correção"] = (tabela[CORRIGIDO] / com_status.where(com_status > 0)) * 100

    pendentes = df[df["STATUS"] == PENDENTE]
    tabela["Pendência mais antiga (dias)"] = pendentes.groupby("DETENTORA")["DIAS_DESDE_ENVIO"].max()
    tabela["Pendentes sem cobrança"] = (
        pendentes[pendentes["DATA_COBRANCA"].isna()].groupby("DETENTORA").size()
    )
    tabela["Pendentes sem cobrança"] = tabela["Pendentes sem cobrança"].fillna(0).astype(int)

    colunas = ["Total", CORRIGIDO, PENDENTE, SEM_STATUS, "Taxa de correção",
               "Pendência mais antiga (dias)", "Pendentes sem cobrança"]
    return (
        tabela[colunas].sort_values([PENDENTE, "Total"], ascending=False)
        .reset_index().rename(columns={"DETENTORA": "Detentora"})
    )


def tabela_de_registros(df):
    return df.rename(columns={
        "UFSIGLA": "Sigla",
        "DETENTORA": "Detentora",
        "ID DETENTORA": "ID Detentora",
        "EMAIL": "E-mails",
        "DATA_ENVIO": "Enviado em",
        "DATA_COBRANCA": "Última cobrança",
        "STATUS": "Status",
        "DIAS_DESDE_ENVIO": "Dias desde o envio",
        "DIAS_SEM_CONTATO": "Dias sem contato",
    })


COLUNAS_DATA = {
    "Enviado em": st.column_config.DatetimeColumn(format=FORMATO_DATA),
    "Última cobrança": st.column_config.DatetimeColumn(format=FORMATO_DATA),
}


def gerar_csv(df):
    return df.to_csv(index=False, sep=";", date_format="%d/%m/%Y %H:%M").encode("utf-8-sig")


def email_com_problema(valor):
    if pd.isna(valor):
        return True
    enderecos = [e.strip() for e in re.split(r"[;,]", str(valor)) if e.strip()]
    return not enderecos or any(not REGEX_EMAIL.fullmatch(e) for e in enderecos)


def verificar_qualidade(df):
    id_vazio = df["ID DETENTORA"].isna() | (
        df["ID DETENTORA"].fillna("").str.upper().isin(["NÃO INFORMADO", "NAO INFORMADO"])
    )
    return {
        "ID da detentora não informado": df[id_vazio],
        "Status de correção vazio ou desconhecido": df[df["STATUS"] == SEM_STATUS],
        "E-mail vazio ou mal formatado": df[df["EMAIL"].map(email_com_problema)],
        "Data de envio não reconhecida": df[df["DATA_ENVIO"].isna()],
        "Cobrança anterior ao envio": df[df["DATA_COBRANCA"] < df["DATA_ENVIO"]],
        "Sigla repetida": df[df["UFSIGLA"].duplicated(keep=False)],
    }


# =====================================================
# DADOS
# =====================================================
df, modificado_em = carregar_ou_parar()
df = calcular_campos_de_prazo(df)

# =====================================================
# TÍTULO
# =====================================================
st.title("📧 Histórico de Envios")
atualizado_em = pd.Timestamp(modificado_em, unit="s", tz="UTC").tz_convert("America/Bahia")
st.caption(
    f"Planilha salva em {atualizado_em:%d/%m/%Y às %H:%M} · "
    f"{len(df)} registros lidos · envios até {df['DATA_ENVIO'].max():%d/%m/%Y às %H:%M}"
)

# =====================================================
# FILTROS
# =====================================================
st.sidebar.header("Filtros")
st.sidebar.button("🔄 Atualizar dados", width="stretch")

data_minima = df["DATA_ENVIO"].min().date()
data_maxima = df["DATA_ENVIO"].max().date()
periodo = st.sidebar.date_input(
    "Período de envio",
    value=(data_minima, data_maxima),
    min_value=data_minima,
    max_value=data_maxima,
    format="DD/MM/YYYY",
)
detentoras = st.sidebar.multiselect(
    "Detentora", sorted(df["DETENTORA"].unique()), placeholder="Todas"
)
ufs = st.sidebar.multiselect(
    "UF", sorted(df["UF"].dropna().unique()), placeholder="Todas"
)
status = st.sidebar.multiselect(
    "Status da correção", ORDEM_STATUS, placeholder="Todos"
)
busca = st.sidebar.text_input("Buscar sigla ou ID", placeholder="Ex.: BACOU").strip()

df_filtrado = df
if isinstance(periodo, (tuple, list)) and len(periodo) == 2:
    datas_envio = df_filtrado["DATA_ENVIO"].dt.date
    df_filtrado = df_filtrado[
        df_filtrado["DATA_ENVIO"].isna() | datas_envio.between(periodo[0], periodo[1])
    ]
if detentoras:
    df_filtrado = df_filtrado[df_filtrado["DETENTORA"].isin(detentoras)]
if ufs:
    df_filtrado = df_filtrado[df_filtrado["UF"].isin(ufs)]
if status:
    df_filtrado = df_filtrado[df_filtrado["STATUS"].isin(status)]
if busca:
    encontrados = (
        df_filtrado["UFSIGLA"].str.contains(busca, case=False, regex=False, na=False)
        | df_filtrado["ID DETENTORA"].str.contains(busca, case=False, regex=False, na=False)
    )
    df_filtrado = df_filtrado[encontrados]

if df_filtrado.empty:
    st.info("Nenhum registro encontrado com esses filtros.")
    st.stop()

pendentes = df_filtrado[df_filtrado["STATUS"] == PENDENTE]

# =====================================================
# KPIs
# =====================================================
total = len(df_filtrado)
corrigidos = int((df_filtrado["STATUS"] == CORRIGIDO).sum())
total_pendentes = len(pendentes)
com_status = corrigidos + total_pendentes
taxa = f"{corrigidos / com_status:.1%}".replace(".", ",") if com_status else "–"
atrasados = int((pendentes["DIAS_DESDE_ENVIO"] > LIMITE_ATRASO_DIAS).sum())
sem_cobranca = int(pendentes["DATA_COBRANCA"].isna().sum())

colunas_kpi = st.columns(6)
colunas_kpi[0].metric("Total de registros", total)
colunas_kpi[1].metric(
    "Detentoras", df_filtrado["DETENTORA"].nunique()
)
colunas_kpi[2].metric(
    "Corrigidos", corrigidos,
    help="Registros com CORRIGIDO = SIM.",
)
colunas_kpi[3].metric(
    "Taxa de correção", taxa,
    help="Corrigidos ÷ (corrigidos + pendentes). Registros sem status ficam de fora.",
)
colunas_kpi[4].metric(
    f"Pendentes +{LIMITE_ATRASO_DIAS}d", f"{atrasados} de {total_pendentes}",
    help=(
        f"Registros com CORRIGIDO = NAO enviados há mais de {LIMITE_ATRASO_DIAS} dias, "
        "em relação ao total de pendentes."
    ),
)
colunas_kpi[5].metric(
    "Sem cobrança", sem_cobranca,
    help="Registros com CORRIGIDO = NAO e DATA_COBRANCA vazia.",
)

st.divider()

# =====================================================
# GRÁFICOS
# =====================================================
CONFIG_GRAFICO = {"displayModeBar": False}

coluna_1, coluna_2 = st.columns([3, 2])
with coluna_1:
    st.plotly_chart(
        grafico_status_horizontal(df_filtrado, "DETENTORA", "Status por detentora"),
        width="stretch", config=CONFIG_GRAFICO,
    )
with coluna_2:
    if pendentes.empty:
        st.info("Não há pendências nesta seleção.")
    else:
        st.plotly_chart(grafico_atraso(pendentes), width="stretch", config=CONFIG_GRAFICO)

coluna_3, coluna_4 = st.columns([3, 2])
with coluna_3:
    if df_filtrado["DATA_ENVIO"].notna().any():
        st.plotly_chart(grafico_envios_semana(df_filtrado), width="stretch", config=CONFIG_GRAFICO)
with coluna_4:
    st.plotly_chart(
        grafico_status_horizontal(df_filtrado, "UF", "Status por UF"),
        width="stretch", config=CONFIG_GRAFICO,
    )

# =====================================================
# DETALHAMENTO
# =====================================================
problemas = verificar_qualidade(df)
total_alertas = sum(1 for registros in problemas.values() if not registros.empty)

aba_pendencias, aba_detentoras, aba_registros, aba_qualidade = st.tabs([
    f"Pendências ({total_pendentes})",
    "Resumo por detentora",
    f"Todos os registros ({total})",
    f"Qualidade dos dados ({total_alertas})",
])

with aba_pendencias:
    st.caption("As pendências mais antigas aparecem primeiro.")
    lista = tabela_de_registros(
        pendentes.sort_values("DIAS_DESDE_ENVIO", ascending=False)
    )[["Sigla", "UF", "Detentora", "ID Detentora", "Enviado em", "Dias desde o envio",
       "Última cobrança", "Dias sem contato", "E-mails"]]
    st.dataframe(
        lista, hide_index=True, width="stretch", height=420,
        column_config={
            **COLUNAS_DATA,
            "Dias sem contato": st.column_config.NumberColumn(
                help="Dias desde a última cobrança ou, se não houver, desde o envio."
            ),
        },
    )
    st.download_button(
        "📥 Baixar lista de pendências", gerar_csv(lista),
        "pendencias.csv", "text/csv", key="baixar_pendencias",
    )

with aba_detentoras:
    st.dataframe(
        resumo_por_detentora(df_filtrado), hide_index=True, width="stretch",
        column_config={
            "Taxa de correção": st.column_config.ProgressColumn(
                format="%.1f%%", min_value=0, max_value=100,
                help="Corrigidos ÷ (corrigidos + pendentes). Compare junto com o total.",
            ),
        },
    )

with aba_registros:
    todos = tabela_de_registros(df_filtrado)[
        ["Sigla", "UF", "Detentora", "ID Detentora", "Status", "Enviado em",
         "Última cobrança", "Dias desde o envio", "E-mails"]
    ]
    st.dataframe(
        todos, hide_index=True, width="stretch", height=500, column_config=COLUNAS_DATA
    )
    st.download_button(
        "📥 Baixar resultado filtrado", gerar_csv(todos),
        "historico_filtrado.csv", "text/csv", key="baixar_filtrado",
    )

with aba_qualidade:
    st.caption("Verifica a planilha inteira, sem considerar os filtros.")
    if total_alertas == 0:
        st.success("Nenhum problema encontrado.")
    for descricao, registros in problemas.items():
        if registros.empty:
            continue
        with st.expander(f"{descricao}: {len(registros)}"):
            st.dataframe(
                tabela_de_registros(registros)[
                    ["Sigla", "Detentora", "ID Detentora", "Status", "Enviado em",
                     "Última cobrança", "E-mails"]
                ],
                hide_index=True, width="stretch", column_config=COLUNAS_DATA,
            )