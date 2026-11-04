import streamlit as st
import pandas as pd
import plotly.express as px

# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="Histórico de Envios",
    page_icon="📧",
    layout="wide"
)

# =====================================================
# CARREGAMENTO
# =====================================================
@st.cache_data
def carregar_dados():
    df = pd.read_excel("historico_envios.xlsx")

    # Ajustar nomes das colunas
    df.columns = [col.strip() for col in df.columns]

    # Converter datas
    for coluna in ["DATA_ENVIO", "DATA_COBRANCA"]:
        if coluna in df.columns:
            df[coluna] = pd.to_datetime(
                df[coluna],
                errors="coerce"
            )

    # Padronizar CORRIGIDO
    if "CORRIGIDO" in df.columns:
        df["CORRIGIDO"] = (
            df["CORRIGIDO"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

    return df

df = carregar_dados()

# =====================================================
# TÍTULO
# =====================================================
st.title("📧 Histórico de Envios")
st.markdown("---")

# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.header("Filtros")

detentoras = sorted(
    df["DETENTORA"].dropna().unique()
) if "DETENTORA" in df.columns else []

detentora = st.sidebar.multiselect(
    "Detentora",
    options=detentoras,
    default=detentoras
)

status = st.sidebar.multiselect(
    "Status Correção",
    options=["SIM", "NAO", "NÃO"],
    default=["SIM", "NAO", "NÃO"]
)

# =====================================================
# FILTROS
# =====================================================
df_filtrado = df.copy()

if detentora:
    df_filtrado = df_filtrado[
        df_filtrado["DETENTORA"].isin(detentora)
    ]

if status:
    df_filtrado = df_filtrado[
        df_filtrado["CORRIGIDO"].isin(status)
    ]

# =====================================================
# KPIs
# =====================================================
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Total de Registros",
        len(df_filtrado)
    )

with col2:
    corrigidos = len(
        df_filtrado[
            df_filtrado["CORRIGIDO"].isin(["SIM"])
        ]
    )
    st.metric(
        "Corrigidos",
        corrigidos
    )

with col3:
    nao_corrigidos = len(
        df_filtrado[
            df_filtrado["CORRIGIDO"].isin(["NAO", "NÃO"])
        ]
    )
    st.metric(
        "Não Corrigidos",
        nao_corrigidos
    )

with col4:
    st.metric(
        "Detentoras",
        df_filtrado["DETENTORA"].nunique()
    )

st.markdown("---")

# =====================================================
# GRÁFICO DETENTORAS
# =====================================================
col1, col2 = st.columns(2)

with col1:
    graf_det = (
        df_filtrado["DETENTORA"]
        .value_counts()
        .reset_index()
    )

    graf_det.columns = [
        "DETENTORA",
        "QUANTIDADE"
    ]

    fig = px.bar(
        graf_det,
        x="DETENTORA",
        y="QUANTIDADE",
        color="QUANTIDADE",
        title="Envios por Detentora",
        color_continuous_scale="Purples"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

# =====================================================
# STATUS CORRIGIDO
# =====================================================
with col2:
    graf_status = (
        df_filtrado["CORRIGIDO"]
        .value_counts()
        .reset_index()
    )

    graf_status.columns = [
        "STATUS",
        "QUANTIDADE"
    ]

    fig2 = px.pie(
        graf_status,
        names="STATUS",
        values="QUANTIDADE",
        title="Status das Correções",
        hole=0.5,
        color="STATUS",
        color_discrete_map={
            "SIM": "#63D600",
            "NAO": "#EB3C7D",
            "NÃO": "#EB3C7D",
        }
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

# =====================================================
# EVOLUÇÃO DOS ENVIOS
# =====================================================
if "DATA_ENVIO" in df_filtrado.columns:

    evolucao = (
        df_filtrado
        .dropna(subset=["DATA_ENVIO"])
        .groupby(
            df_filtrado["DATA_ENVIO"].dt.date
        )
        .size()
        .reset_index(name="QUANTIDADE")
    )

    fig3 = px.line(
        evolucao,
        x="DATA_ENVIO",
        y="QUANTIDADE",
        markers=True,
        title="Evolução dos Envios"
    )

    st.plotly_chart(
        fig3,
        use_container_width=True
    )

# =====================================================
# TABELA
# =====================================================
st.subheader("Detalhamento")

st.dataframe(
    df_filtrado,
    use_container_width=True,
    height=600
)

# =====================================================
# DOWNLOAD
# =====================================================
csv = df_filtrado.to_csv(
    index=False,
    sep=";"
).encode("utf-8-sig")

st.download_button(
    "📥 Baixar Resultado Filtrado",
    csv,
    "historico_filtrado.csv",
    "text/csv"
)