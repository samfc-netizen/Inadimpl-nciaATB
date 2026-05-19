# app_inadimplencia_autobras.py
# Dashboard Streamlit para análise de inadimplência da Autobras
# Como rodar:
# 1) pip install streamlit pandas openpyxl plotly xlsxwriter
# 2) streamlit run app_inadimplencia_autobras.py

import re
from datetime import datetime, date
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Inadimplência Autobras",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Estilo visual
# -----------------------------
st.markdown(
    """
    <style>
        .main .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
        div[data-testid="stMetricValue"] {font-size: 1.75rem;}
        .kpi-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 16px;
            padding: 16px;
            box-shadow: 0 1px 8px rgba(0,0,0,0.04);
        }
        .section-title {
            margin-top: 10px;
            padding-top: 8px;
            border-top: 1px solid #e5e7eb;
        }
        .insight-box {
            background: #f8fafc;
            border-left: 5px solid #334155;
            border-radius: 10px;
            padding: 14px 16px;
            margin-bottom: 10px;
            color: #111827;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Funções auxiliares
# -----------------------------
def limpar_moeda(valor):
    """Converte valores no padrão brasileiro, ex: 1.234,56, para float."""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if texto in ["", "------", "nan", "None"]:
        return 0.0
    texto = texto.replace("R$", "").replace(" ", "")
    texto = re.sub(r"[^0-9,.-]", "", texto)
    # Se houver ponto e vírgula, assume pt-BR: ponto milhar e vírgula decimal
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def formatar_reais(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def formatar_pct(valor):
    try:
        return f"{valor:.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"



MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}
MESES_PT_INV = {v: k for k, v in MESES_PT.items()}


def ordenar_meses_nomes(meses):
    return sorted(meses, key=lambda x: MESES_PT_INV.get(x, 99))


def detectar_linha_cabecalho(uploaded_file):
    """Detecta a linha onde está o cabeçalho real do relatório.

    O relatório da Autobras pode vir com uma linha de título acima do cabeçalho
    e pode chamar a data de vencimento de "Vencimento" ou "Data de vencimento".
    """
    uploaded_file.seek(0)
    preview = pd.read_excel(uploaded_file, header=None, nrows=15)
    for i in range(len(preview)):
        row = [str(x).strip().lower() for x in preview.iloc[i].tolist()]
        tem_cliente = "destinado à" in row or "destinado a" in row
        tem_vencimento = "data de vencimento" in row or "vencimento" in row
        tem_valor = "valor total" in row or "valor" in row
        if tem_cliente and tem_vencimento and tem_valor:
            return i
    return 0


def primeira_coluna_existente(df, opcoes):
    """Retorna o primeiro nome de coluna encontrado na planilha."""
    mapa = {str(c).strip().lower(): c for c in df.columns}
    for opcao in opcoes:
        chave = opcao.strip().lower()
        if chave in mapa:
            return mapa[chave]
    return None


def carregar_excel(uploaded_file):
    header_row = detectar_linha_cabecalho(uploaded_file)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_row)

    # Remove linhas vazias e linhas de totalização que não sejam títulos válidos
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Aceita variações de nomes vindas do relatório do ERP.
    col_cliente = primeira_coluna_existente(df, ["Destinado à", "Destinado a", "Cliente", "Nome"])
    col_vencimento = primeira_coluna_existente(df, ["Data de vencimento", "Vencimento", "Data Vencimento"])
    col_competencia = primeira_coluna_existente(df, ["Data de competência", "Competência", "Data Competência"])
    col_confirmacao = primeira_coluna_existente(df, ["Data de confirmação", "Confirmação", "Data Confirmação"])
    valor_col = primeira_coluna_existente(df, ["Valor total", "Valor", "Valor Aberto", "Saldo"])

    faltantes = []
    if not col_cliente:
        faltantes.append("Destinado à / Cliente")
    if not col_vencimento:
        faltantes.append("Vencimento / Data de vencimento")
    if not valor_col:
        faltantes.append("Valor total / Valor")
    if faltantes:
        st.error(f"A planilha não contém as colunas obrigatórias: {', '.join(faltantes)}")
        st.write("Colunas encontradas:", list(df.columns))
        st.stop()

    # Padronização básica
    df["Cliente"] = df[col_cliente].astype(str).str.strip()
    df["Loja"] = df.get("Loja", "Não informado").fillna("Não informado").astype(str).str.strip()
    df["Situação"] = df.get("Situação", "").fillna("").astype(str).str.strip()
    df["Forma de pagamento"] = df.get("Forma de pagamento", "Não informado").fillna("Não informado").astype(str).str.strip()
    df["Conta bancária"] = df.get("Conta bancária", "Não informado").fillna("Não informado").astype(str).str.strip()
    df["Descrição"] = df.get("Descrição", "").fillna("").astype(str).str.strip()

    df["Data Vencimento"] = pd.to_datetime(df[col_vencimento], dayfirst=True, errors="coerce")
    df["Data Competência"] = pd.to_datetime(df[col_competencia], dayfirst=True, errors="coerce") if col_competencia else pd.NaT
    df["Data Confirmação"] = pd.to_datetime(df[col_confirmacao], dayfirst=True, errors="coerce") if col_confirmacao else pd.NaT

    df["Valor Aberto"] = df[valor_col].apply(limpar_moeda)
    df = df[df["Valor Aberto"] > 0].copy()

    hoje = pd.Timestamp(date.today())
    df["Dias em atraso"] = (hoje - df["Data Vencimento"]).dt.days
    df["Está vencido"] = df["Dias em atraso"] > 0
    df["Está confirmado"] = df["Data Confirmação"].notna()

    # Inadimplência: títulos vencidos e não confirmados. Mantém também situação Atrasado como reforço.
    df["É inadimplente"] = ((df["Está vencido"]) & (~df["Está confirmado"])) | (df["Situação"].str.lower().str.contains("atras", na=False))

    df["Faixa de atraso"] = pd.cut(
        df["Dias em atraso"].fillna(0),
        bins=[-99999, 0, 7, 15, 30, 60, 90, 180, 99999],
        labels=["A vencer", "1-7 dias", "8-15 dias", "16-30 dias", "31-60 dias", "61-90 dias", "91-180 dias", "+180 dias"],
    )
    df["Mês Vencimento"] = df["Data Vencimento"].dt.to_period("M").astype(str)
    df["Ano"] = df["Data Vencimento"].dt.year
    df["Mês Número"] = df["Data Vencimento"].dt.month
    df["Mês Nome"] = df["Mês Número"].map(MESES_PT)

    return df


def tabela_download_excel(dfs: dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for nome, data in dfs.items():
            sheet = nome[:31]
            data.to_excel(writer, index=False, sheet_name=sheet)
            workbook = writer.book
            worksheet = writer.sheets[sheet]
            money_fmt = workbook.add_format({"num_format": "R$ #,##0.00"})
            pct_fmt = workbook.add_format({"num_format": "0.0%"})
            header_fmt = workbook.add_format({"bold": True, "bg_color": "#E5E7EB", "border": 1})
            for col_num, value in enumerate(data.columns.values):
                worksheet.write(0, col_num, value, header_fmt)
                largura = min(max(len(str(value)) + 2, 12), 38)
                worksheet.set_column(col_num, col_num, largura)
                if "Valor" in str(value):
                    worksheet.set_column(col_num, col_num, 16, money_fmt)
                if "%" in str(value) or "Participação" in str(value):
                    worksheet.set_column(col_num, col_num, 14, pct_fmt)
            worksheet.freeze_panes(1, 0)
    return output.getvalue()


# -----------------------------
# App
# -----------------------------
st.title("Dashboard Financeiro | Autobras")
st.caption("Análise de inadimplência e títulos a receber: visão por período, cliente, loja, aging e melhores práticas de cobrança.")

with st.sidebar:
    st.header("Upload e filtros")
    arquivo = st.file_uploader("Anexe o Excel de contas a receber", type=["xlsx", "xls"])
    st.info("O app aceita relatórios com cabeçalho na linha 1 ou abaixo do título do relatório.")

if not arquivo:
    st.warning("Anexe o arquivo Excel para gerar a análise.")
    st.stop()

base = carregar_excel(arquivo)
base_inad = base[base["É inadimplente"]].copy()

with st.sidebar:
    pagina = st.radio(
        "Multipage",
        ["Inadimplência", "Títulos a Receber"],
        index=0
    )

    st.subheader("Período de vencimento")

    anos_disponiveis = sorted(
        [int(a) for a in base["Ano"].dropna().unique().tolist()]
    )

    if anos_disponiveis:
        anos_sel = st.multiselect(
            "Ano",
            anos_disponiveis,
            default=anos_disponiveis,
            help="Selecione um ou mais anos de vencimento."
        )

        base_anos = base[base["Ano"].isin(anos_sel)].copy() if anos_sel else base.iloc[0:0].copy()

        meses_disponiveis = ordenar_meses_nomes(
            base_anos["Mês Nome"].dropna().unique().tolist()
        )

        meses_sel = st.multiselect(
            "Mês",
            meses_disponiveis,
            default=meses_disponiveis,
            help="Selecione um ou mais meses de vencimento."
        )
    else:
        anos_sel = []
        meses_sel = []

    lojas = sorted(base["Loja"].dropna().unique().tolist())
    lojas_sel = st.multiselect("Loja", lojas, default=lojas)

    clientes = sorted(base["Cliente"].dropna().unique().tolist())
    clientes_sel = st.multiselect("Cliente", clientes, default=[])

    faixa_sel = st.multiselect(
        "Faixa de atraso",
        ["A vencer", "1-7 dias", "8-15 dias", "16-30 dias", "31-60 dias", "61-90 dias", "91-180 dias", "+180 dias"],
        default=[]
    )

# Aplica filtros
filtro = base.copy()

if anos_sel:
    filtro = filtro[filtro["Ano"].isin(anos_sel)]
else:
    filtro = filtro.iloc[0:0].copy()

if meses_sel:
    filtro = filtro[filtro["Mês Nome"].isin(meses_sel)]
else:
    filtro = filtro.iloc[0:0].copy()

if lojas_sel:
    filtro = filtro[filtro["Loja"].isin(lojas_sel)]
if clientes_sel:
    filtro = filtro[filtro["Cliente"].isin(clientes_sel)]
if faixa_sel:
    filtro = filtro[filtro["Faixa de atraso"].astype(str).isin(faixa_sel)]

hoje_ref = pd.Timestamp(date.today()).normalize()

# Regra:
# - Inadimplência considera somente títulos com vencimento até a data vigente.
# - Títulos a Receber considera vencimentos de hoje para frente e que ainda não estão inadimplentes.
base_vencida_ate_hoje = filtro[filtro["Data Vencimento"].notna() & (filtro["Data Vencimento"] <= hoje_ref)].copy()
inad = base_vencida_ate_hoje[base_vencida_ate_hoje["É inadimplente"]].copy()

titulos_receber = filtro[
    filtro["Data Vencimento"].notna()
    & (filtro["Data Vencimento"] >= hoje_ref)
    & (~filtro["É inadimplente"])
].copy()

# ============================================================
# MULTIPAGE: TÍTULOS A RECEBER
# ============================================================

if pagina == "Títulos a Receber":
    st.subheader("Títulos a Receber")
    st.caption("Títulos com vencimento de hoje para frente e que ainda não estão inadimplentes.")

    valor_receber = titulos_receber["Valor Aberto"].sum()
    qtd_receber = len(titulos_receber)
    clientes_receber = titulos_receber["Cliente"].nunique()
    ticket_medio_receber = valor_receber / qtd_receber if qtd_receber else 0
    proximo_vencimento = titulos_receber["Data Vencimento"].min() if not titulos_receber.empty else pd.NaT

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Valor a receber", formatar_reais(valor_receber), f"{qtd_receber} títulos")
    c2.metric("Clientes com títulos a vencer", f"{clientes_receber}")
    c3.metric("Ticket médio", formatar_reais(ticket_medio_receber))
    c4.metric("Próximo vencimento", proximo_vencimento.strftime("%d/%m/%Y") if pd.notna(proximo_vencimento) else "-")

    if titulos_receber.empty:
        st.success("Não há títulos a receber no filtro selecionado.")
        st.stop()

    por_periodo_receber = (
        titulos_receber.groupby("Mês Vencimento", dropna=False)
        .agg(
            Valor_a_Receber=("Valor Aberto", "sum"),
            Titulos=("Valor Aberto", "count"),
            Clientes=("Cliente", "nunique"),
            Primeiro_Vencimento=("Data Vencimento", "min"),
            Ultimo_Vencimento=("Data Vencimento", "max"),
        )
        .reset_index()
        .sort_values("Mês Vencimento")
    )
    por_periodo_receber["Ticket Médio"] = por_periodo_receber["Valor_a_Receber"] / por_periodo_receber["Titulos"]

    por_cliente_receber = (
        titulos_receber.groupby("Cliente", dropna=False)
        .agg(
            Valor_a_Receber=("Valor Aberto", "sum"),
            Titulos=("Valor Aberto", "count"),
            Primeiro_Vencimento=("Data Vencimento", "min"),
            Ultimo_Vencimento=("Data Vencimento", "max"),
            Loja_Principal=("Loja", lambda x: x.mode().iat[0] if not x.mode().empty else ""),
        )
        .reset_index()
        .sort_values("Valor_a_Receber", ascending=False)
    )
    por_cliente_receber["Participação"] = por_cliente_receber["Valor_a_Receber"] / valor_receber if valor_receber else 0

    por_loja_receber = (
        titulos_receber.groupby("Loja", dropna=False)
        .agg(
            Valor_a_Receber=("Valor Aberto", "sum"),
            Titulos=("Valor Aberto", "count"),
            Clientes=("Cliente", "nunique"),
        )
        .reset_index()
        .sort_values("Valor_a_Receber", ascending=False)
    )
    por_loja_receber["Participação"] = por_loja_receber["Valor_a_Receber"] / valor_receber if valor_receber else 0

    aba_r1, aba_r2, aba_r3, aba_r4 = st.tabs(["Visão Geral", "Clientes", "Período", "Base analítica"])

    with aba_r1:
        c1, c2 = st.columns([1.2, 1])
        with c1:
            fig_loja_rec = px.bar(
                por_loja_receber,
                x="Loja",
                y="Valor_a_Receber",
                text=por_loja_receber["Valor_a_Receber"].apply(formatar_reais),
                title="Títulos a receber por loja",
            )
            fig_loja_rec.update_layout(yaxis_title="Valor a receber", xaxis_title="Loja")
            st.plotly_chart(fig_loja_rec, use_container_width=True)
        with c2:
            fig_tree_rec = px.treemap(
                por_cliente_receber.head(30),
                path=["Cliente"],
                values="Valor_a_Receber",
                title="Top clientes por valor a receber",
            )
            st.plotly_chart(fig_tree_rec, use_container_width=True)

        st.subheader("Resumo por loja")
        st.dataframe(
            por_loja_receber.assign(
                Valor_a_Receber=por_loja_receber["Valor_a_Receber"].apply(formatar_reais),
                Participação=por_loja_receber["Participação"].apply(lambda x: formatar_pct(x * 100)),
            ),
            use_container_width=True,
            hide_index=True,
        )

    with aba_r2:
        st.subheader("Ranking de clientes com títulos a receber")
        top_n_rec = st.slider("Quantidade de clientes no gráfico", 5, 30, 15, key="top_receber")
        top_clientes_rec = por_cliente_receber.head(top_n_rec).sort_values("Valor_a_Receber")
        fig_cli_rec = px.bar(
            top_clientes_rec,
            x="Valor_a_Receber",
            y="Cliente",
            orientation="h",
            text=top_clientes_rec["Valor_a_Receber"].apply(formatar_reais),
            title=f"Top {top_n_rec} clientes por valor a receber",
        )
        fig_cli_rec.update_layout(xaxis_title="Valor a receber", yaxis_title="Cliente")
        st.plotly_chart(fig_cli_rec, use_container_width=True)

        st.dataframe(
            por_cliente_receber.assign(
                Valor_a_Receber=por_cliente_receber["Valor_a_Receber"].apply(formatar_reais),
                Participação=por_cliente_receber["Participação"].apply(lambda x: formatar_pct(x * 100)),
                Primeiro_Vencimento=por_cliente_receber["Primeiro_Vencimento"].dt.strftime("%d/%m/%Y"),
                Ultimo_Vencimento=por_cliente_receber["Ultimo_Vencimento"].dt.strftime("%d/%m/%Y"),
            ),
            use_container_width=True,
            hide_index=True,
        )

    with aba_r3:
        st.subheader("Títulos a receber por período de vencimento")
        fig_periodo_rec = px.line(
            por_periodo_receber,
            x="Mês Vencimento",
            y="Valor_a_Receber",
            markers=True,
            text=por_periodo_receber["Valor_a_Receber"].apply(formatar_reais),
            title="Evolução mensal dos títulos a receber",
        )
        fig_periodo_rec.update_layout(yaxis_title="Valor a receber", xaxis_title="Mês de vencimento")
        st.plotly_chart(fig_periodo_rec, use_container_width=True)

        st.dataframe(
            por_periodo_receber.assign(
                Valor_a_Receber=por_periodo_receber["Valor_a_Receber"].apply(formatar_reais),
                **{"Ticket Médio": por_periodo_receber["Ticket Médio"].apply(formatar_reais)},
                Primeiro_Vencimento=por_periodo_receber["Primeiro_Vencimento"].dt.strftime("%d/%m/%Y"),
                Ultimo_Vencimento=por_periodo_receber["Ultimo_Vencimento"].dt.strftime("%d/%m/%Y"),
            ),
            use_container_width=True,
            hide_index=True,
        )

    with aba_r4:
        st.subheader("Base analítica dos títulos a receber")
        colunas_exibir_rec = [
            "Código", "Cliente", "CPF/CNPJ", "Descrição", "Loja", "Forma de pagamento",
            "Data Vencimento", "Valor Aberto", "Situação"
        ]
        colunas_exibir_rec = [c for c in colunas_exibir_rec if c in titulos_receber.columns]
        detalhe_rec = titulos_receber[colunas_exibir_rec].sort_values(["Data Vencimento", "Valor Aberto"], ascending=[True, False]).copy()
        detalhe_rec_view = detalhe_rec.copy()
        detalhe_rec_view["Data Vencimento"] = detalhe_rec_view["Data Vencimento"].dt.strftime("%d/%m/%Y")
        detalhe_rec_view["Valor Aberto"] = detalhe_rec_view["Valor Aberto"].apply(formatar_reais)
        st.dataframe(detalhe_rec_view, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Exportar títulos a receber")
    export_receber = tabela_download_excel({
        "Resumo por Cliente": por_cliente_receber,
        "Resumo por Loja": por_loja_receber,
        "Resumo por Periodo": por_periodo_receber,
        "Titulos a Receber": titulos_receber[[c for c in ["Código", "Cliente", "CPF/CNPJ", "Descrição", "Loja", "Forma de pagamento", "Data Vencimento", "Valor Aberto", "Situação"] if c in titulos_receber.columns]],
    })
    st.download_button(
        "Baixar títulos a receber em Excel",
        data=export_receber,
        file_name="titulos_a_receber_autobras.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.stop()

# ============================================================
# MULTIPAGE: INADIMPLÊNCIA
# ============================================================

valor_total_carteira = base_vencida_ate_hoje["Valor Aberto"].sum()
valor_inad = inad["Valor Aberto"].sum()
qtd_titulos = len(base_vencida_ate_hoje)
qtd_inad = len(inad)
taxa_inad = (valor_inad / valor_total_carteira * 100) if valor_total_carteira else 0
clientes_inad = inad["Cliente"].nunique()
maior_atraso = int(inad["Dias em atraso"].max()) if not inad.empty else 0
media_atraso_pond = (inad["Dias em atraso"].clip(lower=0).mul(inad["Valor Aberto"]).sum() / valor_inad) if valor_inad else 0

st.subheader("Indicadores executivos")
st.caption("Inadimplência considera somente títulos com vencimento até a data vigente.")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Carteira vencida filtrada", formatar_reais(valor_total_carteira), f"{qtd_titulos} títulos")
col2.metric("Inadimplência", formatar_reais(valor_inad), f"{qtd_inad} títulos")
col3.metric("Taxa de inadimplência", formatar_pct(taxa_inad))
col4.metric("Clientes inadimplentes", f"{clientes_inad}")
col5.metric("Atraso médio ponderado", f"{media_atraso_pond:.0f} dias")

if inad.empty:
    st.success("Não há títulos inadimplentes no filtro selecionado.")
    st.stop()

st.markdown('<div class="section-title"></div>', unsafe_allow_html=True)

# Tabelas resumidas
por_periodo = (
    inad.groupby("Mês Vencimento", dropna=False)
    .agg(Valor_Inadimplente=("Valor Aberto", "sum"), Titulos=("Valor Aberto", "count"), Clientes=("Cliente", "nunique"), Maior_Atraso=("Dias em atraso", "max"))
    .reset_index()
    .sort_values("Mês Vencimento")
)
por_periodo["Ticket Médio"] = por_periodo["Valor_Inadimplente"] / por_periodo["Titulos"]

por_cliente = (
    inad.groupby("Cliente", dropna=False)
    .agg(Valor_Inadimplente=("Valor Aberto", "sum"), Titulos=("Valor Aberto", "count"), Primeiro_Vencimento=("Data Vencimento", "min"), Ultimo_Vencimento=("Data Vencimento", "max"), Maior_Atraso=("Dias em atraso", "max"), Loja_Principal=("Loja", lambda x: x.mode().iat[0] if not x.mode().empty else ""))
    .reset_index()
    .sort_values("Valor_Inadimplente", ascending=False)
)
por_cliente["Participação"] = por_cliente["Valor_Inadimplente"] / valor_inad
por_cliente["Risco"] = pd.cut(
    por_cliente["Maior_Atraso"],
    bins=[-1, 15, 30, 60, 90, 99999],
    labels=["Baixo", "Atenção", "Médio", "Alto", "Crítico"]
)

por_loja = (
    inad.groupby("Loja", dropna=False)
    .agg(Valor_Inadimplente=("Valor Aberto", "sum"), Titulos=("Valor Aberto", "count"), Clientes=("Cliente", "nunique"), Maior_Atraso=("Dias em atraso", "max"))
    .reset_index()
    .sort_values("Valor_Inadimplente", ascending=False)
)
por_loja["Participação"] = por_loja["Valor_Inadimplente"] / valor_inad

por_faixa = (
    inad.groupby("Faixa de atraso", observed=False)
    .agg(Valor_Inadimplente=("Valor Aberto", "sum"), Titulos=("Valor Aberto", "count"), Clientes=("Cliente", "nunique"))
    .reset_index()
)
por_faixa["Participação"] = por_faixa["Valor_Inadimplente"] / valor_inad

# Gráficos
aba1, aba2, aba3, aba4, aba5 = st.tabs(["Visão Geral", "Clientes", "Período", "Títulos em aberto", "Melhores práticas"])

with aba1:
    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig_loja = px.bar(
            por_loja,
            x="Loja",
            y="Valor_Inadimplente",
            text=por_loja["Valor_Inadimplente"].apply(formatar_reais),
            title="Inadimplência por loja",
        )
        fig_loja.update_layout(yaxis_title="Valor inadimplente", xaxis_title="Loja")
        st.plotly_chart(fig_loja, use_container_width=True)
    with c2:
        fig_faixa = px.pie(
            por_faixa[por_faixa["Valor_Inadimplente"] > 0],
            values="Valor_Inadimplente",
            names="Faixa de atraso",
            title="Aging da inadimplência",
            hole=0.45,
        )
        st.plotly_chart(fig_faixa, use_container_width=True)

    st.subheader("Resumo por loja")
    st.dataframe(
        por_loja.assign(
            Valor_Inadimplente=por_loja["Valor_Inadimplente"].apply(formatar_reais),
            Participação=por_loja["Participação"].apply(lambda x: formatar_pct(x * 100)),
        ),
        use_container_width=True,
        hide_index=True,
    )

with aba2:
    st.subheader("Ranking de clientes inadimplentes")
    top_n = st.slider("Quantidade de clientes no gráfico", 5, 30, 15)
    top_clientes = por_cliente.head(top_n).sort_values("Valor_Inadimplente")
    fig_cli = px.bar(
        top_clientes,
        x="Valor_Inadimplente",
        y="Cliente",
        orientation="h",
        text=top_clientes["Valor_Inadimplente"].apply(formatar_reais),
        title=f"Top {top_n} clientes por valor inadimplente",
    )
    fig_cli.update_layout(xaxis_title="Valor inadimplente", yaxis_title="Cliente")
    st.plotly_chart(fig_cli, use_container_width=True)

    st.dataframe(
        por_cliente.assign(
            Valor_Inadimplente=por_cliente["Valor_Inadimplente"].apply(formatar_reais),
            Participação=por_cliente["Participação"].apply(lambda x: formatar_pct(x * 100)),
            Primeiro_Vencimento=por_cliente["Primeiro_Vencimento"].dt.strftime("%d/%m/%Y"),
            Ultimo_Vencimento=por_cliente["Ultimo_Vencimento"].dt.strftime("%d/%m/%Y"),
        ),
        use_container_width=True,
        hide_index=True,
    )

with aba3:
    st.subheader("Inadimplência por período de vencimento")
    fig_periodo = px.line(
        por_periodo,
        x="Mês Vencimento",
        y="Valor_Inadimplente",
        markers=True,
        text=por_periodo["Valor_Inadimplente"].apply(formatar_reais),
        title="Evolução mensal da inadimplência",
    )
    fig_periodo.update_layout(yaxis_title="Valor inadimplente", xaxis_title="Mês de vencimento")
    st.plotly_chart(fig_periodo, use_container_width=True)

    st.dataframe(
        por_periodo.assign(
            Valor_Inadimplente=por_periodo["Valor_Inadimplente"].apply(formatar_reais),
            **{"Ticket Médio": por_periodo["Ticket Médio"].apply(formatar_reais)}
        ),
        use_container_width=True,
        hide_index=True,
    )

with aba4:
    st.subheader("Base analítica dos títulos inadimplentes")
    colunas_exibir = [
        "Código", "Cliente", "CPF/CNPJ", "Descrição", "Loja", "Forma de pagamento",
        "Data Vencimento", "Dias em atraso", "Faixa de atraso", "Valor Aberto", "Situação"
    ]
    colunas_exibir = [c for c in colunas_exibir if c in inad.columns]
    detalhe = inad[colunas_exibir].sort_values(["Dias em atraso", "Valor Aberto"], ascending=[False, False]).copy()
    detalhe_view = detalhe.copy()
    detalhe_view["Data Vencimento"] = detalhe_view["Data Vencimento"].dt.strftime("%d/%m/%Y")
    detalhe_view["Valor Aberto"] = detalhe_view["Valor Aberto"].apply(formatar_reais)
    st.dataframe(detalhe_view, use_container_width=True, hide_index=True)

with aba5:
    st.subheader("Leitura de analista sênior e boas práticas")

    maior_cliente = por_cliente.iloc[0]
    conc_top5 = por_cliente.head(5)["Valor_Inadimplente"].sum() / valor_inad * 100 if valor_inad else 0
    aging_60 = inad[inad["Dias em atraso"] > 60]["Valor Aberto"].sum()
    pct_aging_60 = aging_60 / valor_inad * 100 if valor_inad else 0

    st.markdown(
        f"""
        <div class="insight-box">
        <b>Diagnóstico financeiro:</b> a inadimplência filtrada soma <b>{formatar_reais(valor_inad)}</b>, concentrada em <b>{clientes_inad}</b> clientes e <b>{qtd_inad}</b> títulos. 
        O maior cliente em aberto é <b>{maior_cliente['Cliente']}</b>, com <b>{formatar_reais(maior_cliente['Valor_Inadimplente'])}</b>, representando <b>{formatar_pct(maior_cliente['Participação'] * 100)}</b> da inadimplência.
        </div>

        <div class="insight-box">
        <b>Concentração de risco:</b> os 5 maiores clientes representam <b>{formatar_pct(conc_top5)}</b> do valor inadimplente. 
        Quando poucos clientes concentram grande parte do saldo vencido, a rotina de cobrança deve ser tratada como gestão de carteira, não apenas como baixa operacional de boletos.
        </div>

        <div class="insight-box">
        <b>Aging:</b> títulos com mais de 60 dias somam <b>{formatar_reais(aging_60)}</b>, equivalente a <b>{formatar_pct(pct_aging_60)}</b> do valor inadimplente. 
        Acima de 60 dias, a probabilidade de recuperação costuma cair e a cobrança deve migrar para régua formal, negociação documentada e bloqueio/limite de crédito.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("""
    #### Melhores práticas recomendadas

    1. **Separar cobrança operacional de cobrança crítica**  
       Títulos até 15 dias podem seguir cobrança leve: WhatsApp, ligação e reenvio de boleto. Acima de 30 dias, tratar como caso de risco, com responsável definido e prazo de resolução.

    2. **Criar régua de cobrança por aging**  
       - 1 a 7 dias: lembrete cordial e confirmação de recebimento da cobrança.  
       - 8 a 15 dias: contato ativo com previsão objetiva de pagamento.  
       - 16 a 30 dias: negociação com data formal e acompanhamento diário.  
       - 31 a 60 dias: bloquear novas vendas a prazo, revisar limite e exigir entrada.  
       - Acima de 60 dias: renegociação formal, confissão de dívida quando aplicável e avaliação de cobrança externa/jurídica.

    3. **Atacar concentração primeiro**  
       Priorize os clientes com maior valor em aberto. A carteira não deve ser cobrada apenas por ordem cronológica; deve ser cobrada por impacto financeiro e risco.

    4. **Implantar política de crédito**  
       Definir limite por cliente, prazo máximo, regra para clientes novos, bloqueio automático por atraso e exceções aprovadas pela liderança.

    5. **Acompanhar indicadores semanalmente**  
       Indicadores mínimos: valor vencido, taxa de inadimplência, clientes inadimplentes, valor acima de 30/60/90 dias, recuperação semanal e reincidência por cliente.

    6. **Usar o histórico para decisão comercial**  
       Cliente recorrente em atraso não deve receber o mesmo prazo/limite de um cliente adimplente. A venda precisa considerar margem, prazo e risco de recebimento.
    """)

# Exportação
st.markdown("---")
st.subheader("Exportar análise")
export = tabela_download_excel({
    "Resumo por Cliente": por_cliente,
    "Resumo por Loja": por_loja,
    "Resumo por Periodo": por_periodo,
    "Aging": por_faixa,
    "Titulos Inadimplentes": inad[[c for c in ["Código", "Cliente", "CPF/CNPJ", "Descrição", "Loja", "Forma de pagamento", "Data Vencimento", "Dias em atraso", "Faixa de atraso", "Valor Aberto", "Situação"] if c in inad.columns]],
})
st.download_button(
    "Baixar análise em Excel",
    data=export,
    file_name="analise_inadimplencia_autobras.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
