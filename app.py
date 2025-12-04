import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="Simulador Financeiro SENAI")

st.title("📊 Simulador de Receitas e Despesas – SENAI")

# ------------------------------------------------------------
# Função para ler aba RECEITA e DESPESA
# ------------------------------------------------------------
def processar_planilha(file):
    try:
        xls = pd.ExcelFile(file)
    except:
        return None, None

    receitas, despesas = None, None

    for aba in xls.sheet_names:
        aba_lower = aba.lower()

        if "receita" in aba_lower:
            df = pd.read_excel(xls, aba)
            df.columns = [str(c).strip().lower() for c in df.columns]

            col_descr = [c for c in df.columns if "descricao" in c or "discr" in c]
            col_valor = [c for c in df.columns if "valor" in c]

            if col_descr and col_valor:
                df = df[[col_descr[0], col_valor[0]]].dropna()
                df.columns = ["descricao", "valor"]
                receitas = df

        if "despesa" in aba_lower:
            df = pd.read_excel(xls, aba)
            df.columns = [str(c).strip().lower() for c in df.columns]

            col_descr = [c for c in df.columns if "descricao" in c or "discr" in c]
            col_valor = [c for c in df.columns if "valor" in c]

            if col_descr and col_valor:
                df = df[[col_descr[0], col_valor[0]]].dropna()
                df.columns = ["descricao", "valor"]
                despesas = df

    return receitas, despesas


# ------------------------------------------------------------
# Upload dos arquivos
# ------------------------------------------------------------
st.sidebar.header("📁 Enviar planilhas")
uploads = st.sidebar.file_uploader(
    "Selecione os arquivos .xlsx dos meses (na ordem correta)",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploads:
    st.info("Envie os arquivos para iniciar.")
    st.stop()

dados_mensais = []
nomes_meses = []

for arquivo in uploads:
    nome = arquivo.name
    nomes_meses.append(nome)

    receitas, despesas = processar_planilha(arquivo)

    if receitas is None or despesas is None:
        st.error(f"Erro ao processar {arquivo.name}. Verifique abas RECEITA e DESPESA.")
        st.stop()

    total_receitas = receitas.groupby("descricao")["valor"].sum()
    total_despesas = despesas.groupby("descricao")["valor"].sum()

    dados_mensais.append({
        "arquivo": nome,
        "receitas": total_receitas,
        "despesas": total_despesas
    })

st.success("Planilhas carregadas com sucesso!")

# ------------------------------------------------------------
# Cálculo das variações
# ------------------------------------------------------------
variacoes = []

for i in range(len(dados_mensais)):
    if i == 0:
        var_receita = dados_mensais[i]["receitas"]
        var_despesa = dados_mensais[i]["despesas"]
    else:
        var_receita = dados_mensais[i]["receitas"] - dados_mensais[i-1]["receitas"]
        var_despesa = dados_mensais[i]["despesas"] - dados_mensais[i-1]["despesas"]

    variacoes.append({
        "mes": nomes_meses[i],
        "receita": var_receita.sum(),
        "despesa": var_despesa.sum()
    })

df_variacoes = pd.DataFrame(variacoes)
df_variacoes["saldo"] = df_variacoes["receita"] - df_variacoes["despesa"]
df_variacoes["acumulado"] = df_variacoes["saldo"].cumsum()

# ------------------------------------------------------------
# Gráfico Plotly
# ------------------------------------------------------------
x = df_variacoes["mes"]
y = df_variacoes["saldo"]
y_ac = df_variacoes["acumulado"]

fig = go.Figure()

fig.add_trace(go.Bar(x=x, y=y, name="Saldo Mensal"))
fig.add_trace(go.Scatter(x=x, y=y_ac, mode="lines+markers", name="Acumulado"))

fig.add_trace(go.Scatter(x=x, y=y_ac * 1.10, mode="lines",
                         line=dict(dash="dash", color="green"), name="+10%"))
fig.add_trace(go.Scatter(x=x, y=y_ac * 0.90, mode="lines",
                         line=dict(dash="dash", color="red"), name="-10%"))

fig.update_layout(
    title="📈 Projeção Financeira Mensal",
    xaxis_title="Mês",
    yaxis_title="Valor (R$)",
    height=600
)

st.plotly_chart(fig, use_container_width=True)

st.subheader("📋 Tabela de Variações Calculadas")
st.dataframe(df_variacoes)
