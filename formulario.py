import streamlit as st
from datetime import date
from openpyxl import load_workbook
import pandas as pd
import os

# --- Inicializações de sessão ---
if "insumos" not in st.session_state:
    st.session_state.insumos = []
if "resetar_insumo" not in st.session_state:
    st.session_state.resetar_insumo = False
if "resetar_pedido" not in st.session_state:
    st.session_state.resetar_pedido = False
if "editar_index" not in st.session_state:
    st.session_state.editar_index = None

# --- Funções auxiliares ---
def resetar_campos_insumo():
    st.session_state.resetar_insumo = True

def resetar_formulario():
    st.session_state.resetar_pedido = True
    resetar_campos_insumo()
    st.session_state.insumos = []

def registrar_historico(numero, obra, data):
    historico_path = "historico_pedidos.csv"
    registro = {"numero": numero, "obra": obra, "data": data.strftime("%Y-%m-%d")}
    df_hist = pd.DataFrame([registro])
    if os.path.exists(historico_path):
        df_antigo = pd.read_csv(historico_path)
        df_hist = pd.concat([df_antigo, df_hist], ignore_index=True)
    df_hist.to_csv(historico_path, index=False)

def carregar_dados():
    df_empreend = pd.read_excel("Empreendimentos.xlsx")
    df_insumos = pd.read_excel("Insumos.xlsx")
    df_empreend.loc[-1] = ["", "", "", ""]
    df_empreend.index = df_empreend.index + 1
    df_empreend = df_empreend.sort_index()
    insumos_vazios = pd.DataFrame({"Código": [""], "Descrição": [""], "Unidade": [""]})
    df_insumos = pd.concat([insumos_vazios, df_insumos], ignore_index=True)
    return df_empreend, df_insumos

# --- Carrega dados ---
df_empreend, df_insumos = carregar_dados()

# --- Logo e título ---
st.image("logo.png", width=300)
st.markdown("""
    <div style='text-align: center;'>
        <h2 style='color: #003366;'>Sistema de Pedidos de Materiais</h2>
        <p style='font-size: 14px; color: #555;'>Preencha os campos com atenção. Evite abreviações desnecessárias.<br>
        Para pedidos novos, utilize sempre códigos oficiais quando disponíveis.</p>
    </div>
""", unsafe_allow_html=True)

# --- Dados do Pedido ---
with st.expander("📋 Dados do Pedido", expanded=True):
    if st.session_state.resetar_pedido:
        st.session_state.pedido_numero = ""
        st.session_state.data_pedido = date.today()
        st.session_state.solicitante = ""
        st.session_state.executivo = ""
        st.session_state.obra_selecionada = ""
        st.session_state.cnpj = ""
        st.session_state.endereco = ""
        st.session_state.cep = ""
        st.session_state.resetar_pedido = False

    col1, col2 = st.columns(2)
    with col1:
        pedido_numero = st.text_input("Pedido Nº", key="pedido_numero")
        solicitante = st.text_input("Solicitante", key="solicitante")
        obra_selecionada = st.selectbox("Obra", df_empreend["NOME"].unique(), index=0, key="obra_selecionada")
    with col2:
        data_pedido = st.date_input("Data", value=st.session_state.get("data_pedido", date.today()), key="data_pedido")
        executivo = st.text_input("Executivo", key="executivo")

    if obra_selecionada:
        dados_obra = df_empreend[df_empreend["NOME"] == obra_selecionada].iloc[0]
        st.session_state.cnpj = dados_obra["EMPRD_CNPJFAT"]
        st.session_state.endereco = dados_obra["ENDEREÇO"]
        st.session_state.cep = dados_obra["Cep"]

    st.text_input("CNPJ/CPF", value=st.session_state.get("cnpj", ""), disabled=True)
    st.text_input("Endereço", value=st.session_state.get("endereco", ""), disabled=True)
    st.text_input("CEP", value=st.session_state.get("cep", ""), disabled=True)

st.divider()

# --- Inicializa variáveis auxiliares para edição ---
descricao_editar = ""
descricao_livre_editar = ""
codigo_editar = ""
unidade_editar = ""
quantidade_editar = 0.0
complemento_editar = ""

if st.session_state.editar_index is not None:
    editar = st.session_state.insumos[st.session_state.editar_index]
    descricao_editar = editar["descricao"]
    codigo_editar = editar["codigo"]
    unidade_editar = editar["unidade"]
    quantidade_editar = editar["quantidade"]
    complemento_editar = editar["complemento"]
    st.session_state.insumos.pop(st.session_state.editar_index)
    st.session_state.editar_index = None

# --- Formulário de Insumo ---
st.markdown("### ➕ Adicionar ou Editar Insumo")

descricao = st.selectbox(
    "Descrição do insumo",
    df_insumos["Descrição"].unique(),
    key="descricao",
    index=df_insumos["Descrição"].tolist().index(descricao_editar)
    if descricao_editar in df_insumos["Descrição"].tolist() else 0
)

# Preenche código e unidade ao selecionar uma descrição, caso não esteja editando
if descricao and descricao_livre_editar == "":
    dados_insumo = df_insumos[df_insumos["Descrição"] == descricao]
    if not dados_insumo.empty:
        codigo_editar = dados_insumo.iloc[0]["Código"]
        unidade_editar = dados_insumo.iloc[0]["Unidade"]

st.write("Ou preencha manualmente se não estiver listado:")
descricao_livre = st.text_input("Nome do insumo (livre)", key="descricao_livre", value=descricao_livre_editar)
st.text_input("Código do insumo", value=codigo_editar, disabled=True)
unidade = st.text_input("Unidade", key="unidade", value=unidade_editar)
quantidade = st.number_input("Quantidade", min_value=0.0, format="%.2f", key="quantidade", value=quantidade_editar)
complemento = st.text_area("Complemento", key="complemento", value=complemento_editar)

if st.button("➕ Adicionar insumo"):
    descricao_final = descricao if descricao else descricao_livre
    if descricao_final and unidade.strip() and quantidade > 0:
        novo_insumo = {
            "descricao": descricao_final,
            "codigo": codigo_editar,
            "unidade": unidade,
            "quantidade": quantidade,
            "complemento": complemento,
        }
        st.session_state.insumos.append(novo_insumo)
        st.success("Insumo adicionado com sucesso!")
        resetar_campos_insumo()
        st.rerun()
    else:
        st.warning("Preencha todos os campos obrigatórios do insumo.")

# --- Renderiza tabela de insumos ---
if st.session_state.insumos:
    st.subheader("📦 Insumos adicionados")
    for i, insumo in enumerate(st.session_state.insumos):
        cols = st.columns([6, 1, 1])
        with cols[0]:
            st.markdown(f"**{i+1}.** {insumo['descricao']} — {insumo['quantidade']} {insumo['unidade']}")
        with cols[1]:
            if st.button("✏️ Editar", key=f"edit_{i}"):
                st.session_state.editar_index = i
                st.rerun()
        with cols[2]:
            if st.button("🗑️", key=f"delete_{i}"):
                st.session_state.insumos.pop(i)
                st.rerun()

# --- Finalização do Pedido ---
if st.button("📤 Enviar Pedido"):
    campos_obrigatorios = [
        st.session_state.pedido_numero,
        st.session_state.data_pedido,
        st.session_state.solicitante,
        st.session_state.executivo,
        st.session_state.obra_selecionada,
        st.session_state.cnpj,
        st.session_state.endereco,
        st.session_state.cep
    ]

    if not all(campos_obrigatorios):
        st.warning("⚠️ Preencha todos os campos obrigatórios antes de enviar o pedido.")
        st.stop()

    try:
        caminho_modelo = "Modelo_Pedido.xlsx"
        wb = load_workbook(caminho_modelo)
        ws = wb["Pedido"]

        ws["H2"] = st.session_state.pedido_numero
        ws["D3"] = st.session_state.data_pedido.strftime("%d/%m/%Y")
        ws["D4"] = st.session_state.solicitante
        ws["D5"] = st.session_state.executivo
        ws["D7"] = st.session_state.obra_selecionada
        ws["D8"] = st.session_state.cnpj
        ws["D9"] = st.session_state.endereco
        ws["D10"] = st.session_state.cep

        linha = 13
        for insumo in st.session_state.insumos:
            ws[f"C{linha}"] = insumo["codigo"]
            ws[f"D{linha}"] = insumo["descricao"]
            ws[f"F{linha}"] = insumo["unidade"]
            ws[f"G{linha}"] = insumo["quantidade"]
            ws[f"H{linha}"] = insumo["complemento"]
            linha += 1

        nome_saida = f"pedido_{st.session_state.pedido_numero or 'sem_numero'}_{st.session_state.obra_selecionada}.xlsx"
        wb.save(nome_saida)

        with open(nome_saida, "rb") as f:
            excel_bytes = f.read()

        st.success("✅ Pedido gerado com sucesso!")
        st.download_button("📥 Baixar Excel", data=excel_bytes, file_name=nome_saida)

        st.info("Para gerar um PDF, abra o arquivo no Excel e use a opção 'Salvar como PDF'.")

        registrar_historico(st.session_state.pedido_numero, st.session_state.obra_selecionada, st.session_state.data_pedido)
        resetar_formulario()

    except Exception as e:
        st.error(f"Erro ao gerar pedido: {e}")

# --- Histórico de pedidos ---
if st.checkbox("📖 Ver histórico de pedidos"):
    historico_path = "historico_pedidos.csv"
    if os.path.exists(historico_path):
        df = pd.read_csv(historico_path)
        df["data"] = pd.to_datetime(df["data"])
        obra_filtro = st.selectbox("Filtrar por obra", ["Todas"] + sorted(df["obra"].unique()))
        mes_filtro = st.selectbox("Filtrar por mês", ["Todos"] + sorted(df["data"].dt.strftime("%Y-%m").unique()))

        if obra_filtro != "Todas":
            df = df[df["obra"] == obra_filtro]
        if mes_filtro != "Todos":
            df = df[df["data"].dt.strftime("%Y-%m") == mes_filtro]

        st.table(df[["numero", "obra", "data"]])
    else:
        st.info("Nenhum pedido registrado ainda.")
