import streamlit as st
from datetime import date
from openpyxl import load_workbook
import pandas as pd
import os

# --- Função para resetar campos dos insumos ---
def resetar_campos_insumo():
    st.session_state["resetar_insumo"] = True

# --- Função para resetar todos os campos do pedido ---
def resetar_formulario():
    st.session_state.resetar_pedido = True
    resetar_campos_insumo()
    st.session_state.insumos = []

# Inicializações
if "insumos" not in st.session_state:
    st.session_state.insumos = []
if "resetar_insumo" not in st.session_state:
    st.session_state.resetar_insumo = False
if "resetar_pedido" not in st.session_state:
    st.session_state.resetar_pedido = False

# --- Carregar dados das planilhas ---
df_empreend = pd.read_excel("Empreendimentos.xlsx")
df_insumos = pd.read_excel("Insumos.xlsx")
df_empreend.loc[-1] = ["", "", "", ""]
df_empreend.index = df_empreend.index + 1
df_empreend = df_empreend.sort_index()

insumos_vazios = pd.DataFrame({"Código": [""], "Descrição": [""], "Unidade": [""]})
df_insumos = pd.concat([insumos_vazios, df_insumos], ignore_index=True)

# --- Dados do Pedido ---
st.subheader("Dados do Pedido")
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

pedido_numero = st.text_input("Pedido de material Nº", key="pedido_numero")
data_pedido = st.date_input("Data", value=st.session_state.get("data_pedido", date.today()), key="data_pedido")
solicitante = st.text_input("Solicitante", key="solicitante")
executivo = st.text_input("Executivo", key="executivo")
obra_selecionada = st.selectbox("Obra", df_empreend["NOME"].unique(), index=0, key="obra_selecionada")

if obra_selecionada:
    dados_obra = df_empreend[df_empreend["NOME"] == obra_selecionada].iloc[0]
    st.session_state["cnpj"] = dados_obra["EMPRD_CNPJFAT"]
    st.session_state["endereco"] = dados_obra["ENDEREÇO"]
    st.session_state["cep"] = dados_obra["Cep"]

st.text_input("CNPJ/CPF", value=st.session_state.get("cnpj", ""), disabled=True)
st.text_input("Endereço", value=st.session_state.get("endereco", ""), disabled=True)
st.text_input("CEP", value=st.session_state.get("cep", ""), disabled=True)

# --- Adicionar Insumo ---
if st.session_state.resetar_insumo:
    st.session_state.descricao = ""
    st.session_state.descricao_livre = ""
    st.session_state.codigo = ""
    st.session_state.unidade = ""
    st.session_state.quantidade = 0.0
    st.session_state.complemento = ""
    st.session_state.resetar_insumo = False

st.subheader("Adicionar Insumo")
descricao = st.selectbox("Descrição do insumo", df_insumos["Descrição"].unique(), index=0, key="descricao")
codigo = ""
unidade = ""
if descricao:
    dados_insumo = df_insumos[df_insumos["Descrição"] == descricao].iloc[0]
    codigo = dados_insumo["Código"]
    unidade = dados_insumo["Unidade"]

st.write("Caso o insumo não esteja listado acima, digite abaixo:")
descricao_livre = st.text_input("Nome do insumo (livre)", key="descricao_livre")

st.text_input("Código do insumo", value=codigo, disabled=True)
unidade = st.text_input("Unidade", value=unidade, key="unidade")
quantidade = st.number_input("Quantidade", min_value=0.0, format="%.2f", key="quantidade")
complemento = st.text_area("Complemento", key="complemento")

if st.button("➕ Adicionar insumo"):
    descricao_final = descricao if descricao else descricao_livre
    if descricao_final and unidade.strip() and quantidade > 0:
        novo_insumo = {
            "descricao": descricao_final,
            "codigo": codigo,
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

# --- Mostrar e excluir insumos adicionados ---
if st.session_state.insumos:
    st.subheader("Insumos adicionados")
    for i, insumo in enumerate(st.session_state.insumos):
        col1, col2 = st.columns([6, 1])
        with col1:
            st.markdown(f"**{i+1}.** {insumo['descricao']} ({insumo['quantidade']} {insumo['unidade']})")
        with col2:
            if st.button("🗑️", key=f"delete_{i}"):
                st.session_state.insumos.pop(i)
                st.rerun()

# --- Botão final para gerar Excel ---
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
        st.markdown("Clique abaixo para baixar o arquivo gerado:")
        st.download_button(
            label="📥 Baixar pedido em Excel",
            data=excel_bytes,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        resetar_formulario()

    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")
