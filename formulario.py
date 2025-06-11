import streamlit as st
from datetime import date
from openpyxl import load_workbook
import pandas as pd
import os
# import win32com.client as win32

# --- Função para resetar campos dos insumos ---
def resetar_campos_insumo():
    st.session_state["resetar_insumo"] = True

# --- Função para resetar todos os campos do pedido ---
def resetar_formulario():
    st.session_state.resetar_pedido = True
    resetar_campos_insumo()
    st.session_state.insumos = []

# Inicializa a lista de insumos na sessão
if "insumos" not in st.session_state:
    st.session_state.insumos = []

# Inicializa o flag de reset para os campos de insumo
if "resetar_insumo" not in st.session_state:
    st.session_state.resetar_insumo = False

# Inicializa o flag de reset para todos os campos
if "resetar_pedido" not in st.session_state:
    st.session_state.resetar_pedido = False


# --- Função para gerar PDF a partir do Excel ---
# def salvar_pdf_do_excel(caminho_excel, nome_pdf_saida):
 #    excel = win32.gencache.EnsureDispatch("Excel.Application")
 #    wb = excel.Workbooks.Open(os.path.abspath(caminho_excel))
 #    ws = wb.Worksheets("Pedido")
 #    ws.PageSetup.Zoom = False
 #    ws.PageSetup.FitToPagesWide = 1
 #    ws.PageSetup.FitToPagesTall = False
 #    wb.ExportAsFixedFormat(0, os.path.abspath(nome_pdf_saida))
 #    wb.Close(SaveChanges=False)
 #    excel.Quit()

# --- Carregar dados das planilhas ---
df_empreend = pd.read_excel("Empreendimentos.xlsx")
df_insumos = pd.read_excel("Insumos.xlsx")

# Adiciona opção vazia no topo
df_empreend.loc[-1] = ["", "", "", ""]
df_empreend.index = df_empreend.index + 1
df_empreend = df_empreend.sort_index()

insumos_vazios = pd.DataFrame({"Código": [""], "Descrição": [""], "Unidade": [""]})
df_insumos = pd.concat([insumos_vazios, df_insumos], ignore_index=True)

# --- Dados do Pedido ---
st.subheader("Dados do Pedido")
# Aplica reset visual dos campos de pedido se necessário
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

cnpj = ""
endereco = ""
cep = ""

if obra_selecionada:
    dados_obra = df_empreend[df_empreend["NOME"] == obra_selecionada].iloc[0]
    cnpj = dados_obra["EMPRD_CNPJFAT"]
    endereco = dados_obra["ENDEREÇO"]
    cep = dados_obra["Cep"]

cnpj = st.text_input("CNPJ/CPF", value=cnpj, key="cnpj")
endereco = st.text_input("Endereço", value=endereco, key="endereco")
cep = st.text_input("CEP", value=cep, key="cep")

# --- Adicionar Insumo ---

# Aplica reset visual dos campos se necessário
if st.session_state.resetar_insumo:
    st.session_state.descricao = ""
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

codigo = st.text_input("Código do insumo", value=codigo, key="codigo")
unidade = st.text_input("Unidade", value=unidade, key="unidade")
quantidade = st.number_input("Quantidade", min_value=0.0, format="%.2f", key="quantidade")
complemento = st.text_area("Complemento", key="complemento")

# --- Botão para adicionar insumo ---
if st.button("➕ Adicionar insumo"):
    if descricao and codigo and unidade and quantidade > 0:
        novo_insumo = {
            "descricao": descricao,
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

# --- Botão final para gerar Excel + PDF ---
if st.button("📤 Enviar Pedido"):
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

        nome_saida = f"pedido_{st.session_state.pedido_numero or 'sem_numero'}.xlsx"
        wb.save(nome_saida)
        st.success("Pedido gerado com sucesso!")

        # PDF desabilitado no Streamlit Cloud por incompatibilidade
        # nome_pdf = nome_saida.replace(".xlsx", ".pdf")
        # salvar_pdf_do_excel(nome_saida, nome_pdf)
        # st.success(f"PDF exportado: {nome_pdf}")

        # Botão para download do Excel gerado
        with open(nome_saida, "rb") as f:
            st.download_button(
                label="📥 Baixar pedido em Excel",
                data=f,
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        resetar_formulario()
        st.rerun()

    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")
