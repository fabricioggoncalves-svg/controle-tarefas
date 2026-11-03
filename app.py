import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime

# ----------------------------------------------------
# CONFIGURAÇÃO GERAL
# ----------------------------------------------------
st.set_page_config(page_title="Controle de Tarefas - Advocacia", layout="wide", page_icon="⚖️")

# Nomes possíveis da planilha no Google Drive
NOMES_PLANILHA_POSSIVEIS = ["Controle de Tarefas", "Tarefas Advog Homecare"]

# ----------------------------------------------------
# CREDENCIAIS GOOGLE SERVICE ACCOUNT (NOVA CHAVE ATIVA)
# ----------------------------------------------------
SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
    "project_id": "agenda-simples-513616",
    "private_key_id": "40fb2cef068b00f130f1d3f73cbe1632e74410c9",
    "private_key": (
        "-----BEGIN PRIVATE KEY-----\n"
        "MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCns6DQdYvAcaGs\n"
        "8rrezJZlf8Ga39NEVS5mDvoP1nUGkKKICMRsFUBnNTwrS7JfUUmMExHEGBhl05l0\n"
        "CzmWZ31mXcMGZQezHhwmrVyLdTlwflEXH/p0CvL8v97dP+81mDDzKsUV1ht+ywx5\n"
        "fRqtW2TwQHxy/Fma63Pqv5Ak/N1Zyf/RIUM2uPCs/BIWMbFDKCzVhwvtUG1Mqc4P\n"
        "SSOFBnExdNzbZVJUnxXSHeziL7581LPAa9Qnbo5GF8U06l4td9be84DLKSSNdsVi\n"
        "qdilowDxhRlGPRgoiWQzxx2oxiaaUbdFImymYiQFZDp2ON0Sbk13RfFmx6+2qllN\n"
        "ClMapHfpAgMBAAECggEACDm8zQy+7pEEQ5XiSZRSWiciiWSqGaNt866PYGB9+/Tz\n"
        "K+X2MHD4aDeJvHvpg1ha2aaDzBD8fV4cHD+CsxhfY7p8mt8wuf3jlpgRPnvg9d4U\n"
        "PySf5bo1weTtWv/MTV4qJHs4EIKgdZ8i/QCAYvllIY2gLQmooXskA+MigFXwy3sT\n"
        "GC0VVp+M3AGN6tmnPTyzrIIMsWomCkhmHaXiiz4wR+0M8yINXXN140YXCKXntMEc\n"
        "Z5oAnuqUSHDfn/9Se4vLc4zflaDumjQs2XeUQp0t5qNNkxxLyl43CZ50VWVFCMbo\ndBipAaNSh6jP6oAX5xfLQkgfD1fIlD7o/MRghJRFwwKBgQDbAWKnFmTrIDTUQBbV\n"
        "fg4I/u64WZ0T+LIkxNEX7rvpT1I1zm4oEamnKQCgJoH0YMbOK+f9APODYYmnr5J+\n"
        "R6+r+sf0uKnCVMEkNMq8zNxIigSSX2ppCJiZ4R3p4b0puaXVCu9r5iPMeUWT/H1l\n"
        "CG4qrCULQkbScdFBp3JSwBufbwKBgQDEB61Jhifc+COCc/vfKc98QsRZ1NumYmMQ\n"
        "D7wpSEQeMW32E+8Kw5b+OzbW41gRz9qtGVxIKir9RdyYLOiAOL5pdc3hfC1qpwcE\n"
        "PSr8nOR24Df/pDcTvFn2GQrsWmYr0ji1CCMRIgOicwvFEmZXpqKSXKOFELm7nnuT\n"
        "1HjhLayyJwKBgD7ynf1ypV0jq6uaJDYG/le9lrXsJAb6g5Gn9IWRe/dqcjG+/NCN\n"
        "ixrd4S0nBFFOg7X8zrhczqFjpqJJB60+Q5MGCPUhIsIVhFI4zlEH/QTojrnXrQk1\n"
        "jE8mcmeoF6vnjW+3retwgbei38nSV+RzUN75Dci3qw2P41qgm26HgykHAoGBAJbe\n"
        "rbCjjrSTSYkSSOxVWnnxTPsyRqbEnCd1ZMSE+OAf/NhiORyYsfISkr9KE7MJvOxy\n"
        "U0K7KYDqLr/z+DEO+wVhhEWhc7bcvBq4FQh8m50eLhsTQMxwrvAoD+8eF/pr/fbe\n"
        "P1blTdTOJGXVJm0ZM4/fun4kBD5w1/1LjokHxR0LAoGBAMVWPA7LjKT7o0fC2CE/\n"
        "TYvcFbUQGnVt80NzSJdEF23YoqmWmr8HM3vpRVuhBWprK1uDx1dCVvsEE7B4dIlw\n"
        "+HCvIww2GdmaaDZ3VK6oEiqX3XpvE4HLFRgDPk9EeNN7ySkgJ4D0x/nN2Ft8hgb1\n"
        "AzOfTA5bmkpIl6tFpFv4iwni\n"
        "-----END PRIVATE KEY-----\n"
    ),
    "client_email": "task-manager@agenda-simples-513616.iam.gserviceaccount.com",
    "client_id": "107227922236372451032",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/task-manager%40agenda-simples-513616.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

# ----------------------------------------------------
# CONEXÃO GOOGLE SHEETS
# ----------------------------------------------------
@st.cache_resource(ttl=60)
def conectar_google_sheets():
    """Autentica e abre a planilha de tarefas existente."""
    try:
        creds_dict = dict(SERVICE_ACCOUNT_INFO)
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        escopos = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credenciais = Credentials.from_service_account_info(creds_dict, scopes=escopos)
        cliente = gspread.authorize(credenciais)

        planilha = None
        for nome in NOMES_PLANILHA_POSSIVEIS:
            try:
                planilha = cliente.open(nome)
                if planilha:
                    break
            except Exception:
                continue

        if not planilha:
            raise Exception(f"Não foi possível abrir a planilha com os nomes: {NOMES_PLANILHA_POSSIVEIS}")

        return planilha
    except Exception as e:
        st.error(f"Erro ao conectar à planilha: {e}")
        return None

def obter_aba_tarefas(planilha):
    """Localiza a aba com os dados (Página1 ou a primeira aba da planilha)."""
    try:
        return planilha.worksheet("Página1")
    except Exception:
        return planilha.get_worksheet(0)

def carregar_tarefas(aba):
    """Lê todas as linhas e colunas reais da planilha."""
    dados = aba.get_all_values()
    if not dados or len(dados) < 2:
        return pd.DataFrame(columns=["Responsável", "Número", "Vencimento", "Categoria", "Descrição", "Status"])
    
    cabecalhos = [str(c).strip() for c in dados[0]]
    linhas = dados[1:]
    
    df = pd.DataFrame(linhas, columns=cabecalhos)
    
    # Remove linhas vazias
    if "Descrição" in df.columns:
        df = df[df["Descrição"].str.strip() != ""].copy()
    elif "Número" in df.columns:
        df = df[df["Número"].str.strip() != ""].copy()

    # Adiciona colunas esperadas caso não existam
    colunas_obrigatorias = ["Responsável", "Número", "Vencimento", "Categoria", "Descrição", "Status"]
    for col in colunas_obrigatorias:
        if col not in df.columns:
            df[col] = ""

    return df[colunas_obrigatorias]

# ----------------------------------------------------
# INTERFACE PRINCIPAL
# ----------------------------------------------------
st.title("⚖️ Controle de Tarefas")
st.caption("Sincronizado em tempo real com a planilha do Google Sheets")

planilha = conectar_google_sheets()

if not planilha:
    st.warning("Verifique o compartilhamento da planilha com o e-mail da conta de serviço como Editor.")
    st.stop()

aba = obter_aba_tarefas(planilha)
df_tarefas = carregar_tarefas(aba)

# ----------------------------------------------------
# MÉTRICAS RÁPIDAS
# ----------------------------------------------------
total_tarefas = len(df_tarefas)
pendentes = len(df_tarefas[df_tarefas["Status"].str.lower().str.contains("pendente", na=False)])
atrasadas = len(df_tarefas[df_tarefas["Status"].str.lower().str.contains("atrasad", na=False)])
concluidas = len(df_tarefas[df_tarefas["Status"].str.lower().str.contains("conclu", na=False)])

m1, m2, m3, m4 = st.columns(4)
m1.metric("Total de Tarefas", total_tarefas)
m2.metric("Pendentes", pendentes)
m3.metric("Atrasadas", atrasadas)
m4.metric("Concluídas", concluidas)

st.divider()

# ----------------------------------------------------
# FORMULÁRIO: INCLUIR NOVA TAREFA
# ----------------------------------------------------
with st.expander("➕ Nova Tarefa", expanded=False):
    # Calcula próximo número sequencial
    proximo_num = 1
    if not df_tarefas.empty and "Número" in df_tarefas.columns:
        nums_validos = pd.to_numeric(df_tarefas["Número"], errors="coerce").dropna()
        if not nums_validos.empty:
            proximo_num = int(nums_validos.max()) + 1

    with st.form("form_nova_tarefa_planilha", clear_on_submit=True):
        c1, c2, c3 = st.columns([1.2, 1, 1.2])
        resp_lista = ["Fabíola", "Fabrício", "Outro"]
        f_resp = c1.selectbox("Responsável", resp_lista)
        f_venc = c2.date_input("Vencimento", value=date.today(), format="DD/MM/YYYY")
        
        # Categorias existentes para sugestão
        cats_existentes = sorted(list(set(df_tarefas["Categoria"].dropna().unique())))
        cats_existentes = [c for c in cats_existentes if str(c).strip()]
        if "Site" not in cats_existentes:
            cats_existentes.append("Site")
        if "Outros" not in cats_existentes:
            cats_existentes.append("Outros")
            
        f_cat = c3.selectbox("Categoria", cats_existentes)
        
        c4, c5 = st.columns([3, 1])
        f_desc = c4.text_input("Descrição da Tarefa", placeholder="Ex: Revisar petição inicial...")
        f_status = c5.selectbox("Status Inicial", ["Pendente", "Concluído", "Atrasado"])

        if st.form_submit_button("Salvar na Planilha", type="primary"):
            if f_desc.strip():
                data_formatada = f_venc.strftime("%d/%m/%Y")
                nova_linha = [
                    f_resp,
                    str(proximo_num),
                    data_formatada,
                    f_cat,
                    f_desc.strip(),
                    f_status
                ]
                aba.append_row(nova_linha)
                st.success(f"Tarefa nº {proximo_num} salva com sucesso!")
                st.cache_resource.clear()
                st.rerun()
            else:
                st.warning("Preencha a descrição da tarefa.")

# ----------------------------------------------------
# FILTROS E VISUALIZAÇÃO DA TABELA
# ----------------------------------------------------
st.subheader("📋 Lista de Tarefas")

col_f1, col_f2, col_f3, col_f4 = st.columns([1.2, 1.2, 1.2, 2])

filtro_resp = col_f1.selectbox("Filtrar Responsável:", ["Todos"] + sorted(list(set(df_tarefas["Responsável"].dropna().unique()))))
filtro_status = col_f2.selectbox("Filtrar Status:", ["Todos"] + sorted(list(set(df_tarefas["Status"].dropna().unique()))))
filtro_cat = col_f3.selectbox("Filtrar Categoria:", ["Todas"] + sorted(list(set(df_tarefas["Categoria"].dropna().unique()))))
busca_desc = col_f4.text_input("Buscar por Descrição:", placeholder="Digite uma palavra...")

# Aplicação dos filtros
df_filtrado = df_tarefas.copy()

if filtro_resp != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Responsável"] == filtro_resp]

if filtro_status != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Status"] == filtro_status]

if filtro_cat != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Categoria"] == filtro_cat]

if busca_desc.strip():
    df_filtrado = df_filtrado[df_filtrado["Descrição"].str.contains(busca_desc, case=False, na=False)]

# Ordenação por Número decrescente
if not df_filtrado.empty:
    df_filtrado["_num"] = pd.to_numeric(df_filtrado["Número"], errors="coerce").fillna(0)
    df_filtrado = df_filtrado.sort_values(by="_num", ascending=True).drop(columns=["_num"])

st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

# ----------------------------------------------------
# ATUALIZAÇÃO RÁPIDA DE STATUS
# ----------------------------------------------------
with st.expander("✏️ Atualizar Status de uma Tarefa"):
    if not df_tarefas.empty:
        opcoes_tarefas = {
            f"Nº {r['Número']} - {r['Descrição'][:50]} ({r['Responsável']})": r['Número']
            for _, r in df_tarefas.iterrows() if str(r['Número']).strip()
        }
        
        tarefa_selecionada = st.selectbox("Selecione a tarefa:", list(opcoes_tarefas.keys()))
        num_tarefa = opcoes_tarefas[tarefa_selecionada]
        
        col_st1, col_st2 = st.columns([1.5, 1])
        novo_status = col_st1.selectbox("Novo Status:", ["Concluído", "Pendente", "Atrasado"])
        
        if col_st2.button("Confirmar Atualização", type="primary"):
            # Localiza a linha correspondente na planilha
            todas_celulas = aba.col_values(2)  # Coluna B: Número
            linha_index = None
            for idx, valor in enumerate(todas_celulas, start=1):
                if str(valor).strip() == str(num_tarefa).strip():
                    linha_index = idx
                    break
            
            if linha_index:
                aba.update_cell(linha_index, 6, novo_status)  # Coluna F: Status
                st.success(f"Status da tarefa nº {num_tarefa} atualizado para '{novo_status}'!")
                st.cache_resource.clear()
                st.rerun()
            else:
                st.error("Não foi possível encontrar a linha da tarefa na planilha.")
