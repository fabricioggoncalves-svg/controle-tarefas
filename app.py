import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta, time

# ----------------------------------------------------
# CONFIGURAÇÃO GERAL
# ----------------------------------------------------
st.set_page_config(page_title="Agenda e Tarefas - Fabíola Guimarães", layout="wide", page_icon="⚖️")

NOME_PLANILHA = "Controle de Tarefas"  # Certifique-se de que o nome no Google Sheets seja exatamente este

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
        "Z5oAnuqUSHDfn/9Se4vLc4zflaDumjQs2XeUQp0t5qNNkxxLyl43CZ50VWVFCMbo\n"
        "dBipAaNSh6jP6oAX5xfLQkgfD1fIlD7o/MRghJRFwwKBgQDbAWKnFmTrIDTUQBbV\n"
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
# CONEXÃO COM GOOGLE SHEETS
# ----------------------------------------------------
@st.cache_resource(ttl=300)
def obter_conexao_sheets():
    """Autentica na API do Google Sheets tratando quebras de linha da private_key."""
    try:
        creds_dict = dict(SERVICE_ACCOUNT_INFO)
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        escopos = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credenciais = Credentials.from_service_account_info(creds_dict, scopes=escopos)
        cliente = gspread.authorize(credenciais)
        planilha = cliente.open(NOME_PLANILHA)
        return planilha
    except Exception as e:
        st.error(f"Erro ao ligar à folha de cálculo: {e}")
        return None

def obter_ou_criar_aba(planilha, nome_aba, cabecalhos):
    """Retorna uma aba da planilha; se não existir, cria com os cabeçalhos padrão."""
    try:
        return planilha.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome_aba, rows="100", cols=str(len(cabecalhos)))
        aba.append_row(cabecalhos)
        return aba

def carregar_dados(planilha, nome_aba, cabecalhos_padrao):
    """Lê os dados da aba e retorna como DataFrame."""
    aba = obter_ou_criar_aba(planilha, nome_aba, cabecalhos_padrao)
    valores = aba.get_all_records()
    if not valores:
        return pd.DataFrame(columns=cabecalhos_padrao)
    return pd.DataFrame(valores)

# ----------------------------------------------------
# INTERFACE PRINCIPAL
# ----------------------------------------------------
st.title("⚖️ Gestão de Tarefas e Agenda")
st.caption("Conectado ao Google Sheets (Persistência em Nuvem)")

planilha = obter_conexao_sheets()

if not planilha:
    st.warning(
        f"Certifique-se de que compartilhou a folha de cálculo **'{NOME_PLANILHA}'** "
        f"com o email da conta de serviço como **Editor**:\n\n"
        f"`{SERVICE_ACCOUNT_INFO['client_email']}`"
    )
    st.stop()

# Cabeçalhos padrão das abas
COLS_TAREFAS = ["ID", "Título", "Prazo", "Prioridade", "Status", "Responsável", "Descrição"]
COLS_AGENDA = ["ID", "Título", "Data Início", "Hora Início", "Data Fim", "Hora Fim", "Local/Link", "Descrição"]

aba_tarefas_sheets = obter_ou_criar_aba(planilha, "Tarefas", COLS_TAREFAS)
aba_agenda_sheets = obter_ou_criar_aba(planilha, "Agenda", COLS_AGENDA)

# Navegação
aba_sel_1, aba_sel_2 = st.tabs(["✅ Gestão de Tarefas", "📅 Agenda de Compromissos"])

# ====================================================
# ABA 1: TAREFAS
# ====================================================
with aba_sel_1:
    st.subheader("Quadro de Tarefas")
    
    with st.expander("➕ Nova Tarefa", expanded=False):
        with st.form("form_nova_tarefa_gs", clear_on_submit=True):
            c1, c2 = st.columns(2)
            t_titulo = c1.text_input("Título da Tarefa")
            t_prazo = c2.date_input("Prazo Limite", value=date.today(), format="DD/MM/YYYY")
            
            c3, c4 = st.columns(2)
            t_prio = c3.selectbox("Prioridade", ["Baixa", "Média", "Alta", "Urgente"], index=1)
            t_resp = c4.selectbox("Responsável", ["Fabrício Gonçalves", "Dra. Fabíola Guimarães", "Outro"])
            
            t_desc = st.text_area("Descrição / Anotações")
            
            if st.form_submit_button("Salvar Tarefa no Google Sheets", type="primary"):
                if t_titulo.strip():
                    df_t = carregar_dados(planilha, "Tarefas", COLS_TAREFAS)
                    novo_id = 1 if df_t.empty else int(df_t["ID"].max()) + 1
                    
                    nova_linha = [
                        novo_id,
                        t_titulo.strip(),
                        str(t_prazo),
                        t_prio,
                        "Não Iniciada",
                        t_resp,
                        t_desc.strip()
                    ]
                    aba_tarefas_sheets.append_row(nova_linha)
                    st.success("Tarefa gravada com sucesso na folha de cálculo!")
                    st.cache_resource.clear()
                    st.rerun()
                else:
                    st.warning("Preencha o título da tarefa.")

    # Listagem de Tarefas
    df_tarefas = carregar_dados(planilha, "Tarefas", COLS_TAREFAS)
    if df_tarefas.empty:
        st.info("Nenhuma tarefa registada até ao momento.")
    else:
        st.dataframe(df_tarefas, use_container_width=True)

# ====================================================
# ABA 2: AGENDA
# ====================================================
with aba_sel_2:
    st.subheader("Agenda de Compromissos")
    
    with st.expander("➕ Novo Compromisso", expanded=False):
        with st.form("form_novo_comp_gs", clear_on_submit=True):
            c1, c2 = st.columns(2)
            comp_tit = c1.text_input("Título do Compromisso")
            comp_loc = c2.text_input("Local ou Link da Reunião")
            
            c3, c4, c5, c6 = st.columns(4)
            comp_d_ini = c3.date_input("Data Início", value=date.today(), format="DD/MM/YYYY")
            comp_h_ini = c4.time_input("Hora Início", value=time(9, 0))
            comp_d_fim = c5.date_input("Data Término", value=date.today(), format="DD/MM/YYYY")
            comp_h_fim = c6.time_input("Hora Término", value=time(10, 0))
            
            comp_desc = st.text_area("Anotações do Compromisso")
            
            if st.form_submit_button("Salvar Compromisso", type="primary"):
                if comp_tit.strip():
                    df_a = carregar_dados(planilha, "Agenda", COLS_AGENDA)
                    novo_comp_id = 1 if df_a.empty else int(df_a["ID"].max()) + 1
                    
                    nova_linha_agenda = [
                        novo_comp_id,
                        comp_tit.strip(),
                        str(comp_d_ini),
                        comp_h_ini.strftime("%H:%M"),
                        str(comp_d_fim),
                        comp_h_fim.strftime("%H:%M"),
                        comp_loc.strip(),
                        comp_desc.strip()
                    ]
                    aba_agenda_sheets.append_row(nova_linha_agenda)
                    st.success("Compromisso gravado com sucesso!")
                    st.cache_resource.clear()
                    st.rerun()
                else:
                    st.warning("Preencha o título do compromisso.")

    # Exibição da Agenda
    df_agenda = carregar_dados(planilha, "Agenda", COLS_AGENDA)
    if df_agenda.empty:
        st.info("Nenhum compromisso agendado.")
    else:
        st.dataframe(df_agenda, use_container_width=True)
