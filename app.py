import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta, time

# ----------------------------------------------------
# CONFIGURAÇÃO GERAL
# ----------------------------------------------------
st.set_page_config(page_title="Agenda e Tarefas - Fabíola Guimarães", layout="wide", page_icon="⚖️")

NOME_PLANILHA = "Agenda e Tarefas"  # Altere para o nome exato da sua planilha no Google Sheets

# ----------------------------------------------------
# CREDENCIAIS GOOGLE SERVICE ACCOUNT EMBUTIDAS
# ----------------------------------------------------
SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
    "project_id": "agenda-simples-513616",
    "private_key_id": "799585a59acb5eaf1fe040463eba6503a2e5599b",
    "private_key": (
        "-----BEGIN PRIVATE KEY-----\n"
        "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC9EJNizZirmFKi\n"
        "Hc5o8kvmrRIrWa4QrNHW3CDYwHTF4pJF9obQOGe1XTSbU6kHP+ZYrkHnPQfxfzgT\n"
        "yklmn6pYRvvnFN8VkdLN0yRMQiF5GjSOpBdsgcTOBILzk2KNaECx8GVbooJNHkmk\n"
        "gyakDk7ZObMOAfogTmblho4cYJMkdbWUNPjwZDrWqFC6rTUR04b5GBhjBMU0vUqR\n"
        "aKu/6Bx0Kiizhqv2o8Pd3yQKt7FgKEG93XqBfMOFvRxUsocxuQVpqD5uJBDaLLAD\n"
        "3gJO+VY9NprCps3P/3PI9pWd5mx6h33NNiQcIvV8VAyp8NnEACQRu6SDGM9uXck3\n"
        "8YxooxmLAgMBAAECggEALTYhiKcRszchy5LFIWOPswIALk2rymEJbRAG6s5bHlco\n"
        "E4L4iiJy6nP3eI1xBkLJmrtRFzMJCz8Y8eKV2uyjl6ih5IihME9G0eGD9Jh72LHj\n"
        "j5i94TzHC10Y05aTIq2jiump01fxQO/VPHoAa4rgYCHkJ51hawhUot6ibTbK+RSj\n"
        "tFD0cDREDPyps2q8wu1i8nqZsVTzjQSmMShSo8o3BXRrLfOrnZW3ItHX8zVpdQH0\n"
        "YnJKq80X0D6mOgQAbFZoFh0GIbtLrQAjHjWH8Jc93UHpOdfJJuQyhGh+3ikJtJ92\n"
        "2GLKtTmecoMYdpBzcd4cjc9krFJMv8lLVojTYulawQKBgQDfHtYbryq4I2rluKbi\n"
        "9KyWBST7jjaLX0VhxVQ4wHSkTQc6IUxU3eeceoNxTZ8JpkC4v/FqNqNrtm7ac9/7\n"
        "t5mWnRKoQAngtmhHnMyJyKgtF3IlxP8PHSupQqLXEIJAVDceMQ/8iw+UjkMZZ5zT\n"
        "VhE6BQHrXy/pPlWIAm2p5WPP5wKBgQDY7P8QdkTPY2IpA4U1ArhKeF4SDSZ8L8Om\nlCxunkb+Gqpvgzf2X6wA1D+wCuSzDkRs0vNMTh6sqfTl3pCDG6k6D8zs1y1wsLUd\nz9y07dP+3V/9W7pAnXBzvnhH6vV9VuhL4UBe1/azMK/wGH78VEt18VlfTZz1DQrR\nOFJ08VUEvQKBgQDF6QBaRiJKe4ASjJvbrNR5rc6EJD1VmwhcFUEDphLQulKyyV5/\n2OXc8v8svYFYedPnDLsDT7vdi1EBDM+c49qksmcePs+EN8h/3SHBBe/HL4jc38Su\nj11vIeB3i/k9PnY1nadkh6UmTJpQKov9Gdbl9BMjyKci6VvM7PqBPSwv+QKBgBJu\n6Sm78xLwbgnwK6Sai5CAc/Xyuh95a5W/J5tqxbuYBn2dV+cOsC4N2y+ZbV+tybC5\nqEUGqmKfupjB9eO+OP4fE5nYmxrlWS6kWCvjvcz9gMelwMj21/Rt7gimdZ/l/61b\nAb5NecA2Ve0AF+JXQ2ZNUPvhO4IzToUC1Tbsvv29AoGAWQkL28gKb6O933u0GUtv\niwsVcOZSjXjF42K+Bo+SANKFfDcYCCxW+RBl03e8a+ikdXNbdTaYhezqJVfxr6YX\nMKN5nqannyBhI8rG+fuqYbThoYD7JjqsB7rkFRLyi5O0rBTi4QIGbSVqAe6N+pWz\nDu4RCJk924LMjRitH5XISjQ=\n"
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
        # Clona o dicionário e garante a substituição correta de escapes literais
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
