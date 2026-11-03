import os
from datetime import date
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

# --- Configuração de Página ---
st.set_page_config(
    page_title="Tarefas",
    page_icon="✅",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# --- Estilos CSS ---
st.markdown(
    """
<style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 5rem;
        max-width: 600px;
    }
    .task-card {
        background-color: var(--secondary-background-color, #f8f9fa);
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 10px;
        display: flex;
        flex-direction: column;
        gap: 6px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }
    .task-title {
        font-size: 1.02rem;
        font-weight: 500;
        color: var(--text-color);
        line-height: 1.35;
    }
    .task-meta {
        font-size: 0.8rem;
        display: flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
    }
    .badge-chip {
        background-color: rgba(66, 133, 244, 0.12);
        color: #1a73e8;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 500;
    }
    .badge-cat {
        background-color: rgba(100, 100, 100, 0.1);
        color: #555;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
    }
    .badge-date {
        color: #d93025;
        font-size: 0.75rem;
        font-weight: 500;
    }
</style>
""",
    unsafe_allow_html=True,
)

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/13NLC3A0QVt3I3DPOqHS5gKijuZtbbzHcaUepq-tMihQ/edit#gid=0"


# --- Autenticação ---
@st.cache_resource
def get_gspread_client():
  scopes = [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
  ]
  if os.path.exists("service_account.json"):
    creds = Credentials.from_service_account_file(
        "service_account.json", scopes=scopes
    )
  elif "gcp_service_account" in st.secrets:
    service_info = dict(st.secrets["gcp_service_account"])
    if "\\n" in service_info.get("private_key", ""):
      service_info["private_key"] = service_info["private_key"].replace(
          "\\n", "\n"
      )
    creds = Credentials.from_service_account_info(service_info, scopes=scopes)
  else:
    raise FileNotFoundError(
        "Coloque o ficheiro service_account.json na raiz do projecto."
    )

  return gspread.authorize(creds)


def load_data():
  client = get_gspread_client()
  sheet = client.open_by_url(SPREADSHEET_URL).worksheet("Página1")
  data = sheet.get_all_records()
  df = pd.DataFrame(data)

  if not df.empty and "Descrição" in df.columns:
    df = df[df["Descrição"].astype(str).str.strip() != ""]
  return sheet, df


# --- Carregar Dados ---
try:
  sheet, df = load_data()
except Exception as e:
  st.error(f"Erro ao ligar à folha de cálculo: {e}")
  st.stop()

st.title("Tarefas")

# Consolidação de Responsáveis
all_responsaveis = set(["Fabíola", "Fabrício"])
if not df.empty and "Responsável" in df.columns:
  for item in df["Responsável"].dropna():
    partes = [p.strip() for p in str(item).split(",") if p.strip()]
    all_responsaveis.update(partes)
sorted_resp = sorted(list(all_responsaveis))

col_f1, col_f2 = st.columns([1, 1])
with col_f1:
  filtro_resp = st.selectbox("Responsável", ["Todos"] + sorted_resp)
with col_f2:
  filtro_status = st.selectbox(
      "Estado", ["Pendentes", "Concluídas", "Todos"], index=0
  )

df_filtered = df.copy()

if filtro_resp != "Todos" and not df_filtered.empty:
  df_filtered = df_filtered[
      df_filtered["Responsável"]
      .astype(str)
      .str.contains(filtro_resp, case=False, na=False)
  ]

if not df_filtered.empty and "Status" in df_filtered.columns:
  if filtro_status == "Pendentes":
    df_filtered = df_filtered[
        df_filtered["Status"].str.strip().str.lower() != "concluído"
    ]
  elif filtro_status == "Concluídas":
    df_filtered = df_filtered[
        df_filtered["Status"].str.strip().str.lower() == "concluído"
    ]

st.caption(f"{len(df_filtered)} tarefa(s) apresentada(s)")

# --- Lista de Tarefas ---
for idx, row in df_filtered.iterrows():
  num_tarefa = row.get("Número", "")
  descricao = row.get("Descrição", "")
  resp = row.get("Responsável", "")
  vencimento = row.get("Vencimento", "")
  categoria = row.get("Categoria", "")
  status = str(row.get("Status", "Pendente")).strip()

  is_concluido = status.lower() == "concluído"

  col_check, col_body = st.columns([0.15, 0.85])

  with col_check:
    marcado = st.checkbox(
        label=f"check_{num_tarefa}_{idx}",
        value=is_concluido,
        label_visibility="collapsed",
        key=f"task_{num_tarefa}_{idx}",
    )
    if marcado != is_concluido:
      novo_status = "Concluído" if marcado else "Pendente"
      cell = sheet.find(str(num_tarefa))
      if cell:
        sheet.update_cell(cell.row, 6, novo_status)
        st.cache_resource.clear()
        st.rerun()

  with col_body:
    style_desc = (
        "text-decoration: line-through; opacity: 0.55;" if is_concluido else ""
    )
    chips_resp = "".join([
        f'<span class="badge-chip">👤 {r.strip()}</span> '
        for r in str(resp).split(",")
        if r.strip()
    ])
    cat_badge = (
        f'<span class="badge-cat">📁 {categoria}</span> ' if categoria else ""
    )
    date_badge = (
        f'<span class="badge-date">📅 {vencimento}</span> '
        if vencimento
        else ""
    )

    st.markdown(
        f"""
        <div class="task-card">
            <div class="task-title" style="{style_desc}">{descricao}</div>
            <div class="task-meta">
                {chips_resp}
                {cat_badge}
                {date_badge}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# --- Formulário de Inclusão ---
st.write("---")
with st.expander("➕ Adicionar Nova Tarefa", expanded=False):
  with st.form("form_nova_tarefa", clear_on_submit=True):
    nova_desc = st.text_input(
        "Descrição da Tarefa", placeholder="Ex: Rever texto do site principal"
    )

    novos_responsaveis = st.multiselect(
        "Responsáveis",
        options=sorted_resp,
        default=["Fabíola"] if "Fabíola" in sorted_resp else [],
    )

    col_c1, col_c2 = st.columns(2)
    with col_c1:
      categorias_existentes = (
          [c for c in df["Categoria"].dropna().unique() if str(c).strip()]
          if not df.empty
          else []
      )
      opcoes_cat = sorted(
          list(set(categorias_existentes + ["Site", "Outros"]))
      )
      nova_categoria = st.selectbox("Categoria", opcoes_cat)
    with col_c2:
      nova_data = st.date_input("Data de Vencimento", value=date.today())

    submetido = st.form_submit_button(
        "Guardar Tarefa", use_container_width=True
    )

    if submetido:
      if not nova_desc.strip():
        st.warning("Por favor, preencha a descrição da tarefa.")
      else:
        try:
          numeros = pd.to_numeric(df["Número"], errors="coerce").dropna()
          proximo_num = int(numeros.max() + 1) if not numeros.empty else 1
        except Exception:
          proximo_num = len(df) + 1

        resp_str = ", ".join(novos_responsaveis) if novos_responsaveis else ""
        venc_str = nova_data.strftime("%d/%m/%Y")

        nova_linha = [
            resp_str,
            proximo_num,
            venc_str,
            nova_categoria,
            nova_desc,
            "Pendente",
        ]
        sheet.append_row(nova_linha)

        st.success("Tarefa adicionada com sucesso!")
        st.cache_resource.clear()
        st.rerun()