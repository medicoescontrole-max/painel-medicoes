import streamlit as st
import pandas as pd
import requests
from io import BytesIO

st.set_page_config(
    page_title="Consulta de Medições",
    page_icon="📊",
    layout="wide"
)

# --- 1. AUTENTICAÇÃO ---
def checar_login():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            st.markdown("## 🔒 Consulta de Medições")
            st.caption("Acesso restrito para acompanhamento de contratos e medições.")
            with st.form("form_login"):
                usuario = st.text_input("Usuário").strip().lower()
                senha = st.text_input("Senha", type="password")
                entrar = st.form_submit_button("Entrar", use_container_width=True)

                if entrar:
                    usuarios_validos = st.secrets.get("usuarios", {})
                    if usuario in usuarios_validos and str(usuarios_validos[usuario]) == senha:
                        st.session_state["autenticado"] = True
                        st.session_state["usuario_atual"] = usuario
                        st.rerun()
                    else:
                        st.error("Usuário ou senha inválidos.")
        return False
    return True

# --- 2. DOWNLOAD E LISTAGEM DE ABAS DO GOOGLE DRIVE ---
@st.cache_data(ttl=60)
def baixar_arquivo_drive():
    file_id = st.secrets["DRIVE_FILE_ID"]
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.content

@st.cache_data(ttl=60)
def listar_abas(conteudo_bytes):
    excel = pd.ExcelFile(BytesIO(conteudo_bytes), engine="openpyxl")
    # Filtra abas que contenham "STATUS" ou retorna todas caso mude o padrão de nome
    abas_status = [aba for aba in excel.sheet_names if "STATUS" in aba.upper()]
    return abas_status if abas_status else excel.sheet_names

@st.cache_data(ttl=60)
def processar_aba(conteudo_bytes, nome_aba):
    df = pd.read_excel(
        BytesIO(conteudo_bytes),
        sheet_name=nome_aba,
        header=2,
        usecols="A:O",
        engine="openpyxl"
    )

    # Limpa quebras de linha (\n) e espaços duplicados nos cabeçalhos
    df.columns = [" ".join(str(c).split()) for c in df.columns]

    if "N° CONTRATO" not in df.columns:
        return pd.DataFrame()

    df = df.dropna(subset=["N° CONTRATO"]).copy()

    cols_texto = [
        "N° CONTRATO", "EMPRESA", "CRITÉRIOS", "OBJETO", "VIGÊNCIA",
        "N° PERÍODO", "STATUS", "ENV. FONECEDOR ?", "LANÇ. NO RM", "OBS.:"
    ]
    for col in cols_texto:
        if col in df.columns:
            df[col] = df[col].fillna("-").astype(str).str.strip()

    for col_data in ["INICIO", "FIM", "DATA"]:
        if col_data in df.columns:
            df[col_data] = pd.to_datetime(df[col_data], errors="coerce").dt.strftime("%d/%m/%Y").fillna("-")

    col_valor = next((c for c in df.columns if "VALOR" in c.upper()), None)
    if col_valor:
        serie_val = df[col_valor].astype(str).str.replace("R$", "", regex=False).str.strip()
        df["VALOR_NUM"] = pd.to_numeric(df[col_valor], errors="coerce")
        mask_na = df["VALOR_NUM"].isna()
        if mask_na.any():
            limpo = serie_val[mask_na].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            df.loc[mask_na, "VALOR_NUM"] = pd.to_numeric(limpo, errors="coerce")
        df["VALOR_NUM"] = df["VALOR_NUM"].fillna(0.0)
        df["VALOR_EXIBICAO"] = df["VALOR_NUM"].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x > 0 else "-"
        )
    else:
        df["VALOR_NUM"] = 0.0
        df["VALOR_EXIBICAO"] = "-"

    return df

# --- 3. PAINEL PRINCIPAL ---
if checar_login():
    topo_esq, topo_dir = st.columns([4, 1])
    with topo_esq:
        st.title("📊 Status de Medições")
    with topo_dir:
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("🔄 Atualizar", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
        with c_btn2:
            if st.button("Sair", use_container_width=True):
                st.session_state["autenticado"] = False
                st.rerun()

    try:
        conteudo_excel = baixar_arquivo_drive()
        abas_disponiveis = listar_abas(conteudo_excel)
    except Exception as e:
        st.error(f"Falha ao conectar com a planilha no Google Drive: {e}")
        st.stop()

    # --- FILTROS LATERAIS ---
    st.sidebar.header("🔎 Filtros")

    aba_padrao = st.secrets.get("ABA_PLANILHA", abas_disponiveis[0])
    idx_padrao = abas_disponiveis.index(aba_padrao) if aba_padrao in abas_disponiveis else 0
    aba_selecionada = st.sidebar.selectbox("Período / Aba", abas_disponiveis, index=idx_padrao)

    df = processar_aba(conteudo_excel, aba_selecionada)
    if df.empty:
        st.warning("Nenhum dado encontrado nesta aba. Verifique se o cabeçalho está na linha 3.")
        st.stop()

    vigencia_sel = st.sidebar.radio("Vigência do Contrato", ["ATIVO", "ENCERRADO", "Todos"], index=0)
    busca = st.sidebar.text_input("Buscar Contrato ou Empresa", placeholder="Ex: 053-0017 ou ARMAC")

    df_base = df if vigencia_sel == "Todos" else df[df["VIGÊNCIA"] == vigencia_sel]

    lista_status = ["Todos"] + sorted(df_base["STATUS"].unique().tolist())
    status_sel = st.sidebar.selectbox("Status / Localização", lista_status)

    lista_objeto = ["Todos"] + sorted([x for x in df_base["OBJETO"].unique().tolist() if x != "-"])
    objeto_sel = st.sidebar.selectbox("Objeto", lista_objeto)

    df_filtrado = df_base.copy()
    if busca:
        termo = busca.lower()
        df_filtrado = df_filtrado[
            df_filtrado["N° CONTRATO"].str.lower().str.contains(termo) |
            df_filtrado["EMPRESA"].str.lower().str.contains(termo)
        ]
    if status_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["STATUS"] == status_sel]
    if objeto_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["OBJETO"] == objeto_sel]

    # --- INDICADORES ---
    total = len(df_filtrado)
    realizadas = len(df_filtrado[df_filtrado["VALOR_NUM"] > 0])
    faltantes = total - realizadas
    valor_total = df_filtrado["VALOR_NUM"].sum()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Contratos Filtrados", total)
    m2.metric("Realizadas (c/ Valor)", realizadas)
    m3.metric("Faltantes / Zeradas", faltantes)
    m4.metric("Valor Realizado", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.divider()

    # --- FORMATAÇÃO DA TABELA ---
    def colorir_status(val):
        v = str(val).upper()
        if "FINANCEIRO" in v:
            return "background-color: #198754; color: white; font-weight: bold;"
        elif "MEDIÇÕES" in v or "PENDÊNCIA" in v:
            return "background-color: #dc3545; color: white; font-weight: bold;"
        elif "GERENTE" in v:
            return "background-color: #0d6efd; color: white; font-weight: bold;"
        elif "FORNECEDOR" in v or "TIAGO" in v or "PEDRO" in v or "ADITIVO" in v:
            return "background-color: #ffc107; color: black; font-weight: bold;"
        return ""

    colunas_desejadas = [
        "N° CONTRATO", "EMPRESA", "OBJETO", "VALOR_EXIBICAO",
        "N° PERÍODO", "STATUS", "DATA", "ENV. FONECEDOR ?", "LANÇ. NO RM", "OBS.:"
    ]
    colunas_existentes = [c for c in colunas_desejadas if c in df_filtrado.columns]

    df_exibir = df_filtrado[colunas_existentes].rename(columns={"VALOR_EXIBICAO": "VALOR REALIZADO"})

    st.dataframe(
        df_exibir.style.map(colorir_status, subset=["STATUS"]),
        use_container_width=True,
        hide_index=True,
        height=600
    )
