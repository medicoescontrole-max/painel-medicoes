import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime, timedelta
import extra_streamlit_components as stx

st.set_page_config(
    page_title="Portal de Medições",
    page_icon="📊",
    layout="wide"
)

# Estilização visual fina para aproximar os cards e dar o acabamento premium do protótipo
st.markdown("""
    <style>
    /* Estilo para a barra compacta de status */
    .status-container {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        padding: 10px 14px;
        background-color: #f0f4f8;
        border-radius: 8px;
        border-left: 4px solid #0d6efd;
        margin-bottom: 20px;
        align-items: center;
    }
    .status-title {
        font-weight: bold;
        color: #1e293b;
        margin-right: 8px;
        font-size: 0.9rem;
    }
    .status-badge {
        display: inline-flex;
        align-items: center;
        background-color: #ffffff;
        color: #475569;
        font-size: 0.8rem;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 6px;
        border: 1px solid #cbd5e1;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    .status-badge .count {
        background-color: #3b82f6;
        color: #ffffff;
        font-size: 0.75rem;
        border-radius: 4px;
        padding: 0px 4px;
        margin-left: 5px;
    }
    /* Estilização para deixar os cards de métricas mais encorpados */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 12px 18px !important;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    [data-testid="stMetricLabel"] {
        color: #64748b !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        color: #1e293b !important;
    }
    /* Reduz espaços em branco excessivos */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

cookie_manager = stx.CookieManager(key="gerenciador_sessao")

# --- 1. AUTENTICAÇÃO ---
def checar_login():
    usuarios_validos = st.secrets.get("usuarios", {})

    usuario_cookie = cookie_manager.get(cookie="usuario_medicoes")
    if usuario_cookie and usuario_cookie in usuarios_validos:
        st.session_state["autenticado"] = True
        st.session_state["usuario_atual"] = usuario_cookie
        return True

    if st.session_state.get("autenticado", False):
        return True

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("## 🔒 Portal de Medições")
        st.caption("Insira seu login e senha para acompanhar o andamento.")
        with st.form("form_login"):
            usuario = st.text_input("Usuário").strip().lower()
            senha = st.text_input("Senha", type="password")
            lembrar = st.checkbox("Manter conectado por 7 dias", value=True)
            entrar = st.form_submit_button("Entrar", use_container_width=True)

            if entrar:
                if usuario in usuarios_validos and str(usuarios_validos[usuario]) == senha:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_atual"] = usuario
                    if lembrar:
                        expira = datetime.now() + timedelta(days=7)
                        cookie_manager.set("usuario_medicoes", usuario, expires_at=expira)
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos.")
    return False

# --- 2. LEITURA DO GOOGLE DRIVE ---
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
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    if "N° CONTRATO" not in df.columns:
        return pd.DataFrame()

    df = df.dropna(subset=["N° CONTRATO"]).copy()
    df = df[~df["N° CONTRATO"].astype(str).str.upper().str.contains("TOTAL")].copy()

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

# --- 3. REGRA DE TRADUÇÃO DE % PARA O FORNECEDOR ---
def calcular_etapa_fornecedor(status_interno):
    s = str(status_interno).upper()
    if "FINANCEIRO" in s:
        return 100, "Concluído / Encaminhado ao Financeiro"
    elif "GERENTE" in s or "VAGNER" in s:
        return 90, "Aprovação Final da Gerência (Penúltima Etapa)"
    elif "MEDIÇÕES / CONTRATOS" in s:
        return 80, "Conferência Final de Documentação"
    elif "FORNECEDOR" in s:
        return 65, "Aguardando Assinatura / Retorno do Fornecedor"
    elif any(k in s for k in ["TIAGO", "PEDRO", "BRUNO", "ORLANDO", "SAMUEL", "JANER", "LEYDSON", "RODRIGO"]):
        return 45, "Em Validação Técnica / Assinaturas Iniciais"
    elif any(k in s for k in ["PENDÊNCIA", "REVISÃO", "ADITIVO", "APROPRIAÇÃO", "ESPELHO", "SETOR"]):
        return 25, "Em Conferência Técnica / Ajustes"
    elif "SEM MEDIÇÃO" in s or "ENCERRADO" in s or s == "-":
        return 0, "Sem Medição no Período"
    return 30, "Em Processamento"

# --- 4. EXECUÇÃO DO PAINEL ---
if checar_login():
    usuario_logado = st.session_state.get("usuario_atual", "")
    permissoes = st.secrets.get("permissoes", {})
    escopo = str(permissoes.get(usuario_logado, "TODOS")).strip()
    eh_interno = (escopo.upper() == "TODOS")

    topo_esq, topo_dir = st.columns([4, 1])
    with topo_esq:
        st.title("📊 Status de Medições" if eh_interno else "📋 Acompanhamento de Medição")
    with topo_dir:
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("🔄 Atualizar", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
        with c_btn2:
            if st.button("Sair", use_container_width=True):
                st.session_state["autenticado"] = False
                cookie_manager.delete("usuario_medicoes")
                st.rerun()

    try:
        conteudo_excel = baixar_arquivo_drive()
        abas_disponiveis = listar_abas(conteudo_excel)
    except Exception as e:
        st.error(f"Falha ao conectar com a planilha no Google Drive: {e}")
        st.stop()

    aba_padrao = st.secrets.get("ABA_PLANILHA", abas_disponiveis[0])
    idx_padrao = abas_disponiveis.index(aba_padrao) if aba_padrao in abas_disponiveis else 0

    # ==========================================================
    # VISÃO A: EQUIPE INTERNA (COMPLETA + FILTRO INCLUIR/EXCLUIR)
    # ==========================================================
    if eh_interno:
        st.sidebar.header("🔎 Filtros")
        aba_selecionada = st.sidebar.selectbox("Período / Aba", abas_disponiveis, index=idx_padrao)
        df = processar_aba(conteudo_excel, aba_selecionada)

        if df.empty:
            st.warning("Nenhum dado encontrado nesta aba.")
            st.stop()

        vigencia_sel = st.sidebar.radio("Vigência do Contrato", ["ATIVO", "ENCERRADO", "Todos"], index=0)
        busca = st.sidebar.text_input("Buscar Contrato ou Empresa", placeholder="Ex: 053-0017 ou ARMAC")
        df_base = df if vigencia_sel == "Todos" else df[df["VIGÊNCIA"] == vigencia_sel]

        # Contagem prévia por status para exibir nas opções do filtro
        contagem_base = df_base["STATUS"].value_counts().to_dict()
        lista_status = sorted(contagem_base.keys())

        st.sidebar.markdown("---")
        modo_status = st.sidebar.radio(
            "Filtro de Status / Localização",
            ["Incluir apenas", "Excluir selecionados"],
            horizontal=True
        )
        status_selecionados = st.sidebar.multiselect(
            "Selecione um ou mais status:",
            options=lista_status,
            default=[],
            format_func=lambda s: f"{s} ({contagem_base.get(s, 0)})",
            placeholder="Todos (nenhum filtro aplicado)"
        )

        st.sidebar.markdown("---")
        lista_objeto = ["Todos"] + sorted([x for x in df_base["OBJETO"].unique().tolist() if x != "-"])
        objeto_sel = st.sidebar.selectbox("Objeto", lista_objeto)

        # Aplicação dos filtros
        df_filtrado = df_base.copy()
        if busca:
            termo = busca.lower()
            df_filtrado = df_filtrado[
                df_filtrado["N° CONTRATO"].str.lower().str.contains(termo) |
                df_filtrado["EMPRESA"].str.lower().str.contains(termo)
            ]

        if status_selecionados:
            if modo_status == "Incluir apenas":
                df_filtrado = df_filtrado[df_filtrado["STATUS"].isin(status_selecionados)]
            else:
                df_filtrado = df_filtrado[~df_filtrado["STATUS"].isin(status_selecionados)]

        if objeto_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado["OBJETO"] == objeto_sel]

        # --- INDICADORES GERAIS ---
        total = len(df_filtrado)
        realizadas = len(df_filtrado[df_filtrado["VALOR_NUM"] > 0])
        faltantes = total - realizadas
        valor_total = df_filtrado["VALOR_NUM"].sum()

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Contratos Filtrados", total)
        m2.metric("Realizadas (c/ Valor)", realizadas)
        m3.metric("Faltantes / Zeradas", faltantes)
        m4.metric("Valor Realizado", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        # --- CONTAGEM DETALHADA EM BADGES COMPACTOS ---
        if not df_filtrado.empty:
            resumo_status = df_filtrado["STATUS"].value_counts().items()
            badges_html = "".join([
                f'<span class="status-badge">{st_nome}<span class="count">{qtd}</span></span>'
                for st_nome, qtd in resumo_status
            ])
            st.markdown(
                f'<div class="status-container">'
                f'<span class="status-title">📍 Quantidades na seleção atual ({total} total):</span>'
                f'{badges_html}'
                f'</div>',
                unsafe_allow_html=True
            )

        st.divider()

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

    # ==========================================================
    # VISÃO B: FORNECEDOR (RESTRITA AO CONTRATO + % PROGRESSO)
    # ==========================================================
    else:
        aba_selecionada = st.selectbox("Selecione o Período de Referência:", abas_disponiveis, index=idx_padrao)
        df = processar_aba(conteudo_excel, aba_selecionada)

        contratos_permitidos = [c.strip().upper() for c in escopo.split(",")]
        df_forn = df[df["N° CONTRATO"].str.upper().isin(contratos_permitidos)].copy()

        if df_forn.empty:
            st.info("Nenhum registro localizado para o seu contrato neste período.")
            st.stop()

        for _, row in df_forn.iterrows():
            pct, descricao_etapa = calcular_etapa_fornecedor(row["STATUS"])

            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.subheader(f"{row['EMPRESA']} — Contrato {row['N° CONTRATO']}")
                    st.caption(f"Período: {row.get('INICIO', '-')} a {row.get('FIM', '-')}  |  Medição Nº: {row.get('N° PERÍODO', '-')}")
                with c2:
                    st.metric("Progresso de Assinaturas", f"{pct}%")

                st.progress(pct / 100.0)
                st.markdown(f"**Situação Atual:** {descricao_etapa} *(Atualizado em {row.get('DATA', '-')})*")

                if pct == 75:
                    st.warning("Ação necessária: A medição já foi enviada e aguarda assinatura/retorno da sua empresa.")
                elif pct == 100:
                    st.success("Processo de assinatura concluído e encaminhado ao setor Financeiro.")
