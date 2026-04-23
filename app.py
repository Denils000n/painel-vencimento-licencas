import streamlit as st
import pandas as pd
import calendar
from datetime import date
from io import BytesIO

st.set_page_config(
    page_title="Painel de Vencimento de Licenças",
    page_icon="📅",
    layout="wide"
)

STATUS_VALIDOS = ["Pendente", "Em andamento", "Renovada"]

COLUNAS_SISTEMA = [
    "Empresa",
    "Colaborador",
    "Centro de Custo",
    "Tipo de Licença",
    "Valor da Licença",
    "Vencimento",
    "Status",
]

COLUNAS_MINIMAS = [
    "Colaborador",
    "Centro de Custo",
]

SINONIMOS = {
    "Empresa": [
        "empresa", "company", "organization", "organizacao", "organização", "tenant", "companhia"
    ],
    "Colaborador": [
        "colaborador", "usuario", "usuário", "user", "user name", "display name",
        "nome", "employee", "funcionario", "funcionário", "user principal name",
        "upn", "email", "mail", "e-mail", "account"
    ],
    "Centro de Custo": [
        "centro de custo", "centro custo", "cc", "cost center", "costcentre",
        "office", "office location", "localizacao", "localização", "obra", "obra codigo"
    ],
    "Tipo de Licença": [
        "tipo de licença", "tipo de licenca", "licenca", "licença", "license",
        "licenses", "assigned licenses", "sku", "sku name", "product", "product name",
        "service", "serviço", "service plan"
    ],
    "Valor da Licença": [
        "valor da licença", "valor da licenca", "valor", "price", "cost", "amount",
        "preco", "preço", "license value", "valor licenca"
    ],
    "Vencimento": [
        "vencimento", "due date", "expiration", "expiry", "renewal date",
        "data vencimento", "expiration date", "next renewal", "valid until"
    ],
    "Status": [
        "status", "situação", "situacao", "state", "renewal status"
    ],
}

OPCOES_ALERTA = [
    "30 dias",
    "60 dias",
    "90 dias",
    "120 dias",
    "12 meses",
    "24 meses",
    "36 meses",
    "48 meses",
    "60 meses",
]


def normalizar_texto(txt):
    return (
        str(txt)
        .strip()
        .lower()
        .replace("_", " ")
        .replace("-", " ")
    )


def formatar_brl(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def listar_abas_excel(arquivo):
    xls = pd.ExcelFile(arquivo)
    return xls.sheet_names


def ler_arquivo(arquivo, aba_escolhida=None, juntar_abas=False):
    if arquivo.name.endswith(".csv"):
        try:
            return pd.read_csv(arquivo)
        except Exception:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=";")

    xls = pd.ExcelFile(arquivo)
    abas = xls.sheet_names

    if juntar_abas:
        dfs = []
        for aba in abas:
            df_aba = pd.read_excel(xls, sheet_name=aba)
            df_aba["Aba_Origem"] = aba
            dfs.append(df_aba)
        return pd.concat(dfs, ignore_index=True)

    if aba_escolhida is None:
        aba_escolhida = abas[0]

    df = pd.read_excel(xls, sheet_name=aba_escolhida)
    df["Aba_Origem"] = aba_escolhida
    return df


def detectar_mapeamento_automatico(colunas):
    mapeamento = {}
    colunas_normalizadas = {col: normalizar_texto(col) for col in colunas}

    for campo_sistema, sinonimos in SINONIMOS.items():
        encontrada = None

        for col_original, col_norm in colunas_normalizadas.items():
            if col_norm == normalizar_texto(campo_sistema):
                encontrada = col_original
                break

        if encontrada is None:
            for col_original, col_norm in colunas_normalizadas.items():
                if col_norm in [normalizar_texto(s) for s in sinonimos]:
                    encontrada = col_original
                    break

        if encontrada is None:
            for col_original, col_norm in colunas_normalizadas.items():
                for s in sinonimos:
                    s_norm = normalizar_texto(s)
                    if s_norm in col_norm or col_norm in s_norm:
                        encontrada = col_original
                        break
                if encontrada is not None:
                    break

        mapeamento[campo_sistema] = encontrada

    return mapeamento


def identificar_empresa_pelo_centro(valor):
    if pd.isna(valor):
        return "Não informada"

    valor = str(valor).strip()
    prefixo = valor[:2]

    mapa = {
        "01": "Afonso França",
        "02": "AFFIT",
        "03": "AFDI",
        "04": "AFSW",
    }

    return mapa.get(prefixo, "Não informada")


def aplicar_mapeamento(df_original, mapeamento):
    df = pd.DataFrame()

    for campo in COLUNAS_SISTEMA:
        coluna_origem = mapeamento.get(campo)
        if coluna_origem and coluna_origem in df_original.columns:
            df[campo] = df_original[coluna_origem]
        else:
            df[campo] = None

    if "Aba_Origem" in df_original.columns:
        df["Aba_Origem"] = df_original["Aba_Origem"]
    else:
        df["Aba_Origem"] = "Sem aba"

    if df["Empresa"].isna().all() or df["Empresa"].astype(str).str.strip().eq("").all():
        df["Empresa"] = df["Centro de Custo"].apply(identificar_empresa_pelo_centro)

    if df["Tipo de Licença"].isna().all():
        df["Tipo de Licença"] = "Não informado"
    else:
        df["Tipo de Licença"] = df["Tipo de Licença"].fillna("Não informado")

    if df["Valor da Licença"].isna().all():
        df["Valor da Licença"] = 0
    else:
        df["Valor da Licença"] = (
            df["Valor da Licença"]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.strip()
        )
        df["Valor da Licença"] = pd.to_numeric(df["Valor da Licença"], errors="coerce").fillna(0)

    if df["Status"].isna().all():
        df["Status"] = "Pendente"
    else:
        df["Status"] = df["Status"].fillna("Pendente")

    df["Status"] = df["Status"].astype(str).str.strip()
    df.loc[~df["Status"].isin(STATUS_VALIDOS), "Status"] = "Pendente"

    df["Empresa"] = df["Empresa"].fillna("Não informada").astype(str).str.strip()
    df["Colaborador"] = df["Colaborador"].fillna("Não informado").astype(str).str.strip()
    df["Centro de Custo"] = df["Centro de Custo"].fillna("Não informado").astype(str).str.strip()
    df["Tipo de Licença"] = df["Tipo de Licença"].fillna("Não informado").astype(str).str.strip()
    df["Vencimento"] = pd.to_datetime(df["Vencimento"], errors="coerce")

    return df


def validar_minimos(df):
    faltantes = []
    for col in COLUNAS_MINIMAS:
        valores = df[col].fillna("").astype(str).str.strip() if col in df.columns else pd.Series(dtype=str)
        if col not in df.columns or valores.eq("").all():
            faltantes.append(col)
    return faltantes


def parse_alerta_config(alerta_str):
    valor, unidade = alerta_str.split()
    valor = int(valor)
    return valor, unidade


def texto_alerta_dinamico(alerta_str):
    return f"Vence em até {alerta_str}"


def esta_na_faixa_alerta(vencimento, status, alerta_str):
    if pd.isna(vencimento):
        return False
    if status in ["Renovada", "Em andamento"]:
        return False

    hoje = pd.Timestamp(date.today())
    if vencimento < hoje:
        return False

    valor, unidade = parse_alerta_config(alerta_str)

    if unidade == "dias":
        limite = hoje + pd.Timedelta(days=valor)
    else:
        limite = hoje + pd.DateOffset(months=valor)

    return hoje <= vencimento <= limite


def atualizar_alertas(df, alerta_str):
    hoje = pd.Timestamp(date.today())
    df["Dias para Vencer"] = (df["Vencimento"] - hoje).dt.days
    alerta_label = texto_alerta_dinamico(alerta_str)

    def definir_alerta(row):
        if pd.isna(row["Vencimento"]):
            return "Sem vencimento"
        if row["Status"] == "Renovada":
            return "Renovada"
        if row["Status"] == "Em andamento":
            return "Em andamento"
        if row["Dias para Vencer"] < 0:
            return "Vencida"
        if esta_na_faixa_alerta(row["Vencimento"], row["Status"], alerta_str):
            return alerta_label
        return "No prazo"

    df["Alerta"] = df.apply(definir_alerta, axis=1)
    return df


def gerar_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Licencas")
    return buffer.getvalue()


def cor_dia(df_dia, alerta_label):
    if df_dia.empty:
        return "#f8fafc"
    if (df_dia["Alerta"] == "Vencida").any():
        return "#fee2e2"
    if (df_dia["Alerta"] == alerta_label).any():
        return "#fecaca"
    if (df_dia["Status"] == "Em andamento").any():
        return "#fef3c7"
    if not df_dia.empty and (df_dia["Status"] == "Renovada").all():
        return "#dcfce7"
    return "#e2e8f0"


def borda_dia(df_dia, alerta_label):
    if df_dia.empty:
        return "#e5e7eb"
    if (df_dia["Alerta"] == "Vencida").any():
        return "#dc2626"
    if (df_dia["Alerta"] == alerta_label).any():
        return "#ef4444"
    if (df_dia["Status"] == "Em andamento").any():
        return "#d97706"
    if not df_dia.empty and (df_dia["Status"] == "Renovada").all():
        return "#16a34a"
    return "#94a3b8"


def resumo_dia(df_dia):
    if df_dia.empty:
        return "Sem itens"
    return f"{len(df_dia)} licença(s)"


def aplicar_filtros(df, empresa, centro, tipo, aba):
    df_filtrado = df.copy()

    if empresa != "Todas":
        df_filtrado = df_filtrado[df_filtrado["Empresa"] == empresa]

    if centro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Centro de Custo"] == centro]

    if tipo != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Tipo de Licença"] == tipo]

    if aba != "Todas":
        df_filtrado = df_filtrado[df_filtrado["Aba_Origem"] == aba]

    return df_filtrado


def adicionar_meses(data_base, meses):
    return pd.Timestamp(data_base) + pd.DateOffset(months=int(meses))


def obter_proxima_data_critica(df, alerta_label):
    criticos = df[df["Alerta"].isin(["Vencida", alerta_label])].copy()
    criticos = criticos.dropna(subset=["Vencimento"]).sort_values("Vencimento")
    if criticos.empty:
        return None
    return criticos.iloc[0]["Vencimento"].date()


st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 1rem;
}
.main-title {
    font-size: 30px;
    font-weight: 700;
    margin-bottom: 4px;
    color: #0f172a;
}
.sub-title {
    font-size: 14px;
    color: #475569;
    margin-bottom: 18px;
}
.card-info {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 14px;
    box-shadow: 0 1px 4px rgba(15,23,42,0.05);
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📅 Painel de Vencimento de Licenças</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Calendário visual, leitura de múltiplas abas, alertas de vencimento, navegação mensal e renovação rápida.</div>',
    unsafe_allow_html=True
)

if "filtro_alerta" not in st.session_state:
    st.session_state.filtro_alerta = "Todos"

if "alerta_config" not in st.session_state:
    st.session_state.alerta_config = "30 dias"

arquivo = st.file_uploader(
    "📁 Suba a planilha revisada do mês (CSV ou Excel)",
    type=["csv", "xlsx"]
)

if arquivo:
    if arquivo.name.endswith(".xlsx"):
        abas_disponiveis = listar_abas_excel(arquivo)

        modo_leitura = st.radio(
            "Como deseja ler o Excel?",
            ["Usar uma aba", "Juntar todas as abas"],
            horizontal=True
        )

        if modo_leitura == "Usar uma aba":
            aba_escolhida = st.selectbox("Selecione a aba", abas_disponiveis)
            juntar_abas = False
        else:
            aba_escolhida = None
            juntar_abas = True
    else:
        abas_disponiveis = []
        aba_escolhida = None
        juntar_abas = False
        modo_leitura = "CSV"

    assinatura_arquivo = f"{arquivo.name}|{modo_leitura}|{aba_escolhida}|{juntar_abas}"

    if (
        "arquivo_base" not in st.session_state
        or st.session_state.arquivo_base != assinatura_arquivo
    ):
        df_original = ler_arquivo(
            arquivo,
            aba_escolhida=aba_escolhida,
            juntar_abas=juntar_abas
        )
        df_original.columns = [str(c).strip() for c in df_original.columns]

        mapeamento_auto = detectar_mapeamento_automatico(df_original.columns)
        df_mapeado = aplicar_mapeamento(df_original, mapeamento_auto)
        faltantes = validar_minimos(df_mapeado)

        if faltantes:
            st.error(
                "Não foi possível identificar automaticamente os campos mínimos necessários. "
                f"Campos esperados: {', '.join(faltantes)}"
            )
            st.stop()

        df_mapeado = atualizar_alertas(df_mapeado, st.session_state.alerta_config)

        st.session_state.df_licencas = df_mapeado.copy()
        st.session_state.df_original = df_original.copy()
        st.session_state.arquivo_base = assinatura_arquivo
        st.session_state.data_selecionada = None
        st.session_state.filtro_alerta = "Todos"
        st.session_state.mes_sel = date.today().month
        st.session_state.ano_sel = date.today().year

    st.session_state.df_licencas = atualizar_alertas(
        st.session_state.df_licencas.copy(),
        st.session_state.alerta_config
    )

    df = st.session_state.df_licencas.copy()
    alerta_label = texto_alerta_dinamico(st.session_state.alerta_config)

    with st.sidebar:
        st.header("Filtros")

        empresas = ["Todas"] + sorted(df["Empresa"].dropna().unique().tolist())
        empresa_sel = st.selectbox("Empresa", empresas)

        centros = ["Todos"] + sorted(df["Centro de Custo"].dropna().unique().tolist())
        centro_sel = st.selectbox("Centro de custo", centros)

        tipos = ["Todos"] + sorted(df["Tipo de Licença"].dropna().unique().tolist())
        tipo_sel = st.selectbox("Tipo de licença", tipos)

        abas_filtro = ["Todas"] + sorted(df["Aba_Origem"].dropna().unique().tolist())
        aba_sel = st.selectbox("Aba", abas_filtro)

        st.divider()

        st.selectbox(
            "Faixa do alerta vermelho",
            options=OPCOES_ALERTA,
            index=OPCOES_ALERTA.index(st.session_state.alerta_config),
            key="alerta_config"
        )

        st.divider()
        st.markdown("### Legenda interativa")

        if st.button("🔴 Vencida", use_container_width=True):
            st.session_state.filtro_alerta = "Vencida"

        if st.button(f"🟥 {alerta_label}", use_container_width=True):
            st.session_state.filtro_alerta = alerta_label

        if st.button("🟡 Em andamento", use_container_width=True):
            st.session_state.filtro_alerta = "Em andamento"

        if st.button("🟢 Renovada", use_container_width=True):
            st.session_state.filtro_alerta = "Renovada"

        if st.button("⚪ Sem vencimento", use_container_width=True):
            st.session_state.filtro_alerta = "Sem vencimento"

        if st.button("Limpar filtro de alerta", use_container_width=True):
            st.session_state.filtro_alerta = "Todos"

    st.session_state.df_licencas = atualizar_alertas(
        st.session_state.df_licencas.copy(),
        st.session_state.alerta_config
    )
    df = st.session_state.df_licencas.copy()
    alerta_label = texto_alerta_dinamico(st.session_state.alerta_config)

    df_exibicao = aplicar_filtros(df, empresa_sel, centro_sel, tipo_sel, aba_sel)

    filtro_alerta = st.session_state.filtro_alerta
    if filtro_alerta != "Todos":
        df_exibicao = df_exibicao[df_exibicao["Alerta"] == filtro_alerta].copy()

    if filtro_alerta in ["Vencida", alerta_label]:
        proxima = obter_proxima_data_critica(df_exibicao, alerta_label)
        if proxima:
            st.session_state.data_selecionada = proxima
            st.session_state.mes_sel = proxima.month
            st.session_state.ano_sel = proxima.year

    hoje = date.today()

    if "mes_sel" not in st.session_state:
        st.session_state.mes_sel = hoje.month

    if "ano_sel" not in st.session_state:
        st.session_state.ano_sel = hoje.year

    nav1, nav2, nav3 = st.columns([1, 2, 1])

    if nav1.button("◀ Mês anterior"):
        if st.session_state.mes_sel == 1:
            st.session_state.mes_sel = 12
            st.session_state.ano_sel -= 1
        else:
            st.session_state.mes_sel -= 1

    nav2.markdown(
        f"""
        <div style="
            text-align:center;
            font-size:22px;
            font-weight:700;
            padding-top:6px;
            color:#0f172a;">
            {calendar.month_name[st.session_state.mes_sel]} / {st.session_state.ano_sel}
        </div>
        """,
        unsafe_allow_html=True
    )

    if nav3.button("Próximo mês ▶"):
        if st.session_state.mes_sel == 12:
            st.session_state.mes_sel = 1
            st.session_state.ano_sel += 1
        else:
            st.session_state.mes_sel += 1

    mes_sel = st.session_state.mes_sel
    ano_sel = st.session_state.ano_sel

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("🔴 Vencidas", int((df_exibicao["Alerta"] == "Vencida").sum()))
    k2.metric(f"🟥 {st.session_state.alerta_config}", int((df_exibicao["Alerta"] == alerta_label).sum()))
    k3.metric("🟡 Em andamento", int((df_exibicao["Status"] == "Em andamento").sum()))
    k4.metric("🟢 Renovadas", int((df_exibicao["Status"] == "Renovada").sum()))
    k5.metric("⚪ Sem vencimento", int((df_exibicao["Alerta"] == "Sem vencimento").sum()))

    if filtro_alerta != "Todos":
        st.info(f"Filtro ativo: {filtro_alerta}")

    st.divider()
    st.subheader(f"Calendário — {calendar.month_name[mes_sel]} / {ano_sel}")

    nomes_dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
    cab = st.columns(7)
    for i, nome in enumerate(nomes_dias):
        cab[i].markdown(f"**{nome}**")

    cal = calendar.Calendar(firstweekday=0)
    semanas = cal.monthdayscalendar(ano_sel, mes_sel)

    for semana in semanas:
        cols = st.columns(7)
        for i, dia in enumerate(semana):
            if dia == 0:
                cols[i].markdown(" ")
                continue

            data_atual = pd.Timestamp(date(ano_sel, mes_sel, dia))
            df_dia = df_exibicao[df_exibicao["Vencimento"].dt.date == data_atual.date()].copy()
            cor = cor_dia(df_dia, alerta_label)
            borda = borda_dia(df_dia, alerta_label)
            resumo = resumo_dia(df_dia)

            cols[i].markdown(
                f"""
                <div style="
                    background:{cor};
                    border:2px solid {borda};
                    border-radius:16px;
                    padding:10px;
                    min-height:110px;
                    box-shadow:0 1px 5px rgba(15,23,42,0.05);
                    margin-bottom:6px;">
                    <div style="font-size:18px;font-weight:700;color:#0f172a;">{dia}</div>
                    <div style="font-size:12px;color:#475569;margin-top:8px;">{resumo}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

            if cols[i].button("Abrir", key=f"dia_{ano_sel}_{mes_sel}_{dia}"):
                st.session_state.data_selecionada = data_atual.date()

    st.divider()

    esquerda, direita = st.columns([1.7, 1])

    with esquerda:
        st.subheader("Painel detalhado")

        data_selecionada = st.session_state.get("data_selecionada", None)

        if filtro_alerta == "Vencida":
            st.markdown("**Tratativa sugerida:** atuar imediatamente nas vencidas.")
        elif filtro_alerta == alerta_label:
            st.markdown(f"**Tratativa sugerida:** priorizar licenças que vencem em até {st.session_state.alerta_config}.")
        elif filtro_alerta == "Em andamento":
            st.markdown("**Tratativa sugerida:** apenas acompanhamento.")
        elif filtro_alerta == "Renovada":
            st.markdown("**Tratativa sugerida:** somente consulta.")
        elif filtro_alerta == "Sem vencimento":
            st.markdown("**Tratativa sugerida:** somente completar dados quando necessário.")

        if filtro_alerta in ["Vencida", alerta_label, "Em andamento", "Renovada", "Sem vencimento"]:
            df_lista = df_exibicao.copy()

            if not df_lista.empty:
                colunas_visual = [
                    "Empresa",
                    "Colaborador",
                    "Tipo de Licença",
                    "Centro de Custo",
                    "Valor da Licença",
                    "Vencimento",
                    "Dias para Vencer",
                    "Status",
                    "Alerta",
                    "Aba_Origem",
                ]

                df_visual_alerta = df_lista[colunas_visual].copy()
                df_visual_alerta["Valor da Licença"] = df_visual_alerta["Valor da Licença"].apply(formatar_brl)
                df_visual_alerta["Vencimento"] = df_visual_alerta["Vencimento"].dt.strftime("%d/%m/%Y")
                df_visual_alerta["Vencimento"] = df_visual_alerta["Vencimento"].fillna("Não informado")
                df_visual_alerta["Dias para Vencer"] = df_visual_alerta["Dias para Vencer"].fillna("Não informado")

                st.dataframe(df_visual_alerta, use_container_width=True)

        if data_selecionada:
            st.markdown(f"**Data selecionada:** {data_selecionada.strftime('%d/%m/%Y')}")
            df_dia = df_exibicao[df_exibicao["Vencimento"].dt.date == data_selecionada].copy()

            if df_dia.empty:
                st.info("Nenhuma licença para esta data com os filtros atuais.")
            else:
                colunas_visual = [
                    "Empresa",
                    "Colaborador",
                    "Tipo de Licença",
                    "Centro de Custo",
                    "Valor da Licença",
                    "Vencimento",
                    "Dias para Vencer",
                    "Status",
                    "Alerta",
                    "Aba_Origem",
                ]

                df_visual = df_dia[colunas_visual].copy()
                df_visual["Valor da Licença"] = df_visual["Valor da Licença"].apply(formatar_brl)
                df_visual["Vencimento"] = df_visual["Vencimento"].dt.strftime("%d/%m/%Y")
                df_visual["Dias para Vencer"] = df_visual["Dias para Vencer"].fillna("Não informado")

                st.dataframe(df_visual, use_container_width=True)
        else:
            st.info("Clique em um dia do calendário para ver os detalhes.")

    with direita:
        st.subheader("✏️ Editor")

        acao_permitida = filtro_alerta in ["Todos", "Vencida", alerta_label]

        if not acao_permitida:
            st.info("Neste filtro o painel fica em modo consulta. Sem ação obrigatória.")
        elif df_exibicao.empty:
            st.info("Nenhum item disponível para edição com os filtros atuais.")
        else:
            opcoes = [
                f"{idx} | {row['Empresa']} | {row['Colaborador']} | {row['Tipo de Licença']} | {row['Aba_Origem']}"
                for idx, row in df_exibicao.iterrows()
            ]

            item_sel = st.selectbox("Selecione a licença", opcoes)
            idx_sel = int(item_sel.split(" | ")[0])

            linha_atual = st.session_state.df_licencas.loc[idx_sel]

            st.markdown('<div class="card-info">', unsafe_allow_html=True)
            st.write(f"**Empresa:** {linha_atual['Empresa']}")
            st.write(f"**Colaborador:** {linha_atual['Colaborador']}")
            st.write(f"**Tipo de Licença:** {linha_atual['Tipo de Licença']}")
            st.write(f"**Centro de Custo:** {linha_atual['Centro de Custo']}")
            st.write(f"**Valor:** {formatar_brl(linha_atual['Valor da Licença'])}")
            st.write(f"**Aba de origem:** {linha_atual['Aba_Origem']}")

            vencimento_atual = linha_atual["Vencimento"]
            if pd.isna(vencimento_atual):
                st.write("**Vencimento atual:** Não informado")
                data_padrao = hoje
                dias_exib = "Não informado"
            else:
                st.write(f"**Vencimento atual:** {vencimento_atual.strftime('%d/%m/%Y')}")
                data_padrao = vencimento_atual.date()
                dias_atual = linha_atual.get("Dias para Vencer", None)
                dias_exib = "Não informado" if pd.isna(dias_atual) else int(dias_atual)

            st.write(f"**Dias para vencer:** {dias_exib}")
            st.write(f"**Status atual:** {linha_atual['Status']}")
            st.markdown('</div>', unsafe_allow_html=True)

            novo_status = st.selectbox(
                "Novo status",
                STATUS_VALIDOS,
                index=STATUS_VALIDOS.index(linha_atual["Status"]) if linha_atual["Status"] in STATUS_VALIDOS else 0
            )

            nova_data = st.date_input("Data de vencimento", value=data_padrao)

            st.markdown("**Renovação rápida**")
            b1, b2, b3, b4 = st.columns(4)
            b5, b6, b7, b8 = st.columns(4)

            def renovar(meses):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, meses)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(
                    st.session_state.df_licencas,
                    st.session_state.alerta_config
                )
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b1.button("+1 mês"):
                renovar(1)
            if b2.button("+3 meses"):
                renovar(3)
            if b3.button("+6 meses"):
                renovar(6)
            if b4.button("+12 meses"):
                renovar(12)
            if b5.button("+24 meses"):
                renovar(24)
            if b6.button("+36 meses"):
                renovar(36)
            if b7.button("+48 meses"):
                renovar(48)
            if b8.button("+60 meses"):
                renovar(60)

            st.divider()

            s1, s2 = st.columns(2)

            if s1.button("Salvar edição manual"):
                st.session_state.df_licencas.loc[idx_sel, "Status"] = novo_status
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = pd.Timestamp(nova_data)
                st.session_state.df_licencas = atualizar_alertas(
                    st.session_state.df_licencas,
                    st.session_state.alerta_config
                )
                st.success("Licença atualizada com sucesso.")

            if s2.button("Informar vencimento"):
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = pd.Timestamp(nova_data)
                st.session_state.df_licencas = atualizar_alertas(
                    st.session_state.df_licencas,
                    st.session_state.alerta_config
                )
                st.success("Vencimento informado com sucesso.")

        st.divider()
        st.subheader("📥 Exportação")
        st.download_button(
            label="Baixar planilha atualizada",
            data=gerar_excel(st.session_state.df_licencas),
            file_name="licencas_atualizadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("Todo mês suba a planilha revisada. O arquivo baixado aqui pode virar a base do próximo ciclo.")
else:
    st.info("Suba a planilha revisada do mês para abrir o painel.")
