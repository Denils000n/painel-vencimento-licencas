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
        "empresa", "company", "organization", "organização", "tenant", "companhia"
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
    except:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def ler_arquivo(arquivo):
    if arquivo.name.endswith(".csv"):
        try:
            return pd.read_csv(arquivo)
        except:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=";")
    return pd.read_excel(arquivo)


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
        if col not in df.columns:
            faltantes.append(col)
    return faltantes


def atualizar_alertas(df):
    hoje = pd.Timestamp(date.today())
    df["Dias para Vencer"] = (df["Vencimento"] - hoje).dt.days

    def definir_alerta(row):
        if pd.isna(row["Vencimento"]):
            return "Sem vencimento"
        if row["Status"] == "Renovada":
            return "Renovada"
        if row["Status"] == "Em andamento":
            return "Em andamento"
        if row["Dias para Vencer"] < 0:
            return "Vencida"
        if row["Dias para Vencer"] <= 30:
            return "Vence em até 30 dias"
        return "No prazo"

    df["Alerta"] = df.apply(definir_alerta, axis=1)
    return df


def gerar_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Licencas")
    return buffer.getvalue()


def cor_dia(df_dia):
    if df_dia.empty:
        return "#f8fafc"
    if (df_dia["Alerta"] == "Vencida").any():
        return "#fee2e2"
    if (df_dia["Alerta"] == "Vence em até 30 dias").any():
        return "#fecaca"
    if (df_dia["Status"] == "Em andamento").any():
        return "#fef3c7"
    if not df_dia.empty and (df_dia["Status"] == "Renovada").all():
        return "#dcfce7"
    return "#e2e8f0"


def borda_dia(df_dia):
    if df_dia.empty:
        return "#e5e7eb"
    if (df_dia["Alerta"] == "Vencida").any():
        return "#dc2626"
    if (df_dia["Alerta"] == "Vence em até 30 dias").any():
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


def aplicar_filtros(df, empresa, centro, tipo):
    df_filtrado = df.copy()

    if empresa != "Todas":
        df_filtrado = df_filtrado[df_filtrado["Empresa"] == empresa]

    if centro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Centro de Custo"] == centro]

    if tipo != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Tipo de Licença"] == tipo]

    return df_filtrado


def adicionar_meses(data_base, meses):
    return pd.Timestamp(data_base) + pd.DateOffset(months=int(meses))


def obter_proxima_data_critica(df):
    criticos = df[df["Alerta"].isin(["Vencida", "Vence em até 30 dias"])].copy()
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
.legenda {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 10px 12px;
    margin-bottom: 8px;
    font-size: 14px;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📅 Painel de Vencimento de Licenças</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Mapeamento automático de colunas, calendário visual e tratamento por criticidade.</div>',
    unsafe_allow_html=True
)

if "filtro_alerta" not in st.session_state:
    st.session_state.filtro_alerta = "Todos"

arquivo = st.file_uploader(
    "📁 Suba a planilha revisada do mês (CSV ou Excel)",
    type=["csv", "xlsx"]
)

if arquivo:
    if (
        "arquivo_base" not in st.session_state
        or st.session_state.arquivo_base != arquivo.name
    ):
        df_original = ler_arquivo(arquivo)
        df_original.columns = [str(c).strip() for c in df_original.columns]
        st.session_state.df_original = df_original.copy()
        st.session_state.arquivo_base = arquivo.name
        st.session_state.mapeamento_auto = detectar_mapeamento_automatico(df_original.columns)
        st.session_state.data_selecionada = None
        st.session_state.filtro_alerta = "Todos"

    df_original = st.session_state.df_original.copy()
    mapeamento_auto = st.session_state.mapeamento_auto.copy()

    st.subheader("Mapeamento de colunas")
    st.caption("O sistema tentou identificar as colunas automaticamente. Se precisar, ajuste abaixo.")

    opcoes_colunas = ["-- Não mapear --"] + list(df_original.columns)
    col_map_1, col_map_2 = st.columns(2)
    mapeamento_final = {}

    campos_esquerda = ["Empresa", "Colaborador", "Centro de Custo", "Tipo de Licença"]
    campos_direita = ["Valor da Licença", "Vencimento", "Status"]

    with col_map_1:
        for campo in campos_esquerda:
            valor_padrao = mapeamento_auto.get(campo)
            idx = opcoes_colunas.index(valor_padrao) if valor_padrao in opcoes_colunas else 0
            escolha = st.selectbox(f"{campo}", opcoes_colunas, index=idx, key=f"map_{campo}")
            mapeamento_final[campo] = None if escolha == "-- Não mapear --" else escolha

    with col_map_2:
        for campo in campos_direita:
            valor_padrao = mapeamento_auto.get(campo)
            idx = opcoes_colunas.index(valor_padrao) if valor_padrao in opcoes_colunas else 0
            escolha = st.selectbox(f"{campo}", opcoes_colunas, index=idx, key=f"map_{campo}")
            mapeamento_final[campo] = None if escolha == "-- Não mapear --" else escolha

    if st.button("Aplicar mapeamento"):
        df_mapeado = aplicar_mapeamento(df_original, mapeamento_final)
        faltantes = validar_minimos(df_mapeado)
        if faltantes:
            st.error(f"Não foi possível continuar. Faltam os campos mínimos: {', '.join(faltantes)}")
            st.stop()

        df_mapeado = atualizar_alertas(df_mapeado)
        st.session_state.df_licencas = df_mapeado.copy()
        st.success("Mapeamento aplicado com sucesso.")

    if "df_licencas" in st.session_state:
        df = st.session_state.df_licencas.copy()

        with st.sidebar:
            st.header("Filtros")

            empresas = ["Todas"] + sorted(df["Empresa"].dropna().unique().tolist())
            empresa_sel = st.selectbox("Empresa", empresas)

            centros = ["Todos"] + sorted(df["Centro de Custo"].dropna().unique().tolist())
            centro_sel = st.selectbox("Centro de custo", centros)

            tipos = ["Todos"] + sorted(df["Tipo de Licença"].dropna().unique().tolist())
            tipo_sel = st.selectbox("Tipo de licença", tipos)

            st.divider()
            st.markdown("### Legenda interativa")

            if st.button("🔴 Vencida", use_container_width=True):
                st.session_state.filtro_alerta = "Vencida"

            if st.button("🟥 Vence em até 30 dias", use_container_width=True):
                st.session_state.filtro_alerta = "Vence em até 30 dias"

            if st.button("🟡 Em andamento", use_container_width=True):
                st.session_state.filtro_alerta = "Em andamento"

            if st.button("🟢 Renovada", use_container_width=True):
                st.session_state.filtro_alerta = "Renovada"

            if st.button("⚪ Sem vencimento", use_container_width=True):
                st.session_state.filtro_alerta = "Sem vencimento"

            if st.button("Limpar filtro de alerta", use_container_width=True):
                st.session_state.filtro_alerta = "Todos"

        df_exibicao = aplicar_filtros(df, empresa_sel, centro_sel, tipo_sel)

        filtro_alerta = st.session_state.filtro_alerta
        if filtro_alerta != "Todos":
            df_exibicao = df_exibicao[df_exibicao["Alerta"] == filtro_alerta].copy()

        if filtro_alerta in ["Vencida", "Vence em até 30 dias"]:
            proxima = obter_proxima_data_critica(df_exibicao)
            if proxima:
                st.session_state.data_selecionada = proxima

        hoje = date.today()
        datas_validas = df_exibicao["Vencimento"].dropna()

        if len(datas_validas) > 0:
            anos = sorted(datas_validas.dt.year.unique().tolist())
        else:
            anos = [hoje.year]

        if hoje.year not in anos:
            anos.append(hoje.year)
            anos = sorted(anos)

        c1, c2 = st.columns(2)
        mes_sel = c1.selectbox(
            "Mês",
            list(range(1, 13)),
            index=hoje.month - 1,
            format_func=lambda x: calendar.month_name[x]
        )
        ano_sel = c2.selectbox(
            "Ano",
            anos,
            index=anos.index(hoje.year) if hoje.year in anos else 0
        )

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("🔴 Vencidas", int((df_exibicao["Alerta"] == "Vencida").sum()))
        k2.metric("🟥 Até 30 dias", int((df_exibicao["Alerta"] == "Vence em até 30 dias").sum()))
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
                cor = cor_dia(df_dia)
                borda = borda_dia(df_dia)
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
            elif filtro_alerta == "Vence em até 30 dias":
                st.markdown("**Tratativa sugerida:** priorizar renovação e planejamento.")
            elif filtro_alerta == "Em andamento":
                st.markdown("**Tratativa sugerida:** apenas acompanhamento.")
            elif filtro_alerta == "Renovada":
                st.markdown("**Tratativa sugerida:** somente consulta.")
            elif filtro_alerta == "Sem vencimento":
                st.markdown("**Tratativa sugerida:** somente completar dados quando necessário.")

            if filtro_alerta in ["Vencida", "Vence em até 30 dias", "Em andamento", "Renovada", "Sem vencimento"]:
                df_lista = df_exibicao.copy()

                if not df_lista.empty:
                    df_visual_alerta = df_lista[
                        [
                            "Empresa",
                            "Colaborador",
                            "Tipo de Licença",
                            "Centro de Custo",
                            "Valor da Licença",
                            "Vencimento",
                            "Dias para Vencer",
                            "Status",
                            "Alerta",
                        ]
                    ].copy()

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
                    df_visual = df_dia[
                        [
                            "Empresa",
                            "Colaborador",
                            "Tipo de Licença",
                            "Centro de Custo",
                            "Valor da Licença",
                            "Vencimento",
                            "Dias para Vencer",
                            "Status",
                            "Alerta",
                        ]
                    ].copy()

                    df_visual["Valor da Licença"] = df_visual["Valor da Licença"].apply(formatar_brl)
                    df_visual["Vencimento"] = df_visual["Vencimento"].dt.strftime("%d/%m/%Y")
                    df_visual["Dias para Vencer"] = df_visual["Dias para Vencer"].fillna("Não informado")

                    st.dataframe(df_visual, use_container_width=True)
            else:
                st.info("Clique em um dia do calendário para ver os detalhes.")

        with direita:
            st.subheader("✏️ Editor")

            acao_permitida = filtro_alerta in ["Todos", "Vencida", "Vence em até 30 dias"]

            if not acao_permitida:
                st.info("Neste filtro o painel fica em modo consulta. Sem ação obrigatória.")
            elif df_exibicao.empty:
                st.info("Nenhum item disponível para edição com os filtros atuais.")
            else:
                opcoes = [
                    f"{idx} | {row['Empresa']} | {row['Colaborador']} | {row['Tipo de Licença']}"
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
                    st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
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
                    st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                    st.success("Licença atualizada com sucesso.")

                if s2.button("Informar vencimento"):
                    st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = pd.Timestamp(nova_data)
                    st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
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
    st.info("Suba a planilha revisada do mês para iniciar o mapeamento automático.")
