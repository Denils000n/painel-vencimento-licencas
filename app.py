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

COLUNAS_OBRIGATORIAS = [
    "Empresa",
    "Colaborador",
    "Tipo de Licença",
    "Centro de Custo",
    "Valor da Licença",
    "Status",
]


def formatar_brl(valor):
    try:
        valor = float(valor)
    except:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def ler_arquivo(arquivo):
    if arquivo.name.endswith(".csv"):
        return pd.read_csv(arquivo)
    return pd.read_excel(arquivo)


def normalizar_dataframe(df):
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltantes:
        st.error(f"Faltam estas colunas obrigatórias: {', '.join(faltantes)}")
        st.stop()

    if "Vencimento" not in df.columns:
        df["Vencimento"] = pd.NaT

    for col in ["Empresa", "Colaborador", "Tipo de Licença", "Centro de Custo", "Status"]:
        df[col] = df[col].astype(str).str.strip()

    df["Valor da Licença"] = (
        df["Valor da Licença"]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )
    df["Valor da Licença"] = pd.to_numeric(df["Valor da Licença"], errors="coerce").fillna(0)

    df["Vencimento"] = pd.to_datetime(df["Vencimento"], errors="coerce")
    df.loc[~df["Status"].isin(STATUS_VALIDOS), "Status"] = "Pendente"

    return df


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
    export_df = df.copy()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Licencas")
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
    if (df_dia["Status"] == "Renovada").all():
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
    if (df_dia["Status"] == "Renovada").all():
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
    '<div class="sub-title">Calendário visual, alertas de vencimento, atualização mensal por planilha revisada e renovação rápida.</div>',
    unsafe_allow_html=True
)

arquivo = st.file_uploader(
    "📁 Suba a planilha revisada do mês (CSV ou Excel)",
    type=["csv", "xlsx"]
)

if arquivo:
    nome_arquivo_atual = arquivo.name

    if (
        "df_licencas" not in st.session_state
        or "arquivo_base" not in st.session_state
        or st.session_state.arquivo_base != nome_arquivo_atual
    ):
        df_inicial = ler_arquivo(arquivo)
        df_inicial = normalizar_dataframe(df_inicial)
        df_inicial = atualizar_alertas(df_inicial)

        st.session_state.df_licencas = df_inicial.copy()
        st.session_state.arquivo_base = nome_arquivo_atual
        st.session_state.data_selecionada = None

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
        st.markdown("### Legenda")
        st.markdown('<div class="legenda">🔴 Vencida</div>', unsafe_allow_html=True)
        st.markdown('<div class="legenda">🟥 Vence em até 30 dias</div>', unsafe_allow_html=True)
        st.markdown('<div class="legenda">🟡 Em andamento</div>', unsafe_allow_html=True)
        st.markdown('<div class="legenda">🟢 Renovada</div>', unsafe_allow_html=True)
        st.markdown('<div class="legenda">⚪ Sem vencimento</div>', unsafe_allow_html=True)

    df_exibicao = aplicar_filtros(df, empresa_sel, centro_sel, tipo_sel)

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

        st.divider()
        st.subheader("Licenças sem vencimento informado")

        df_sem_venc = df_exibicao[df_exibicao["Vencimento"].isna()].copy()

        if df_sem_venc.empty:
            st.success("Não há licenças sem vencimento informado.")
        else:
            df_sem_venc_visual = df_sem_venc[
                [
                    "Empresa",
                    "Colaborador",
                    "Tipo de Licença",
                    "Centro de Custo",
                    "Valor da Licença",
                    "Status",
                    "Alerta",
                ]
            ].copy()
            df_sem_venc_visual["Valor da Licença"] = df_sem_venc_visual["Valor da Licença"].apply(formatar_brl)
            st.dataframe(df_sem_venc_visual, use_container_width=True)

    with direita:
        st.subheader("✏️ Editor")

        if df_exibicao.empty:
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

            if b1.button("+1 mês"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 1)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b2.button("+3 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 3)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b3.button("+6 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 6)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b4.button("+12 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 12)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b5.button("+24 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 24)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b6.button("+36 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 36)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b7.button("+48 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 48)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

            if b8.button("+60 meses"):
                base = linha_atual["Vencimento"] if not pd.isna(linha_atual["Vencimento"]) else pd.Timestamp(nova_data)
                nova = adicionar_meses(base, 60)
                st.session_state.df_licencas.loc[idx_sel, "Status"] = "Renovada"
                st.session_state.df_licencas.loc[idx_sel, "Vencimento"] = nova
                st.session_state.df_licencas = atualizar_alertas(st.session_state.df_licencas)
                st.success(f"Renovada até {nova.strftime('%d/%m/%Y')}")

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
    st.info("Suba a planilha revisada do mês para abrir o calendário e o painel.")