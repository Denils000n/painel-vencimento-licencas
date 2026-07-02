import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import calendar
import unicodedata
from datetime import date, datetime

# ============================================================
# CONFIGURACOES & CONSTANTES
# ============================================================

st.set_page_config(
    page_title="Gerenciador de Licencas",
    page_icon="=",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_PATH = os.environ.get("LICENCAS_DB_PATH", os.path.join(os.path.dirname(os.path.abspath(__file__)), "licencas.db"))

MESES_PT = ["", "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

EMPRESA_PREFIXOS = {
    "01": "Afonso Franca",
    "02": "AFFIT",
    "03": "AFDI",
    "04": "AFSW",
}

COLUNAS_SISTEMA = [
    "Empresa", "Colaborador", "Centro de Custo",
    "Tipo de Licenca", "Valor da Licenca", "Vencimento", "Status"
]

SINONIMOS = {
    "Empresa":         ["empresa", "company", "companhia", "entidade", "filial", "unidade"],
    "Colaborador":     ["colaborador", "funcionario", "employee", "nome", "name",
                        "usuario", "user", "login", "upn", "email", "username"],
    "Centro de Custo": ["centro de custo", "cc", "cost center", "departamento", "setor", "area", "depto"],
    "Tipo de Licenca": ["tipo de licenca", "tipo licenca", "tipo", "licenca",
                        "license", "produto", "product", "software", "sku", "aplicacao"],
    "Valor da Licenca":["valor da licenca", "valor licenca", "valor", "value",
                        "preco", "price", "custo", "cost"],
    "Vencimento":      ["vencimento", "expiration", "expiry", "data vencimento",
                        "data de vencimento", "validade", "valid until", "expires"],
    "Status":          ["status", "situacao", "state", "estado"],
}

STATUS_VALIDOS = ["Pendente", "Em andamento", "Renovada", "Cancelada"]

ALERTA_OPCOES = {
    "30 dias": 30, "60 dias": 60, "90 dias": 90,
    "6 meses": 180, "12 meses": 365, "18 meses": 540,
    "24 meses": 730, "36 meses": 1095, "48 meses": 1460, "60 meses": 1825,
}

COR_ALERTA = {
    "Vencida":  "#FF5252",
    "Critica":  "#FF9800",
    "Atencao":  "#FFC107",
    "Ok":       "#4CAF50",
    "Sem data": "#9E9E9E",
}

BADGE_ALERTA = {
    "Vencida":  "background:#FF5252;color:#fff",
    "Critica":  "background:#FF9800;color:#fff",
    "Atencao":  "background:#FFC107;color:#000",
    "Ok":       "background:#4CAF50;color:#fff",
    "Sem data": "background:#9E9E9E;color:#fff",
}


# ============================================================
# BANCO DE DADOS
# ============================================================

def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS licencas (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa          TEXT    DEFAULT 'Nao informada',
            colaborador      TEXT    NOT NULL,
            centro_custo     TEXT,
            tipo_licenca     TEXT    NOT NULL,
            valor_licenca    REAL,
            vencimento       TEXT,
            status           TEXT    DEFAULT 'Pendente',
            alerta           TEXT    DEFAULT 'Sem data',
            dias_para_vencer INTEGER,
            criado_em        TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            atualizado_em    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE UNIQUE INDEX IF NOT EXISTS idx_unico
            ON licencas(colaborador, tipo_licenca, COALESCE(empresa,''));

        CREATE TABLE IF NOT EXISTS importacoes (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            arquivo               TEXT,
            aba                   TEXT,
            data_importacao       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            registros_novos       INTEGER DEFAULT 0,
            registros_atualizados INTEGER DEFAULT 0,
            registros_total       INTEGER DEFAULT 0
        );
    """)
    conn.commit()
    conn.close()


def carregar_licencas(filtros=None):
    conn = get_conn()
    query = "SELECT * FROM licencas WHERE 1=1"
    params = []
    if filtros:
        if filtros.get("empresa"):
            query += " AND empresa = ?"; params.append(filtros["empresa"])
        if filtros.get("centro_custo"):
            query += " AND centro_custo = ?"; params.append(filtros["centro_custo"])
        if filtros.get("tipo_licenca"):
            query += " AND tipo_licenca = ?"; params.append(filtros["tipo_licenca"])
        if filtros.get("alerta"):
            query += " AND alerta = ?"; params.append(filtros["alerta"])
        if filtros.get("status"):
            query += " AND status = ?"; params.append(filtros["status"])
    query += " ORDER BY vencimento ASC NULLS LAST"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def atualizar_registro(id_, campos):
    conn = get_conn()
    set_parts = [f"{k} = ?" for k in campos] + ["atualizado_em = CURRENT_TIMESTAMP"]
    valores = list(campos.values()) + [id_]
    conn.execute(f"UPDATE licencas SET {', '.join(set_parts)} WHERE id = ?", valores)
    conn.commit()
    conn.close()


def deletar_registro(id_):
    conn = get_conn()
    conn.execute("DELETE FROM licencas WHERE id = ?", (id_,))
    conn.commit()
    conn.close()


def upsert_licencas(df):
    """Insere ou atualiza. Chave: colaborador + tipo_licenca + empresa."""
    conn = get_conn()
    novos = atualizados = 0

    for _, row in df.iterrows():
        colab   = str(row.get("Colaborador", "") or "").strip()
        tipo    = str(row.get("Tipo de Licenca", "") or "").strip()
        empresa = str(row.get("Empresa", "") or "Nao informada").strip()

        if not colab or not tipo:
            continue

        vencimento = None
        v = row.get("Vencimento")
        if v and pd.notna(v):
            try:
                vencimento = pd.to_datetime(v).strftime("%Y-%m-%d")
            except Exception:
                pass

        valor = None
        val_raw = row.get("Valor da Licenca")
        if val_raw and pd.notna(val_raw):
            try:
                s = str(val_raw).replace("R$", "").replace(".", "").replace(",", ".").strip()
                valor = float(s)
            except Exception:
                pass

        status = row.get("Status", "Pendente")
        if status not in STATUS_VALIDOS:
            status = "Pendente"

        centro = str(row.get("Centro de Custo", "") or "").strip()

        cur = conn.execute(
            "SELECT id FROM licencas WHERE colaborador=? AND tipo_licenca=? AND COALESCE(empresa,'')=?",
            (colab, tipo, empresa)
        )
        existing = cur.fetchone()

        if existing:
            conn.execute(
                "UPDATE licencas SET empresa=?, centro_custo=?, valor_licenca=?, "
                "vencimento=?, status=?, atualizado_em=CURRENT_TIMESTAMP WHERE id=?",
                (empresa, centro, valor, vencimento, status, existing[0])
            )
            atualizados += 1
        else:
            conn.execute(
                "INSERT INTO licencas (empresa, colaborador, centro_custo, tipo_licenca, valor_licenca, vencimento, status) "
                "VALUES (?,?,?,?,?,?,?)",
                (empresa, colab, centro, tipo, valor, vencimento, status)
            )
            novos += 1

    conn.commit()
    conn.close()
    return novos, atualizados


def log_importacao(arquivo, aba, novos, atualizados, total):
    conn = get_conn()
    conn.execute(
        "INSERT INTO importacoes (arquivo, aba, registros_novos, registros_atualizados, registros_total) VALUES (?,?,?,?,?)",
        (arquivo, aba, novos, atualizados, total)
    )
    conn.commit()
    conn.close()


def get_historico():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT arquivo, aba, data_importacao, registros_novos, registros_atualizados, registros_total "
        "FROM importacoes ORDER BY data_importacao DESC LIMIT 50",
        conn
    )
    conn.close()
    return df


def recalcular_alertas(dias_alerta):
    """Recalcula alerta e dias_para_vencer para todos os registros."""
    conn = get_conn()
    rows = conn.execute("SELECT id, vencimento FROM licencas").fetchall()
    hoje = date.today()
    limite_atencao = int(dias_alerta)
    limite_critica = int(dias_alerta * 0.33)

    updates = []
    for id_, venc_str in rows:
        if not venc_str:
            updates.append(("Sem data", None, id_))
            continue
        try:
            venc = datetime.strptime(venc_str, "%Y-%m-%d").date()
            dias = (venc - hoje).days
            if dias < 0:
                alerta = "Vencida"
            elif dias <= limite_critica:
                alerta = "Critica"
            elif dias <= limite_atencao:
                alerta = "Atencao"
            else:
                alerta = "Ok"
            updates.append((alerta, dias, id_))
        except Exception:
            updates.append(("Sem data", None, id_))

    conn.executemany(
        "UPDATE licencas SET alerta=?, dias_para_vencer=? WHERE id=?", updates
    )
    conn.commit()
    conn.close()


# ============================================================
# UTILITARIOS DE IMPORTACAO
# ============================================================

def normalizar(txt):
    txt = str(txt).lower().strip()
    return unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode()


def identificar_empresa_pelo_cc(valor):
    if not valor or (isinstance(valor, float)):
        return "Nao informada"
    prefixo = str(valor).strip()[:2]
    return EMPRESA_PREFIXOS.get(prefixo, f"CC-{prefixo}")


def listar_abas(arquivo):
    try:
        arquivo.seek(0)
        xls = pd.ExcelFile(arquivo)
        return xls.sheet_names
    except Exception:
        return []


def ler_arquivo(arquivo, aba=None):
    nome = arquivo.name.lower()
    if nome.endswith(".csv"):
        for sep, enc in [(",", "utf-8"), (";", "utf-8"), (";", "latin1"), (",", "latin1")]:
            try:
                arquivo.seek(0)
                return pd.read_csv(arquivo, sep=sep, encoding=enc)
            except Exception:
                continue
        raise ValueError("Nao foi possivel ler o CSV.")
    else:
        arquivo.seek(0)
        return pd.read_excel(arquivo, sheet_name=aba)


def detectar_mapeamento(df):
    mapa = {}
    colunas_norm = [normalizar(c) for c in df.columns]
    colunas_orig = list(df.columns)
    for col_sis, sinonimos in SINONIMOS.items():
        for i, col_n in enumerate(colunas_norm):
            if any(s in col_n or col_n in s for s in sinonimos):
                if col_sis not in mapa:
                    mapa[col_sis] = colunas_orig[i]
                break
    return mapa


def aplicar_mapeamento(df, mapa):
    rename = {v: k for k, v in mapa.items() if v in df.columns}
    df = df.rename(columns=rename)
    for col in COLUNAS_SISTEMA:
        if col not in df.columns:
            df[col] = None

    mask_sem_empresa = df["Empresa"].isna() | (df["Empresa"].astype(str).str.strip() == "")
    if mask_sem_empresa.all():
        df["Empresa"] = df["Centro de Custo"].apply(identificar_empresa_pelo_cc)
    else:
        df.loc[mask_sem_empresa, "Empresa"] = df.loc[mask_sem_empresa, "Centro de Custo"].apply(
            identificar_empresa_pelo_cc
        )

    df["Status"] = df["Status"].apply(
        lambda x: x if x in STATUS_VALIDOS else "Pendente"
    )
    return df[COLUNAS_SISTEMA]


# ============================================================
# UTILITARIOS GERAIS
# ============================================================

def formatar_brl(valor):
    try:
        return "R$ {:,.2f}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "-"


def adicionar_meses(dt, n):
    mes = ((dt.month - 1 + n) % 12) + 1
    ano = dt.year + ((dt.month - 1 + n) // 12)
    import calendar as _cal
    ultimo_dia = _cal.monthrange(ano, mes)[1]
    dia = min(dt.day, ultimo_dia)
    return dt.replace(year=ano, month=mes, day=dia)


def gerar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Licencas")
    return output.getvalue()


# ============================================================
# INICIALIZACAO
# ============================================================

init_db()

defaults = {
    "pagina": "Painel",
    "mes_sel": date.today().month,
    "ano_sel": date.today().year,
    "data_sel": None,
    "dias_alerta": 90,
    "alerta_recalc": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("## Gerenciador de Licencas")
    st.markdown("---")

    pagina = st.radio(
        "Navegacao",
        ["Painel", "Importar", "Licencas", "Exportar"],
        index=["Painel", "Importar", "Licencas", "Exportar"].index(st.session_state.pagina),
        key="radio_pagina"
    )
    st.session_state.pagina = pagina

    st.markdown("---")
    st.markdown("**Configuracao de Alertas**")
    alerta_idx = list(ALERTA_OPCOES.keys()).index("90 dias")
    alerta_sel = st.selectbox("Faixa de alerta", list(ALERTA_OPCOES.keys()), index=alerta_idx, key="alerta_sel")
    dias_alerta = ALERTA_OPCOES[alerta_sel]

    if dias_alerta != st.session_state.dias_alerta:
        st.session_state.dias_alerta = dias_alerta
        recalcular_alertas(dias_alerta)
    else:
        recalcular_alertas(dias_alerta)

    st.markdown("---")
    df_all = carregar_licencas()
    total_lic = len(df_all)
    if total_lic > 0:
        vencidas  = int((df_all["alerta"] == "Vencida").sum())
        criticas  = int((df_all["alerta"] == "Critica").sum())
        atencao   = int((df_all["alerta"] == "Atencao").sum())
        st.metric("Total de licencas", total_lic)
        c1, c2 = st.columns(2)
        c1.metric("Vencidas", vencidas)
        c2.metric("Criticas", criticas)
        st.metric("Atencao", atencao)
    else:
        st.info("Nenhuma licenca cadastrada.")


# ============================================================
# PAGINA: IMPORTAR
# ============================================================

if st.session_state.pagina == "Importar":
    st.title("Importar Planilhas")
    st.markdown(
        "Faca upload de uma ou mais planilhas (.xlsx, .xls, .csv). "
        "O sistema detecta as colunas automaticamente e salva no banco de dados.  \n"
        "Registros existentes (mesmo **Colaborador + Tipo de Licenca + Empresa**) serao **atualizados**."
    )

    arquivos = st.file_uploader(
        "Selecione as planilhas",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="uploader"
    )

    if arquivos:
        for arquivo in arquivos:
            with st.expander(f"Arquivo: {arquivo.name}", expanded=True):
                abas = listar_abas(arquivo)

                # Selecao de aba
                aba_sel = None
                juntar_abas = False
                if abas:
                    if len(abas) > 1:
                        opcao_aba = st.radio(
                            "Qual aba importar?",
                            ["Todas as abas"] + abas,
                            key=f"aba_{arquivo.name}"
                        )
                        juntar_abas = opcao_aba == "Todas as abas"
                        aba_sel = None if juntar_abas else opcao_aba
                    else:
                        aba_sel = abas[0]
                        st.caption(f"Aba: {aba_sel}")

                # Leitura
                try:
                    if juntar_abas:
                        dfs = []
                        for aba in abas:
                            arquivo.seek(0)
                            dfs.append(pd.read_excel(arquivo, sheet_name=aba))
                        df_raw = pd.concat(dfs, ignore_index=True)
                        abas_label = "todas"
                    else:
                        df_raw = ler_arquivo(arquivo, aba_sel)
                        abas_label = aba_sel or "csv"

                    df_raw = df_raw.dropna(how="all")
                    st.caption(f"{len(df_raw)} linhas | Colunas: {', '.join(df_raw.columns.tolist()[:10])}")

                    # Mapeamento automatico
                    mapa_auto = detectar_mapeamento(df_raw)
                    st.markdown("**Mapeamento de colunas** (ajuste se necessario)")

                    cols_disp = ["(nao mapear)"] + df_raw.columns.tolist()
                    mapa_user = {}
                    grid = st.columns(2)
                    for idx, col_sis in enumerate(COLUNAS_SISTEMA):
                        container = grid[idx % 2]
                        default = mapa_auto.get(col_sis, "(nao mapear)")
                        if default not in cols_disp:
                            default = "(nao mapear)"
                        sel = container.selectbox(
                            col_sis,
                            cols_disp,
                            index=cols_disp.index(default),
                            key=f"map_{arquivo.name}_{col_sis}"
                        )
                        if sel != "(nao mapear)":
                            mapa_user[col_sis] = sel

                    if "Colaborador" not in mapa_user or "Tipo de Licenca" not in mapa_user:
                        st.warning("Mapeie pelo menos **Colaborador** e **Tipo de Licenca** para importar.")
                    else:
                        df_mapeado = aplicar_mapeamento(df_raw.copy(), mapa_user)
                        df_mapeado = df_mapeado.dropna(subset=["Colaborador", "Tipo de Licenca"])

                        sem_empresa = (df_mapeado["Empresa"].isin(["Nao informada"])).sum()
                        if sem_empresa > 0:
                            st.warning(f"{sem_empresa} registros sem empresa identificada pelo prefixo do CC.")

                        sem_venc = df_mapeado["Vencimento"].isna().sum()
                        if sem_venc > 0:
                            st.info(f"{sem_venc} registros sem data de vencimento.")

                        with st.expander("Preview (primeiros 5 registros)"):
                            st.dataframe(df_mapeado.head(5), use_container_width=True)

                        if st.button(f"Importar {arquivo.name} ({len(df_mapeado)} registros)", 
                                     key=f"btn_{arquivo.name}", type="primary"):
                            novos, atualizados = upsert_licencas(df_mapeado)
                            log_importacao(arquivo.name, abas_label, novos, atualizados, len(df_mapeado))
                            recalcular_alertas(st.session_state.dias_alerta)
                            st.success(f"Concluido: {novos} novos + {atualizados} atualizados")
                            st.rerun()

                except Exception as e:
                    st.error(f"Erro ao processar {arquivo.name}: {e}")

    st.markdown("---")
    st.subheader("Historico de importacoes")
    hist = get_historico()
    if len(hist) > 0:
        hist.columns = ["Arquivo", "Aba", "Data", "Novos", "Atualizados", "Total"]
        st.dataframe(hist, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma importacao realizada ainda.")

    st.markdown("---")
    st.subheader("Gerenciar banco de dados")
    with st.expander("Opcoes avancadas"):
        st.warning("As acoes abaixo sao irreversiveis.")
        if st.button("Exportar backup completo (Excel)", key="backup_btn"):
            df_bk = carregar_licencas()
            st.download_button(
                "Baixar backup",
                data=gerar_excel(df_bk),
                file_name=f"backup_licencas_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ============================================================
# PAGINA: PAINEL
# ============================================================

elif st.session_state.pagina == "Painel":
    st.title("Painel de Vencimentos")

    df = carregar_licencas()

    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada. Va para **Importar** para adicionar planilhas.")
        st.stop()

    # Filtros
    with st.expander("Filtros", expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        empresas_list  = ["Todas"] + sorted(df["empresa"].dropna().unique().tolist())
        ccs_list       = ["Todos"] + sorted(df["centro_custo"].dropna().unique().tolist())
        tipos_list     = ["Todos"] + sorted(df["tipo_licenca"].dropna().unique().tolist())
        alertas_list   = ["Todos", "Vencida", "Critica", "Atencao", "Ok", "Sem data"]

        emp_f   = fc1.selectbox("Empresa",          empresas_list, key="pan_emp")
        cc_f    = fc2.selectbox("Centro de Custo",  ccs_list,      key="pan_cc")
        tipo_f  = fc3.selectbox("Tipo de Licenca",  tipos_list,    key="pan_tipo")
        alerta_f= fc4.selectbox("Alerta",           alertas_list,  key="pan_alerta")

    df_fil = df.copy()
    if emp_f   != "Todas": df_fil = df_fil[df_fil["empresa"]      == emp_f]
    if cc_f    != "Todos": df_fil = df_fil[df_fil["centro_custo"] == cc_f]
    if tipo_f  != "Todos": df_fil = df_fil[df_fil["tipo_licenca"] == tipo_f]
    if alerta_f!= "Todos": df_fil = df_fil[df_fil["alerta"]       == alerta_f]

    # Metricas
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Registros (filtro)", len(df_fil))
    m2.metric("Vencidas",  int((df_fil["alerta"] == "Vencida").sum()))
    m3.metric("Criticas",  int((df_fil["alerta"] == "Critica").sum()))
    m4.metric("Atencao",   int((df_fil["alerta"] == "Atencao").sum()))
    val_tot = df_fil["valor_licenca"].sum()
    m5.metric("Valor Total", formatar_brl(val_tot) if val_tot > 0 else "-")

    st.markdown("---")

    col_cal, col_det = st.columns([3, 2])

    with col_cal:
        # Navegacao de mes
        nav_a, nav_b, nav_c = st.columns([1, 4, 1])
        if nav_a.button("<<", key="prev_m"):
            if st.session_state.mes_sel == 1:
                st.session_state.mes_sel = 12; st.session_state.ano_sel -= 1
            else:
                st.session_state.mes_sel -= 1
            st.session_state.data_sel = None
            st.rerun()

        nav_b.markdown(
            f"<h3 style='text-align:center;margin:0'>{MESES_PT[st.session_state.mes_sel]} / {st.session_state.ano_sel}</h3>",
            unsafe_allow_html=True
        )

        if nav_c.button(">>", key="next_m"):
            if st.session_state.mes_sel == 12:
                st.session_state.mes_sel = 1; st.session_state.ano_sel += 1
            else:
                st.session_state.mes_sel += 1
            st.session_state.data_sel = None
            st.rerun()

        # Indice de datas com vencimentos
        df_fil["_vd"] = pd.to_datetime(df_fil["vencimento"], errors="coerce").dt.date

        mes = st.session_state.mes_sel
        ano = st.session_state.ano_sel
        hoje = date.today()

        def info_dia(dia):
            dt = date(ano, mes, dia)
            rows = df_fil[df_fil["_vd"] == dt]
            if len(rows) == 0:
                return None, 0
            alertas_dia = rows["alerta"].tolist()
            for a in ["Vencida", "Critica", "Atencao", "Ok"]:
                if a in alertas_dia:
                    return a, len(rows)
            return "Sem data", len(rows)

        # Cabecalho
        dias_semana = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
        hcols = st.columns(7)
        for i, d in enumerate(dias_semana):
            hcols[i].markdown(f"<center><small><b>{d}</b></small></center>", unsafe_allow_html=True)

        cal_mat = calendar.monthcalendar(ano, mes)
        for semana in cal_mat:
            scols = st.columns(7)
            for i, dia in enumerate(semana):
                if dia == 0:
                    scols[i].write(" ")
                    continue
                alerta_d, n_lic = info_dia(dia)
                eh_hoje = (date(ano, mes, dia) == hoje)
                eh_sel  = (st.session_state.data_sel == date(ano, mes, dia))

                if alerta_d:
                    lbl = f"**{dia}**\n{n_lic}"
                    btn_type = "primary"
                elif eh_hoje:
                    lbl = f"**{dia}**"
                    btn_type = "secondary"
                else:
                    lbl = str(dia)
                    btn_type = "secondary"

                if scols[i].button(lbl, key=f"d_{ano}_{mes}_{dia}",
                                   use_container_width=True, type=btn_type):
                    st.session_state.data_sel = date(ano, mes, dia)
                    st.rerun()

    with col_det:
        if st.session_state.data_sel:
            dt_sel = st.session_state.data_sel
            st.markdown(f"### {dt_sel.strftime('%d/%m/%Y')}")

            df_dia = df_fil[df_fil["_vd"] == dt_sel].copy()

            if len(df_dia) == 0:
                st.info("Nenhuma licenca vence nesta data (com filtros atuais).")
            else:
                for _, row in df_dia.iterrows():
                    cor = COR_ALERTA.get(row["alerta"], "#9E9E9E")
                    dias_txt = f"{int(row['dias_para_vencer'])}d" if pd.notna(row["dias_para_vencer"]) else "?"
                    st.markdown(f"""
<div style="border-left:4px solid {cor};padding:8px 12px;margin:6px 0;background:#FAFAFA;border-radius:4px">
  <b>{row['colaborador']}</b><br>
  <small>{row['tipo_licenca']}</small><br>
  <small>Empresa: {row['empresa']} | CC: {row['centro_custo'] or '-'}</small><br>
  <small>Valor: {formatar_brl(row['valor_licenca'])} | <b style="color:{cor}">{row['alerta']} ({dias_txt})</b></small>
</div>""", unsafe_allow_html=True)

                    with st.expander(f"Editar #{row['id']} - {row['colaborador']}"):
                        e1, e2 = st.columns(2)
                        novo_status = e1.selectbox(
                            "Status", STATUS_VALIDOS,
                            index=STATUS_VALIDOS.index(row["status"]) if row["status"] in STATUS_VALIDOS else 0,
                            key=f"st_{row['id']}"
                        )
                        try:
                            vd_atual = datetime.strptime(row["vencimento"], "%Y-%m-%d").date()
                        except Exception:
                            vd_atual = date.today()
                        nova_data = e2.date_input("Vencimento", value=vd_atual, key=f"dt_{row['id']}")

                        st.markdown("Renovacao rapida:")
                        rcols = st.columns(4)
                        for j, meses_n in enumerate([1, 3, 6, 12]):
                            if rcols[j].button(f"+{meses_n}m", key=f"ren_{row['id']}_{meses_n}"):
                                nova = adicionar_meses(vd_atual, meses_n)
                                atualizar_registro(row["id"], {
                                    "vencimento": nova.strftime("%Y-%m-%d"),
                                    "status": "Renovada"
                                })
                                recalcular_alertas(st.session_state.dias_alerta)
                                st.success(f"Renovado para {nova.strftime('%d/%m/%Y')}")
                                st.rerun()

                        rcols2 = st.columns(4)
                        for j, meses_n in enumerate([24, 36, 48, 60]):
                            if rcols2[j].button(f"+{meses_n}m", key=f"ren2_{row['id']}_{meses_n}"):
                                nova = adicionar_meses(vd_atual, meses_n)
                                atualizar_registro(row["id"], {
                                    "vencimento": nova.strftime("%Y-%m-%d"),
                                    "status": "Renovada"
                                })
                                recalcular_alertas(st.session_state.dias_alerta)
                                st.success(f"Renovado para {nova.strftime('%d/%m/%Y')}")
                                st.rerun()

                        if st.button("Salvar alteracoes", key=f"save_{row['id']}", type="primary"):
                            atualizar_registro(row["id"], {
                                "status": novo_status,
                                "vencimento": nova_data.strftime("%Y-%m-%d")
                            })
                            recalcular_alertas(st.session_state.dias_alerta)
                            st.success("Salvo!")
                            st.rerun()
        else:
            st.markdown("### Proximas a vencer")
            df_urg = df_fil[df_fil["alerta"].isin(["Vencida", "Critica", "Atencao"])].sort_values("dias_para_vencer")
            if len(df_urg) > 0:
                for _, row in df_urg.head(10).iterrows():
                    cor = COR_ALERTA.get(row["alerta"], "#9E9E9E")
                    dias_txt = f"{int(row['dias_para_vencer'])}d" if pd.notna(row["dias_para_vencer"]) else "vencida"
                    st.markdown(f"""
<div style="border-left:4px solid {cor};padding:5px 10px;margin:3px 0;font-size:13px">
  <b>{row['colaborador']}</b> — {row['tipo_licenca']}<br>
  <small>{row['empresa']} | <b style="color:{cor}">{dias_txt}</b></small>
</div>""", unsafe_allow_html=True)
            else:
                st.success("Nenhuma licenca critica no periodo filtrado.")


# ============================================================
# PAGINA: LICENCAS
# ============================================================

elif st.session_state.pagina == "Licencas":
    st.title("Gerenciar Licencas")

    df = carregar_licencas()

    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada.")
        st.stop()

    # Filtros
    fc1, fc2, fc3, fc4 = st.columns(4)
    empresas_l = ["Todas"] + sorted(df["empresa"].dropna().unique().tolist())
    tipos_l    = ["Todos"] + sorted(df["tipo_licenca"].dropna().unique().tolist())
    alertas_l  = ["Todos", "Vencida", "Critica", "Atencao", "Ok", "Sem data"]
    status_l   = ["Todos"] + STATUS_VALIDOS

    emp_f2    = fc1.selectbox("Empresa",          empresas_l, key="lic_emp")
    tipo_f2   = fc2.selectbox("Tipo de Licenca",  tipos_l,    key="lic_tipo")
    alerta_f2 = fc3.selectbox("Alerta",           alertas_l,  key="lic_alerta")
    status_f2 = fc4.selectbox("Status",           status_l,   key="lic_status")

    df_fil2 = df.copy()
    if emp_f2    != "Todas": df_fil2 = df_fil2[df_fil2["empresa"]      == emp_f2]
    if tipo_f2   != "Todos": df_fil2 = df_fil2[df_fil2["tipo_licenca"] == tipo_f2]
    if alerta_f2 != "Todos": df_fil2 = df_fil2[df_fil2["alerta"]       == alerta_f2]
    if status_f2 != "Todos": df_fil2 = df_fil2[df_fil2["status"]       == status_f2]

    st.caption(f"{len(df_fil2)} registros exibidos de {len(df)} total")

    # Tabela editavel
    colunas_edit = ["id", "colaborador", "empresa", "centro_custo", "tipo_licenca",
                    "valor_licenca", "vencimento", "status", "alerta", "dias_para_vencer"]

    edited = st.data_editor(
        df_fil2[colunas_edit].reset_index(drop=True),
        column_config={
            "id":               st.column_config.NumberColumn("ID", disabled=True),
            "colaborador":      st.column_config.TextColumn("Colaborador"),
            "empresa":          st.column_config.SelectboxColumn("Empresa", options=list(EMPRESA_PREFIXOS.values()) + ["Nao informada"]),
            "centro_custo":     st.column_config.TextColumn("Centro de Custo"),
            "tipo_licenca":     st.column_config.TextColumn("Tipo de Licenca"),
            "valor_licenca":    st.column_config.NumberColumn("Valor R$", format="%.2f"),
            "vencimento":       st.column_config.TextColumn("Vencimento (AAAA-MM-DD)"),
            "status":           st.column_config.SelectboxColumn("Status", options=STATUS_VALIDOS),
            "alerta":           st.column_config.TextColumn("Alerta", disabled=True),
            "dias_para_vencer": st.column_config.NumberColumn("Dias", disabled=True),
        },
        use_container_width=True,
        height=420,
        key="tabela_edit",
        num_rows="fixed"
    )

    if st.button("Salvar todas as alteracoes", type="primary", key="salvar_tabela"):
        original = df_fil2[colunas_edit].reset_index(drop=True)
        alterados = 0
        for i in range(len(edited)):
            row_edit = edited.iloc[i]
            row_orig = original.iloc[i]
            campos_mudar = {}
            for col in ["colaborador", "empresa", "centro_custo", "tipo_licenca",
                        "valor_licenca", "vencimento", "status"]:
                if str(row_edit[col]) != str(row_orig[col]):
                    campos_mudar[col] = row_edit[col]
            if campos_mudar:
                atualizar_registro(int(row_edit["id"]), campos_mudar)
                alterados += 1
        if alterados > 0:
            recalcular_alertas(st.session_state.dias_alerta)
            st.success(f"{alterados} registro(s) atualizado(s).")
            st.rerun()
        else:
            st.info("Nenhuma alteracao detectada.")

    st.markdown("---")
    st.subheader("Deletar registro")
    del_id = st.number_input("ID do registro a deletar", min_value=1, step=1, key="del_id")
    if st.button("Deletar registro", type="secondary", key="del_btn"):
        deletar_registro(int(del_id))
        st.warning(f"Registro #{del_id} removido.")
        st.rerun()


# ============================================================
# PAGINA: EXPORTAR
# ============================================================

elif st.session_state.pagina == "Exportar":
    st.title("Exportar Dados")

    df = carregar_licencas()

    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada.")
        st.stop()

    st.subheader("Filtros de exportacao")
    fc1, fc2, fc3, fc4 = st.columns(4)
    empresas_e = ["Todas"] + sorted(df["empresa"].dropna().unique().tolist())
    tipos_e    = ["Todos"] + sorted(df["tipo_licenca"].dropna().unique().tolist())
    alertas_e  = ["Todos", "Vencida", "Critica", "Atencao", "Ok", "Sem data"]
    status_e   = ["Todos"] + STATUS_VALIDOS

    emp_e    = fc1.selectbox("Empresa",         empresas_e, key="exp_emp")
    tipo_e   = fc2.selectbox("Tipo",            tipos_e,    key="exp_tipo")
    alerta_e = fc3.selectbox("Alerta",          alertas_e,  key="exp_alerta")
    status_e2= fc4.selectbox("Status",          status_e,   key="exp_status")

    df_exp = df.copy()
    if emp_e    != "Todas": df_exp = df_exp[df_exp["empresa"]      == emp_e]
    if tipo_e   != "Todos": df_exp = df_exp[df_exp["tipo_licenca"] == tipo_e]
    if alerta_e != "Todos": df_exp = df_exp[df_exp["alerta"]       == alerta_e]
    if status_e2!= "Todos": df_exp = df_exp[df_exp["status"]       == status_e2]

    st.metric("Registros a exportar", len(df_exp))
    st.dataframe(df_exp.head(10), use_container_width=True, hide_index=True)

    if len(df_exp) > 0:
        st.download_button(
            label=f"Baixar Excel ({len(df_exp)} registros)",
            data=gerar_excel(df_exp),
            file_name=f"licencas_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    st.markdown("---")
    st.subheader("Estatisticas gerais")

    s1, s2, s3 = st.columns(3)

    with s1:
        st.markdown("**Por empresa**")
        emp_stats = df.groupby("empresa").size().reset_index(name="Licencas").sort_values("Licencas", ascending=False)
        st.dataframe(emp_stats, use_container_width=True, hide_index=True)

    with s2:
        st.markdown("**Por status de alerta**")
        al_stats = df.groupby("alerta").size().reset_index(name="Registros")
        ordem = ["Vencida","Critica","Atencao","Ok","Sem data"]
        al_stats["_ord"] = al_stats["alerta"].map({a: i for i, a in enumerate(ordem)})
        al_stats = al_stats.sort_values("_ord").drop(columns="_ord")
        st.dataframe(al_stats, use_container_width=True, hide_index=True)

    with s3:
        st.markdown("**Top 10 por valor total**")
        top = (
            df.groupby("tipo_licenca")["valor_licenca"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        top.columns = ["Tipo de Licenca", "Valor Total R$"]
        top["Valor Total R$"] = top["Valor Total R$"].apply(lambda x: formatar_brl(x) if x else "-")
        st.dataframe(top, use_container_width=True, hide_index=True)

