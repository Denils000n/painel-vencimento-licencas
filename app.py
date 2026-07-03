import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import calendar
import unicodedata
import requests
import json
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

DB_PATH = os.environ.get("LICENCAS_DB_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "licencas.db"))

CREDS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ms365_creds.json")

MESES_PT = ["","Janeiro","Fevereiro","Marco","Abril","Maio","Junho",
            "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

EMPRESA_PREFIXOS = {"01":"Afonso Franca","02":"AFFIT","03":"AFDI","04":"AFSW"}

EMPRESA_MAP = {
    "AF":"Afonso Franca","af":"Afonso Franca",
    "Afonso Franca":"Afonso Franca","Afonso Franca":"Afonso Franca",
    "AFFIT":"AFFIT","Affit":"AFFIT",
    "AFDI":"AFDI","AFSW":"AFSW",
}

COLUNAS_SISTEMA = ["Empresa","Colaborador","Centro de Custo",
                   "Tipo de Licenca","Valor da Licenca","Vencimento","Status"]

SINONIMOS = {
    "Empresa":         ["empresa","company","companhia","entidade","filial","unidade"],
    "Colaborador":     ["colaborador","funcionario","employee","nome","name",
                        "usuario","user","login","upn","username"],
    "Centro de Custo": ["centro de custo","cc","cost center","departamento","setor","area","depto","centrocusto"],
    "Tipo de Licenca": ["tipo de licenca","tipo licenca","tipo","licenca",
                        "license","produto","product","software","sku","aplicacao"],
    "Valor da Licenca":["valor da licenca","valor licenca","valor por licenca"],
    "Vencimento":      ["vencimento","expiration","expiry","data vencimento",
                        "data de vencimento","validade","valid until","expires"],
    "Status":          ["status","situacao","state","estado"],
}

STATUS_VALIDOS = ["Pendente","Em andamento","Renovada","Cancelada"]

ALERTA_OPCOES = {
    "30 dias":30,"60 dias":60,"90 dias":90,
    "6 meses":180,"12 meses":365,"18 meses":540,
    "24 meses":730,"36 meses":1095,"48 meses":1460,"60 meses":1825,
}

COR_ALERTA = {
    "Vencida":"#FF5252","Critica":"#FF9800",
    "Atencao":"#FFC107","Ok":"#4CAF50","Sem data":"#9E9E9E",
}

# Nomes amigaveis para SKUs Microsoft 365
MS365_SKU_NAMES = {
    "ENTERPRISEPACK":               "Office 365 E3",
    "ENTERPRISEPREMIUM":            "Office 365 E5",
    "STANDARDPACK":                 "Office 365 E1",
    "DESKLESSPACK":                 "Office 365 F3",
    "O365_BUSINESS_PREMIUM":        "Microsoft 365 Business Premium",
    "SPB":                          "Microsoft 365 Business Premium",
    "SMB_BUSINESS":                 "Microsoft 365 Business Basic",
    "O365_BUSINESS":                "Microsoft 365 Apps for Business",
    "OFFICESUBSCRIPTION":           "Microsoft 365 Apps for Enterprise",
    "TEAMS_EXPLORATORY":            "Microsoft Teams Exploratory",
    "TEAMS1":                       "Microsoft Teams Enterprise",
    "MCOSTANDARD":                  "Skype for Business Online (Plan 2)",
    "POWER_BI_STANDARD":            "Power BI (Free)",
    "POWER_BI_PRO":                 "Power BI Pro",
    "POWER_BI_PREMIUM_P1":          "Power BI Premium P1",
    "PBI_PREMIUM_PER_USER":         "Power BI Premium Per User",
    "FLOW_FREE":                    "Power Automate Free",
    "POWERFLOW_P1":                 "Power Automate Plan 1",
    "POWERFLOW_P2":                 "Power Automate Plan 2",
    "POWERAPPS_VIRAL":              "Power Apps Developer Plan",
    "POWERAPPS_DEV":                "Power Apps Developer Plan",
    "PROJECTPROFESSIONAL":          "Project Online Professional",
    "PROJECTPREMIUM":               "Project Online Premium",
    "PROJECT_PLAN1":                "Project Plan 1",
    "PROJECT_PLAN3":                "Project Plan 3",
    "VISIOCLIENT":                  "Visio Plan 2",
    "VISIOONLINE_PLAN1":            "Visio Plan 1",
    "WINDOWS_STORE":                "Windows Store",
    "WIN10_PRO_ENT_SUB":            "Windows 10/11 Enterprise E3",
    "WIN_DEF_ATP":                  "Microsoft Defender for Endpoint",
    "INTUNE_A":                     "Microsoft Intune",
    "EMS":                          "Enterprise Mobility + Security E3",
    "EMSPREMIUM":                   "Enterprise Mobility + Security E5",
    "AAD_PREMIUM":                  "Azure Active Directory Premium P1",
    "AAD_PREMIUM_P2":               "Azure Active Directory Premium P2",
    "RIGHTSMANAGEMENT":             "Azure Information Protection Plan 1",
    "STREAM":                       "Microsoft Stream",
    "ONEDRIVE_BASIC":               "OneDrive for Business (Basic)",
    "SHAREPOINT_S_DEVELOPER":       "SharePoint Online (Plan 1)",
    "SHAREPOINTENTERPRISE":         "SharePoint Online (Plan 2)",
    "EXCHANGESTANDARD":             "Exchange Online (Plan 1)",
    "EXCHANGEENTERPRISE":           "Exchange Online (Plan 2)",
    "MCOMEETADV":                   "Microsoft 365 Audio Conferencing",
    "PHONESYSTEM_VIRTUALUSER":      "Microsoft Teams Phone Resource Account",
    "MCOEV":                        "Microsoft Teams Phone Standard",
    "TEAMS_PREMIUM":                "Microsoft Teams Premium",
    "COPILOT_STUDIO_VIRAL_TRIAL":   "Power Virtual Agents (Trial)",
    "Microsoft_365_Copilot":        "Microsoft 365 Copilot",
    "DYN365_ENTERPRISE_SALES":      "Dynamics 365 Sales Premium",
    "DYN365_ENTERPRISE_P1":         "Dynamics 365 Customer Engagement Plan",
    "Dynamics_365_Sales_Premium":   "Dynamics 365 Sales Premium",
    "CRM_ONLINE_ENTERPRISE":        "Dynamics 365 Customer Voice",
    "FORMS_PRO":                    "Dynamics 365 Customer Voice",
    "ENTERPRISEPACK_B_PILOT":       "Office 365 E3 (sem Teams)",
    "STANDARDPACK_STUDENT":         "Office 365 A1 for Students",
    "POWER_BI_ADDON":               "Power BI for Office 365 Add-On",
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
            fonte            TEXT    DEFAULT 'planilha',
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

    # Migracoes: adiciona colunas ausentes em bancos criados com versoes antigas
    migrations = [
        "ALTER TABLE licencas ADD COLUMN fonte TEXT DEFAULT 'planilha'",
        "ALTER TABLE licencas ADD COLUMN alerta TEXT DEFAULT 'Sem data'",
        "ALTER TABLE licencas ADD COLUMN dias_para_vencer INTEGER",
        "ALTER TABLE licencas ADD COLUMN criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE licencas ADD COLUMN atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
    ]
    for sql in migrations:
        try:
            conn.execute(sql)
            conn.commit()
        except Exception:
            pass  # coluna ja existe, ignorar

    conn.close()


def carregar_licencas(filtros=None):
    conn = get_conn()
    q = "SELECT * FROM licencas WHERE 1=1"
    params = []
    if filtros:
        if filtros.get("empresa"):   q += " AND empresa=?";      params.append(filtros["empresa"])
        if filtros.get("tipo"):      q += " AND tipo_licenca=?";  params.append(filtros["tipo"])
        if filtros.get("alerta"):    q += " AND alerta=?";        params.append(filtros["alerta"])
        if filtros.get("status"):    q += " AND status=?";        params.append(filtros["status"])
    q += " ORDER BY vencimento ASC NULLS LAST"
    df = pd.read_sql_query(q, conn, params=params)
    conn.close()
    return df


def atualizar_registro(id_, campos):
    conn = get_conn()
    sets = [f"{k}=?" for k in campos] + ["atualizado_em=CURRENT_TIMESTAMP"]
    conn.execute(f"UPDATE licencas SET {','.join(sets)} WHERE id=?",
                 list(campos.values()) + [id_])
    conn.commit()
    conn.close()


def upsert_licencas(df, fonte="planilha"):
    conn = get_conn()
    novos = atualizados = 0
    for _, row in df.iterrows():
        colab   = str(row.get("Colaborador","") or "").strip()
        tipo    = str(row.get("Tipo de Licenca","") or "").strip()
        empresa = str(row.get("Empresa","") or "Nao informada").strip()
        if not colab or not tipo:
            continue
        vencimento = None
        v = row.get("Vencimento")
        if v and pd.notna(v) and str(v).strip() not in ("","None","nan","NaT"):
            try: vencimento = pd.to_datetime(v).strftime("%Y-%m-%d")
            except Exception: pass
        valor = None
        val_raw = row.get("Valor da Licenca")
        if val_raw and pd.notna(val_raw):
            try:
                s = str(val_raw).replace("R$","").replace(".","").replace(",",".").strip()
                valor = float(s)
            except Exception: pass
        status = row.get("Status","Pendente")
        if status not in STATUS_VALIDOS: status = "Pendente"
        centro = str(row.get("Centro de Custo","") or "").strip()
        cur = conn.execute(
            "SELECT id FROM licencas WHERE colaborador=? AND tipo_licenca=? AND COALESCE(empresa,'')=?",
            (colab, tipo, empresa))
        ex = cur.fetchone()
        if ex:
            conn.execute(
                "UPDATE licencas SET empresa=?,centro_custo=?,valor_licenca=?,"
                "vencimento=?,status=?,fonte=?,atualizado_em=CURRENT_TIMESTAMP WHERE id=?",
                (empresa,centro,valor,vencimento,status,fonte,ex[0]))
            atualizados += 1
        else:
            conn.execute(
                "INSERT INTO licencas (empresa,colaborador,centro_custo,tipo_licenca,"
                "valor_licenca,vencimento,status,fonte) VALUES(?,?,?,?,?,?,?,?)",
                (empresa,colab,centro,tipo,valor,vencimento,status,fonte))
            novos += 1
    conn.commit()
    conn.close()
    return novos, atualizados


def log_importacao(arquivo, aba, novos, atualizados, total):
    conn = get_conn()
    conn.execute(
        "INSERT INTO importacoes (arquivo,aba,registros_novos,registros_atualizados,registros_total)"
        " VALUES(?,?,?,?,?)", (arquivo,aba,novos,atualizados,total))
    conn.commit()
    conn.close()


def get_historico():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT arquivo,aba,data_importacao,registros_novos,registros_atualizados,registros_total "
        "FROM importacoes ORDER BY data_importacao DESC LIMIT 50", conn)
    conn.close()
    return df


def recalcular_alertas(dias_alerta):
    conn = get_conn()
    rows = conn.execute("SELECT id,vencimento FROM licencas").fetchall()
    hoje = date.today()
    lim_atencao = int(dias_alerta)
    lim_critica  = max(7, int(dias_alerta * 0.33))
    updates = []
    for id_, venc_str in rows:
        if not venc_str:
            updates.append(("Sem data", None, id_))
            continue
        try:
            venc = datetime.strptime(venc_str, "%Y-%m-%d").date()
            dias = (venc - hoje).days
            if dias < 0:            alerta = "Vencida"
            elif dias <= lim_critica:  alerta = "Critica"
            elif dias <= lim_atencao:  alerta = "Atencao"
            else:                      alerta = "Ok"
            updates.append((alerta, dias, id_))
        except Exception:
            updates.append(("Sem data", None, id_))
    conn.executemany("UPDATE licencas SET alerta=?,dias_para_vencer=? WHERE id=?", updates)
    conn.commit()
    conn.close()

# ============================================================
# UTILITARIOS DE IMPORTACAO
# ============================================================

def normalizar(txt):
    txt = str(txt).lower().strip()
    txt = " ".join(txt.split())  # colapsa espacos duplos (ex: "valor  da licenca")
    return unicodedata.normalize("NFKD", txt).encode("ascii","ignore").decode()


def fix_enc(s):
    if isinstance(s, str):
        try: return s.encode('latin1').decode('utf-8')
        except Exception: return s
    return s


def norm_empresa(v):
    if pd.isna(v): return "Nao informada"
    s = fix_enc(str(v).strip())
    return EMPRESA_MAP.get(s, s)


def dedup_columns(df):
    """Remove colunas duplicadas mantendo a primeira ocorrencia."""
    cols = pd.Series(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    df.columns = new_cols
    return df


def listar_abas(arquivo):
    try:
        arquivo.seek(0)
        return pd.ExcelFile(arquivo).sheet_names
    except Exception:
        return []


def ler_arquivo(arquivo, aba=None):
    nome = arquivo.name.lower()
    if nome.endswith(".csv"):
        for sep, enc in [(",","utf-8"),(";","utf-8"),(";","latin1"),(",","latin1")]:
            try:
                arquivo.seek(0)
                df = pd.read_csv(arquivo, sep=sep, encoding=enc)
                return dedup_columns(df)
            except Exception:
                continue
        raise ValueError("Nao foi possivel ler o CSV.")
    else:
        arquivo.seek(0)
        df = pd.read_excel(arquivo, sheet_name=aba)
        return dedup_columns(df)


def detectar_mapeamento(df):
    """Cada coluna-fonte so pode ser mapeada UMA vez. Correspondencia exata ou substring > 4 chars."""
    mapa = {}
    colunas_norm = [normalizar(c) for c in df.columns]
    colunas_orig = list(df.columns)
    fontes_usadas = set()
    for col_sis, sinonimos in SINONIMOS.items():
        for i, col_n in enumerate(colunas_norm):
            src_col = colunas_orig[i]
            if src_col in fontes_usadas:
                continue
            if any(
                s == col_n or
                (len(s) > 4 and s in col_n) or
                (len(col_n) > 4 and col_n in s)
                for s in sinonimos
            ):
                mapa[col_sis] = src_col
                fontes_usadas.add(src_col)
                break
    return mapa


def aplicar_mapeamento(df, mapa):
    # Construir rename sem conflitos
    rename = {}
    destinos_usados = set()
    for sys_col, src_col in mapa.items():
        if src_col in df.columns and sys_col not in destinos_usados:
            if src_col != sys_col:  # apenas renomear se diferente
                rename[src_col] = sys_col
            destinos_usados.add(sys_col)

    # Verificar que o rename nao vai criar duplicatas
    rename_final = {}
    for src, dst in rename.items():
        if dst not in df.columns or src == dst:
            rename_final[src] = dst
        # se dst ja existe como coluna E src != dst → pular (evita duplicata)

    df = df.rename(columns=rename_final)
    df = dedup_columns(df)  # seguranca extra

    for col in COLUNAS_SISTEMA:
        if col not in df.columns:
            df[col] = None

    # Auto-detectar empresa pelo CC se vazio
    mask = df["Empresa"].isna() | (df["Empresa"].astype(str).str.strip() == "")
    if mask.all():
        df["Empresa"] = df["Centro de Custo"].apply(
            lambda x: EMPRESA_PREFIXOS.get(str(x).strip()[:2], "Nao informada") if pd.notna(x) else "Nao informada")
    else:
        df.loc[mask, "Empresa"] = df.loc[mask, "Centro de Custo"].apply(
            lambda x: EMPRESA_PREFIXOS.get(str(x).strip()[:2], "Nao informada") if pd.notna(x) else "Nao informada")

    df["Status"] = df["Status"].apply(lambda x: x if x in STATUS_VALIDOS else "Pendente")
    return df[COLUNAS_SISTEMA]


def eh_mapeamento_exato(mapa):
    """Retorna True se todas as 7 colunas foram mapeadas."""
    return all(c in mapa for c in COLUNAS_SISTEMA)

# ============================================================
# MICROSOFT 365 GRAPH API
# ============================================================

def salvar_creds_ms365(tenant, client_id, client_secret):
    with open(CREDS_PATH, "w") as f:
        json.dump({"tenant_id": tenant, "client_id": client_id,
                   "client_secret": client_secret}, f)


def carregar_creds_ms365():
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def obter_token_ms365(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    data = resp.json()
    if "access_token" not in data:
        raise ValueError(f"Erro ao obter token: {data.get('error_description','Verifique as credenciais.')}")
    return data["access_token"]


def buscar_skus_ms365(token):
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus",
                        headers=headers, timeout=30).json()
    mapa = {}
    for sku in resp.get("value", []):
        part = sku.get("skuPartNumber","")
        nome = MS365_SKU_NAMES.get(part, part)
        mapa[sku["skuId"]] = nome
    return mapa


def buscar_usuarios_ms365(token, sku_map):
    headers = {"Authorization": f"Bearer {token}"}
    url = ("https://graph.microsoft.com/v1.0/users"
           "?$select=displayName,department,companyName,userPrincipalName,assignedLicenses"
           "&$top=999")
    registros = []
    paginas = 0
    while url and paginas < 50:  # max 50 paginas = 49.999 usuarios
        resp = requests.get(url, headers=headers, timeout=60).json()
        for u in resp.get("value", []):
            empresa = u.get("companyName") or "Afonso Franca"
            empresa = EMPRESA_MAP.get(empresa, empresa)
            depto   = u.get("department") or ""
            nome    = u.get("displayName") or u.get("userPrincipalName","")
            for lic in u.get("assignedLicenses", []):
                if lic.get("disabledPlans"):
                    continue  # licenca com planos desativados = provavel nao-usada
                sku_id = lic["skuId"]
                tipo   = sku_map.get(sku_id, sku_id)
                registros.append({
                    "Empresa":          empresa,
                    "Colaborador":      nome,
                    "Centro de Custo":  depto,
                    "Tipo de Licenca":  tipo,
                    "Valor da Licenca": None,
                    "Vencimento":       "",
                    "Status":           "Pendente",
                })
        url = resp.get("@odata.nextLink")
        paginas += 1
    return pd.DataFrame(registros) if registros else pd.DataFrame(columns=COLUNAS_SISTEMA)

# ============================================================
# UTILITARIOS GERAIS
# ============================================================

def formatar_brl(valor):
    try:
        return "R$ {:,.2f}".format(float(valor)).replace(",","X").replace(".",",").replace("X",".")
    except Exception:
        return "-"


def adicionar_meses(dt, n):
    import calendar as _cal
    mes = ((dt.month - 1 + n) % 12) + 1
    ano = dt.year + ((dt.month - 1 + n) // 12)
    dia = min(dt.day, _cal.monthrange(ano, mes)[1])
    return dt.replace(year=ano, month=mes, day=dia)


def gerar_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Licencas")
    return out.getvalue()

# ============================================================
# INICIALIZACAO
# ============================================================

init_db()

for k, v in {"pagina":"Painel","mes_sel":date.today().month,"ano_sel":date.today().year,
             "data_sel":None,"dias_alerta":30,"importado":False}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("## Licencas - Afonso Franca")
    st.markdown("---")
    pagina = st.radio("Menu", ["Painel","Importar","MS365 Sync","Licencas","Exportar"],
                      index=["Painel","Importar","MS365 Sync","Licencas","Exportar"]
                      .index(st.session_state.pagina), key="radio_pagina")
    st.session_state.pagina = pagina

    st.markdown("---")
    alerta_idx = list(ALERTA_OPCOES.keys()).index("30 dias")
    alerta_sel = st.selectbox("Alerta de vencimento", list(ALERTA_OPCOES.keys()),
                              index=alerta_idx, key="alerta_sel")
    dias_alerta = ALERTA_OPCOES[alerta_sel]
    if dias_alerta != st.session_state.dias_alerta:
        st.session_state.dias_alerta = dias_alerta

    recalcular_alertas(dias_alerta)

    st.markdown("---")
    df_all = carregar_licencas()
    if len(df_all) > 0:
        venc  = int((df_all["alerta"]=="Vencida").sum())
        crit  = int((df_all["alerta"]=="Critica").sum())
        atenc = int((df_all["alerta"]=="Atencao").sum())
        total_lic = len(df_all)
        st.metric("Total licencas", total_lic)
        c1,c2 = st.columns(2)
        c1.metric("Vencidas", venc, delta=None)
        c2.metric("Criticas", crit, delta=None)
        if atenc > 0:
            st.warning(f"{atenc} licencas vencem em {alerta_sel}")
        if venc > 0:
            st.error(f"{venc} licencas JA VENCIDAS")
    else:
        st.info("Sem licencas. Importe uma planilha.")

# ============================================================
# PAGINA: PAINEL
# ============================================================

if st.session_state.pagina == "Painel":
    st.title("Painel de Licencas")

    df = carregar_licencas()
    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada ainda.")
        st.markdown("**Como comecar:**")
        st.markdown("1. Va em **Importar** e faca upload da planilha `licencas_padronizado.xlsx`")
        st.markdown("2. Ou va em **MS365 Sync** para buscar automaticamente as licencas do Microsoft 365")
        st.stop()

    # ── Filtros ───────────────────────────────────────────────────────────
    with st.expander("Filtros", expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        emp_f    = fc1.selectbox("Empresa", ["Todas"] + sorted(df["empresa"].dropna().unique().tolist()),     key="pan_emp")
        tipo_f   = fc2.selectbox("Tipo",    ["Todos"] + sorted(df["tipo_licenca"].dropna().unique().tolist()),key="pan_tipo")
        fonte_f  = fc3.selectbox("Fonte",   ["Todos", "planilha", "ms365_sync"],                             key="pan_fonte")
        status_f = fc4.selectbox("Status",  ["Todos"] + STATUS_VALIDOS,                                      key="pan_st")

    df_fil = df.copy()
    if emp_f    != "Todas": df_fil = df_fil[df_fil["empresa"]      == emp_f]
    if tipo_f   != "Todos": df_fil = df_fil[df_fil["tipo_licenca"] == tipo_f]
    if fonte_f  != "Todos": df_fil = df_fil[df_fil["fonte"]        == fonte_f]
    if status_f != "Todos": df_fil = df_fil[df_fil["status"]       == status_f]

    n_venc  = int((df_fil["alerta"] == "Vencida").sum())
    n_crit  = int((df_fil["alerta"] == "Critica").sum())
    n_atenc = int((df_fil["alerta"] == "Atencao").sum())
    n_ok    = int((df_fil["alerta"] == "Ok").sum())
    n_sd    = int((df_fil["alerta"] == "Sem data").sum())
    val_tot = df_fil["valor_licenca"].dropna().sum()
    n_ms365 = int((df_fil["fonte"] == "ms365_sync").sum())
    n_plan  = int((df_fil["fonte"] == "planilha").sum())

    # ── KPI row ───────────────────────────────────────────────────────────
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total", len(df_fil))
    m2.metric("Vencidas",  n_venc)
    m3.metric("Criticas",  n_crit)
    m4.metric("Atencao",   n_atenc)
    m5.metric("OK",        n_ok)
    m6.metric("Valor total", formatar_brl(val_tot) if val_tot > 0 else "-")

    st.caption(f"Fontes: Planilha {n_plan} | MS365 Sync {n_ms365}")
    st.markdown("---")

    # ── Helpers de card ───────────────────────────────────────────────────
    def _card_dias_txt(row):
        dias = row.get("dias_para_vencer")
        try:
            dias_int = int(dias)
        except (TypeError, ValueError):
            return "Sem data de vencimento"
        if dias_int < 0:
            return f"Vencida ha {abs(dias_int)} dias"
        elif dias_int == 0:
            return "Vence HOJE"
        else:
            return f"Vence em {dias_int} dias"

    def _render_single_card(row):
        cor = COR_ALERTA.get(row["alerta"], "#9E9E9E")
        dias_txt = _card_dias_txt(row)
        fonte_icon = "MS365" if row["fonte"] == "ms365_sync" else "Planilha"
        venc_fmt = ""
        if row.get("vencimento"):
            try:
                venc_fmt = datetime.strptime(row["vencimento"], "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                pass
        bg_map = {
            "Vencida": "#FFF0F0",
            "Critica": "#FFF5EC",
            "Atencao": "#FFFCE8",
            "Ok":      "#F0FFF4",
            "Sem data":"#F8F8F8",
        }
        bg = bg_map.get(row["alerta"], "#FFFFFF")
        nome  = str(row["colaborador"])
        tipo  = str(row["tipo_licenca"])
        nome_disp = (nome[:28] + "...") if len(nome) > 28 else nome
        tipo_disp = (tipo[:34] + "...") if len(tipo) > 34 else tipo
        venc_line = (
            f'<div style="font-size:11px;color:#888">Vencimento: {venc_fmt}</div>'
            if venc_fmt else ""
        )
        st.markdown(
            f'<div style="border-left:4px solid {cor};background:{bg};'
            f'padding:10px 14px;border-radius:6px;margin-bottom:2px">'
            f'<div style="font-weight:700;font-size:14px;color:#1a1a1a" title="{nome}">{nome_disp}</div>'
            f'<div style="font-size:12px;color:#444;margin:2px 0">{tipo_disp}</div>'
            f'<div style="font-size:11px;color:#777">{row["empresa"]} &nbsp;|&nbsp; {fonte_icon}</div>'
            f'<div style="font-size:12px;font-weight:600;color:{cor};margin-top:4px">'
            f'{row["alerta"]} &mdash; {dias_txt}</div>'
            f'{venc_line}'
            f'</div>',
            unsafe_allow_html=True
        )

        with st.expander(f"Detalhes / Renovar  #{row['id']}"):
            dc1, dc2 = st.columns(2)
            dc1.markdown(f"**Colaborador:** {row['colaborador']}")
            dc1.markdown(f"**Empresa:** {row['empresa']}")
            dc1.markdown(f"**Centro de Custo:** {row['centro_custo'] or '-'}")
            dc1.markdown(f"**Fonte:** {fonte_icon}")
            dc2.markdown(f"**Tipo de Licenca:** {row['tipo_licenca']}")
            dc2.markdown(f"**Valor:** {formatar_brl(row['valor_licenca'])}")
            dc2.markdown(f"**Vencimento:** {venc_fmt or 'Sem data'}")
            dc2.markdown(f"**Status:** {row['status']}  |  **Alerta:** {row['alerta']}")

            if row.get("vencimento"):
                st.markdown("**Renovar rapido:**")
                try:
                    vd_atual = datetime.strptime(row["vencimento"], "%Y-%m-%d").date()
                except Exception:
                    vd_atual = date.today()
                rc = st.columns(8)
                for jj, mn in enumerate([1, 3, 6, 12, 24, 36, 48, 60]):
                    if rc[jj].button(f"+{mn}m", key=f"ren_{row['id']}_{mn}"):
                        nova = adicionar_meses(vd_atual, mn)
                        atualizar_registro(row["id"], {
                            "vencimento": nova.strftime("%Y-%m-%d"),
                            "status": "Renovada"
                        })
                        recalcular_alertas(st.session_state.dias_alerta)
                        st.success(f"Renovado para {nova.strftime('%d/%m/%Y')}")
                        st.rerun()

            ec1, ec2, ec3 = st.columns([2, 2, 1])
            ns = ec1.selectbox(
                "Status", STATUS_VALIDOS,
                index=STATUS_VALIDOS.index(row["status"]) if row["status"] in STATUS_VALIDOS else 0,
                key=f"st_{row['id']}"
            )
            try:
                _vd_def = datetime.strptime(row["vencimento"], "%Y-%m-%d").date() if row.get("vencimento") else date.today()
            except Exception:
                _vd_def = date.today()
            nd = ec2.date_input("Vencimento", value=_vd_def, key=f"dt_{row['id']}")
            if ec3.button("Salvar", key=f"sv_{row['id']}", type="primary"):
                atualizar_registro(row["id"], {"status": ns, "vencimento": nd.strftime("%Y-%m-%d")})
                recalcular_alertas(st.session_state.dias_alerta)
                st.success("Salvo!")
                st.rerun()

    def render_cards(df_sub, max_cards=150):
        if len(df_sub) == 0:
            st.success("Nenhum registro nesta categoria.")
            return
        if len(df_sub) > max_cards:
            st.caption(f"Mostrando os {max_cards} primeiros de {len(df_sub)} registros. Use os filtros para refinar.")
            df_sub = df_sub.head(max_cards)
        for ii in range(0, len(df_sub), 3):
            cols = st.columns(3)
            for jj, (_, row) in enumerate(df_sub.iloc[ii:ii+3].iterrows()):
                with cols[jj]:
                    _render_single_card(row)

    def _sort_alerta(df_s):
        order = {"Vencida": 0, "Critica": 1, "Atencao": 2, "Ok": 3, "Sem data": 4}
        df_s = df_s.copy()
        df_s["_ord"] = df_s["alerta"].map(order).fillna(5)
        return df_s.sort_values(["_ord", "dias_para_vencer"], na_position="last").drop(columns="_ord")

    # ── Tabs ──────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        f"Vencidas ({n_venc})",
        f"Criticas ({n_crit})",
        f"Atencao ({n_atenc})",
        f"OK ({n_ok})",
        f"Sem data ({n_sd})",
        "Todas",
    ])

    with tab1:
        render_cards(
            df_fil[df_fil["alerta"] == "Vencida"].sort_values("dias_para_vencer", na_position="last")
        )
    with tab2:
        render_cards(
            df_fil[df_fil["alerta"] == "Critica"].sort_values("dias_para_vencer", na_position="last")
        )
    with tab3:
        render_cards(
            df_fil[df_fil["alerta"] == "Atencao"].sort_values("dias_para_vencer", na_position="last")
        )
    with tab4:
        render_cards(
            df_fil[df_fil["alerta"] == "Ok"].sort_values("dias_para_vencer", na_position="last"),
            max_cards=60
        )
    with tab5:
        render_cards(
            df_fil[df_fil["alerta"] == "Sem data"].sort_values("colaborador"),
            max_cards=60
        )
    with tab6:
        render_cards(_sort_alerta(df_fil), max_cards=150)

    # ── Calendario (colapsado) ────────────────────────────────────────────
    with st.expander("Calendario de vencimentos", expanded=False):
        nav_a, nav_b, nav_c = st.columns([1, 4, 1])
        if nav_a.button("<<", key="prev_m"):
            if st.session_state.mes_sel == 1:
                st.session_state.mes_sel = 12; st.session_state.ano_sel -= 1
            else:
                st.session_state.mes_sel -= 1
            st.session_state.data_sel = None; st.rerun()
        nav_b.markdown(
            f"<h3 style='text-align:center;margin:0'>"
            f"{MESES_PT[st.session_state.mes_sel]} / {st.session_state.ano_sel}</h3>",
            unsafe_allow_html=True)
        if nav_c.button(">>", key="next_m"):
            if st.session_state.mes_sel == 12:
                st.session_state.mes_sel = 1; st.session_state.ano_sel += 1
            else:
                st.session_state.mes_sel += 1
            st.session_state.data_sel = None; st.rerun()

        df_fil["_vd"] = pd.to_datetime(df_fil["vencimento"], errors="coerce").dt.date
        mes  = st.session_state.mes_sel
        ano  = st.session_state.ano_sel
        hoje = date.today()

        def info_dia(dia):
            dt   = date(ano, mes, dia)
            rows = df_fil[df_fil["_vd"] == dt]
            if len(rows) == 0: return None, 0
            for a in ["Vencida", "Critica", "Atencao", "Ok"]:
                if a in rows["alerta"].tolist(): return a, len(rows)
            return "Sem data", len(rows)

        for i, d in enumerate(["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]):
            st.columns(7)[i].markdown(f"<center><small><b>{d}</b></small></center>", unsafe_allow_html=True)

        for semana in calendar.monthcalendar(ano, mes):
            scols = st.columns(7)
            for i, dia in enumerate(semana):
                if dia == 0: scols[i].write(" "); continue
                al, n_lic = info_dia(dia)
                eh_hoje   = date(ano, mes, dia) == hoje
                if al:         lbl = f"**{dia}**\n{n_lic}"; tp = "primary"
                elif eh_hoje:  lbl = f"**{dia}**";          tp = "secondary"
                else:          lbl = str(dia);              tp = "secondary"
                if scols[i].button(lbl, key=f"d_{ano}_{mes}_{dia}",
                                   use_container_width=True, type=tp):
                    st.session_state.data_sel = date(ano, mes, dia); st.rerun()

        if st.session_state.data_sel:
            dt_sel = st.session_state.data_sel
            st.markdown(f"### {dt_sel.strftime('%d/%m/%Y')}")
            df_dia = df_fil[df_fil["_vd"] == dt_sel]
            if len(df_dia) == 0:
                st.info("Nenhuma licenca vence nesta data.")
            else:
                for _, row in df_dia.iterrows():
                    cor2     = COR_ALERTA.get(row["alerta"], "#9E9E9E")
                    dias_txt2= f"{int(row['dias_para_vencer'])}d" if pd.notna(row.get("dias_para_vencer")) else "?"
                    st.markdown(
                        f'<div style="border-left:4px solid {cor2};padding:8px 12px;margin:6px 0;'
                        f'background:#FAFAFA;border-radius:4px">'
                        f'<b>{row["colaborador"]}</b><br>'
                        f'<small>{row["tipo_licenca"]}</small><br>'
                        f'<small>Empresa: {row["empresa"]} | CC: {row["centro_custo"] or "-"}</small><br>'
                        f'<small>Valor: {formatar_brl(row["valor_licenca"])} | '
                        f'<b style="color:{cor2}">{row["alerta"]} ({dias_txt2})</b></small>'
                        f'</div>',
                        unsafe_allow_html=True
                    )


# ============================================================
# PAGINA: IMPORTAR
# ============================================================

elif st.session_state.pagina == "Importar":
    st.title("Importar Planilha")
    st.markdown(
        "Faca upload da planilha. O sistema detecta as colunas automaticamente e salva no banco.  \n"
        "Registros com mesmo **Colaborador + Tipo + Empresa** serao **atualizados** automaticamente.")

    arquivo = st.file_uploader("Selecione a planilha (.xlsx ou .csv)", type=["xlsx","xls","csv"])

    if arquivo:
        abas = listar_abas(arquivo)
        aba_sel = None
        juntar = False

        if abas:
            if len(abas) > 1:
                opcao = st.radio("Qual aba?", ["Importar todas as abas"] + abas, key="aba_sel")
                juntar = opcao == "Importar todas as abas"
                aba_sel = None if juntar else opcao
            else:
                aba_sel = abas[0]
                st.caption(f"Aba: {aba_sel}")

        try:
            if juntar:
                dfs = []
                for aba in abas:
                    try:
                        arquivo.seek(0)
                        df_aba = pd.read_excel(arquivo, sheet_name=aba)
                        df_aba = dedup_columns(df_aba)
                        df_aba = df_aba.dropna(how="all")
                        if len(df_aba) > 0 and len(df_aba.columns) > 0:
                            dfs.append(df_aba)
                    except Exception:
                        pass
                if not dfs:
                    st.error("Nenhuma aba com dados validos encontrada.")
                    st.stop()
                df_raw = pd.concat(dfs, ignore_index=True)
                df_raw = dedup_columns(df_raw)
                abas_label = "todas"
            else:
                df_raw = ler_arquivo(arquivo, aba_sel)
                abas_label = aba_sel or "csv"

            df_raw = df_raw.dropna(how="all")
            st.caption(f"{len(df_raw)} linhas | {len(df_raw.columns)} colunas: {', '.join(df_raw.columns.tolist()[:8])}")

            # Detectar mapeamento automatico
            mapa_auto = detectar_mapeamento(df_raw)
            mapeado_completo = eh_mapeamento_exato(mapa_auto)

            if mapeado_completo:
                # IMPORTACAO DIRETA - sem necessidade de ajuste manual
                st.success("Colunas detectadas automaticamente. Pronto para importar!")
                df_prev = aplicar_mapeamento(df_raw.copy(), mapa_auto)
                df_prev = df_prev.dropna(subset=["Colaborador","Tipo de Licenca"])

                col_info1, col_info2, col_info3 = st.columns(3)
                col_info1.metric("Registros", len(df_prev))
                col_info2.metric("Com vencimento", df_prev["Vencimento"].notna().sum())
                col_info3.metric("Com valor", df_prev["Valor da Licenca"].notna().sum())

                with st.expander("Preview (5 primeiros registros)"):
                    st.dataframe(df_prev.head(5), use_container_width=True)

                if st.button(f"Importar {len(df_prev)} registros", type="primary", key="btn_import_direto"):
                    novos, atualizados = upsert_licencas(df_prev, fonte="planilha")
                    log_importacao(arquivo.name, abas_label, novos, atualizados, len(df_prev))
                    recalcular_alertas(st.session_state.dias_alerta)
                    st.success(f"Concluido: {novos} novos + {atualizados} atualizados")
                    st.session_state.pagina = "Painel"
                    st.rerun()
            else:
                # Mapeamento parcial - mostrar UI de ajuste
                faltando = [c for c in COLUNAS_SISTEMA if c not in mapa_auto]
                st.warning(f"Ajuste o mapeamento para: {', '.join(faltando)}")

                cols_disp = ["(nao mapear)"] + df_raw.columns.tolist()
                mapa_user = {}
                grid = st.columns(2)
                for idx, col_sis in enumerate(COLUNAS_SISTEMA):
                    container = grid[idx % 2]
                    default = mapa_auto.get(col_sis, "(nao mapear)")
                    if default not in cols_disp: default = "(nao mapear)"
                    sel = container.selectbox(col_sis, cols_disp,
                        index=cols_disp.index(default), key=f"map_{col_sis}")
                    if sel != "(nao mapear)":
                        mapa_user[col_sis] = sel

                if "Colaborador" in mapa_user and "Tipo de Licenca" in mapa_user:
                    df_prev = aplicar_mapeamento(df_raw.copy(), mapa_user)
                    df_prev = df_prev.dropna(subset=["Colaborador","Tipo de Licenca"])
                    with st.expander("Preview"):
                        st.dataframe(df_prev.head(5), use_container_width=True)
                    if st.button(f"Importar {len(df_prev)} registros", type="primary", key="btn_import_manual"):
                        novos, atualizados = upsert_licencas(df_prev, fonte="planilha")
                        log_importacao(arquivo.name, abas_label, novos, atualizados, len(df_prev))
                        recalcular_alertas(st.session_state.dias_alerta)
                        st.success(f"Concluido: {novos} novos + {atualizados} atualizados")
                        st.session_state.pagina = "Painel"
                        st.rerun()
                else:
                    st.error("Mapeie pelo menos Colaborador e Tipo de Licenca.")

        except Exception as e:
            st.error(f"Erro: {e}")

    st.markdown("---")
    st.subheader("Historico de importacoes")
    hist = get_historico()
    if len(hist) > 0:
        hist.columns = ["Arquivo","Aba","Data","Novos","Atualizados","Total"]
        st.dataframe(hist, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma importacao realizada ainda.")

# ============================================================
# PAGINA: MS365 SYNC
# ============================================================

elif st.session_state.pagina == "MS365 Sync":
    st.title("Sincronizar Microsoft 365")
    st.markdown(
        "Busca automaticamente todos os usuarios e suas licencas atribuidas no seu tenant MS365.  \n"
        "Os dados sao inseridos/atualizados no banco com a mesma logica das planilhas importadas.")

    with st.expander("Como obter as credenciais?", expanded=False):
        st.markdown("""
**Passos no portal Azure (portal.azure.com):**

1. Va em **Azure Active Directory** → **Registros de app** → **Novo registro**
2. Nome: ex. `Gerenciador Licencas`; Conta: *Este diretorio apenas*
3. Apos criar: copie o **ID do aplicativo (client_id)** e o **ID do diretorio (tenant_id)**
4. Va em **Certificados e segredos** → **Novo segredo** → copie o valor (client_secret)
5. Va em **Permissoes de API** → **Adicionar permissao** → **Microsoft Graph** → **Permissoes de aplicativo**
6. Adicione: `User.Read.All` e `Organization.Read.All`
7. Clique em **Conceder consentimento do administrador**

As credenciais sao salvas localmente no arquivo `.ms365_creds.json` ao lado do app.
        """)

    creds = carregar_creds_ms365()
    with st.form("form_creds"):
        st.subheader("Credenciais Azure AD")
        tenant_id     = st.text_input("Tenant ID (ID do diretorio)", value=creds.get("tenant_id",""))
        client_id     = st.text_input("Client ID (ID do aplicativo)", value=creds.get("client_id",""))
        client_secret = st.text_input("Client Secret", value=creds.get("client_secret",""), type="password")
        col_save, col_sync = st.columns(2)
        salvar_btn = col_save.form_submit_button("Salvar credenciais")
        sincronizar_btn = col_sync.form_submit_button("Sincronizar agora", type="primary")

    if salvar_btn:
        if tenant_id and client_id and client_secret:
            salvar_creds_ms365(tenant_id, client_id, client_secret)
            st.success("Credenciais salvas.")
        else:
            st.error("Preencha todos os campos.")

    if sincronizar_btn:
        if not (tenant_id and client_id and client_secret):
            st.error("Preencha e salve as credenciais antes de sincronizar.")
        else:
            salvar_creds_ms365(tenant_id, client_id, client_secret)
            with st.spinner("Conectando ao Microsoft Graph..."):
                try:
                    token = obter_token_ms365(tenant_id, client_id, client_secret)
                    st.info("Token obtido. Buscando SKUs...")
                    sku_map = buscar_skus_ms365(token)
                    st.info(f"{len(sku_map)} tipos de licenca encontrados. Buscando usuarios...")
                    df_ms = buscar_usuarios_ms365(token, sku_map)

                    if len(df_ms) == 0:
                        st.warning("Nenhum usuario com licenca atribuida encontrado.")
                    else:
                        st.info(f"{len(df_ms)} registros encontrados. Importando...")
                        novos, atualizados = upsert_licencas(df_ms, fonte="ms365_sync")
                        log_importacao("MS365 Graph API", "sync", novos, atualizados, len(df_ms))
                        recalcular_alertas(st.session_state.dias_alerta)
                        st.success(f"Sincronizacao concluida: {novos} novos + {atualizados} atualizados")

                        col1, col2, col3 = st.columns(3)
                        col1.metric("Total registros", len(df_ms))
                        col2.metric("Usuarios unicos", df_ms["Colaborador"].nunique())
                        col3.metric("Tipos de licenca", df_ms["Tipo de Licenca"].nunique())

                        with st.expander("Preview"):
                            st.dataframe(df_ms.head(10), use_container_width=True)

                        if st.button("Ir para o Painel", type="primary"):
                            st.session_state.pagina = "Painel"; st.rerun()
                except Exception as e:
                    st.error(f"Erro na sincronizacao: {e}")

    st.markdown("---")
    st.subheader("Status da ultima sincronizacao")
    conn = get_conn()
    last_sync = pd.read_sql_query(
        "SELECT data_importacao,registros_novos,registros_atualizados,registros_total "
        "FROM importacoes WHERE arquivo='MS365 Graph API' ORDER BY data_importacao DESC LIMIT 1", conn)
    conn.close()
    if len(last_sync) > 0:
        r = last_sync.iloc[0]
        st.info(f"Ultima sync: {r['data_importacao']} | {r['registros_total']} registros "
                f"({r['registros_novos']} novos, {r['registros_atualizados']} atualizados)")
    else:
        st.info("Nenhuma sincronizacao realizada ainda.")

# ============================================================
# PAGINA: LICENCAS
# ============================================================

elif st.session_state.pagina == "Licencas":
    st.title("Gerenciar Licencas")
    df = carregar_licencas()
    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada."); st.stop()

    fc1,fc2,fc3,fc4 = st.columns(4)
    emp_f2    = fc1.selectbox("Empresa",  ["Todas"]+sorted(df["empresa"].dropna().unique().tolist()),    key="lic_emp")
    tipo_f2   = fc2.selectbox("Tipo",     ["Todos"]+sorted(df["tipo_licenca"].dropna().unique().tolist()),key="lic_tipo")
    alerta_f2 = fc3.selectbox("Alerta",   ["Todos","Vencida","Critica","Atencao","Ok","Sem data"],        key="lic_al")
    status_f2 = fc4.selectbox("Status",   ["Todos"]+STATUS_VALIDOS,                                      key="lic_st")

    df_fil2 = df.copy()
    if emp_f2    != "Todas": df_fil2 = df_fil2[df_fil2["empresa"]      == emp_f2]
    if tipo_f2   != "Todos": df_fil2 = df_fil2[df_fil2["tipo_licenca"] == tipo_f2]
    if alerta_f2 != "Todos": df_fil2 = df_fil2[df_fil2["alerta"]       == alerta_f2]
    if status_f2 != "Todos": df_fil2 = df_fil2[df_fil2["status"]       == status_f2]

    st.caption(f"{len(df_fil2)} de {len(df)} registros")
    COLS_ED = ["id","colaborador","empresa","centro_custo","tipo_licenca",
               "valor_licenca","vencimento","status","alerta","dias_para_vencer"]

    edited = st.data_editor(
        df_fil2[COLS_ED].reset_index(drop=True),
        column_config={
            "id":               st.column_config.NumberColumn("ID",       disabled=True),
            "colaborador":      st.column_config.TextColumn("Colaborador"),
            "empresa":          st.column_config.SelectboxColumn("Empresa", options=list(EMPRESA_PREFIXOS.values())+["Nao informada"]),
            "centro_custo":     st.column_config.TextColumn("CC"),
            "tipo_licenca":     st.column_config.TextColumn("Tipo"),
            "valor_licenca":    st.column_config.NumberColumn("Valor R$", format="%.2f"),
            "vencimento":       st.column_config.TextColumn("Vencimento (AAAA-MM-DD)"),
            "status":           st.column_config.SelectboxColumn("Status", options=STATUS_VALIDOS),
            "alerta":           st.column_config.TextColumn("Alerta",    disabled=True),
            "dias_para_vencer": st.column_config.NumberColumn("Dias",    disabled=True),
        },
        use_container_width=True, height=420, num_rows="fixed", key="tabela_ed")

    if st.button("Salvar alteracoes", type="primary"):
        orig = df_fil2[COLS_ED].reset_index(drop=True)
        alt = 0
        for i in range(len(edited)):
            diffs = {c: edited.iloc[i][c] for c in ["colaborador","empresa","centro_custo",
                     "tipo_licenca","valor_licenca","vencimento","status"]
                     if str(edited.iloc[i][c]) != str(orig.iloc[i][c])}
            if diffs:
                atualizar_registro(int(edited.iloc[i]["id"]), diffs); alt += 1
        if alt > 0:
            recalcular_alertas(st.session_state.dias_alerta)
            st.success(f"{alt} registro(s) atualizado(s)."); st.rerun()
        else:
            st.info("Sem alteracoes.")

# ============================================================
# PAGINA: EXPORTAR
# ============================================================

elif st.session_state.pagina == "Exportar":
    st.title("Exportar Dados")
    df = carregar_licencas()
    if len(df) == 0:
        st.info("Nenhuma licenca cadastrada."); st.stop()

    fc1,fc2,fc3,fc4 = st.columns(4)
    emp_e    = fc1.selectbox("Empresa", ["Todas"]+sorted(df["empresa"].dropna().unique().tolist()), key="exp_e")
    tipo_e   = fc2.selectbox("Tipo",    ["Todos"]+sorted(df["tipo_licenca"].dropna().unique().tolist()), key="exp_t")
    alerta_e = fc3.selectbox("Alerta",  ["Todos","Vencida","Critica","Atencao","Ok","Sem data"],   key="exp_a")
    status_e = fc4.selectbox("Status",  ["Todos"]+STATUS_VALIDOS, key="exp_s")

    df_exp = df.copy()
    if emp_e    != "Todas": df_exp = df_exp[df_exp["empresa"]      == emp_e]
    if tipo_e   != "Todos": df_exp = df_exp[df_exp["tipo_licenca"] == tipo_e]
    if alerta_e != "Todos": df_exp = df_exp[df_exp["alerta"]       == alerta_e]
    if status_e != "Todos": df_exp = df_exp[df_exp["status"]       == status_e]

    st.metric("Registros a exportar", len(df_exp))
    st.dataframe(df_exp.head(10), use_container_width=True, hide_index=True)

    if len(df_exp) > 0:
        st.download_button(
            label=f"Baixar Excel ({len(df_exp)} registros)",
            data=gerar_excel(df_exp),
            file_name=f"licencas_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary")

    st.markdown("---")
    st.subheader("Estatisticas")
    s1,s2,s3 = st.columns(3)
    with s1:
        st.markdown("**Por empresa**")
        st.dataframe(df.groupby("empresa").size().reset_index(name="Licencas")
                     .sort_values("Licencas",ascending=False), use_container_width=True, hide_index=True)
    with s2:
        st.markdown("**Por alerta**")
        ord_al = {"Vencida":0,"Critica":1,"Atencao":2,"Ok":3,"Sem data":4}
        al = df.groupby("alerta").size().reset_index(name="Registros")
        al["_o"] = al["alerta"].map(ord_al)
        st.dataframe(al.sort_values("_o").drop(columns="_o"), use_container_width=True, hide_index=True)
    with s3:
        st.markdown("**Top 10 por valor**")
        top = df.groupby("tipo_licenca")["valor_licenca"].sum().sort_values(ascending=False).head(10).reset_index()
        top.columns = ["Tipo","Valor Total R$"]
        top["Valor Total R$"] = top["Valor Total R$"].apply(lambda x: formatar_brl(x) if x else "-")
        st.dataframe(top, use_container_width=True, hide_index=True)

