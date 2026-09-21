from pathlib import Path
from io import BytesIO
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse
from datetime import datetime
import html
import unicodedata

import pandas as pd
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# =========================================================
# CONFIGURACAO
# =========================================================
st.set_page_config(page_title="Radar Pre-Op", page_icon="📡", layout="wide")

# Atualiza automaticamente a cada 5 minutos.
st_autorefresh(interval=5 * 60 * 1000, key="radar_atualizacao_automatica")

SHAREPOINT_URL = "https://trigogrupo-my.sharepoint.com/:x:/g/personal/isabel_mendes_grupotrigo_com_br/IQBQBfuaDt2BRZ3DnDhUAB64AXlXAjiIByO_jURmo2U7syo?e=k04bNY"
ARQUIVO_PADRAO = "PROJETOS PREOP 2026 - DASH.xlsx"
ABA_GERAL = "Geral"
LOGO_ARQUIVO = "trigo_logo.webp"

# Colunas da aba Geral informadas pelo usuario (indice Excel -> indice Python)
COL_STATUS = 0       # A - usado somente para manter unidades PREVISTAS
COL_DATA = 1         # B - Data prevista
COL_MARCA = 9        # J - Marca
COL_RESTAURANTE = 10 # K - Restaurante
COL_TIPO = 11        # L - Tipo de projeto
COL_COMENTARIOS = 20 # U - Comentarios
COL_RISCO = 21       # V - Farol riscos

ORANGE = "#F58220"
DARK = "#222222"
RED = "#C93C3C"
GOLD = "#D5A61C"
GREEN = "#2E8B57"
GRAY = "#8B8B8B"
BLUE = "#356A9A"

RISCO_ORDER = ["🔴 ALTO RISCO", "🟡 MÉDIO RISCO", "🟢 BAIXO RISCO", "⚪ SEM RISCO INFORMADO"]
RISCO_COLORS = {
    "🔴 ALTO RISCO": RED,
    "🟡 MÉDIO RISCO": GOLD,
    "🟢 BAIXO RISCO": GREEN,
    "⚪ SEM RISCO INFORMADO": GRAY,
}

TIPO_ORDER = ["Inauguração", "Reforma", "Repasse", "Mudança de Ponto", "Virada Digital", "Migração", "Não informado"]
TIPO_COLORS = {
    "Inauguração": ORANGE,
    "Reforma": BLUE,
    "Repasse": "#187A7B",
    "Mudança de Ponto": "#7655A6",
    "Virada Digital": GOLD,
    "Migração": "#D6527C",
    "Não informado": GRAY,
}

# =========================================================
# FUNCOES
# =========================================================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return " ".join(texto.split())


def padronizar_marca(valor):
    n = normalizar_texto(valor)
    mapa = {
        "CIB": "CHINA IN BOX",
        "CHINA IN BOX": "CHINA IN BOX",
        "CPQ": "CASA DO PÃO DE QUEIJO",
        "CASA DO PAO DE QUEIJO": "CASA DO PÃO DE QUEIJO",
        "ASA": "ASA AÇAÍ",
        "ASA ACAI": "ASA AÇAÍ",
        "SPOLETO": "SPOLETO",
        "PIZZARIA SPOLETO": "PIZZARIA SPOLETO",
        "GENDAI": "GENDAI",
        "KONI": "KONI",
        "2V": "2V",
        "2V RJ": "2V",
    }
    return mapa.get(n, str(valor).strip() if pd.notna(valor) else "SEM MARCA")


def padronizar_tipo(valor):
    n = normalizar_texto(valor)
    if "INAUGUR" in n or "IMPLANT" in n:
        return "Inauguração"
    if "REFORMA" in n:
        return "Reforma"
    if "REPASSE" in n:
        return "Repasse"
    if "VIRADA DIGITAL" in n:
        return "Virada Digital"
    if "MUDANCA DE PONTO" in n:
        return "Mudança de Ponto"
    if "MIGRACAO" in n:
        return "Migração"
    return str(valor).strip() if pd.notna(valor) else "Não informado"


def classificar_risco(valor):
    bruto = "" if pd.isna(valor) else str(valor).strip()
    txt = normalizar_texto(bruto)

    if "🔴" in bruto or any(t in txt for t in ["ALTO RISCO", "ALTO", "VERMELH", "CRITIC", "URGENT"]):
        return "🔴 ALTO RISCO"
    if "🟡" in bruto or "⚠" in bruto or any(t in txt for t in ["MEDIO RISCO", "MEDIO", "AMAREL", "ATENCAO"]):
        return "🟡 MÉDIO RISCO"
    if "🟢" in bruto or any(t in txt for t in ["BAIXO RISCO", "BAIXO", "VERDE", "OK"]):
        return "🟢 BAIXO RISCO"
    return "⚪ SEM RISCO INFORMADO"


def url_download_sharepoint(url: str) -> str:
    p = urlparse(url)
    query = dict(parse_qsl(p.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunparse((p.scheme, p.netloc, p.path, p.params, urlencode(query), p.fragment))


@st.cache_data(ttl=300, show_spinner=False)
def baixar_sharepoint(url: str):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    }
    r = requests.get(url_download_sharepoint(url), headers=headers, timeout=30, allow_redirects=True)
    r.raise_for_status()
    conteudo = r.content
    content_type = (r.headers.get("content-type") or "").lower()

    if not conteudo.startswith(b"PK"):
        if "text/html" in content_type or b"<html" in conteudo[:500].lower():
            raise PermissionError("O SharePoint solicitou autenticação.")
        raise ValueError("O conteúdo recebido do SharePoint não parece ser um arquivo XLSX.")
    return conteudo


@st.cache_data(ttl=300, show_spinner=False)
def carregar_excel(fonte: str, arquivo_local: str):
    if fonte == "SharePoint":
        try:
            return baixar_sharepoint(SHAREPOINT_URL), "SharePoint (online)", None
        except Exception as e:
            caminho = Path(arquivo_local)
            if caminho.exists():
                return caminho.read_bytes(), "Arquivo local (contingência)", str(e)
            raise

    caminho = Path(arquivo_local)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return caminho.read_bytes(), "Arquivo local", None


@st.cache_data(ttl=300, show_spinner=False)
def carregar_radar_geral(fonte: str, arquivo_local: str):
    conteudo, fonte_usada, aviso_fonte = carregar_excel(fonte, arquivo_local)

    df = pd.read_excel(BytesIO(conteudo), sheet_name=ABA_GERAL, engine="openpyxl")

    if df.shape[1] <= COL_RISCO:
        raise ValueError(
            f"A aba '{ABA_GERAL}' possui {df.shape[1]} colunas, mas o Radar precisa chegar até a coluna V."
        )

    radar = pd.DataFrame({
        "Status": df.iloc[:, COL_STATUS],
        "Data Prevista": df.iloc[:, COL_DATA],
        "Marca": df.iloc[:, COL_MARCA],
        "Restaurante": df.iloc[:, COL_RESTAURANTE],
        "Tipo de Projeto": df.iloc[:, COL_TIPO],
        "Comentários": df.iloc[:, COL_COMENTARIOS],
        "Farol Riscos": df.iloc[:, COL_RISCO],
    })

    radar["Status_norm"] = radar["Status"].map(normalizar_texto)
    radar["Data Prevista"] = pd.to_datetime(radar["Data Prevista"], errors="coerce")
    radar["Marca"] = radar["Marca"].map(padronizar_marca)
    radar["Tipo de Projeto"] = radar["Tipo de Projeto"].map(padronizar_tipo)
    radar["Risco"] = radar["Farol Riscos"].map(classificar_risco)

    # Mantem somente projetos PREVISTOS, como no conceito original do Radar do Pre-Op.
    radar = radar[radar["Status_norm"] == "PREVISTO"].copy()

    # Remove linhas sem unidade/restaurante.
    radar = radar[
        radar["Restaurante"].notna()
        & radar["Restaurante"].astype(str).str.strip().ne("")
        & radar["Restaurante"].astype(str).str.strip().str.lower().ne("nan")
    ].copy()

    radar["Restaurante"] = radar["Restaurante"].astype(str).str.strip()
    radar["Comentários"] = radar["Comentários"].fillna("").astype(str).str.strip()
    radar["Farol Riscos"] = radar["Farol Riscos"].fillna("").astype(str).str.strip()

    return radar.reset_index(drop=True), fonte_usada, aviso_fonte


def card_unidade(row):
    risco = row["Risco"]
    cor = RISCO_COLORS.get(risco, GRAY)
    data = row["Data Prevista"].strftime("%d/%m/%Y") if pd.notna(row["Data Prevista"]) else "Sem data prevista"
    restaurante = html.escape(str(row["Restaurante"]))
    tipo = html.escape(str(row["Tipo de Projeto"]))
    comentario = html.escape(row["Comentários"] or "Sem comentários cadastrados").replace("\n", "<br>")

    return f"""
    <div class="radar-card" style="border-left:6px solid {cor};">
        <div class="radar-topline">
            <div>
                <div class="radar-name">{restaurante}</div>
                <div class="radar-meta">{tipo} &nbsp;•&nbsp; Previsão: <b>{data}</b></div>
            </div>
            <div class="risk-pill" style="background:{cor}18;color:{cor};border:1px solid {cor}45;">{risco}</div>
        </div>
        <div class="comment-title">COMENTÁRIOS</div>
        <div class="comment-box">{comentario}</div>
    </div>
    """


# =========================================================
# ESTILO
# =========================================================
st.markdown(f"""
<style>
.stApp {{ background: linear-gradient(180deg,#FFF8F1 0,#F7F7F7 260px); }}
.block-container {{ padding-top: 1rem; padding-bottom: 2.5rem; max-width: 1500px; }}
h1,h2,h3 {{ color: {DARK}; letter-spacing:-0.02em; }}
[data-testid="stSidebar"] {{ background:#FFFFFF; border-right:1px solid #ECECEC; }}
.source-pill {{display:inline-block;background:#fff;border:1px solid #e8e8e8;border-radius:20px;padding:6px 12px;color:#555;font-size:12px;}}
.radar-card {{background:#fff;border-radius:14px;padding:16px 18px;margin:10px 0 14px 0;box-shadow:0 2px 11px rgba(0,0,0,.055);}}
.radar-topline {{display:flex;justify-content:space-between;gap:18px;align-items:flex-start;}}
.radar-name {{font-size:17px;font-weight:800;color:#262626;line-height:1.25;}}
.radar-meta {{font-size:12.5px;color:#6b7280;margin-top:5px;}}
.risk-pill {{font-size:11px;font-weight:800;border-radius:999px;padding:6px 10px;white-space:nowrap;}}
.comment-title {{font-size:10px;font-weight:800;color:#8a8a8a;letter-spacing:.8px;margin-top:13px;margin-bottom:4px;}}
.comment-box {{font-size:13.5px;color:#444;line-height:1.45;background:#F8F9FA;border-radius:9px;padding:10px 12px;}}
.brand-count {{color:#777;font-size:12px;margin-top:-7px;margin-bottom:8px;}}
.tipo-section {{background:#FFFFFF;border:1px solid #ECECEC;border-radius:14px;padding:12px 14px 6px 14px;margin:10px 0 16px 0;}}
.tipo-header {{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:6px;}}
.tipo-title {{font-size:15px;font-weight:800;color:#2c2c2c;}}
.tipo-count {{font-size:11px;font-weight:700;color:#6b7280;background:#F4F4F5;border-radius:999px;padding:4px 8px;white-space:nowrap;}}
.sidebar-note {{font-size:11.5px;color:#777;line-height:1.4;margin-top:-4px;}}
</style>
""", unsafe_allow_html=True)

# =========================================================
# CABECALHO
# =========================================================
logo_path = Path(__file__).parent / LOGO_ARQUIVO
head_logo, head_title = st.columns([1.15, 4], vertical_alignment="center")
with head_logo:
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)
    else:
        st.caption("GRUPO TRIGO")

with head_title:
    st.markdown(f"""
    <div style="background:{ORANGE};padding:18px 24px;border-radius:18px;margin-bottom:8px;">
      <div style="font-size:13px;color:white;font-weight:700;letter-spacing:1.2px;">PRÉ-OP | GESTÃO DE PROJETOS</div>
      <div style="font-size:32px;color:white;font-weight:800;">Radar do Pré-Op</div>
      <div style="font-size:14px;color:white;opacity:.92;">Unidades previstas, riscos, datas e comentários por marca</div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# FONTE DE DADOS
# =========================================================
caminho_default = Path(__file__).parent / ARQUIVO_PADRAO
with st.sidebar:
    st.header("Fonte de dados")
    fonte_escolhida = st.radio("Origem", ["SharePoint", "Arquivo local"], index=0)
    st.caption("O painel tenta atualizar a planilha online a cada 5 minutos.")
    caminho_digitado = st.text_input("Arquivo local / contingência", value=str(caminho_default))
    if st.button("🔄 Atualizar dados agora", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    radar, fonte_usada, aviso_fonte = carregar_radar_geral(fonte_escolhida, caminho_digitado)
except Exception as e:
    st.error(f"Erro ao carregar a planilha: {e}")
    st.stop()

status_col1, status_col2 = st.columns([3, 2])
with status_col1:
    st.markdown(f"<span class='source-pill'>Fonte carregada: <b>{fonte_usada}</b></span>", unsafe_allow_html=True)
with status_col2:
    st.caption(f"Última leitura: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}")

if aviso_fonte:
    st.warning(
        "O acesso direto ao SharePoint não foi autenticado pelo Python. "
        "O painel está usando o arquivo local como contingência. "
        f"Detalhe: {aviso_fonte}"
    )

# =========================================================
# FILTROS - NA LATERAL
# =========================================================
st.subheader("Radar")
st.caption("Base: aba Geral | somente unidades com Status = PREVISTO.")

marcas = sorted(radar["Marca"].dropna().unique().tolist())
tipos = sorted(radar["Tipo de Projeto"].dropna().unique().tolist())
riscos_existentes = [r for r in RISCO_ORDER if r in set(radar["Risco"])]

with st.sidebar:
    st.divider()
    st.header("Filtros do Radar")
    st.markdown("<div class='sidebar-note'>Deixe um filtro sem seleção para exibir <b>todas</b> as opções daquele campo.</div>", unsafe_allow_html=True)

    marcas_sel = st.multiselect(
        "Marca",
        marcas,
        default=[],
        placeholder="Todas as marcas",
        key="filtro_marca_radar",
    )
    tipos_sel = st.multiselect(
        "Tipo de projeto",
        tipos,
        default=[],
        placeholder="Todos os tipos",
        key="filtro_tipo_radar",
    )
    riscos_sel = st.multiselect(
        "Farol de risco",
        riscos_existentes,
        default=[],
        placeholder="Todos os riscos",
        key="filtro_risco_radar",
    )

    busca = st.text_input(
        "Pesquisar unidade ou comentário",
        placeholder="Ex.: Norte Shopping, atraso, Cielo...",
        key="filtro_busca_radar",
    )

    st.caption("Período previsto")
    data_de = st.date_input(
        "A partir de",
        value=None,
        format="DD/MM/YYYY",
        key="filtro_data_de_radar",
    )
    data_ate = st.date_input(
        "Até",
        value=None,
        format="DD/MM/YYYY",
        key="filtro_data_ate_radar",
    )

    if st.button("Limpar filtros", use_container_width=True, key="limpar_filtros_radar"):
        for chave in [
            "filtro_marca_radar", "filtro_tipo_radar", "filtro_risco_radar",
            "filtro_busca_radar", "filtro_data_de_radar", "filtro_data_ate_radar"
        ]:
            if chave in st.session_state:
                del st.session_state[chave]
        st.rerun()

# Começa sempre com toda a base. Cada filtro só é aplicado quando há seleção.
filtrado = radar.copy()

if marcas_sel:
    filtrado = filtrado[filtrado["Marca"].isin(marcas_sel)].copy()

if tipos_sel:
    filtrado = filtrado[filtrado["Tipo de Projeto"].isin(tipos_sel)].copy()

if riscos_sel:
    filtrado = filtrado[filtrado["Risco"].isin(riscos_sel)].copy()

if busca.strip():
    termo = normalizar_texto(busca)
    alvo = (
        filtrado["Restaurante"].fillna("").astype(str)
        + " "
        + filtrado["Comentários"].fillna("").astype(str)
    ).map(normalizar_texto)
    filtrado = filtrado[alvo.str.contains(termo, regex=False, na=False)].copy()

if data_de is not None:
    filtrado = filtrado[filtrado["Data Prevista"].dt.date >= data_de].copy()
if data_ate is not None:
    filtrado = filtrado[filtrado["Data Prevista"].dt.date <= data_ate].copy()

# =========================================================
# RESUMO
# =========================================================
total = len(filtrado)
alto = int((filtrado["Risco"] == "🔴 ALTO RISCO").sum())
medio = int((filtrado["Risco"] == "🟡 MÉDIO RISCO").sum())
baixo = int((filtrado["Risco"] == "🟢 BAIXO RISCO").sum())

m1, m2, m3, m4 = st.columns(4)
m1.metric("Unidades previstas", total)
m2.metric("Alto risco", alto)
m3.metric("Médio risco", medio)
m4.metric("Baixo risco", baixo)

st.divider()

# =========================================================
# RADAR POR MARCA E TIPO DE PROJETO
# =========================================================
if filtrado.empty:
    st.info("Nenhuma unidade encontrada com os filtros selecionados.")
else:
    prioridade_risco = {r: i for i, r in enumerate(RISCO_ORDER)}
    prioridade_tipo = {t: i for i, t in enumerate(TIPO_ORDER)}

    filtrado["_risco_ord"] = filtrado["Risco"].map(prioridade_risco).fillna(99)
    filtrado["_tipo_ord"] = filtrado["Tipo de Projeto"].map(prioridade_tipo).fillna(99)
    filtrado["_data_ord"] = filtrado["Data Prevista"].fillna(pd.Timestamp.max)
    filtrado = filtrado.sort_values(["Marca", "_tipo_ord", "_risco_ord", "_data_ord", "Restaurante"])

    marcas_visiveis = filtrado["Marca"].drop_duplicates().tolist()
    abrir_marca = len(marcas_visiveis) == 1

    for marca in marcas_visiveis:
        bloco_marca = filtrado[filtrado["Marca"] == marca].copy()
        qtd_alto = int((bloco_marca["Risco"] == "🔴 ALTO RISCO").sum())
        qtd_medio = int((bloco_marca["Risco"] == "🟡 MÉDIO RISCO").sum())
        qtd_baixo = int((bloco_marca["Risco"] == "🟢 BAIXO RISCO").sum())
        qtd_sem_risco = int((bloco_marca["Risco"] == "⚪ SEM RISCO INFORMADO").sum())

        titulo = f"{marca}  •  {len(bloco_marca)} unidade(s)"
        if qtd_alto:
            titulo += f"  •  🔴 {qtd_alto}"
        if qtd_medio:
            titulo += f"  •  🟡 {qtd_medio}"
        if qtd_baixo:
            titulo += f"  •  🟢 {qtd_baixo}"
        if qtd_sem_risco:
            titulo += f"  •  ⚪ {qtd_sem_risco}"

        with st.expander(titulo, expanded=abrir_marca):
            tipos_presentes = bloco_marca["Tipo de Projeto"].drop_duplicates().tolist()
            tipos_presentes = sorted(
                tipos_presentes,
                key=lambda x: prioridade_tipo.get(x, 99),
            )

            for tipo in tipos_presentes:
                bloco_tipo = bloco_marca[bloco_marca["Tipo de Projeto"] == tipo].copy()
                cor_tipo = TIPO_COLORS.get(tipo, GRAY)

                st.markdown(
                    f"""
                    <div class="tipo-section" style="border-top:4px solid {cor_tipo};">
                        <div class="tipo-header">
                            <div class="tipo-title">{html.escape(str(tipo))}</div>
                            <div class="tipo-count">{len(bloco_tipo)} unidade(s)</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                for risco in RISCO_ORDER:
                    bloco_risco = bloco_tipo[bloco_tipo["Risco"] == risco]
                    if bloco_risco.empty:
                        continue

                    st.markdown(f"**{risco} · {len(bloco_risco)} unidade(s)**")
                    for _, row in bloco_risco.iterrows():
                        st.markdown(card_unidade(row), unsafe_allow_html=True)

st.caption(
    "Colunas utilizadas da aba Geral: J = Marca | B = Data prevista | L = Tipo de projeto | "
    "K = Restaurante | U = Comentários | V = Farol riscos. A coluna A é usada apenas para manter Status = PREVISTO."
)

