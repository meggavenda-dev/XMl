
# =========================================================
# app.py — TISS XML + Conciliação + Auditoria (desativado) + Analytics
# =========================================================
from __future__ import annotations

import io
import os
import re
import json
from pathlib import Path
from typing import List, Dict, Optional, Union, IO, Tuple
from decimal import Decimal
from datetime import datetime
import xml.etree.ElementTree as ET
import unicodedata

import pandas as pd
import streamlit as st

# =========================================================
# Configuração da página (UI)
# =========================================================
st.set_page_config(
    page_title="TISS • Conciliação & Analytics (Auditoria desativada)",
    layout="wide"
)
st.title("TISS — Itens por Guia (XML) + Conciliação com Demonstrativo + Analytics")
st.caption("Lê XML TISS (Consulta / SADT), concilia com Demonstrativo itemizado (AMHP), gera rankings e analytics — sem editor de XML. Auditoria mantida no código, porém desativada.")

# =========================================================
# Helpers gerais
# =========================================================
ANS_NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}
DEC_ZERO = Decimal('0')

def dec(txt: Optional[str]) -> Decimal:
    if txt is None:
        return DEC_ZERO
    s = str(txt).strip().replace(',', '.')
    return Decimal(s) if s else DEC_ZERO

def tx(el: Optional[ET.Element]) -> str:
    return (el.text or '').strip() if (el is not None and el.text) else ''

def f_currency(v: Union[int, float, Decimal, str]) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    neg = v < 0
    v = abs(v)
    inteiro = int(v)
    cent = int(round((v - inteiro) * 100))
    s = f"R$ {inteiro:,}".replace(",", ".") + f",{cent:02d}"
    return f"-{s}" if neg else s

def apply_currency(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    d = df.copy()
    for c in cols:
        if c in d.columns:
            d[c] = d[c].apply(f_currency)
    return d

def parse_date_flex(s: str) -> Optional[datetime]:
    if s is None or not isinstance(s, str):
        return None
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None

def normalize_code(s: str, strip_zeros: bool = False) -> str:
    if s is None:
        return ""
    s2 = re.sub(r'[\.\-_/ \t]', '', str(s)).strip()
    return s2.lstrip('0') if strip_zeros else s2

def _normtxt(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.lower().strip()
    return re.sub(r"\s+", " ", s)

# Persistência de mapeamento (JSON)
MAP_FILE = "demo_mappings.json"

def load_demo_mappings() -> dict:
    if os.path.exists(MAP_FILE):
        try:
            with open(MAP_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_demo_mappings(mappings: dict):
    try:
        with open(MAP_FILE, "w", encoding="utf-8") as f:
            json.dump(mappings, f, indent=2, ensure_ascii=False)
    except Exception as e:
        st.error(f"Erro ao salvar mapeamentos: {e}")

# Carrega mapeamentos persistidos na inicialização
if "demo_mappings" not in st.session_state:
    st.session_state["demo_mappings"] = load_demo_mappings()

# Cache
@st.cache_data(show_spinner=False)
def _cached_read_excel(file, sheet_name=0) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")

@st.cache_data(show_spinner=False)
def _cached_xml_bytes(b: bytes) -> List[Dict]:
    # Apenas para cachear parsing; será chamado com bytes do upload
    from io import BytesIO
    return parse_itens_tiss_xml(BytesIO(b))

# =========================================================
# PARTE 2 — XML TISS → Itens por guia
# =========================================================

def _get_numero_lote(root: ET.Element) -> str:
    el = root.find('.//ans:prestadorParaOperadora/ans:loteGuias/ans:numeroLote', ANS_NS)
    if el is not None and tx(el):
        return tx(el)
    el = root.find('.//ans:prestadorParaOperadora/ans:recursoGlosa/ans:guiaRecursoGlosa/ans:numeroLote', ANS_NS)
    if el is not None and tx(el):
        return tx(el)
    return ""

def _itens_consulta(guia: ET.Element) -> List[Dict]:
    proc = guia.find('.//ans:procedimento', ANS_NS)
    codigo_tabela = tx(proc.find('ans:codigoTabela', ANS_NS)) if proc is not None else ''
    codigo_proc   = tx(proc.find('ans:codigoProcedimento', ANS_NS)) if proc is not None else ''
    descricao     = tx(proc.find('ans:descricaoProcedimento', ANS_NS)) if proc is not None else ''
    valor         = dec(tx(proc.find('ans:valorProcedimento', ANS_NS))) if proc is not None else DEC_ZERO
    return [{
        'tipo_item': 'procedimento',
        'identificadorDespesa': '',
        'codigo_tabela': codigo_tabela,
        'codigo_procedimento': codigo_proc,
        'descricao_procedimento': descricao,
        'quantidade': Decimal('1'),
        'valor_unitario': valor,
        'valor_total': valor
    }]

def _itens_sadt(guia: ET.Element) -> List[Dict]:
    out = []
    for it in guia.findall('.//ans:procedimentosExecutados/ans:procedimentoExecutado', ANS_NS):
        proc = it.find('ans:procedimento', ANS_NS)
        codigo_tabela = tx(proc.find('ans:codigoTabela', ANS_NS)) if proc is not None else ''
        codigo_proc   = tx(proc.find('ans:codigoProcedimento', ANS_NS)) if proc is not None else ''
        descricao     = tx(proc.find('ans:descricaoProcedimento', ANS_NS)) if proc is not None else ''
        qtd  = dec(tx(it.find('ans:quantidadeExecutada', ANS_NS)))
        vuni = dec(tx(it.find('ans:valorUnitario', ANS_NS)))
        vtot = dec(tx(it.find('ans:valorTotal', ANS_NS)))
        if vtot == DEC_ZERO and (vuni > DEC_ZERO and qtd > DEC_ZERO):
            vtot = vuni * qtd
        out.append({
            'tipo_item': 'procedimento',
            'identificadorDespesa': '',
            'codigo_tabela': codigo_tabela,
            'codigo_procedimento': codigo_proc,
            'descricao_procedimento': descricao,
            'quantidade': qtd if qtd > DEC_ZERO else Decimal('1'),
            'valor_unitario': vuni if vuni > DEC_ZERO else vtot,
            'valor_total': vtot,
        })
    for desp in guia.findall('.//ans:outrasDespesas/ans:despesa', ANS_NS):
        ident = tx(desp.find('ans:identificadorDespesa', ANS_NS))
        sv = desp.find('ans:servicosExecutados', ANS_NS)
        codigo_tabela = tx(sv.find('ans:codigoTabela', ANS_NS)) if sv is not None else ''
        codigo_proc   = tx(sv.find('ans:codigoProcedimento', ANS_NS)) if sv is not None else ''
        descricao     = tx(sv.find('ans:descricaoProcedimento', ANS_NS)) if sv is not None else ''
        qtd  = dec(tx(sv.find('ans:quantidadeExecutada', ANS_NS))) if sv is not None else DEC_ZERO
        vuni = dec(tx(sv.find('ans:valorUnitario', ANS_NS)))      if sv is not None else DEC_ZERO
        vtot = dec(tx(sv.find('ans:valorTotal', ANS_NS)))         if sv is not None else DEC_ZERO
        if vtot == DEC_ZERO and (vuni > DEC_ZERO and qtd > DEC_ZERO):
            vtot = vuni * qtd
        out.append({
            'tipo_item': 'outra_despesa',
            'identificadorDespesa': ident,
            'codigo_tabela': codigo_tabela,
            'codigo_procedimento': codigo_proc,
            'descricao_procedimento': descricao,
            'quantidade': qtd if qtd > DEC_ZERO else Decimal('1'),
            'valor_unitario': vuni if vuni > DEC_ZERO else vtot,
            'valor_total': vtot,
        })
    return out

def parse_itens_tiss_xml(source: Union[str, Path, IO[bytes]]) -> List[Dict]:
    if hasattr(source, 'read'):
        if hasattr(source, 'seek'):
            source.seek(0)
        root = ET.parse(source).getroot()
        nome = getattr(source, "name", "upload.xml")
    else:
        p = Path(source)
        root = ET.parse(p).getroot()
        nome = p.name

    numero_lote = _get_numero_lote(root)
    out: List[Dict] = []

    # CONSULTA
    for guia in root.findall('.//ans:guiaConsulta', ANS_NS):
        cab = guia.find('ans:cabecalhoGuia', ANS_NS)
        numero_guia_prest = tx(guia.find('ans:numeroGuiaPrestador', ANS_NS))
        paciente = tx(guia.find('.//ans:dadosBeneficiario/ans:nomeBeneficiario', ANS_NS))
        medico   = tx(guia.find('.//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional', ANS_NS))
        data_atd = tx(guia.find('.//ans:dataAtendimento', ANS_NS))
        for it in _itens_consulta(guia):
            it.update({
                'arquivo': nome,
                'numero_lote': numero_lote,
                'tipo_guia': 'CONSULTA',
                'numeroGuiaPrestador': numero_guia_prest,
                'numeroGuiaOperadora': '',
                'paciente': paciente,
                'medico': medico,
                'data_atendimento': data_atd,
            })
            out.append(it)

    # --- PARTE 2: XML TISS (SADT) ---
    for guia in root.findall('.//ans:guiaSP-SADT', ANS_NS):
        cab = guia.find('ans:cabecalhoGuia', ANS_NS)
        
        # Tenta buscar a guia primeiro na raiz (como está no seu arquivo)
        numero_guia_prest = tx(guia.find('ans:numeroGuiaPrestador', ANS_NS))
        
        # Se não encontrou, tenta dentro do cabeçalho
        if not numero_guia_prest:
            if cab is not None:
                numero_guia_prest = tx(cab.find('ans:numeroGuiaPrestador', ANS_NS))
        
        numero_guia_oper = tx(cab.find('ans:numeroGuiaOperadora', ANS_NS)) if cab is not None else ''
        paciente = tx(guia.find('.//ans:dadosBeneficiario/ans:nomeBeneficiario', ANS_NS))
        medico   = tx(guia.find('.//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional', ANS_NS))
        data_atd = tx(guia.find('.//ans:dataAtendimento', ANS_NS))
        
        for it in _itens_sadt(guia):
            it.update({
                'arquivo': nome,
                'numero_lote': numero_lote,
                'tipo_guia': 'SADT',
                'numeroGuiaPrestador': numero_guia_prest,
                'numeroGuiaOperadora': numero_guia_oper,
                'paciente': paciente,
                'medico': medico,
                'data_atendimento': data_atd,
            })
            out.append(it)

    return out

# =========================================================
# PARTE 3 — Demonstrativo (.xlsx)
#  - Leitor AMHP automático (sem wizard)
#  - Persistência de mapeamentos (JSON)
#  - Wizard apenas quando necessário
# =========================================================

def tratar_codigo_glosa(df: pd.DataFrame) -> pd.DataFrame:
    if "Código Glosa" not in df.columns:
        return df
    gl = df["Código Glosa"].astype(str).fillna("")
    df["motivo_glosa_codigo"]    = gl.str.extract(r"^(\d+)")
    df["motivo_glosa_descricao"] = gl.str.extract(r"^\s*\d+\s*-\s*(.*)$")
    df["motivo_glosa_codigo"]    = df["motivo_glosa_codigo"].fillna("").str.strip()
    df["motivo_glosa_descricao"] = df["motivo_glosa_descricao"].fillna("").str.strip()
    return df




def ler_demo_amhp_fixado(path, strip_zeros_codes: bool = False) -> pd.DataFrame:
    # 1) Lê o arquivo bruto para localizar o cabeçalho
    # Se for CSV (como o detectado), usa read_csv; se for Excel, read_excel
    try:
        df_raw = pd.read_excel(path, header=None, engine="openpyxl")
    except:
        df_raw = pd.read_csv(path, header=None)

    # 2) Localiza a linha do cabeçalho (onde está a coluna CPF/CNPJ)
    header_row = None
    for i in range(min(20, len(df_raw))):
        row_values = df_raw.iloc[i].astype(str).tolist()
        if any("CPF/CNPJ" in str(val).upper() for val in row_values):
            header_row = i
            break
    
    if header_row is None:
        raise ValueError("Não foi possível localizar a linha de cabeçalho 'CPF/CNPJ' no demonstrativo.")

    # 3) Lê novamente a partir do cabeçalho correto
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = df_raw.iloc[header_row]
    
    # Remove colunas sem nome (Unnamed)
    df = df.loc[:, df.columns.notna()]

    # 4) Renomeia para o padrão interno do seu código
    ren = {
        "Guia": "numeroGuiaPrestador",
        "Cod. Procedimento": "codigo_procedimento",
        "Descrição": "descricao_procedimento",
        "Valor Apresentado": "valor_apresentado",
        "Valor Apurado": "valor_pago",
        "Valor Glosa": "valor_glosa",
        "Quant. Exec.": "quantidade_apresentada",
        "Código Glosa": "codigo_glosa_bruto", # Para processar depois
    }
    df = df.rename(columns=ren)

    # 5) Limpeza Crítica: Guia e Código
    def clean_guia(val):
        s = str(val).strip().split('.')[0] # Remove .0
        return s.lstrip('0') # Remove zeros à esquerda para alinhar com XML

    # Limpeza da Guia para evitar que o Pandas leia como 8524664.0
    df["numeroGuiaPrestador"] = (
        df["numeroGuiaPrestador"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.lstrip("0")
    )
    df["codigo_procedimento"] = df["codigo_procedimento"].astype(str).str.strip()
    
    # Normalização de códigos (procedimentos e materiais)
    df["codigo_procedimento_norm"] = df["codigo_procedimento"].map(
        lambda s: normalize_code(s, strip_zeros=strip_zeros_codes)
    )

    # 6) Conversão Numérica
    for c in ["valor_apresentado", "valor_pago", "valor_glosa", "quantidade_apresentada"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', '.'), errors="coerce").fillna(0)

    # 7) Criação das Chaves de Conciliação
    df["numeroGuiaOperadora"] = df["numeroGuiaPrestador"]
    df["chave_prest"] = df["numeroGuiaPrestador"] + "__" + df["codigo_procedimento_norm"]
    df["chave_oper"] = df["chave_prest"]

    # 8) Tratamento da Glosa (separar código de texto)
    if "codigo_glosa_bruto" in df.columns:
        df["motivo_glosa_codigo"] = df["codigo_glosa_bruto"].astype(str).str.extract(r"^(\d+)")
        df["motivo_glosa_descricao"] = df["codigo_glosa_bruto"].astype(str).str.extract(r"^\d+\s*-\s*(.*)")
        df["motivo_glosa_codigo"] = df["motivo_glosa_codigo"].fillna("").str.strip()
        df["motivo_glosa_descricao"] = df["motivo_glosa_descricao"].fillna("").str.strip()

    return df.reset_index(drop=True)


# Auto-detecção genérica (fallback)
_COLMAPS = {
    "lote": [r"\blote\b"],
    "competencia": [r"compet|m[eê]s|refer"],
    "guia_prest": [r"\bguia\b"],
    "guia_oper": [r"\bguia\b"],
    "cod_proc": [r"cod.*proced|proced.*cod|tuss"],
    "desc_proc": [r"descr"],
    "qtd_apres": [r"quant|qtd"],
    "qtd_paga": [r"quant|qtd"],
    "val_apres": [r"apres|cobrado"],
    "val_glosa": [r"glosa"],
    "val_pago": [r"pago|liberado|apurado"],
    "motivo_cod": [r"glosa"],
    "motivo_desc": [r"glosa"],
}

def _match_col(cols, pats):
    norm = {c: _normtxt(c) for c in cols}
    for c, cn in norm.items():
        if all(re.search(p, cn) for p in pats):
            return c
    return None

def _apply_manual_map(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    def pick(k):
        c = mapping.get(k)
        if not c or c == "(não usar)" or c not in df.columns:
            return None
        return df[c]
    out = pd.DataFrame({
        "numero_lote": pick("lote"),
        "competencia": pick("competencia"),
        "numeroGuiaPrestador": pick("guia_prest"),
        "numeroGuiaOperadora": pick("guia_oper"),
        "codigo_procedimento": pick("cod_proc"),
        "descricao_procedimento": pick("desc_proc"),
        "quantidade_apresentada": pd.to_numeric(pick("qtd_apres"), errors="coerce") if pick("qtd_apres") is not None else 0,
        "quantidade_paga": pd.to_numeric(pick("qtd_paga"), errors="coerce") if pick("qtd_paga") is not None else 0,
        "valor_apresentado": pd.to_numeric(pick("val_apres"), errors="coerce") if pick("val_apres") is not None else 0,
        "valor_glosa": pd.to_numeric(pick("val_glosa"), errors="coerce") if pick("val_glosa") is not None else 0,
        "valor_pago": pd.to_numeric(pick("val_pago"), errors="coerce") if pick("val_pago") is not None else 0,
        "motivo_glosa_codigo": pick("motivo_cod"),
        "motivo_glosa_descricao": pick("motivo_desc"),
    })
    for c in ["numero_lote","numeroGuiaPrestador","numeroGuiaOperadora","codigo_procedimento"]:
        out[c] = out[c].astype(str).str.strip()
    for c in ["valor_apresentado","valor_glosa","valor_pago","quantidade_apresentada","quantidade_paga"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    out["codigo_procedimento_norm"] = out["codigo_procedimento"].map(lambda s: normalize_code(s))
    out["chave_prest"] = out["numeroGuiaPrestador"] + "__" + out["codigo_procedimento_norm"]
    out["chave_oper"]  = out["numeroGuiaOperadora"] + "__" + out["codigo_procedimento_norm"]
    return out

def _mapping_wizard_for_demo(uploaded_file):
    st.warning(f"Mapeamento manual pode ser necessário para: **{uploaded_file.name}**")
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Erro abrindo arquivo: {e}")
        return None
    sheet = st.selectbox(
        f"Aba (sheet) do demonstrativo {uploaded_file.name}",
        xls.sheet_names,
        key=f"map_sheet_{uploaded_file.name}"
    )
    df_raw = _cached_read_excel(uploaded_file, sheet)
    st.dataframe(df_raw.head(15), use_container_width=True)
    cols = [str(c) for c in df_raw.columns]
    fields = [
        ("lote", "Lote"), ("competencia", "Competência"),
        ("guia_prest", "Guia Prestador"), ("guia_oper", "Guia Operadora"),
        ("cod_proc", "Código Procedimento"), ("desc_proc", "Descrição Procedimento"),
        ("qtd_apres", "Quantidade Apresentada"), ("qtd_paga", "Quantidade Paga"),
        ("val_apres", "Valor Apresentado"), ("val_glosa", "Valor Glosa"), ("val_pago", "Valor Pago"),
        ("motivo_cod", "Código Glosa"), ("motivo_desc", "Descrição Motivo Glosa"),
    ]
    def _default(k):
        pats = _COLMAPS.get(k, [])
        for i, c in enumerate(cols):
            if any(re.search(p, _normtxt(c)) for p in pats):
                return i + 1
        return 0
    mapping = {}
    for k, label in fields:
        opt = ["(não usar)"] + cols
        sel = st.selectbox(label, opt, index=_default(k), key=f"{uploaded_file.name}_{k}")
        mapping[k] = None if sel == "(não usar)" else sel

    if st.button(f"Salvar mapeamento de {uploaded_file.name}", type="primary"):
        st.session_state["demo_mappings"][uploaded_file.name] = {
            "sheet": sheet,
            "columns": mapping
        }
        save_demo_mappings(st.session_state["demo_mappings"])
        try:
            df = _apply_manual_map(df_raw, mapping)
            df = tratar_codigo_glosa(df)
            st.success("Mapeamento salvo com sucesso!")
            return df
        except Exception as e:
            st.error(f"Erro aplicando mapeamento: {e}")
            return None
    return None

def build_demo_df(demo_files, strip_zeros_codes=False) -> pd.DataFrame:
    if not demo_files:
        return pd.DataFrame()
    parts: List[pd.DataFrame] = []
    st.session_state.setdefault("demo_mappings", load_demo_mappings())
    for f in demo_files:
        fname = f.name
        # 1) leitor AMHP automático
        try:
            df_demo = ler_demo_amhp_fixado(f, strip_zeros_codes=strip_zeros_codes)
            parts.append(df_demo)
            continue
        except Exception:
            pass
        # 2) mapeamento persistido
        mapping_info = st.session_state["demo_mappings"].get(fname)
        if mapping_info:
            try:
                df_demo = ler_demo_amhp_fixado(f, strip_zeros_codes=strip_zeros_codes)
            except:
                df_raw = _cached_read_excel(f, mapping_info["sheet"])
                df_demo = _apply_manual_map(df_raw, mapping_info["columns"])
            df_demo = tratar_codigo_glosa(df_demo)
            parts.append(df_demo)
            continue
        # 3) auto-detecção suave
        try:
            xls = pd.ExcelFile(f, engine="openpyxl")
            sheet = xls.sheet_names[0]
            df_raw = _cached_read_excel(f, sheet)
            cols = [str(c) for c in df_raw.columns]
            pick = {k: _match_col(cols, v) for k, v in _COLMAPS.items()}
            if pick.get("cod_proc"):
                df_demo = _apply_manual_map(df_raw, pick)
                df_demo = tratar_codigo_glosa(df_demo)
                parts.append(df_demo)
                continue
        except:
            pass
        # 4) wizard
        with st.expander(f"⚙️ Mapear manualmente: {fname}", expanded=True):
            df_manual = _mapping_wizard_for_demo(f)
            if df_manual is not None:
                parts.append(df_manual)
            else:
                st.error(f"Não foi possível mapear o demonstrativo '{fname}'.")
    if parts:
        return pd.concat(parts, ignore_index=True)
    return pd.DataFrame()

# =========================================================
# PARTE 4 — Conciliação (XML × Demonstrativo) + Analytics
# =========================================================

def build_xml_df(xml_files, strip_zeros_codes: bool = False) -> pd.DataFrame:
    linhas: List[Dict] = []
    for f in xml_files:
        if hasattr(f, 'seek'):
            f.seek(0)
        try:
            if hasattr(f, 'read'):
                bts = f.read()
                linhas.extend(_cached_xml_bytes(bts))
            else:
                linhas.extend(parse_itens_tiss_xml(f))
        except Exception as e:
            linhas.append({'arquivo': getattr(f, 'name', 'upload.xml'), 'erro': str(e)})

    df = pd.DataFrame(linhas)
    if df.empty:
        return df

    for c in ['quantidade', 'valor_unitario', 'valor_total']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
    df['codigo_procedimento_norm'] = df['codigo_procedimento'].astype(str).map(
        lambda s: normalize_code(s, strip_zeros=strip_zeros_codes)
    )
    df['chave_prest'] = (df['numeroGuiaPrestador'].fillna('').astype(str).str.strip()
                         + '__' + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip())
    df['chave_oper'] = (df['numeroGuiaOperadora'].fillna('').astype(str).str.strip()
                        + '__' + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip())
    return df

# helper para padronizar nomes do "lado XML" após merges com sufixos
_XML_CORE_COLS = [
    'arquivo', 'numero_lote', 'tipo_guia',
    'numeroGuiaPrestador', 'numeroGuiaOperadora',
    'paciente', 'medico', 'data_atendimento',
    'tipo_item', 'identificadorDespesa',
    'codigo_tabela', 'codigo_procedimento', 'codigo_procedimento_norm',
    'descricao_procedimento',
    'quantidade', 'valor_unitario', 'valor_total',
    'chave_oper', 'chave_prest',
]

def _alias_xml_cols(df: pd.DataFrame, cols: List[str] = None, prefer_suffix: str = '_xml') -> pd.DataFrame:
    if cols is None:
        cols = _XML_CORE_COLS
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            cand = f'{c}{prefer_suffix}'
            if cand in out.columns:
                out[c] = out[cand]
    return out

def conciliar_itens(
    df_xml: pd.DataFrame,
    df_demo: pd.DataFrame,
    tolerance_valor: float = 0.02,
    fallback_por_descricao: bool = False,
) -> Dict[str, pd.DataFrame]:

    # 1ª tentativa — pela chave do Prestador
    m1 = df_xml.merge(df_demo, on="chave_prest", how="left", suffixes=("_xml", "_demo"))
    m1 = _alias_xml_cols(m1)
    m1["matched_on"] = m1["valor_apresentado"].notna().map({True: "prestador", False: ""})

    # Registros ainda sem match (sempre do lado do XML)
    restante = m1[m1["matched_on"] == ""].copy()
    restante = _alias_xml_cols(restante)

    # 2ª tentativa — pela chave da Operadora
    cols_for_second_join = [c for c in _XML_CORE_COLS if c in restante.columns]
    still_xml = restante[cols_for_second_join].copy()

    m2 = still_xml.merge(df_demo, on="chave_oper", how="left", suffixes=("_xml", "_demo"))
    m2 = _alias_xml_cols(m2)
    m2["matched_on"] = m2["valor_apresentado"].notna().map({True: "operadora", False: ""})

    conc = pd.concat([m1[m1["matched_on"] != ""], m2[m2["matched_on"] != ""]], ignore_index=True)

    # Fallback opcional por descrição + valor (ainda partindo do XML)
    fallback_matches = pd.DataFrame()
    if fallback_por_descricao:
        rem1 = m1[m1["matched_on"] == ""].copy()
        rem2 = m2[m2["matched_on"] == ""].copy()
        rem1 = _alias_xml_cols(rem1)
        rem2 = _alias_xml_cols(rem2)
        rem_xml = pd.concat([rem1, rem2], ignore_index=True)
        if not rem_xml.empty:
            rem_xml["guia_join"] = rem_xml.apply(
                lambda r: r["numeroGuiaPrestador"] if str(r.get("numeroGuiaPrestador","")).strip()
                else str(r.get("numeroGuiaOperadora","")).strip(), axis=1
            )
            df_demo2 = df_demo.copy()
            df_demo2["guia_join"] = df_demo2.apply(
                lambda r: r["numeroGuiaPrestador"] if str(r.get("numeroGuiaPrestador","")).strip()
                else str(r.get("numeroGuiaOperadora","")).strip(), axis=1
            )
            if "descricao_procedimento" in rem_xml.columns and "descricao_procedimento" in df_demo2.columns:
                tmp = rem_xml.merge(df_demo2, on=["guia_join","descricao_procedimento"], how="left", suffixes=("_xml","_demo"))
                tol = float(tolerance_valor)
                keep = (tmp["valor_apresentado"].notna() & ((tmp["valor_total"] - tmp["valor_apresentado"]).abs() <= tol))
                fallback_matches = tmp[keep].copy()
                if not fallback_matches.empty:
                    fallback_matches["matched_on"] = "descricao+valor"
                    conc = pd.concat([conc, fallback_matches], ignore_index=True)

    # Apenas XML sem match (DEMO extra fica ignorado por construção)
    unmatch = pd.concat([
        m1[m1["matched_on"] == ""],
        m2[m2["matched_on"] == ""],
        fallback_matches[fallback_matches.get("matched_on","") == ""] if not fallback_matches.empty else pd.DataFrame()
    ], ignore_index=True)
    unmatch = _alias_xml_cols(unmatch)
    if not unmatch.empty:
        subset_cols = [c for c in ["arquivo","numero_lote","tipo_guia","numeroGuiaPrestador","codigo_procedimento","valor_total"] if c in unmatch.columns]
        if subset_cols:
            unmatch = unmatch.drop_duplicates(subset=subset_cols)

    if not conc.empty:
        conc = _alias_xml_cols(conc)
        conc["apresentado_diff"] = conc["valor_total"] - conc["valor_apresentado"]
        conc["glosa_pct"] = conc.apply(
            lambda r: (r["valor_glosa"]/r["valor_apresentado"]) if r.get("valor_apresentado",0) and r["valor_apresentado"]>0 else 0.0,
            axis=1
        )

    return {"conciliacao": conc, "nao_casados": unmatch}

# -----------------------------
# Analytics (derivados do conciliado)
# -----------------------------
def kpis_por_competencia(df_conc: pd.DataFrame) -> pd.DataFrame:
    """
    KPIs agora são calculados APENAS com base nos itens conciliados (df_conc),
    garantindo que itens presentes apenas no demonstrativo não afetem os resultados.
    """
    base = df_conc.copy()
    if base.empty:
        return base
    # 'competencia' vem do demonstrativo via merge; se não existir, cria vazia
    if 'competencia' not in base.columns and 'Competência' in base.columns:
        base['competencia'] = base['Competência'].astype(str)
    elif 'competencia' not in base.columns:
        base['competencia'] = ""

    grp = (base.groupby('competencia', dropna=False, as_index=False)
           .agg(valor_apresentado=('valor_apresentado','sum'),
                valor_pago=('valor_pago','sum'),
                valor_glosa=('valor_glosa','sum'))
          )
    grp['glosa_pct'] = grp.apply(
        lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1
    )
    return grp.sort_values('competencia')

def ranking_itens_glosa(df_conc: pd.DataFrame, min_apresentado: float = 500.0, topn: int = 20) -> Tuple[pd.DataFrame, pd.DataFrame]:
    base = df_conc.copy()
    if base.empty:
        return base, base
    grp = (base.groupby(['codigo_procedimento','descricao_procedimento'], dropna=False, as_index=False)
           .agg(valor_apresentado=('valor_apresentado','sum'),
                valor_glosa=('valor_glosa','sum'),
                valor_pago=('valor_pago','sum'),
                itens=('arquivo','count')))
    grp['glosa_pct'] = grp.apply(
        lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1
    )
    top_valor = grp.sort_values('valor_glosa', ascending=False).head(topn)
    top_pct   = grp[grp['valor_apresentado']>=min_apresentado].sort_values('glosa_pct', ascending=False).head(topn)
    return top_valor, top_pct

def motivos_glosa(df_conc: pd.DataFrame, competencia: Optional[str] = None) -> pd.DataFrame:
    base = df_conc.copy()
    if base.empty:
        return base
    if competencia and 'competencia' in base.columns:
        base = base[base['competencia'] == competencia]
    mot = (base.groupby(['motivo_glosa_codigo','motivo_glosa_descricao'], dropna=False, as_index=False)
           .agg(valor_glosa=('valor_glosa','sum'),
                valor_apresentado=('valor_apresentado','sum'),
                itens=('codigo_procedimento','count')))
    mot['glosa_pct'] = mot.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
    return mot.sort_values(['valor_glosa','glosa_pct'], ascending=[False, False])

def outliers_por_procedimento(df_conc: pd.DataFrame, k: float = 1.5) -> pd.DataFrame:
    base = df_conc[['codigo_procedimento','descricao_procedimento','valor_apresentado']].dropna().copy()
    if base.empty:
        return base
    stats = (base.groupby(['codigo_procedimento','descricao_procedimento'])
             .agg(p50=('valor_apresentado','median'),
                  q1=('valor_apresentado', lambda x: x.quantile(0.25)),
                  q3=('valor_apresentado', lambda x: x.quantile(0.75))))
    stats['iqr'] = stats['q3'] - stats['q1']
    base = base.merge(stats.reset_index(), on=['codigo_procedimento','descricao_procedimento'], how='left')
    base['is_outlier'] = (base['valor_apresentado'] > base['q3'] + k*base['iqr']) | (base['valor_apresentado'] < base['q1'] - k*base['iqr'])
    return base[base['is_outlier']].copy()

def simulador_glosa(df_conc: pd.DataFrame, ajustes: Dict[str, float]) -> pd.DataFrame:
    """
    ajustes: dict { '1705': 0.80 } → aplica fator 0.80 (redução de 20%) no valor_glosa dos itens com motivo 1705
    """
    sim = df_conc.copy()
    if sim.empty or 'motivo_glosa_codigo' not in sim.columns:
        return sim
    sim['valor_glosa_sim'] = sim['valor_glosa']
    for cod, fator in ajustes.items():
        mask = sim['motivo_glosa_codigo'].astype(str) == str(cod)
        sim.loc[mask, 'valor_glosa_sim'] = sim.loc[mask, 'valor_glosa'] * float(fator)
    sim['valor_glosa_sim'] = sim['valor_glosa_sim'].clip(lower=0)
    sim['valor_pago_sim'] = sim['valor_apresentado'] - sim['valor_glosa_sim']
    sim['valor_pago_sim'] = sim['valor_pago_sim'].clip(lower=0)
    sim['glosa_pct_sim'] = sim.apply(
        lambda r: (r['valor_glosa_sim']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1
    )
    return sim

# =========================================================
# PARTE 5 — Auditoria de Guias (duplicidade, retorno, indicadores)
# =========================================================
# auditoria (desativado): função mantida para uso futuro; não é chamada na interface nem na exportação.
def build_chave_guia(tipo: str, numeroGuiaPrestador: str, numeroGuiaOperadora: str) -> Optional[str]:
    tipo = (tipo or "").upper()
    if tipo not in ("CONSULTA", "SADT"):
        return None
    guia = (numeroGuiaPrestador or "").strip() or (numeroGuiaOperadora or "").strip()
    return guia if guia else None

def _parse_dt_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def auditar_guias(df_xml_itens: pd.DataFrame, prazo_retorno: int = 30) -> pd.DataFrame:
    if df_xml_itens is None or df_xml_itens.empty:
        return pd.DataFrame()
    req = ["arquivo","numero_lote","tipo_guia","numeroGuiaPrestador","numeroGuiaOperadora","paciente","medico","data_atendimento","valor_total"]
    for c in req:
        if c not in df_xml_itens.columns:
            df_xml_itens[c] = None

    df = df_xml_itens.copy()
    df["data_atendimento_dt"] = _parse_dt_series(df["data_atendimento"])

    agg = (df.groupby(["tipo_guia","numeroGuiaPrestador","numeroGuiaOperadora","paciente","medico"], dropna=False, as_index=False)
           .agg(arquivo=("arquivo", lambda x: sorted(set(str(a) for a in x if str(a).strip()))),
                numero_lote=("numero_lote", lambda x: sorted(set(str(a) for a in x if str(a).strip()))),
                data_atendimento=("data_atendimento_dt","min"),
                itens_na_guia=("valor_total","count"),
                valor_total_xml=("valor_total","sum")))
    agg["arquivo(s)"] = agg["arquivo"].apply(lambda L: ", ".join(L))
    agg["numero_lote(s)"] = agg["numero_lote"].apply(lambda L: ", ".join(L))
    agg.drop(columns=["arquivo","numero_lote"], inplace=True)

    agg["chave_guia"] = agg.apply(lambda r: build_chave_guia(r["tipo_guia"], r["numeroGuiaPrestador"], r["numeroGuiaOperadora"]), axis=1)
    agg["duplicada"] = False
    agg["arquivos_duplicados"] = ""
    agg["lotes_duplicados"] = ""
    agg["retorno_no_periodo"] = False
    agg["retorno_ref"] = ""
    agg["status_auditoria"] = ""

    # ... lógica mantida, porém a função não é utilizada (desativada)
    return agg

# =========================================================
# PARTE 6 — Interface (Uploads, Parâmetros, Processamento, Analytics, Export)
# =========================================================

# -----------------------------
# Sidebar de parâmetros
# -----------------------------
with st.sidebar:
    st.header("Parâmetros")
    prazo_retorno = st.number_input("Prazo de retorno (dias) — (auditoria desativada)", min_value=0, value=30, step=1)
    tolerance_valor = st.number_input("Tolerância p/ fallback por descrição (R$)", min_value=0.00, value=0.02, step=0.01, format="%.2f")
    fallback_desc = st.toggle("Fallback por descrição + valor (quando código não casar)", value=False)
    strip_zeros_codes = st.toggle("Normalizar códigos removendo zeros à esquerda", value=True)

# -----------------------------
# Upload dos arquivos
# -----------------------------
st.subheader("📤 Upload de arquivos")
xml_files = st.file_uploader("XML TISS (um ou mais):", type=['xml'], accept_multiple_files=True)
demo_files = st.file_uploader("Demonstrativos de Pagamento (.xlsx) — itemizado:", type=['xlsx'], accept_multiple_files=True)

# --------------------------------------------------------------
# PROCESSAMENTO DO DEMONSTRATIVO (SEMPRE) — para permitir wizard
# (Não exibimos a tabela completa do demonstrativo para evitar mostrar itens extras)
# --------------------------------------------------------------
df_demo = build_demo_df(demo_files or [], strip_zeros_codes=strip_zeros_codes)

if not df_demo.empty:
    st.info("Demonstrativo carregado e mapeado. A conciliação considerará **somente** os itens presentes nos XMLs. Itens presentes apenas no demonstrativo serão **ignorados**.")
else:
    if demo_files:
        st.info("Carregue um Demonstrativo válido ou conclua o mapeamento manual.")

st.markdown("---")
if st.button("🚀 Processar Conciliação & Analytics", type="primary"):

    # 1) XML
    df_xml = build_xml_df(xml_files or [], strip_zeros_codes=strip_zeros_codes)
    if df_xml.empty:
        st.warning("Nenhum item extraído do(s) XML(s). Verifique os arquivos.")
        st.stop()

    st.subheader("📄 Itens extraídos dos XML (Consulta / SADT)")
    st.dataframe(apply_currency(df_xml, ['valor_unitario','valor_total']), use_container_width=True, height=360)

    if df_demo.empty:
        st.warning("Nenhum demonstrativo válido para conciliar.")
        st.stop()

    # 2) Conciliação (somente a partir do XML)
    result = conciliar_itens(
        df_xml=df_xml,
        df_demo=df_demo,
        tolerance_valor=float(tolerance_valor),
        fallback_por_descricao=fallback_desc
    )
    conc = result["conciliacao"]
    unmatch = result["nao_casados"]

    st.subheader("🔗 Conciliação Item a Item (XML × Demonstrativo)")
    conc_disp = apply_currency(
        conc.copy(),
        ['valor_unitario','valor_total','valor_apresentado','valor_glosa','valor_pago','apresentado_diff']
    )
    st.dataframe(conc_disp, use_container_width=True, height=460)

    c1, c2 = st.columns(2)
    c1.metric("Itens conciliados", len(conc))
    c2.metric("Itens não conciliados (somente XML)", len(unmatch))

    if not unmatch.empty:
        st.subheader("❗ Itens (do XML) não conciliados")
        st.dataframe(apply_currency(unmatch.copy(), ['valor_unitario','valor_total']), use_container_width=True, height=300)
        st.download_button("Baixar Não Conciliados (CSV)", data=unmatch.to_csv(index=False).encode("utf-8"), file_name="nao_conciliados.csv", mime="text/csv")

    # 3) Analytics — APENAS COM BASE NO CONCILIADO
    st.markdown("---")
    st.subheader("📊 Analytics de Glosa (apenas itens conciliados)")

    # 3.1 KPI por Competência (a partir do conc)
    st.markdown("### 📈 Tendência por competência")
    kpi_comp = kpis_por_competencia(conc)
    st.dataframe(apply_currency(kpi_comp, ['valor_apresentado','valor_pago','valor_glosa']), use_container_width=True)
    try:
        st.line_chart(kpi_comp.set_index('competencia')[['valor_apresentado','valor_pago','valor_glosa']])
    except Exception:
        pass

    # 3.2 Top Itens Glosados (valor e %)
    st.markdown("### 🏆 TOP itens glosados (valor e %)")
    min_apres = st.number_input("Corte mínimo de Apresentado para ranking por % (R$)", min_value=0.0, value=500.0, step=50.0)
    top_valor, top_pct = ranking_itens_glosa(conc, min_apresentado=min_apres, topn=20)
    t1, t2 = st.columns(2)
    with t1:
        st.markdown("**Por valor de glosa (TOP 20)**")
        st.dataframe(apply_currency(top_valor, ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)
    with t2:
        st.markdown("**Por % de glosa (TOP 20)**")
        st.dataframe(apply_currency(top_pct, ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)

    # 3.3 Motivos de Glosa (filtro por competência)
    st.markdown("### 🧩 Motivos de glosa — análise")
    comp_opts = ['(todas)']
    if 'competencia' in conc.columns:
        comp_opts += sorted(conc['competencia'].dropna().astype(str).unique().tolist())
    comp_sel = st.selectbox("Filtrar por competência", comp_opts)
    motdf = motivos_glosa(conc, None if comp_sel=='(todas)' else comp_sel)
    st.dataframe(apply_currency(motdf, ['valor_glosa','valor_apresentado']), use_container_width=True)

    # 3.4 Médicos (filtro por competência)
    st.markdown("### 👩‍⚕️ Médicos — ranking por glosa")
    if 'competencia' in conc.columns:
        comp_med = st.selectbox("Competência (médicos)", 
        ['(todas)'] + sorted(conc['competencia'].dropna().astype(str).unique().tolist())
        )
        med_base = conc if comp_med == '(todas)' else conc[conc['competencia'] == comp_med]
    else:
        med_base = conc
    med_rank = (med_base.groupby(['medico'], dropna=False, as_index=False)
                .agg(valor_apresentado=('valor_apresentado','sum'),
                     valor_glosa=('valor_glosa','sum'),
                     valor_pago=('valor_pago','sum'),
                     itens=('arquivo','count')))
    med_rank['glosa_pct'] = med_rank.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
    st.dataframe(apply_currency(med_rank.sort_values(['glosa_pct','valor_glosa'], ascending=[False,False]), ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)

    # 3.5 Glosa por Tabela (22/19) — se existir no conciliado
    st.markdown("### 🧾 Glosa por Tabela (22/19)")
    if 'Tabela' in conc.columns:
        tab = (conc.groupby('Tabela', as_index=False)
               .agg(valor_apresentado=('valor_apresentado','sum'),
                    valor_glosa=('valor_glosa','sum'),
                    valor_pago=('valor_pago','sum')))
        tab['glosa_pct'] = tab.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
        st.dataframe(apply_currency(tab, ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)
    else:
        st.info("Coluna 'Tabela' não encontrada nos itens conciliados (opcional no demonstrativo).")

    # 3.6 Qualidade da Conciliação (origem)
    if 'matched_on' in conc.columns:
        st.markdown("### 🧪 Qualidade da conciliação (origem do match)")
        match_dist = conc['matched_on'].value_counts(dropna=False).rename_axis('origem').reset_index(name='itens')
        st.bar_chart(match_dist.set_index('origem'))
        st.dataframe(match_dist, use_container_width=True)

    # 3.7 Outliers em valor apresentado
    st.markdown("### 🚩 Outliers em valor apresentado (por procedimento)")
    out_df = outliers_por_procedimento(conc, k=1.5)
    if out_df.empty:
        st.info("Nenhum outlier identificado com o critério atual (IQR).")
    else:
        st.dataframe(out_df, use_container_width=True, height=280)
        st.download_button("Baixar Outliers (CSV)", data=out_df.to_csv(index=False).encode("utf-8"), file_name="outliers_valor_apresentado.csv", mime="text/csv")

    # 3.8 Simulador de Faturamento (what-if)
    st.markdown("### 🧮 Simulador de faturamento (what‑if por motivo de glosa)")
    motivos_disponiveis = sorted(conc['motivo_glosa_codigo'].dropna().astype(str).unique().tolist()) if 'motivo_glosa_codigo' in conc.columns else []
    if motivos_disponiveis:
        cols_sim = st.columns(min(4, max(1, len(motivos_disponiveis))))
        ajustes = {}
        for i, cod in enumerate(motivos_disponiveis):
            col = cols_sim[i % len(cols_sim)]
            with col:
                fator = st.slider(f"Motivo {cod} → fator (0–1)", 0.0, 1.0, 1.0, 0.05, help="Ex.: 0,8 reduz a glosa em 20% para esse motivo.")
                ajustes[cod] = fator
        sim = simulador_glosa(conc, ajustes)
        st.write("**Resumo do cenário simulado:**")
        res = (sim.agg(
            total_apres=('valor_apresentado','sum'),
            glosa=('valor_glosa','sum'),
            glosa_sim=('valor_glosa_sim','sum'),
            pago=('valor_pago','sum'),
            pago_sim=('valor_pago_sim','sum')
        ))
        st.json({k: f_currency(v) for k, v in res.to_dict().items()})
    else:
        st.info("Sem motivos de glosa identificados para simulação.")

    # 4) Auditoria por guia — DESATIVADA
    # ----------------------------------------------------------
    # auditoria (desativado): seção intencionalmente desabilitada.
    # df_aud = auditar_guias(df_xml, prazo_retorno=prazo_retorno)
    # st.markdown("---")
    # st.subheader("🔎 Auditoria por Guia (Duplicidade e Retorno)")
    # if df_aud.empty:
    #     st.info("Sem dados para auditoria.")
    # else:
    #     st.dataframe(df_aud, use_container_width=True, height=360)
    # ----------------------------------------------------------

    # 5) Exportação Excel consolidado (apenas itens a partir do XML e os conciliados)
    st.markdown("---")
    st.subheader("📥 Exportar Excel Consolidado")

    # Itens_Demo (somente os que casaram com XML)
    demo_cols_for_export = [c for c in [
        'numero_lote','competencia','numeroGuiaPrestador','numeroGuiaOperadora',
        'codigo_procedimento','descricao_procedimento',
        'quantidade_apresentada','valor_apresentado','valor_glosa','valor_pago',
        'motivo_glosa_codigo','motivo_glosa_descricao','Tabela'
    ] if c in conc.columns]
    itens_demo_match = pd.DataFrame()
    if demo_cols_for_export:
        itens_demo_match = conc[demo_cols_for_export].drop_duplicates().copy()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as wr:
        # Sempre origem XML
        df_xml.to_excel(wr, index=False, sheet_name='Itens_XML')

        # Apenas demonstrativo que teve MATCH com XML
        if not itens_demo_match.empty:
            itens_demo_match.to_excel(wr, index=False, sheet_name='Itens_Demo')

        # Conciliações e não casados
        conc.to_excel(wr, index=False, sheet_name='Conciliação')
        unmatch.to_excel(wr, index=False, sheet_name='Nao_Casados')

        # Motivos
        mot_x = motivos_glosa(conc, None)
        mot_x.to_excel(wr, index=False, sheet_name='Motivos_Glosa')

        # Procedimentos
        proc_x = (conc.groupby(['codigo_procedimento','descricao_procedimento'], dropna=False, as_index=False)
                  .agg(valor_apresentado=('valor_apresentado','sum'),
                       valor_glosa=('valor_glosa','sum'),
                       valor_pago=('valor_pago','sum'),
                       itens=('arquivo','count')))
        proc_x['glosa_pct'] = proc_x.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
        proc_x.to_excel(wr, index=False, sheet_name='Procedimentos_Glosa')

        # Médicos
        med_x = (conc.groupby(['medico'], dropna=False, as_index=False)
                 .agg(valor_apresentado=('valor_apresentado','sum'),
                      valor_glosa=('valor_glosa','sum'),
                      valor_pago=('valor_pago','sum'),
                      itens=('arquivo','count')))
        med_x['glosa_pct'] = med_x.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
        med_x.to_excel(wr, index=False, sheet_name='Medicos')

        # Lotes
        if 'numero_lote' in conc.columns:
            lot_x = (conc.groupby(['numero_lote'], dropna=False, as_index=False)
                     .agg(valor_apresentado=('valor_apresentado','sum'),
                          valor_glosa=('valor_glosa','sum'),
                          valor_pago=('valor_pago','sum'),
                          itens=('arquivo','count')))
            lot_x['glosa_pct'] = lot_x.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
            lot_x.to_excel(wr, index=False, sheet_name='Lotes')

        # KPIs por competência
        kpi_comp.to_excel(wr, index=False, sheet_name='KPIs_Competencia')

        # Top Itens
        top_valor.to_excel(wr, index=False, sheet_name='Top_Itens_Glosa_Valor')
        top_pct.to_excel(wr, index=False, sheet_name='Top_Itens_Glosa_Pct')

        # Qualidade conciliação
        if 'matched_on' in conc.columns:
            match_dist = conc['matched_on'].value_counts(dropna=False).rename_axis('origem').reset_index(name='itens')
            match_dist.to_excel(wr, index=False, sheet_name='Qualidade_Conciliacao')

        # Outliers
        if not out_df.empty:
            out_df.to_excel(wr, index=False, sheet_name='Outliers')

        # Auditoria — DESATIVADO (não exporta)
        # if not df_aud.empty:
        #     df_aud.to_excel(wr, index=False, sheet_name='Auditoria_Guias')

        # Ajustes de largura e congelamento
        for name in wr.sheets:
            ws = wr.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                try:
                    col_letter = col[0].column_letter
                except Exception:
                    continue
                max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    st.download_button(
        "⬇️ Baixar Excel consolidado",
        data=buf.getvalue(),
        file_name="tiss_conciliacao_analytics.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
