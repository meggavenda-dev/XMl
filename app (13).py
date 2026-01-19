
# -*- coding: utf-8 -*-
# =========================================================
# app.py — TISS XML + Conciliação & Analytics + Leitor de Glosas (XLSX)
# (Versão: Detalhes só com glosa + Busca AMHPTISS + SEM gráficos na aba Faturas Glosadas + coluna "Cobrança")
# =========================================================
from __future__ import annotations

import io
import os
import re
import json
import time
import shutil
import xml.etree.ElementTree as ET
import unicodedata
from pathlib import Path
from typing import List, Dict, Optional, Union, IO, Tuple
from decimal import Decimal
from datetime import datetime

import pandas as pd
import numpy as np
import streamlit as st

# =========================================================
# Configuração da página (UI)
# =========================================================
st.set_page_config(page_title="TISS • Conciliação & Analytics", layout="wide")
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

def categorizar_motivo_ans(codigo: str) -> str:
    codigo = str(codigo).strip()
    if codigo in ['1001','1002','1003','1006','1009']: return "Cadastro/Elegibilidade"
    if codigo in ['1201','1202','1205','1209']: return "Autorização/SADT"
    if codigo in ['1801','1802','1805','1806']: return "Tabela/Preços"
    if codigo.startswith('20') or codigo.startswith('22'): return "Auditoria Médica/Técnica"
    if codigo in ['2501','2505','2509']: return "Documentação/Físico"
    return "Outros/Administrativa"

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

if "demo_mappings" not in st.session_state:
    st.session_state["demo_mappings"] = load_demo_mappings()

# Cache
@st.cache_data(show_spinner=False)
def _cached_read_excel(file, sheet_name=0) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")

@st.cache_data(show_spinner=False)
def _cached_xml_bytes(b: bytes) -> List[Dict]:
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
        numero_guia_prest = tx(guia.find('ans:numeroGuiaPrestador', ANS_NS))
        numero_guia_oper  = tx(guia.find('ans:numeroGuiaOperadora', ANS_NS)) or numero_guia_prest
        paciente = tx(guia.find('.//ans:dadosBeneficiario/ans:nomeBeneficiario', ANS_NS))
        medico   = tx(guia.find('.//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional', ANS_NS))
        data_atd = tx(guia.find('.//ans:dataAtendimento', ANS_NS))
        for it in _itens_consulta(guia):
            it.update({
                'arquivo': nome,
                'numero_lote': numero_lote,
                'tipo_guia': 'CONSULTA',
                'numeroGuiaPrestador': numero_guia_prest,
                'numeroGuiaOperadora': numero_guia_oper,
                'paciente': paciente,
                'medico': medico,
                'data_atendimento': data_atd,
            })
            out.append(it)

    # SADT
    for guia in root.findall('.//ans:guiaSP-SADT', ANS_NS):
        cab = guia.find('ans:cabecalhoGuia', ANS_NS)
        aut = guia.find('ans:dadosAutorizacao', ANS_NS)

        numero_guia_prest = tx(guia.find('ans:numeroGuiaPrestador', ANS_NS))
        if not numero_guia_prest and cab is not None:
            numero_guia_prest = tx(cab.find('ans:numeroGuiaPrestador', ANS_NS))

        numero_guia_oper = ""
        if aut is not None:
            numero_guia_oper = tx(aut.find('ans:numeroGuiaOperadora', ANS_NS))
        if not numero_guia_oper and cab is not None:
            numero_guia_oper = tx(cab.find('ans:numeroGuiaOperadora', ANS_NS))
        if not numero_guia_oper:
            numero_guia_oper = numero_guia_prest

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
    try:
        df_raw = pd.read_excel(path, header=None, engine="openpyxl")
    except:
        df_raw = pd.read_csv(path, header=None)

    header_row = None
    for i in range(min(20, len(df_raw))):
        row_values = df_raw.iloc[i].astype(str).tolist()
        if any("CPF/CNPJ" in str(val).upper() for val in row_values):
            header_row = i
            break
    if header_row is None:
        raise ValueError("Não foi possível localizar a linha de cabeçalho 'CPF/CNPJ' no demonstrativo.")

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = df_raw.iloc[header_row]
    df = df.loc[:, df.columns.notna()]

    ren = {
        "Guia": "numeroGuiaPrestador",
        "Cod. Procedimento": "codigo_procedimento",
        "Descrição": "descricao_procedimento",
        "Valor Apresentado": "valor_apresentado",
        "Valor Apurado": "valor_pago",
        "Valor Glosa": "valor_glosa",
        "Quant. Exec.": "quantidade_apresentada",
        "Código Glosa": "codigo_glosa_bruto",
    }
    df = df.rename(columns=ren)

    df["numeroGuiaPrestador"] = (
        df["numeroGuiaPrestador"]
        .astype(str).str.replace(".0", "", regex=False).str.strip().str.lstrip("0")
    )
    df["codigo_procedimento"] = df["codigo_procedimento"].astype(str).str.strip()

    df["codigo_procedimento_norm"] = df["codigo_procedimento"].map(
        lambda s: normalize_code(s, strip_zeros=strip_zeros_codes)
    )

    for c in ["valor_apresentado", "valor_pago", "valor_glosa", "quantidade_apresentada"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', '.'), errors="coerce").fillna(0)

    df["chave_demo"] = df["numeroGuiaPrestador"].astype(str) + "__" + df["codigo_procedimento_norm"].astype(str)

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
    "guia_oper": [r"^\bguia\b"],
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

    df['chave_oper'] = (
        df['numeroGuiaOperadora'].fillna('').astype(str).str.strip()
        + '__' + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip()
    )

    return df

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

    m1 = df_xml.merge(df_demo, left_on="chave_prest", right_on="chave_demo", how="left", suffixes=("_xml", "_demo"))
    m1 = _alias_xml_cols(m1)
    m1["matched_on"] = m1["valor_apresentado"].notna().map({True: "prestador", False: ""})

    restante = m1[m1["matched_on"] == ""].copy()
    restante = _alias_xml_cols(restante)
    cols_xml = df_xml.columns.tolist()
    m2 = restante[cols_xml].merge(df_demo, left_on="chave_oper", right_on="chave_demo", how="left", suffixes=("_xml", "_demo"))
    m2 = _alias_xml_cols(m2)
    m2["matched_on"] = m2["valor_apresentado"].notna().map({True: "operadora", False: ""})

    conc = pd.concat([m1[m1["matched_on"] != ""], m2[m2["matched_on"] != ""]], ignore_index=True)

    fallback_matches = pd.DataFrame()
    if fallback_por_descricao:
        ainda_sem_match = m2[m2["matched_on"] == ""].copy()
        ainda_sem_match = _alias_xml_cols(ainda_sem_match)
        if not ainda_sem_match.empty:
            ainda_sem_match["guia_join"] = ainda_sem_match.apply(
                lambda r: str(r.get("numeroGuiaPrestador", "")).strip() or str(r.get("numeroGuiaOperadora", "")).strip(), axis=1
            )
            df_demo2 = df_demo.copy()
            df_demo2["guia_join"] = df_demo2["numeroGuiaPrestador"].astype(str).str.strip()
            if "descricao_procedimento" in ainda_sem_match.columns and "descricao_procedimento" in df_demo2.columns:
                tmp = ainda_sem_match[cols_xml + ["guia_join"]].merge(
                    df_demo2, on=["guia_join", "descricao_procedimento"], how="left", suffixes=("_xml", "_demo")
                )
                tol = float(tolerance_valor)
                keep = (tmp["valor_apresentado"].notna() & ((tmp["valor_total"] - tmp["valor_apresentado"]).abs() <= tol))
                fallback_matches = tmp[keep].copy()
                if not fallback_matches.empty:
                    fallback_matches["matched_on"] = "descricao+valor"
                    conc = pd.concat([conc, fallback_matches], ignore_index=True)

    if not fallback_matches.empty:
        chaves_resolvidas = fallback_matches["chave_prest"].unique()
        unmatch = m2[(m2["matched_on"] == "") & (~m2["chave_prest"].isin(chaves_resolvidas))].copy()
    else:
        unmatch = m2[m2["matched_on"] == ""].copy()
    unmatch = _alias_xml_cols(unmatch)
    if not unmatch.empty:
        subset_cols = [c for c in ["arquivo", "numeroGuiaPrestador", "codigo_procedimento", "valor_total"] if c in unmatch.columns]
        if subset_cols:
            unmatch = unmatch.drop_duplicates(subset=subset_cols)

    if not conc.empty:
        conc = _alias_xml_cols(conc)
        conc["apresentado_diff"] = conc["valor_total"] - conc["valor_apresentado"]
        conc["glosa_pct"] = conc.apply(
            lambda r: (r["valor_glosa"] / r["valor_apresentado"]) if r.get("valor_apresentado", 0) > 0 else 0.0,
            axis=1
        )

    return {"conciliacao": conc, "nao_casados": unmatch}

# -----------------------------
# Analytics
# -----------------------------
def kpis_por_competencia(df_conc: pd.DataFrame) -> pd.DataFrame:
    base = df_conc.copy()
    if base.empty:
        return base
    if 'competencia' not in base.columns and 'Competência' in base.columns:
        base['competencia'] = base['Competência'].astype(str)
    elif 'competencia' not in base.columns:
        base['competencia'] = ""
    grp = (base.groupby('competencia', dropna=False, as_index=False)
           .agg(valor_apresentado=('valor_apresentado','sum'),
                valor_pago=('valor_pago','sum'),
                valor_glosa=('valor_glosa','sum')))
    grp['glosa_pct'] = grp.apply(
        lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1
    )
    return grp.sort_values('competencia')


def ranking_itens_glosa(df_conc: pd.DataFrame, min_apresentado: float = 0.0, topn: int = 20) -> Tuple[pd.DataFrame, pd.DataFrame]:
    base = df_conc.copy()
    if base.empty:
        return base, base
    grp = (base.groupby(['codigo_procedimento','descricao_procedimento'], dropna=False, as_index=False)
           .agg(valor_apresentado=('valor_apresentado','sum'),
                valor_glosa=('valor_glosa','sum'),
                valor_pago=('valor_pago','sum'),
                qtd_glosada=('valor_glosa', lambda x: (x > 0).sum())))
    grp_com_glosa = grp[grp['valor_glosa'] > 0].copy()
    if grp_com_glosa.empty:
        return pd.DataFrame(), pd.DataFrame()
    grp_com_glosa['glosa_pct'] = (grp_com_glosa['valor_glosa'] / grp_com_glosa['valor_apresentado']) * 100
    top_valor = grp_com_glosa.sort_values('valor_glosa', ascending=False).head(topn)
    top_pct = grp_com_glosa[grp_com_glosa['valor_apresentado'] >= min_apresentado].sort_values('glosa_pct', ascending=False).head(topn)
    return top_valor, top_pct

def motivos_glosa(df_conc: pd.DataFrame, competencia: Optional[str] = None) -> pd.DataFrame:
    base = df_conc.copy()
    if base.empty:
        return base
    base = base[base['valor_glosa'] > 0]
    if competencia and 'competencia' in base.columns:
        base = base[base['competencia'] == competencia]
    if base.empty: return pd.DataFrame()
    mot = (base.groupby(['motivo_glosa_codigo','motivo_glosa_descricao'], dropna=False, as_index=False)
           .agg(valor_glosa=('valor_glosa','sum'),
                itens=('codigo_procedimento','count')))
    mot['categoria'] = mot['motivo_glosa_codigo'].apply(categorizar_motivo_ans)
    total_glosa = mot['valor_glosa'].sum()
    mot['glosa_pct'] = (mot['valor_glosa'] / total_glosa) * 100 if total_glosa > 0 else 0
    return mot.sort_values('valor_glosa', ascending=False)

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
# PARTE 5 — Auditoria de Guias (DESATIVADA)
# =========================================================
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
    return agg

# =========================================================
# PARTE 5.1 — Helpers da aba "Faturas Glosadas (XLSX)"
# =========================================================
def _pick_col(df: pd.DataFrame, *candidates):
    """Retorna o primeiro nome de coluna que existir no DF dentre os candidatos."""
    for cand in candidates:
        for c in df.columns:
            if str(c).strip().lower() == str(cand).strip().lower():
                return c
            lc = str(c).lower()
            if isinstance(cand, str) and all(w in lc for w in cand.lower().split()):
                return c
    return None


@st.cache_data(show_spinner=False)
def read_glosas_xlsx(files) -> tuple[pd.DataFrame, dict]:
    """
    Lê 1..N arquivos .xlsx de Faturas Glosadas (AMHP ou similar),
    concatena e retorna (df, colmap) com mapeamento de colunas.
    Cria sempre colunas de Pagamento derivadas (_pagto_dt/_ym/_mes_br).

    Correções:
      • "Valor Cobrado" passa a usar "Valor Original" (override)
      • "Realizado": NÃO combinar com "Horário". Só coluna exatamente "Realizado".
        Se houver duplicatas, usa a ÚLTIMA.
    """
    if not files:
        return pd.DataFrame(), {}

    parts = []
    for f in files:
        df = pd.read_excel(f, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        parts.append(df)

    df = pd.concat(parts, ignore_index=True)
    cols = df.columns

    # ---------- Mapeamento inicial ----------
    colmap = {
        "valor_cobrado": next((c for c in cols if "Valor Cobrado" in str(c)), None),
        "valor_glosa": next((c for c in cols if "Valor Glosa" in str(c)), None),
        "valor_recursado": next((c for c in cols if "Valor Recursado" in str(c)), None),
        "data_pagamento": next((c for c in cols if "Pagamento" in str(c)), None),
        "data_realizado": None,  # será definido com critério robusto abaixo
        "motivo": next((c for c in cols if "Motivo Glosa" in str(c)), None),
        "desc_motivo": next((c for c in cols if "Descricao Glosa" in str(c) or "Descrição Glosa" in str(c)), None),
        "tipo_glosa": next((c for c in cols if "Tipo de Glosa" in str(c)), None),
        "descricao": _pick_col(df, "descrição", "descricao", "descrição do item", "descricao do item"),
        "convenio": next((c for c in cols if "Convênio" in str(c) or "Convenio" in str(c)), None),
        "prestador": next((c for c in cols if "Nome Clínica" in str(c) or "Nome Clinica" in str(c) or "Prestador" in str(c)), None),
        "amhptiss": next((
            c for c in cols
            if str(c).strip().lower() in {
                "amhptiss", "amhp tiss", "nº amhptiss", "numero amhptiss", "número amhptiss"
            } or "amhptiss" in str(c).strip().lower() or str(c).strip() == "Amhptiss"
        ), None),
        "cobranca": next((c for c in cols if str(c).strip().lower() == "cobrança" or "cobranca" in str(c).lower()), None),
    }

    # ---------- "Realizado" robusto (sem "Horário") ----------
    norm_cols = [(c, re.sub(r"\s+", " ", str(c)).strip().lower()) for c in cols]
    realizado_exact = [c for c, n in norm_cols if n == "realizado"]
    if not realizado_exact:
        realizado_contains = [c for c, n in norm_cols if ("realizado" in n) and ("horar" not in n)]
    else:
        realizado_contains = []
    if realizado_exact:
        col_data_realizado = realizado_exact[-1]
    elif realizado_contains:
        col_data_realizado = realizado_contains[-1]
    else:
        col_data_realizado = None
    colmap["data_realizado"] = col_data_realizado

    # ---------- "Valor Cobrado" ← "Valor Original" ----------
    col_valor_original = next((c for c in cols if str(c).strip().lower() == "valor original"), None)
    if col_valor_original:
        colmap["valor_original"] = col_valor_original
        if colmap["valor_cobrado"] and colmap["valor_cobrado"] in df.columns:
            df[colmap["valor_cobrado"]] = df[col_valor_original]
        else:
            colmap["valor_cobrado"] = col_valor_original

    # ---------- Normalização AMHPTISS ----------
    amhp_col = colmap.get("amhptiss")
    if amhp_col and amhp_col in df.columns:
        df[amhp_col] = (
            df[amhp_col]
            .astype(str)
            .str.replace(r"[^\d]", "", regex=True)
            .str.strip()
        )

    # ---------- Números ----------
    for c in [colmap.get("valor_cobrado"), colmap.get("valor_glosa"), colmap.get("valor_recursado")]:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # ---------- Datas ----------
    if colmap.get("data_realizado") and colmap["data_realizado"] in df.columns:
        df[colmap["data_realizado"]] = pd.to_datetime(
            df[colmap["data_realizado"]], errors="coerce", dayfirst=True
        )
    if colmap.get("data_pagamento") and colmap["data_pagamento"] in df.columns:
        df["_pagto_dt"] = pd.to_datetime(df[colmap["data_pagamento"]], errors="coerce", dayfirst=True)
    else:
        df["_pagto_dt"] = pd.NaT

    if "_pagto_dt" in df.columns and df["_pagto_dt"].notna().any():
        df["_pagto_ym"] = df["_pagto_dt"].dt.to_period("M")
        df["_pagto_mes_br"] = df["_pagto_dt"].dt.strftime("%m/%Y")
    else:
        df["_pagto_ym"] = pd.NaT
        df["_pagto_mes_br"] = ""

    # ---------- Flags de glosa ----------
    if colmap.get("valor_glosa") in df.columns:
        df["_is_glosa"] = df[colmap["valor_glosa"]] < 0
        df["_valor_glosa_abs"] = df[colmap["valor_glosa"]].abs()
    else:
        df["_is_glosa"] = False
        df["_valor_glosa_abs"] = 0.0

    return df, colmap

def build_glosas_analytics(df: pd.DataFrame, colmap: dict) -> dict:
    """
    KPIs e agrupamentos para a aba de glosas (respeita filtros aplicados previamente).
    """
    if df.empty or not colmap:
        return {}

    cm = colmap
    m = df["_is_glosa"].fillna(False)

    total_linhas = len(df)
    periodo_ini = df[cm["data_realizado"]].min() if cm["data_realizado"] in df.columns else None
    periodo_fim = df[cm["data_realizado"]].max() if cm["data_realizado"] in df.columns else None
    valor_cobrado = float(df[cm["valor_cobrado"]].fillna(0).sum()) if cm["valor_cobrado"] in df.columns else 0.0
    valor_glosado = float(df.loc[m, "_valor_glosa_abs"].sum())
    taxa_glosa = (valor_glosado / valor_cobrado) if valor_cobrado else 0.0
    convenios = int(df[cm["convenio"]].nunique()) if cm["convenio"] in df.columns else 0
    prestadores = int(df[cm["prestador"]].nunique()) if cm["prestador"] in df.columns else 0

    base = df.loc[m].copy()

    def _agg(df_, keys):
        if df_.empty:
            return df_
        out = (df_.groupby(keys, dropna=False, as_index=False)
               .agg(Qtd=('_is_glosa', 'size'),
                    Valor_Glosado=('_valor_glosa_abs', 'sum')))
        return out.sort_values(["Valor_Glosado","Qtd"], ascending=False)

    top_motivos = _agg(base, [cm["motivo"], cm["desc_motivo"]]) if cm.get("motivo") and cm.get("desc_motivo") else pd.DataFrame()
    by_tipo     = _agg(base, [cm["tipo_glosa"]]) if cm.get("tipo_glosa") else pd.DataFrame()
    top_itens   = _agg(base, [cm["descricao"]]) if cm.get("descricao") else pd.DataFrame()
    by_convenio = _agg(base, [cm["convenio"]]) if cm.get("convenio") else pd.DataFrame()

    if not top_motivos.empty:
        top_motivos = top_motivos.rename(columns={
            cm["motivo"]: "Motivo",
            cm["desc_motivo"]: "Descrição do Motivo",
            "Valor_Glosado": "Valor Glosado (R$)"
        })
    if not by_tipo.empty:
        by_tipo = by_tipo.rename(columns={cm["tipo_glosa"]: "Tipo de Glosa", "Valor_Glosado":"Valor Glosado (R$)"})
    if not top_itens.empty:
        top_itens = top_itens.rename(columns={cm["descricao"]:"Descrição do Item", "Valor Glosado":"Valor Glosado (R$)"})
    if not by_convenio.empty:
        by_convenio = by_convenio.rename(columns={cm["convenio"]:"Convênio", "Valor Glosado":"Valor Glosado (R$)"})

    return dict(
        kpis=dict(
            linhas=total_linhas,
            periodo_ini=periodo_ini,
            periodo_fim=periodo_fim,
            convenios=convenios,
            prestadores=prestadores,
            valor_cobrado=valor_cobrado,
            valor_glosado=valor_glosado,
            taxa_glosa=taxa_glosa
        ),
        top_motivos=top_motivos,
        by_tipo=by_tipo,
        top_itens=top_itens,
        by_convenio=by_convenio
    )

# =========================================================
# PARTE 6 — Interface (Uploads, Parâmetros, Processamento, Analytics, Export)
# =========================================================

with st.sidebar:
    # Agora a área de parâmetros está sempre FECHADA por padrão
    with st.expander("⚙️ Parâmetros", expanded=False):
        prazo_retorno = st.number_input(
            "Prazo de retorno (dias) — (auditoria desativada)",
            min_value=0, value=30, step=1
        )
        tolerance_valor = st.number_input(
            "Tolerância p/ fallback por descrição (R$)",
            min_value=0.00, value=0.02, 
            step=0.01, format="%.2f"
        )
        fallback_desc = st.toggle(
            "Fallback por descrição + valor (quando código não casar)",
            value=False
        )
        strip_zeros_codes = st.toggle(
            "Normalizar códigos removendo zeros à esquerda",
            value=True
        )

tab_conc, tab_glosas = st.tabs(["🔗 Conciliação TISS", "📑 Faturas Glosadas (XLSX)"])

# =========================================================
# ABA 1 — Conciliação TISS (mantida com gráficos)
# =========================================================
with tab_conc:
    st.subheader("📤 Upload de arquivos")
    xml_files = st.file_uploader("XML TISS (um ou mais):", type=['xml'], accept_multiple_files=True, key="xml_up")
    demo_files = st.file_uploader("Demonstrativos de Pagamento (.xlsx) — itemizado:", type=['xlsx'], accept_multiple_files=True, key="demo_up")

    # PROCESSAMENTO DO DEMONSTRATIVO (sempre) — permite wizard
    df_demo = build_demo_df(demo_files or [], strip_zeros_codes=strip_zeros_codes)
    if not df_demo.empty:
        st.info("Demonstrativo carregado e mapeado. A conciliação considerará **somente** os itens presentes nos XMLs. Itens presentes apenas no demonstrativo serão **ignorados**.")
    else:
        if demo_files:
            st.info("Carregue um Demonstrativo válido ou conclua o mapeamento manual.")

    st.markdown("---")
    if st.button("🚀 Processar Conciliação & Analytics", type="primary", key="btn_conc"):
        df_xml = build_xml_df(xml_files or [], strip_zeros_codes=strip_zeros_codes)
        if df_xml.empty:
            st.warning("Nenhum item extraído do(s) XML(s). Verifique os arquivos.")
            st.stop()

        st.subheader("📄 Itens extraídos dos XML (Consulta / SADT)")
        st.dataframe(apply_currency(df_xml, ['valor_unitario','valor_total']), use_container_width=True, height=360)

        if df_demo.empty:
            st.warning("Nenhum demonstrativo válido para conciliar.")
            st.stop()

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
            st.download_button("Baixar Não Conciliados (CSV)", data=unmatch.to_csv(index=False).encode("utf-8"),
                               file_name="nao_conciliados.csv", mime="text/csv")

        # Analytics (conciliado)
        st.markdown("---")
        st.subheader("📊 Analytics de Glosa (apenas itens conciliados)")

        st.markdown("### 📈 Tendência por competência")
        kpi_comp = kpis_por_competencia(conc)
        st.dataframe(apply_currency(kpi_comp, ['valor_apresentado','valor_pago','valor_glosa']), use_container_width=True)
        try:
            st.line_chart(kpi_comp.set_index('competencia')[['valor_apresentado','valor_pago','valor_glosa']])
        except Exception:
            pass

        st.markdown("### 🏆 TOP itens glosados (valor e %)")
        min_apres = st.number_input("Corte mínimo de Apresentado para ranking por % (R$)", min_value=0.0, value=500.0, step=50.0, key="min_apres_pct")
        top_valor, top_pct = ranking_itens_glosa(conc, min_apresentado=min_apres, topn=20)
        t1, t2 = st.columns(2)
        with t1:
            st.markdown("**Por valor de glosa (TOP 20)**")
            st.dataframe(apply_currency(top_valor, ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)
        with t2:
            st.markdown("**Por % de glosa (TOP 20)**")
            st.dataframe(apply_currency(top_pct, ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)

        st.markdown("### 🧩 Motivos de glosa — análise")
        comp_opts = ['(todas)']
        if 'competencia' in conc.columns:
            comp_opts += sorted(conc['competencia'].dropna().astype(str).unique().tolist())
        comp_sel = st.selectbox("Filtrar por competência", comp_opts, key="comp_mot")
        motdf = motivos_glosa(conc, None if comp_sel=='(todas)' else comp_sel)
        st.dataframe(apply_currency(motdf, ['valor_glosa','valor_apresentado']), use_container_width=True)

        st.markdown("### 👩‍⚕️ Médicos — ranking por glosa")
        if 'competencia' in conc.columns:
            comp_med = st.selectbox("Competência (médicos)",
                                    ['(todas)'] + sorted(conc['competencia'].dropna().astype(str).unique().tolist()),
                                    key="comp_med")
            med_base = conc if comp_med == '(todas)' else conc[conc['competencia'] == comp_med]
        else:
            med_base = conc
        med_rank = (med_base.groupby(['medico'], dropna=False, as_index=False)
                    .agg(valor_apresentado=('valor_apresentado','sum'),
                         valor_glosa=('valor_glosa','sum'),
                         valor_pago=('valor_pago','sum'),
                         itens=('arquivo','count')))
        med_rank['glosa_pct'] = med_rank.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
        st.dataframe(apply_currency(med_rank.sort_values(['glosa_pct','valor_glosa'], ascending=[False,False]),
                                    ['valor_apresentado','valor_glosa','valor_pago']), use_container_width=True)

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

        if 'matched_on' in conc.columns:
            st.markdown("### 🧪 Qualidade da conciliação (origem do match)")
            match_dist = conc['matched_on'].value_counts(dropna=False).rename_axis('origem').reset_index(name='itens')
            st.bar_chart(match_dist.set_index('origem'))
            st.dataframe(match_dist, use_container_width=True)

        st.markdown("### 🚩 Outliers em valor apresentado (por procedimento)")
        out_df = outliers_por_procedimento(conc, k=1.5)
        if out_df.empty:
            st.info("Nenhum outlier identificado com o critério atual (IQR).")
        else:
            st.dataframe(out_df, use_container_width=True, height=280)
            st.download_button("Baixar Outliers (CSV)", data=out_df.to_csv(index=False).encode("utf-8"),
                               file_name="outliers_valor_apresentado.csv", mime="text/csv")

        st.markdown("### 🧮 Simulador de faturamento (what‑if por motivo de glosa)")
        motivos_disponiveis = sorted(conc['motivo_glosa_codigo'].dropna().astype(str).unique().tolist()) if 'motivo_glosa_codigo' in conc.columns else []
        if motivos_disponiveis:
            cols_sim = st.columns(min(4, max(1, len(motivos_disponiveis))))
            ajustes = {}
            for i, cod in enumerate(motivos_disponiveis):
                col = cols_sim[i % len(cols_sim)]
                with col:
                    fator = st.slider(f"Motivo {cod} → fator (0–1)", 0.0, 1.0, 1.0, 0.05,
                                      help="Ex.: 0,8 reduz a glosa em 20% para esse motivo.", key=f"sim_{cod}")
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

        # Export Excel consolidado
        st.markdown("---")
        st.subheader("📥 Exportar Excel Consolidado")

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
            df_xml.to_excel(wr, index=False, sheet_name='Itens_XML')
            if not itens_demo_match.empty:
                itens_demo_match.to_excel(wr, index=False, sheet_name='Itens_Demo')
            conc.to_excel(wr, index=False, sheet_name='Conciliação')
            unmatch.to_excel(wr, index=False, sheet_name='Nao_Casados')

            mot_x = motivos_glosa(conc, None)
            mot_x.to_excel(wr, index=False, sheet_name='Motivos_Glosa')

            proc_x = (conc.groupby(['codigo_procedimento','descricao_procedimento'], dropna=False, as_index=False)
                      .agg(valor_apresentado=('valor_apresentado','sum'),
                           valor_glosa=('valor_glosa','sum'),
                           valor_pago=('valor_pago','sum'),
                           itens=('arquivo','count')))
            proc_x['glosa_pct'] = proc_x.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
            proc_x.to_excel(wr, index=False, sheet_name='Procedimentos_Glosa')

            med_x = (conc.groupby(['medico'], dropna=False, as_index=False)
                     .agg(valor_apresentado=('valor_apresentado','sum'),
                          valor_glosa=('valor_glosa','sum'),
                          valor_pago=('valor_pago','sum'),
                          itens=('arquivo','count')))
            med_x['glosa_pct'] = med_x.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
            med_x.to_excel(wr, index=False, sheet_name='Medicos')

            if 'numero_lote' in conc.columns:
                lot_x = (conc.groupby(['numero_lote'], dropna=False, as_index=False)
                         .agg(valor_apresentado=('valor_apresentado','sum'),
                              valor_glosa=('valor_glosa','sum'),
                              valor_pago=('valor_pago','sum'),
                              itens=('arquivo','count')))
                lot_x['glosa_pct'] = lot_x.apply(lambda r: (r['valor_glosa']/r['valor_apresentado']) if r['valor_apresentado']>0 else 0, axis=1)
                lot_x.to_excel(wr, index=False, sheet_name='Lotes')

            kpi_comp.to_excel(wr, index=False, sheet_name='KPIs_Competencia')

        st.download_button(
            "⬇️ Baixar Excel consolidado",
            data=buf.getvalue(),
            file_name="tiss_conciliacao_analytics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =========================================================
# ABA 2 — Faturas Glosadas (XLSX) (SEM gráficos)
# =========================================================
with tab_glosas:
    st.subheader("Leitor de Faturas Glosadas (XLSX) — independente do XML/Demonstrativo")
    st.caption("A análise respeita filtros por **Convênio** e por **mês de Pagamento**. O processamento é persistido com session_state.")

    if "glosas_ready" not in st.session_state:
        st.session_state.glosas_ready = False
        st.session_state.glosas_data = None
        st.session_state.glosas_colmap = None
        st.session_state.glosas_files_sig = None

    glosas_files = st.file_uploader(
        "Relatórios de Faturas Glosadas (.xlsx):",
        type=["xlsx"],
        accept_multiple_files=True,
        key="glosas_xlsx_up"
    )

    def _files_signature(files):
        if not files:
            return None
        return tuple(sorted((getattr(f, "name", ""), getattr(f, "size", 0)) for f in files))

    a1, a2 = st.columns(2)
    with a1:
        proc_click = st.button("📊 Processar Faturas Glosadas", type="primary", key="proc_glosas_btn")
    with a2:
        clear_click = st.button("🧹 Limpar / Resetar", key="clear_glosas_btn")

    if clear_click:
        st.session_state.glosas_ready = False
        st.session_state.glosas_data = None
        st.session_state.glosas_colmap = None
        st.session_state.glosas_files_sig = None
        st.rerun()

    if proc_click:
        if not glosas_files:
            st.warning("Selecione pelo menos um arquivo .xlsx antes de processar.")
        else:
            files_sig = _files_signature(glosas_files)
            df_g, colmap = read_glosas_xlsx(glosas_files)
            st.session_state.glosas_data = df_g
            st.session_state.glosas_colmap = colmap
            st.session_state.glosas_ready = True
            st.session_state.glosas_files_sig = files_sig
            st.rerun()

    if st.session_state.glosas_ready and st.session_state.glosas_data is not None:
        current_sig = _files_signature(glosas_files)
        if (glosas_files and current_sig != st.session_state.glosas_files_sig):
            st.info("Os arquivos enviados mudaram desde o último processamento. Clique em **Processar Faturas Glosadas** para atualizar.")

        df_g   = st.session_state.glosas_data
        colmap = st.session_state.glosas_colmap

        # Diagnóstico
        with st.expander("🔧 Diagnóstico (debug rápido)", expanded=False):
            st.write("**Colunas do DataFrame:**", list(df_g.columns))
            st.write("**Mapeamento detectado (colmap):**")
            st.json({k: v for k, v in colmap.items() if v})
            st.write("**Amostra (5 linhas):**")
            st.dataframe(df_g.head(5), use_container_width=True)
            flags = {
                "_pagto_dt": "_pagto_dt" in df_g.columns,
                "_pagto_ym": "_pagto_ym" in df_g.columns,
                "_pagto_mes_br": "_pagto_mes_br" in df_g.columns,
            }
            st.write("**Flags de Pagamento criadas?**", flags)

        # Filtros
        has_pagto = ("_pagto_dt" in df_g.columns) and df_g["_pagto_dt"].notna().any()
        if not has_pagto:
            st.warning("Coluna 'Pagamento' não encontrada ou sem dados válidos. Recursos mensais ficarão limitados.")

        conv_opts = ["(todos)"]
        if colmap.get("convenio") and colmap["convenio"] in df_g.columns:
            conv_unique = sorted(df_g[colmap["convenio"]].dropna().astype(str).unique().tolist())
            conv_opts += conv_unique
        conv_sel = st.selectbox("Convênio", conv_opts, index=0, key="conv_glosas")

        if has_pagto:
            meses_df = (df_g.loc[df_g["_pagto_ym"].notna(), ["_pagto_ym","_pagto_mes_br"]]
                          .drop_duplicates().sort_values("_pagto_ym"))
            meses_labels = meses_df["_pagto_mes_br"].tolist()
            modo_periodo = st.radio("Período (por **Pagamento**):",
                                    ["Todos os meses (agrupado)", "Um mês"],
                                    horizontal=False, key="modo_periodo")
            mes_sel_label = None
            if modo_periodo == "Um mês" and meses_labels:
                mes_sel_label = st.selectbox("Escolha o mês (Pagamento)", meses_labels, key="mes_pagto_sel")
        else:
            modo_periodo = "Todos os meses (agrupado)"
            mes_sel_label = None

        # Aplicar filtros
        df_view = df_g.copy()
        amhp_col = colmap.get("amhptiss")
        if amhp_col and amhp_col in df_view.columns:
            df_view[amhp_col] = (
                df_view[amhp_col]
                .astype(str)
                .str.replace(r"[^\d]", "", regex=True)
                .str.strip()
            )
        if conv_sel != "(todos)" and colmap.get("convenio") and colmap["convenio"] in df_view.columns:
            df_view = df_view[df_view[colmap["convenio"]].astype(str) == conv_sel]
        if has_pagto and mes_sel_label:
            df_view = df_view[df_view["_pagto_mes_br"] == mes_sel_label]

        # Série mensal (Pagamento) — SEM gráficos (sempre soma o Valor Cobrado = Valor Original)
        st.markdown("### 📅 Glosa por **mês de pagamento**")
        has_pagto = ("_pagto_dt" in df_view.columns) and df_view["_pagto_dt"].notna().any()
        if has_pagto:
            base_m = df_view[df_view["_is_glosa"] == True].copy()
            if base_m.empty:
                st.info("Sem glosas no recorte atual.")
            else:                
                # ========================
                # 📅 Glosa por mês de pagamento — versão personalizada
                # ========================
                mensal = (
                    base_m.groupby(["_pagto_ym", "_pagto_mes_br"], as_index=False)
                          .agg(
                              Valor_Glosado=("_valor_glosa_abs", "sum"),
                              Valor_Cobrado=(colmap["valor_cobrado"], "sum"),
                              Valor_Recursado=(colmap["valor_recursado"], "sum") if colmap.get("valor_recursado") in base_m.columns else ("_valor_glosa_abs", "size")
                          )
                          .sort_values("_pagto_ym")
                )
                
                # 1) Renomear colunas
                mensal = mensal.rename(columns={
                    "_pagto_mes_br": "Mês de Pagamento",
                    "Valor_Glosado": "Valor Glosado (R$)",
                    "Valor_Cobrado": "Valor Cobrado (R$)",
                    "Valor_Recursado": "Valor Recursado (R$)",
                })
                
                # 2) Selecionar somente as 4 colunas desejadas
                cols_final = [
                    "Mês de Pagamento",
                    "Valor Cobrado (R$)",
                    "Valor Glosado (R$)",
                    "Valor Recursado (R$)"
                ]
                mensal = mensal[cols_final]
                
                # 3) Formatar moeda
                mensal_fmt = apply_currency(
                    mensal,
                    ["Valor Cobrado (R$)", "Valor Glosado (R$)", "Valor Recursado (R$)"]
                )
                
                # 4) Exibir
                st.dataframe(mensal_fmt, use_container_width=True, height=260)

        else:
            st.info("Sem 'Pagamento' válido para montar série mensal.")

        # ==========================================
        # Seções seguintes
        # ==========================================
        analytics = build_glosas_analytics(df_view, colmap)

        
        st.markdown("### 🏥 Convênios com maior valor glosado")
        
        by_conv = analytics["by_convenio"] if analytics else pd.DataFrame()
        if by_conv.empty:
            st.info("Coluna de 'Convênio' não encontrada.")
        else:
            # 1) Base de Valor Cobrado por convênio (no recorte atual: df_view)
            if colmap.get("convenio") in df_view.columns and colmap.get("valor_cobrado") in df_view.columns:
                cob_df = (
                    df_view.groupby(colmap["convenio"], as_index=False)
                           .agg(Valor_Cobrado=(colmap["valor_cobrado"], "sum"))
                           .rename(columns={colmap["convenio"]: "Convênio"})
                )
            else:
                cob_df = pd.DataFrame(columns=["Convênio", "Valor_Cobrado"])
        
            # 2) Unificar com o ranking de glosa vindo do analytics
            conv_df = by_conv.copy()
        
            # Nome da coluna de glosa (pode ser "Valor Glosado (R$)" ou "Valor_Glosado")
            glosa_col = "Valor Glosado (R$)" if "Valor Glosado (R$)" in conv_df.columns else (
                "Valor_Glosado" if "Valor_Glosado" in conv_df.columns else None
            )
        
            # Renomear para o padrão pedido pelo Guilherme
            ren_map = {}
            if glosa_col:
                ren_map[glosa_col] = "Valor Glosado"
            conv_df = conv_df.rename(columns=ren_map)
        
            # 3) Trazer Valor Cobrado e manter apenas as 4 colunas desejadas
            conv_df = conv_df.merge(cob_df, on="Convênio", how="left")
            # Se quisermos nome sem “(R$)”, usar "Valor Cobrado" mesmo:
            conv_df = conv_df.rename(columns={"Valor_Cobrado": "Valor Cobrado"})
        
            # 4) Selecionar e ordenar colunas
            cols_final = ["Convênio", "Qtd", "Valor Cobrado", "Valor Glosado"]
            for c in cols_final:
                if c not in conv_df.columns:
                    conv_df[c] = 0
            conv_df = conv_df[cols_final].copy()
        
            # 5) Formatar moeda nas duas colunas financeiras
            conv_df_fmt = apply_currency(conv_df, ["Valor Cobrado", "Valor Glosado"])
        
            # 6) Mostrar TOP 20 (por Valor Glosado desc, depois Qtd)
            conv_df_fmt = (
                conv_df_fmt
                .assign(_ord_glosa = conv_df["Valor Glosado"].astype(float),
                        _ord_qtd   = conv_df["Qtd"].astype(int))
                .sort_values(["_ord_glosa", "_ord_qtd"], ascending=[False, False])
                .drop(columns=["_ord_glosa", "_ord_qtd"])
                .head(20)
            )
        
            st.dataframe(conv_df_fmt, use_container_width=True, height=320)


        # ---------- Itens/descrições com maior valor glosado (Detalhes só com glosa) ----------
        st.markdown("### 🧩 Itens/descrições com maior valor glosado")
        top_itens = analytics["top_itens"] if analytics else pd.DataFrame()
        if top_itens.empty:
            st.info("Coluna de 'Descrição' não encontrada.")
        else:
            df_items = top_itens.copy()
            if "Descrição do Item" not in df_items.columns:
                desc_col = colmap.get("descricao")
                if desc_col and desc_col in df_items.columns:
                    df_items = df_items.rename(columns={desc_col: "Descrição do Item"})
            if "Valor Glosado (R$)" not in df_items.columns and "Valor_Glosado" in df_items.columns:
                df_items = df_items.rename(columns={"Valor_Glosado": "Valor Glosado (R$)"})

            if "Qtd" not in df_items.columns:
                if colmap and "_is_glosa" in df_view.columns and colmap.get("descricao") in df_view.columns:
                    qtd_series = (
                        df_view[df_view["_is_glosa"] == True]
                        .groupby(colmap["descricao"])["_is_glosa"]
                        .size()
                    )
                    df_items["Qtd"] = df_items["Descrição do Item"].map(qtd_series).fillna(0).astype(int)
                else:
                    df_items["Qtd"] = 0

            df_items_top = df_items.head(20).copy()
            df_items_show = apply_currency(df_items_top.copy(), ["Valor Glosado (R$)"])

            sel_state_key = "top_itens_editor_selected"
            ver_key       = "top_itens_editor_version"
            if ver_key not in st.session_state:
                st.session_state[ver_key] = 0
            if sel_state_key not in st.session_state:
                st.session_state[sel_state_key] = None

            selected_item_name = st.session_state[sel_state_key]
            prev_series = (df_items_show.get("Descrição do Item", "").astype(str) == str(selected_item_name))
            df_items_show["🔍 Detalhes"] = prev_series

            st.caption("Clique em **🔍 Detalhes** para abrir a relação das guias (somente com glosa) deste item.")
            editor_key = f"top_itens_editor__v{st.session_state[ver_key]}"

            edited = st.data_editor(
                df_items_show,
                use_container_width=True,
                height=420,
                disabled=[c for c in df_items_show.columns if c != "🔍 Detalhes"],
                column_config={
                    "🔍 Detalhes": st.column_config.CheckboxColumn(
                        help="Mostrar detalhes deste item logo abaixo",
                        default=False
                    )
                },
                key=editor_key
            )

            if "Descrição do Item" not in edited.columns:
                new_selected_item = None
            else:
                curr_series = edited["🔍 Detalhes"].astype(bool).reindex(prev_series.index, fill_value=False)
                turned_on  = (curr_series & ~prev_series)
                if turned_on.any():
                    idx = turned_on[turned_on].index[-1]
                    new_selected_item = edited.loc[idx, "Descrição do Item"]
                elif not curr_series.any():
                    new_selected_item = None
                elif curr_series.sum() == 1:
                    idx = curr_series.idxmax()
                    new_selected_item = edited.loc[idx, "Descrição do Item"]
                else:
                    candidates = curr_series[curr_series].index.tolist()
                    prev_idx = prev_series[prev_series].index.tolist()
                    pick = [i for i in candidates if i not in prev_idx]
                    idx = (pick[-1] if pick else candidates[-1])
                    new_selected_item = edited.loc[idx, "Descrição do Item"]

            if new_selected_item != selected_item_name:
                st.session_state[sel_state_key] = new_selected_item
                st.session_state[ver_key] += 1
                st.rerun()

            selected_item_name = st.session_state[sel_state_key]

            # ============ BUSCA POR Nº AMHPTISS ============
            amhp_col = colmap.get("amhptiss")
            if amhp_col and amhp_col in df_g.columns:
                @st.cache_data
                def normalize_and_index(df, col):
                    df2 = df.copy()
                    df2["_amhp_digits"] = (
                        df2[col].astype(str).str.replace(r"[^\d]", "", regex=True).str.strip()
                    )
                    index = {}
                    for i, v in df2["_amhp_digits"].items():
                        if v not in index:
                            index[v] = []
                        index[v].append(i)
                    return df2, index
                df_g, amhp_index = normalize_and_index(df_g, amhp_col)

            st.session_state.setdefault("amhp_query", "")
            st.session_state.setdefault("amhp_result", None)

            st.markdown("## 🔎 Buscar por **Nº AMHPTISS**")
            st.markdown("---")

            if not amhp_col or amhp_col not in df_g.columns:
                st.info("Não foi possível identificar a coluna de **AMHPTISS** nos arquivos enviados.")
            else:
                col1, col2 = st.columns([0.65, 0.35])
                with col1:
                    numero_input = st.text_input(
                        "Informe o Nº AMHPTISS",
                        value=st.session_state.amhp_query,
                        placeholder="Ex.: 61916098"
                    )
                    cbt1, cbt2 = st.columns(2)
                    with cbt1:
                        clique_buscar = st.button("🔍 Buscar", key="btn_buscar_amhp")
                    with cbt2:
                        clique_fechar = st.button("❌ Fechar resultados", key="btn_fechar_amhp")
                with col2:
                    ignorar_filtros = st.checkbox(
                        "Ignorar filtros de Convênio/Mês",
                        False,
                        help="Busca no dataset completo, ignorando filtros ativos."
                    )

                def digits(s): return re.sub(r"\D+", "", str(s or ""))


                if clique_fechar:
                    st.session_state.amhp_query = ""
                    st.session_state.amhp_result = None
                    st.rerun()   # 🔄 não interrompe a página, apenas recarrega


                
                if clique_buscar:
                    num = digits(numero_input)
                    if not num:
                        st.warning("Digite um Nº AMHPTISS válido.")
                    else:
                        st.session_state.amhp_query = num
                        base = df_g if ignorar_filtros else df_view
                
                        if num in amhp_index:
                            idx = amhp_index[num]
                
                            # ✅ mantém só os índices existentes no DF base (evita KeyError)
                            #    Obs.: a ordem é preservada como no índice da guia (idx)
                            idx_validos = [i for i in idx if i in base.index]
                
                            if idx_validos:
                                result = base.loc[idx_validos]
                            else:
                                # A guia existe no dataset completo, mas saiu com os filtros atuais
                                result = pd.DataFrame()
                        else:
                            # Nº AMHPTISS inexistente no dataset
                            result = pd.DataFrame()
                
                        # ✅ SALVA o resultado no estado para ser lido abaixo
                        st.session_state.amhp_result = result



                result = st.session_state.amhp_result
                numero_alvo = st.session_state.amhp_query

                if result is not None:
                    st.markdown("---")
                    st.subheader(f"🧾 Itens da guia — AMHPTISS **{numero_alvo}**")

                    if result.empty:
                        msg = "" if ignorar_filtros else " com os filtros atuais"
                        st.info(f"Nenhuma linha encontrada para esse AMHPTISS{msg}.")
                    else:
                        motivo_col = colmap.get("motivo")
                        if motivo_col and motivo_col in result.columns:
                            result = result.assign(
                                **{motivo_col: result[motivo_col].astype(str).str.replace(r"[^\d]", "", regex=True).str.strip()}
                            )

                        col_vc = colmap.get("valor_cobrado")
                        col_vg = colmap.get("valor_glosa")
                        qtd_cobrados = len(result)
                        total_cobrado = float(pd.to_numeric(result[col_vc], errors="coerce").fillna(0).sum()) if col_vc in result else 0.0
                        total_glosado = float(pd.to_numeric(result[col_vg], errors="coerce").abs().fillna(0).sum()) if col_vg in result else 0.0
                        qtd_glosados = int((result["_is_glosa"] == True).sum()) if "_is_glosa" in result.columns else 0

                        st.markdown("### 📌 Resumo da guia")
                        st.write(f"**Total Cobrado:** {f_currency(total_cobrado)}")
                        st.write(f"**Total Glosado:** {f_currency(total_glosado)}")
                        st.write(f"**Itens cobrados:** {qtd_cobrados}")
                        st.write(f"**Itens glosados:** {qtd_glosados}")
                        st.markdown("---")

                        ren = {}
                        if col_vc and col_vc in result.columns: ren[col_vc] = "Valor Cobrado (R$)"
                        if col_vg and col_vg in result.columns: ren[col_vg] = "Valor Glosado (R$)"
                        col_vr = colmap.get("valor_recursado")
                        if col_vr and col_vr in result.columns: ren[col_vr] = "Valor Recursado (R$)"
                        result_show = result.rename(columns=ren)

                        exibir_cols = [
                            amhp_col,
                            colmap.get("convenio"),
                            colmap.get("prestador"),
                            colmap.get("descricao"),
                            motivo_col,
                            colmap.get("desc_motivo"),
                            colmap.get("tipo_glosa"),
                            colmap.get("data_realizado"),
                            colmap.get("data_pagamento"),
                            colmap.get("cobranca"),
                            "Valor Cobrado (R$)",
                            "Valor Glosado (R$)",
                            "Valor Recursado (R$)",
                        ]
                        exibir_cols = [c for c in exibir_cols if c in result_show.columns]

                        st.dataframe(
                            apply_currency(result_show[exibir_cols], ["Valor Cobrado (R$)", "Valor Glosado (R$)", "Valor Recursado (R$)"]),
                            use_container_width=True,
                            height=420
                        )

                        st.download_button(
                            "⬇️ Baixar resultado (CSV)",
                            result_show[exibir_cols].to_csv(index=False).encode("utf-8"),
                            file_name=f"itens_AMHPTISS_{numero_alvo}.csv",
                            mime="text/csv"
                        )

                        if not ignorar_filtros:
                            st.caption("Dica: se algum item não aparecer, marque **“Ignorar filtros de Convênio/Mês”**.")

        # === DETALHES DO ITEM SELECIONADO ===
        if selected_item_name:
            st.markdown("---")
            st.markdown(f"#### 🔎 Detalhes — {selected_item_name}")

            
            if st.button("❌ Fechar detalhes", key="btn_fechar_detalhes_item"):
                st.session_state[sel_state_key] = None
                st.session_state[ver_key] += 1
                st.rerun()   # 🔄 não afeta outros blocos


            desc_col_map = colmap.get("descricao")
            if not desc_col_map or desc_col_map not in df_view.columns:
                st.warning("Não foi possível localizar a coluna de descrição original no dataset. Verifique o mapeamento.")
            else:
                sel_name_str = str(selected_item_name)
                mask_item = (df_view[desc_col_map].astype(str) == sel_name_str)
                mask_glosa = (mask_item & (df_view["_is_glosa"] == True)) if "_is_glosa" in df_view.columns else mask_item

                amhp_col2 = colmap.get("amhptiss")
                if not amhp_col2:
                    for cand in ["Amhptiss", "AMHPTISS", "AMHP TISS", "Nº AMHPTISS", "Numero AMHPTISS", "Número AMHPTISS"]:
                        if cand in df_view.columns:
                            amhp_col2 = cand
                            break

                possiveis = [
                    amhp_col2,
                    colmap.get("convenio"),
                    colmap.get("prestador"),
                    colmap.get("data_pagamento"),
                    colmap.get("data_realizado"),
                    colmap.get("motivo"),
                    colmap.get("desc_motivo"),
                    colmap.get("cobranca"),
                    colmap.get("valor_cobrado"),
                    colmap.get("valor_glosa"),
                    colmap.get("valor_recursado"),
                ]
                show_cols = [c for c in possiveis if c and c in df_view.columns]

                df_item = df_view.loc[mask_glosa, show_cols]

                vc = colmap.get("valor_cobrado")
                vg = colmap.get("valor_glosa")
                vr = colmap.get("valor_recursado")

                cols_min = [c for c in [vc, vg] if c and c in df_view.columns]
                df_item_all = df_view.loc[mask_item, cols_min] if cols_min else df_view.loc[mask_item, []]

                qtd_itens_cobrados = int(mask_item.sum())
                total_cobrado = float(df_item_all[vc].sum()) if vc in df_item_all.columns else 0.0

                if "_valor_glosa_abs" in df_view.columns:
                    total_glosado = float(df_view.loc[mask_glosa, "_valor_glosa_abs"].sum())
                elif vg and vg in df_view.columns:
                    total_glosado = float(df_view.loc[mask_glosa, vg].abs().sum())
                else:
                    total_glosado = 0.0

                st.markdown("### 📌 Resumo do item")
                st.write(f"**Itens cobrados:** {qtd_itens_cobrados}")
                st.write(f"**Total cobrado:** {f_currency(total_cobrado)}")
                st.write(f"**Total glosado:** {f_currency(total_glosado)}")
                st.markdown("---")

                if "_valor_glosa_abs" in df_view.columns:
                    order_series = df_view.loc[mask_glosa, "_valor_glosa_abs"]
                elif vg and vg in df_view.columns:
                    order_series = df_view.loc[mask_glosa, vg].abs()
                else:
                    order_series = None

                if order_series is not None and not order_series.empty:
                    df_item = df_item.loc[order_series.sort_values(ascending=False).index]

                money_cols_fmt = [c for c in [vc, vg, vr] if c in df_item.columns]

                if not df_item.empty:
                    st.dataframe(
                        apply_currency(df_item, money_cols_fmt),
                        use_container_width=True,
                        height=420,
                    )
                else:
                    st.info(
                        "Nenhuma **guia com glosa** encontrada para este item no recorte atual. "
                        "Se quiser verificar todas as guias cobradas, use a busca por Nº AMHPTISS."
                    )

                base_cols = df_item.columns.tolist()
                st.download_button(
                    "⬇️ Baixar relação (CSV) — apenas guias com glosa",
                    data=df_item[base_cols].to_csv(index=False).encode("utf-8"),
                    file_name=f"guias_com_glosa_item_{re.sub(r'[^A-Za-z0-9_-]+','_', selected_item_name)[:40]}.csv",
                    mime="text/csv",
                )

        # Export análise XLSX (glosas) — mensal somando Valor Cobrado (Valor Original)
        st.markdown("---")
        st.subheader("📥 Exportar análise de Faturas Glosadas (XLSX)")
        from io import BytesIO
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as wr:
            k = analytics["kpis"] if analytics else dict(
                linhas=len(df_view), periodo_ini=None, periodo_fim=None,
                convenios=df_view[colmap["convenio"]].nunique() if colmap.get("convenio") in df_view.columns else 0,
                prestadores=df_view[colmap["prestador"]].nunique() if colmap.get("prestador") in df_view.columns else 0,
                valor_cobrado=float(df_view[colmap["valor_cobrado"]].sum()) if colmap.get("valor_cobrado") in df_view.columns else 0.0,
                valor_glosado=float(df_view["_valor_glosa_abs"].sum()) if "_valor_glosa_abs" in df_view.columns else 0.0,
                taxa_glosa=0.0
            )
            conv_sel = st.session_state.get("conv_glosas", "(todos)")
            modo_periodo = st.session_state.get("modo_periodo", "Todos os meses (agrupado)")
            mes_sel_label = st.session_state.get("mes_pagto_sel", "")

            kpi_df = pd.DataFrame([{
                "Convênio (filtro)": conv_sel,
                "Modo Período": modo_periodo,
                "Mês (se aplicado)": mes_sel_label or "",
                "Registros": k.get("linhas", ""),
                "Período Início": k.get("periodo_ini").strftime("%d/%m/%Y") if k.get("periodo_ini") else "",
                "Período Fim": k.get("periodo_fim").strftime("%d/%m/%Y") if k.get("periodo_fim") else "",
                "Convênios": k.get("convenios", ""),
                "Prestadores": k.get("prestadores", ""),
                "Valor Cobrado (R$)": round(k.get("valor_cobrado", 0.0), 2),
                "Valor Glosado (R$)": round(k.get("valor_glosado", 0.0), 2),
                "Taxa de Glosa (%)": round(k.get("taxa_glosa", 0.0) * 100, 2),
            }])
            kpi_df.to_excel(wr, index=False, sheet_name="KPIs")

            has_pagto = ("_pagto_dt" in df_view.columns) and df_view["_pagto_dt"].notna().any()
            if has_pagto:
                base_m = df_view[df_view["_is_glosa"] == True].copy()
                mensal = (base_m.groupby(["_pagto_ym","_pagto_mes_br"], as_index=False)
                                  .agg(Valor_Glosado=("_valor_glosa_abs","sum"),
                                       Valor_Cobrado=(colmap["valor_cobrado"], "sum"))
                         ).sort_values("_pagto_ym")
                mensal.rename(columns={"_pagto_ym":"YYYY-MM","_pagto_mes_br":"Mês/Ano"}, inplace=True)
                mensal.to_excel(wr, index=False, sheet_name="Mensal_Pagamento")

            if analytics and not analytics["top_motivos"].empty:
                analytics["top_motivos"].to_excel(wr, index=False, sheet_name="Top_Motivos")
            if analytics and not analytics["by_tipo"].empty:
                analytics["by_tipo"].to_excel(wr, index=False, sheet_name="Tipo_Glosa")
            if analytics and not analytics["top_itens"].empty:
                analytics["top_itens"].to_excel(wr, index=False, sheet_name="Top_Itens")
            if analytics and not analytics["by_convenio"].empty:
                analytics["by_convenio"].to_excel(wr, index=False, sheet_name="Convenios")

            col_export = [c for c in [
                colmap.get("amhptiss"),
                colmap.get("data_pagamento"),
                colmap.get("data_realizado"),
                colmap.get("convenio"), colmap.get("prestador"),
                colmap.get("descricao"), colmap.get("tipo_glosa"),
                colmap.get("motivo"), colmap.get("desc_motivo"),
                colmap.get("cobranca"),
                colmap.get("valor_cobrado"), colmap.get("valor_glosa"), colmap.get("valor_recursado")
            ] if c and c in df_view.columns]
            raw = df_view[col_export].copy() if col_export else pd.DataFrame()
            if not raw.empty:
                raw.to_excel(wr, index=False, sheet_name="Bruto_Selecionado")

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
            "⬇️ Baixar análise (XLSX)",
            data=buf.getvalue(),
            file_name="analise_faturas_glosadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if not glosas_files and not st.session_state.glosas_ready:
        st.info("Envie os arquivos e clique em **Processar Faturas Glosadas**.")

