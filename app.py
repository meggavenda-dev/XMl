
# file: app.py
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import List, Dict, Optional, Union, IO
from decimal import Decimal
from datetime import datetime
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st

# =========================================================
# Config & Header
# =========================================================
st.set_page_config(page_title="TISS • Itens por Guia + Conciliação + Auditoria", layout="wide")
st.title("TISS — Itens por Guia (XML) + Conciliação com Demonstrativo + Auditoria")
st.caption("Lê XML TISS (Consulta e SP‑SADT), concilia com Demonstrativo itemizado (motivos de glosa), gera rankings e auditoria — sem editor de XML.")

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
    if not s or not isinstance(s, str):
        return None
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def normalize_code(s: str, strip_zeros: bool = False) -> str:
    """Remove pontuação e espaços; opcionalmente retira zeros à esquerda."""
    if s is None:
        return ""
    s2 = re.sub(r'[\.\-_/ \t]', '', str(s)).strip()
    return s2.lstrip('0') if strip_zeros else s2


# =========================================================
# XML TISS → Itens por guia
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
    out: List[Dict] = []

    # procedimentosExecutados
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
            'valor_unitario': vuni if vuni > DEC_ZERO else (vtot if (vtot > DEC_ZERO) else DEC_ZERO),
            'valor_total': vtot if vtot > DEC_ZERO else (vuni * qtd if (vuni > DEC_ZERO and qtd > DEC_ZERO) else DEC_ZERO),
        })

    # outrasDespesas
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
            'valor_unitario': vuni if vuni > DEC_ZERO else (vtot if (vtot > DEC_ZERO) else DEC_ZERO),
            'valor_total': vtot if vtot > DEC_ZERO else (vuni * qtd if (vuni > DEC_ZERO and qtd > DEC_ZERO) else DEC_ZERO),
        })

    return out


def parse_itens_tiss_xml(source: Union[str, Path, IO[bytes]]) -> List[Dict]:
    """Extrai itens por guia (Consulta e SP‑SADT)."""
    if hasattr(source, 'read'):
        if hasattr(source, 'seek'):
            source.seek(0)
        root = ET.parse(source).getroot()
        arquivo_nome = Path(getattr(source, 'name', 'upload.xml')).name
    else:
        p = Path(source)
        root = ET.parse(p).getroot()
        arquivo_nome = p.name

    numero_lote = _get_numero_lote(root)
    out: List[Dict] = []

    # CONSULTA
    for guia in root.findall('.//ans:guiaConsulta', ANS_NS):
        numero_guia_prest = tx(guia.find('ans:numeroGuiaPrestador', ANS_NS))
        paciente = tx(guia.find('.//ans:dadosBeneficiario/ans:nomeBeneficiario', ANS_NS))
        medico   = tx(guia.find('.//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional', ANS_NS))
        data_atd = tx(guia.find('.//ans:dataAtendimento', ANS_NS))

        for it in _itens_consulta(guia):
            it.update({
                'arquivo': arquivo_nome,
                'numero_lote': numero_lote,
                'tipo_guia': 'CONSULTA',
                'numeroGuiaPrestador': numero_guia_prest,
                'numeroGuiaOperadora': '',
                'paciente': paciente,
                'medico': medico,
                'data_atendimento': data_atd,
            })
            out.append(it)

    # SADT
    for guia in root.findall('.//ans:guiaSP-SADT', ANS_NS):
        cab = guia.find('ans:cabecalhoGuia', ANS_NS)
        numero_guia_prest = tx(cab.find('ans:numeroGuiaPrestador', ANS_NS)) if cab is not None else ''
        numero_guia_oper  = tx(cab.find('ans:numeroGuiaOperadora', ANS_NS)) if cab is not None else ''
        paciente = tx(guia.find('.//ans:dadosBeneficiario/ans:nomeBeneficiario', ANS_NS))
        medico   = tx(guia.find('.//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional', ANS_NS))
        data_atd = tx(guia.find('.//ans:dataAtendimento', ANS_NS))

        for it in _itens_sadt(guia):
            it.update({
                'arquivo': arquivo_nome,
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
# Demonstrativo (.xlsx) → Itens + motivos de glosa
# =========================================================
_COLMAPS = {
    'lote'       : [r'^lote$'],
    'competencia': [r'^compet', r'^m[êe]s|^mes/?ano'],
    'guia_prest' : [r'prestador|guia\s*prest'],
    'guia_oper'  : [r'operadora|guia\s*oper'],
    'cod_proc'   : [r'c[oó]d.*proced|proced.*c[oó]d'],
    'desc_proc'  : [r'descri|proced.*descri'],
    'qtd_apres'  : [r'qtde|quant', r'apresent'],
    'qtd_paga'   : [r'qtde|quant', r'(paga|autori)'],
    'val_apres'  : [r'valor', r'apresent'],
    'val_glosa'  : [r'glosa'],
    'val_pago'   : [r'(valor.*(pago|apurado))|(pago$)|(apurado$)'],
    'motivo_cod' : [r'(motivo.*c[oó]d)|(c[oó]d.*motivo)'],
    'motivo_desc': [r'(descri.*motivo)|(motivo.*descri)'],
}


def _match_col(cols: List[str], pats: List[str]) -> Optional[str]:
    for c in cols:
        s = str(c).strip()
        s_norm = re.sub(r'\s+', ' ', s.lower())
        ok = True
        for p in pats:
            if not re.search(p, s_norm):
                ok = False
                break
        if ok:
            return s
    return None


def _find_header_row(df_raw: pd.DataFrame) -> int:
    # Heurística: mesma usada por você (linha com "CPF/CNPJ" na primeira coluna)
    s0 = df_raw.iloc[:, 0].astype(str).str.strip().str.lower()
    mask = s0.eq('cpf/cnpj')
    if mask.any():
        return int(mask.idxmax()) + 1  # header é a linha seguinte
    return 0


def ler_demo_itens_pagto_xlsx(source, strip_zeros_codes: bool = False) -> pd.DataFrame:
    """Lê planilha itemizada do Demonstrativo; detecta colunas por regex/heurística."""
    xls = pd.ExcelFile(source, engine='openpyxl')
    # escolher a planilha com "item" ou "análise"
    sheet = None
    for s in xls.sheet_names:
        s_norm = s.strip().lower()
        if 'item' in s_norm or 'analise' in s_norm or 'análise' in s_norm:
            sheet = s
            break
    if sheet is None:
        sheet = xls.sheet_names[0]

    df_raw = pd.read_excel(source, sheet_name=sheet, engine='openpyxl')
    hdr = _find_header_row(df_raw)
    df = df_raw.copy()
    if hdr > 0:
        df.columns = df_raw.iloc[hdr]
        df = df_raw.iloc[hdr + 1:].reset_index(drop=True)

    cols = [str(c) for c in df.columns]
    pick = {k: _match_col(cols, v) for k, v in _COLMAPS.items()}

    if not any(pick.get(c) for c in ['val_apres', 'val_glosa', 'val_pago', 'cod_proc']):
        raise ValueError("Não identifiquei colunas itemizadas no Demonstrativo. Anexe um exemplo para mapeamento.")

    def col(c): return pick.get(c)

    out = pd.DataFrame({
        'numero_lote': df[col('lote')] if col('lote') else None,
        'competencia': df[col('competencia')] if col('competencia') else None,
        'numeroGuiaPrestador': (df[col('guia_prest')] if col('guia_prest') else None),
        'numeroGuiaOperadora': (df[col('guia_oper')] if col('guia_oper') else None),
        'codigo_procedimento': df[col('cod_proc')] if col('cod_proc') else None,
        'descricao_procedimento': df[col('desc_proc')] if col('desc_proc') else None,
        'quantidade_apresentada': pd.to_numeric(df[col('qtd_apres')], errors='coerce') if col('qtd_apres') else 0.0,
        'quantidade_paga': pd.to_numeric(df[col('qtd_paga')], errors='coerce') if col('qtd_paga') else 0.0,
        'valor_apresentado': pd.to_numeric(df[col('val_apres')], errors='coerce') if col('val_apres') else 0.0,
        'valor_glosa': pd.to_numeric(df[col('val_glosa')], errors='coerce') if col('val_glosa') else 0.0,
        'valor_pago': pd.to_numeric(df[col('val_pago')], errors='coerce') if col('val_pago') else 0.0,
        'motivo_glosa_codigo': df[col('motivo_cod')] if col('motivo_cod') else None,
        'motivo_glosa_descricao': df[col('motivo_desc')] if col('motivo_desc') else None,
    })

    # normalizações
    for c in ['numero_lote', 'numeroGuiaPrestador', 'numeroGuiaOperadora', 'codigo_procedimento']:
        if c in out.columns and out[c] is not None:
            out[c] = out[c].astype(str).str.strip()

    # códigos normalizados
    out['codigo_procedimento_norm'] = out['codigo_procedimento'].astype(str).map(
        lambda s: normalize_code(s, strip_zeros=strip_zeros_codes)
    )

    for c in ['valor_apresentado', 'valor_glosa', 'valor_pago', 'quantidade_apresentada', 'quantidade_paga']:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors='coerce').fillna(0.0)

    # chaves
    out['chave_prest'] = (out.get('numeroGuiaPrestador', '').astype(str).str.strip()
                          + '__' + out['codigo_procedimento_norm'].astype(str).str.strip())
    out['chave_oper']  = (out.get('numeroGuiaOperadora', '').astype(str).str.strip()
                          + '__' + out['codigo_procedimento_norm'].astype(str).str.strip())
    return out


# =========================================================
# Conciliação item a item
# =========================================================
def build_xml_df(xml_files, strip_zeros_codes: bool = False) -> pd.DataFrame:
    linhas: List[Dict] = []
    for f in xml_files:
        if hasattr(f, 'seek'):
            f.seek(0)
        try:
            linhas.extend(parse_itens_tiss_xml(f))
        except Exception as e:
            linhas.append({'arquivo': getattr(f, 'name', 'upload.xml'), 'erro': str(e)})

    df = pd.DataFrame(linhas)
    if df.empty:
        return df

    # tipos numéricos
    for c in ['quantidade', 'valor_unitario', 'valor_total']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

    # normaliza código para chave
    df['codigo_procedimento_norm'] = df['codigo_procedimento'].astype(str).map(
        lambda s: normalize_code(s, strip_zeros=strip_zeros_codes)
    )

    # chaves
    df['chave_prest'] = (df['numeroGuiaPrestador'].fillna('').astype(str).str.strip()
                         + '__' + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip())
    df['chave_oper'] = (df['numeroGuiaOperadora'].fillna('').astype(str).str.strip()
                        + '__' + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip())
    return df


def build_demo_df(demo_files, strip_zeros_codes: bool = False) -> pd.DataFrame:
    parts = []
    for f in demo_files:
        if hasattr(f, 'seek'):
            f.seek(0)
        parts.append(ler_demo_itens_pagto_xlsx(f, strip_zeros_codes=strip_zeros_codes))
    if parts:
        return pd.concat(parts, ignore_index=True)
    return pd.DataFrame()


def conciliar_itens(
    df_xml: pd.DataFrame,
    df_demo: pd.DataFrame,
    tolerance_valor: float = 0.02,
    fallback_por_descricao: bool = False,
) -> Dict[str, pd.DataFrame]:
    """
    1) match por chave_prest
    2) não casados → match por chave_oper
    3) opcional: fallback por descrição + tolerância de valor
    """
    # 1) Prestador
    m1 = df_xml.merge(df_demo, on='chave_prest', how='left', suffixes=('_xml', '_demo'))
    m1['matched_on'] = m1['valor_apresentado'].notna().map({True: 'prestador', False: ''})

    # 2) Operadora
    not_match = m1[m1['matched_on'].eq('')].copy()
    still_xml = not_match[['arquivo', 'numero_lote', 'tipo_guia', 'numeroGuiaPrestador', 'numeroGuiaOperadora',
                           'paciente', 'medico', 'data_atendimento', 'tipo_item', 'identificadorDespesa',
                           'codigo_tabela', 'codigo_procedimento', 'codigo_procedimento_norm',
                           'descricao_procedimento', 'quantidade', 'valor_unitario', 'valor_total',
                           'chave_oper', 'chave_prest']]
    m2 = still_xml.merge(df_demo, on='chave_oper', how='left', suffixes=('_xml', '_demo'))
    m2['matched_on'] = m2['valor_apresentado'].notna().map({True: 'operadora', False: ''})

    conc = pd.concat([m1[m1['matched_on'] != ''], m2[m2['matched_on'] != '']], ignore_index=True)

    # 3) Fallback por descrição (opcional)
    fallback_matches = pd.DataFrame()
    if fallback_por_descricao:
        rem_xml = pd.concat([m1[m1['matched_on'] == ''], m2[m2['matched_on'] == '']], ignore_index=True)
        if not rem_xml.empty:
            # chaves por guia + descrição — tolerância por valor
            # Fazemos join por guia (prestador primeiro, senão operadora) + descricao_procedimento
            rem_xml['guia_join'] = rem_xml.apply(
                lambda r: r['numeroGuiaPrestador'] if str(r.get('numeroGuiaPrestador', '')).strip() else str(r.get('numeroGuiaOperadora', '')).strip(),
                axis=1
            )
            df_demo2 = df_demo.copy()
            df_demo2['guia_join'] = df_demo2.apply(
                lambda r: r['numeroGuiaPrestador'] if str(r.get('numeroGuiaPrestador', '')).strip() else str(r.get('numeroGuiaOperadora', '')).strip(),
                axis=1
            )
            tmp = rem_xml.merge(
                df_demo2,
                on=['guia_join', 'descricao_procedimento'],
                how='left',
                suffixes=('_xml', '_demo')
            )
            # filtro por tolerância de valor apresentado x valor_total
            tol = float(tolerance_valor)
            keep = (tmp['valor_apresentado'].notna()) & ((tmp['valor_total'] - tmp['valor_apresentado']).abs() <= tol)
            fallback_matches = tmp[keep].copy()
            if not fallback_matches.empty:
                fallback_matches['matched_on'] = 'descricao+valor'
                conc = pd.concat([conc, fallback_matches], ignore_index=True)

    # não casados finais
    unmatch = pd.concat(
        [
            m1[m1['matched_on'] == ''],
            m2[m2['matched_on'] == ''],
            fallback_matches[fallback_matches.get('matched_on', '') == ''].copy() if not fallback_matches.empty else pd.DataFrame()
        ],
        ignore_index=True
    )
    if not unmatch.empty:
        unmatch = unmatch.drop_duplicates(
            subset=['arquivo', 'numero_lote', 'tipo_guia', 'numeroGuiaPrestador', 'codigo_procedimento', 'valor_total']
        )

    # diffs
    if not conc.empty:
        conc['apresentado_diff'] = conc['valor_total'] - conc['valor_apresentado']
        conc['glosa_pct'] = conc.apply(
            lambda r: (r['valor_glosa'] / r['valor_apresentado']) if (r.get('valor_apresentado', 0) and r['valor_apresentado'] > 0) else 0.0,
            axis=1
        )

    return {
        'conciliacao': conc,
        'nao_casados': unmatch
    }


# =========================================================
# Auditoria (duplicidade de guia & retorno em dias)
# =========================================================
def build_chave_guia(tipo: str, numeroGuiaPrestador: str, numeroGuiaOperadora: str) -> Optional[str]:
    t = (tipo or '').upper()
    if t in ('CONSULTA', 'SADT'):
        return str(numeroGuiaPrestador).strip() if numeroGuiaPrestador else None
    return None  # (Recurso não entra aqui; auditoria baseada em XML de guias assistenciais)


def auditar_guias(df_xml_itens: pd.DataFrame, prazo_retorno: int = 30) -> pd.DataFrame:
    """Reduz itens → nível guia e audita duplicidade/retorno."""
    if df_xml_itens.empty:
        return pd.DataFrame()

    base_cols = [
        'arquivo', 'numero_lote', 'tipo_guia',
        'numeroGuiaPrestador', 'numeroGuiaOperadora',
        'paciente', 'medico', 'data_atendimento'
    ]
    base = df_xml_itens[base_cols].drop_duplicates().copy()
    base['chave_guia'] = base.apply(
        lambda r: build_chave_guia(r['tipo_guia'], r['numeroGuiaPrestador'], r['numeroGuiaOperadora']),
        axis=1
    )
    base['duplicada'] = False
    base['arquivos_duplicados'] = ''
    base['lotes_duplicados'] = ''
    base['retorno_no_periodo'] = False
    base['retorno_ref'] = ''
    base['status_auditoria'] = ''

    # mapa por chave
    mapa = {}
    for i, r in base.iterrows():
        k = r.get('chave_guia')
        if not k:
            continue
        mapa.setdefault(k, []).append((i, r.get('numero_lote', ''), r.get('arquivo', '')))

    # duplicidade
    for i, r in base.iterrows():
        k = r.get('chave_guia')
        if not k or k not in mapa:
            continue
        outros = [(j, lot, arq) for (j, lot, arq) in mapa[k] if j != i]
        if outros:
            base.loc[i, 'duplicada'] = True
            base.loc[i, 'lotes_duplicados'] = ",".join(sorted({o[1] for o in outros if o[1]}))
            base.loc[i, 'arquivos_duplicados'] = ",".join(sorted({o[2] for o in outros if o[2]}))

    # retorno (paciente + médico ± prazo em dias)
    datas = {i: parse_date_flex(str(r.get('data_atendimento') or '').strip()) for i, r in base.iterrows()}
    if prazo_retorno and prazo_retorno > 0:
        for i, r in base.iterrows():
            pac = (r.get('paciente') or '').strip()
            med = (r.get('medico') or '').strip()
            d0 = datas.get(i)
            if not pac or not med or not d0:
                continue
            candidatos = base[(base.index != i)
                              & (base['paciente'].fillna('').str.strip() == pac)
                              & (base['medico'].fillna('').str.strip() == med)]
            refs = []
            for j, rr in candidatos.iterrows():
                dj = datas.get(j)
                if not dj:
                    continue
                if abs((d0 - dj).days) <= prazo_retorno:
                    refs.append(f"{rr.get('numero_lote','')}@{rr.get('arquivo','')}@{(rr.get('data_atendimento') or '')}")
            if refs:
                base.loc[i, 'retorno_no_periodo'] = True
                base.loc[i, 'retorno_ref'] = " | ".join(refs)

    # status
    for i, r in base.iterrows():
        status = []
        if r['duplicada']:
            status.append("Duplicidade")
        if r['retorno_no_periodo']:
            status.append("Retorno")
        base.loc[i, 'status_auditoria'] = " + ".join(status) if status else "OK"

    return base


# =========================================================
# UI — Sidebar (parâmetros)
# =========================================================
with st.sidebar:
    st.header("Parâmetros")
    prazo_retorno = st.number_input("Prazo de retorno (dias)", min_value=0, value=30, step=1)
    tolerance_valor = st.number_input("Tolerância p/ fallback por descrição (R$)", min_value=0.00, value=0.02, step=0.01, format="%.2f")
    fallback_desc = st.toggle("Fallback de conciliação por descrição + valor (quando código não casar)", value=False)
    strip_zeros_codes = st.toggle("Normalizar códigos removendo zeros à esquerda", value=True)

# =========================================================
# UI — Uploads
# =========================================================
st.subheader("Upload de arquivos")
xml_files = st.file_uploader("XML TISS (um ou mais):", type=['xml'], accept_multiple_files=True)
demo_files = st.file_uploader("Demonstrativo(s) de Pagamento (.xlsx) — itemizado:", type=['xlsx'], accept_multiple_files=True)

if st.button("Processar", type="primary"):
    # 1) XML → Itens
    df_xml = build_xml_df(xml_files or [], strip_zeros_codes=strip_zeros_codes)
    if df_xml.empty:
        st.warning("Nenhum item extraído do(s) XML(s).")
    else:
        st.subheader("Itens extraídos do XML (Consulta / SP‑SADT)")
        st.dataframe(apply_currency(df_xml, ['valor_unitario', 'valor_total']), use_container_width=True, height=420)

    # 2) Demonstrativo → Itens
    df_demo = build_demo_df(demo_files or [], strip_zeros_codes=strip_zeros_codes)
    if df_demo.empty:
        st.info("Nenhum demonstrativo válido carregado (ou sem colunas itemizadas).")
    else:
        st.subheader("Itens do Demonstrativo (detectados)")
        st.dataframe(apply_currency(df_demo, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True, height=420)

    # 3) Conciliação
    if not df_xml.empty and not df_demo.empty:
        result = conciliar_itens(
            df_xml=df_xml,
            df_demo=df_demo,
            tolerance_valor=float(tolerance_valor),
            fallback_por_descricao=fallback_desc
        )
        conc = result['conciliacao']
        unmatch = result['nao_casados']

        st.subheader("Conciliação item a item (XML × Demonstrativo)")
        conc_disp = apply_currency(conc.copy(), ['valor_unitario', 'valor_total', 'valor_apresentado', 'valor_glosa', 'valor_pago', 'apresentado_diff'])
        st.dataframe(conc_disp, use_container_width=True, height=480)

        c1, c2 = st.columns(2)
        with c1:
            st.metric("Itens casados", len(conc))
        with c2:
            st.metric("Itens não casados", len(unmatch))

        if not unmatch.empty:
            st.markdown("**Itens não casados**")
            st.dataframe(apply_currency(unmatch.copy(), ['valor_unitario', 'valor_total']), use_container_width=True, height=320)

        # 4) Indicadores
        st.markdown("---")
        st.subheader("📊 Indicadores & Rankings")

        colA, colB = st.columns(2)

        with colA:
            st.markdown("#### Motivos de Glosa (Top)")
            if 'motivo_glosa_codigo' in conc.columns or 'motivo_glosa_descricao' in conc.columns:
                mot = (conc.groupby(['motivo_glosa_codigo', 'motivo_glosa_descricao'], dropna=False, as_index=False)
                       .agg(valor_glosa=('valor_glosa', 'sum'),
                            valor_apresentado=('valor_apresentado', 'sum'),
                            itens=('codigo_procedimento', 'count')))
                mot['glosa_pct'] = mot.apply(
                    lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1
                )
                mot = mot.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
                st.dataframe(apply_currency(mot, ['valor_glosa', 'valor_apresentado']), use_container_width=True)
            else:
                st.info("Motivo de glosa não presente no demonstrativo.")

        with colB:
            st.markdown("#### Procedimentos com maior índice de Glosa")
            if not conc.empty:
                proc = (conc.groupby(['codigo_procedimento', 'descricao_procedimento'], dropna=False, as_index=False)
                        .agg(valor_apresentado=('valor_apresentado', 'sum'),
                             valor_glosa=('valor_glosa', 'sum'),
                             valor_pago=('valor_pago', 'sum'),
                             itens=('arquivo', 'count')))
                proc['glosa_pct'] = proc.apply(
                    lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1
                )
                proc = proc.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
                st.dataframe(apply_currency(proc, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True)
            else:
                st.info("Sem dados conciliados para ranking de procedimentos.")

        # Rankings adicionais
        colC, colD = st.columns(2)
        with colC:
            st.markdown("#### Médicos com maior glosa")
            med = (conc.groupby(['medico'], dropna=False, as_index=False)
                   .agg(valor_apresentado=('valor_apresentado', 'sum'),
                        valor_glosa=('valor_glosa', 'sum'),
                        valor_pago=('valor_pago', 'sum'),
                        itens=('arquivo', 'count')))
            med['glosa_pct'] = med.apply(lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1)
            med = med.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
            st.dataframe(apply_currency(med, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True)

        with colD:
            st.markdown("#### Lotes com maior glosa")
            lot = (conc.groupby(['numero_lote'], dropna=False, as_index=False)
                   .agg(valor_apresentado=('valor_apresentado', 'sum'),
                        valor_glosa=('valor_glosa', 'sum'),
                        valor_pago=('valor_pago', 'sum'),
                        itens=('arquivo', 'count')))
            lot['glosa_pct'] = lot.apply(lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1)
            lot = lot.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
            st.dataframe(apply_currency(lot, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True)

        # 5) Auditoria por guia (duplicidade e retorno)
        st.markdown("---")
        st.subheader("🔎 Auditoria por Guia (duplicidade e retorno)")
        df_aud = auditar_guias(df_xml, prazo_retorno=prazo_retorno)
        if df_aud.empty:
            st.info("Sem dados para auditoria (verifique os XML).")
        else:
            st.dataframe(df_aud, use_container_width=True, height=420)

        # 6) Export Excel consolidado
        st.markdown("---")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as wr:
            (df_xml if not df_xml.empty else pd.DataFrame()).to_excel(wr, index=False, sheet_name='Itens_XML')
            (df_demo if not df_demo.empty else pd.DataFrame()).to_excel(wr, index=False, sheet_name='Itens_Demo')
            (conc if not conc.empty else pd.DataFrame()).to_excel(wr, index=False, sheet_name='Conciliação')
            (unmatch if not unmatch.empty else pd.DataFrame()).to_excel(wr, index=False, sheet_name='Nao_Casados')

            if not conc.empty:
                mot_x = (conc.groupby(['motivo_glosa_codigo', 'motivo_glosa_descricao'], dropna=False, as_index=False)
                         .agg(valor_glosa=('valor_glosa', 'sum'),
                              valor_apresentado=('valor_apresentado', 'sum'),
                              itens=('codigo_procedimento', 'count')))
                mot_x['glosa_pct'] = mot_x.apply(lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1)
                mot_x.to_excel(wr, index=False, sheet_name='Motivos_Glosa')

                proc_x = (conc.groupby(['codigo_procedimento', 'descricao_procedimento'], dropna=False, as_index=False)
                          .agg(valor_apresentado=('valor_apresentado', 'sum'),
                               valor_glosa=('valor_glosa', 'sum'),
                               valor_pago=('valor_pago', 'sum'),
                               itens=('arquivo', 'count')))
                proc_x['glosa_pct'] = proc_x.apply(lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1)
                proc_x.to_excel(wr, index=False, sheet_name='Procedimentos_Glosa')

                med_x = (conc.groupby(['medico'], dropna=False, as_index=False)
                         .agg(valor_apresentado=('valor_apresentado', 'sum'),
                              valor_glosa=('valor_glosa', 'sum'),
                              valor_pago=('valor_pago', 'sum'),
                              itens=('arquivo', 'count')))
                med_x['glosa_pct'] = med_x.apply(lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1)
                med_x.to_excel(wr, index=False, sheet_name='Medicos')

                lot_x = (conc.groupby(['numero_lote'], dropna=False, as_index=False)
                         .agg(valor_apresentado=('valor_apresentado', 'sum'),
                              valor_glosa=('valor_glosa', 'sum'),
                              valor_pago=('valor_pago', 'sum'),
                              itens=('arquivo', 'count')))
                lot_x['glosa_pct'] = lot_x.apply(lambda r: (r['valor_glosa'] / r['valor_apresentado']) if r['valor_apresentado'] > 0 else 0.0, axis=1)
                lot_x.to_excel(wr, index=False, sheet_name='Lotes')

            (df_aud if not df_aud.empty else pd.DataFrame()).to_excel(wr, index=False, sheet_name='Auditoria_Guias')

            # ajustes (largura + congela cabeçalho)
            for name in wr.sheets:
                ws = wr.sheets[name]
                ws.freeze_panes = "A2"
                for col in ws.columns:
                    try:
                        col_letter = col[0].column_letter
                    except Exception:
                        continue
                    max_len = 12
                    for cell in col:
                        v = cell.value
                        if v is None:
                            continue
                        s = str(v)
                        if len(s) > max_len:
                            max_len = len(s)
                    ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

        st.download_button(
            "Baixar Excel consolidado",
            data=buf.getvalue(),
            file_name="tiss_itens_conciliacao_auditoria.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.info("Para conciliação, carregue ao menos um XML e um Demonstrativo.")
