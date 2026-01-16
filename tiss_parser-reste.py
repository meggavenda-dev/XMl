
# file: tiss_items_parser.py
from __future__ import annotations

from typing import IO, Union, List, Dict, Optional, Tuple
from pathlib import Path
from decimal import Decimal
import re
import pandas as pd
import xml.etree.ElementTree as ET

ANS_NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}
DEC_ZERO = Decimal('0')

# ----------------------------
# Helpers XML
# ----------------------------
def _dec(txt: Optional[str]) -> Decimal:
    if txt is None:
        return DEC_ZERO
    s = str(txt).strip().replace(',', '.')
    return Decimal(s) if s else DEC_ZERO

def _tx(el: Optional[ET.Element]) -> str:
    return (el.text or '').strip() if (el is not None and el.text) else ''

def _find(root: ET.Element, xp: str) -> Optional[ET.Element]:
    return root.find(xp, ANS_NS)

def _findall(root: ET.Element, xp: str) -> List[ET.Element]:
    return root.findall(xp, ANS_NS)

def _get_numero_lote(root: ET.Element) -> str:
    el = _find(root, './/ans:prestadorParaOperadora/ans:loteGuias/ans:numeroLote')
    if el is not None and _tx(el):
        return _tx(el)
    el = _find(root, './/ans:prestadorParaOperadora/ans:recursoGlosa/ans:guiaRecursoGlosa/ans:numeroLote')
    if el is not None and _tx(el):
        return _tx(el)
    return ""

def _is_consulta(root: ET.Element) -> bool:
    return _find(root, './/ans:guiaConsulta') is not None

def _is_sadt(root: ET.Element) -> bool:
    return _find(root, './/ans:guiaSP-SADT') is not None

# ----------------------------
# Extração item a item — CONSULTA
# ----------------------------
def _itens_consulta(guia: ET.Element) -> List[Dict]:
    # Em consulta, em geral há 1 procedimento (procedimento → codigoTabela/ codigoProcedimento/ descricaoProcedimento / valorProcedimento)
    proc = _find(guia, './/ans:procedimento')
    codigo_tabela = _tx(_find(proc, 'ans:codigoTabela')) if proc is not None else ''
    codigo_proc   = _tx(_find(proc, 'ans:codigoProcedimento')) if proc is not None else ''
    descricao     = _tx(_find(proc, 'ans:descricaoProcedimento')) if proc is not None else ''
    valor         = _dec(_tx(_find(proc, 'ans:valorProcedimento'))) if proc is not None else DEC_ZERO

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

# ----------------------------
# Extração item a item — SADT
# ----------------------------
def _itens_sadt(guia: ET.Element) -> List[Dict]:
    out: List[Dict] = []

    # procedimentosExecutados/procedimentoExecutado
    for it in _findall(guia, './/ans:procedimentosExecutados/ans:procedimentoExecutado'):
        proc = _find(it, 'ans:procedimento')
        codigo_tabela = _tx(_find(proc, 'ans:codigoTabela')) if proc is not None else ''
        codigo_proc   = _tx(_find(proc, 'ans:codigoProcedimento')) if proc is not None else ''
        descricao     = _tx(_find(proc, 'ans:descricaoProcedimento')) if proc is not None else ''

        qtd  = _dec(_tx(_find(it, 'ans:quantidadeExecutada')))
        vuni = _dec(_tx(_find(it, 'ans:valorUnitario')))
        vtot = _dec(_tx(_find(it, 'ans:valorTotal')))
        if vtot == DEC_ZERO and (vuni > DEC_ZERO and qtd > DEC_ZERO):
            vtot = vuni * qtd

        out.append({
            'tipo_item': 'procedimento',
            'identificadorDespesa': '',
            'codigo_tabela': codigo_tabela,
            'codigo_procedimento': codigo_proc,
            'descricao_procedimento': descricao,
            'quantidade': qtd if qtd > DEC_ZERO else Decimal('1'),
            'valor_unitario': vuni if vuni > DEC_ZERO else (vtot if (vtot>DEC_ZERO) else DEC_ZERO),
            'valor_total': vtot if vtot > DEC_ZERO else (vuni*qtd if (vuni>DEC_ZERO and qtd>DEC_ZERO) else DEC_ZERO),
        })

    # outrasDespesas/despesa/servicosExecutados
    for desp in _findall(guia, './/ans:outrasDespesas/ans:despesa'):
        ident = _tx(_find(desp, 'ans:identificadorDespesa'))
        sv = _find(desp, 'ans:servicosExecutados')
        codigo_tabela = _tx(_find(sv, 'ans:codigoTabela')) if sv is not None else ''
        codigo_proc   = _tx(_find(sv, 'ans:codigoProcedimento')) if sv is not None else ''
        descricao     = _tx(_find(sv, 'ans:descricaoProcedimento')) if sv is not None else ''
        qtd  = _dec(_tx(_find(sv, 'ans:quantidadeExecutada'))) if sv is not None else DEC_ZERO
        vuni = _dec(_tx(_find(sv, 'ans:valorUnitario')))      if sv is not None else DEC_ZERO
        vtot = _dec(_tx(_find(sv, 'ans:valorTotal')))         if sv is not None else DEC_ZERO
        if vtot == DEC_ZERO and (vuni > DEC_ZERO and qtd > DEC_ZERO):
            vtot = vuni * qtd

        out.append({
            'tipo_item': 'outra_despesa',
            'identificadorDespesa': ident,
            'codigo_tabela': codigo_tabela,
            'codigo_procedimento': codigo_proc,
            'descricao_procedimento': descricao,
            'quantidade': qtd if qtd > DEC_ZERO else Decimal('1'),
            'valor_unitario': vuni if vuni > DEC_ZERO else (vtot if (vtot>DEC_ZERO) else DEC_ZERO),
            'valor_total': vtot if vtot > DEC_ZERO else (vuni*qtd if (vuni>DEC_ZERO and qtd>DEC_ZERO) else DEC_ZERO),
        })

    return out

# ----------------------------
# API pública — Itens por guia (XML)
# ----------------------------
def parse_itens_tiss_xml(source: Union[str, Path, IO[bytes]]) -> List[Dict]:
    """Extrai itens por guia (Consulta e SP-SADT). Recurso de Glosa não gera itens aqui."""
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
    for guia in _findall(root, './/ans:guiaConsulta'):
        numero_guia_prest = _tx(_find(guia, 'ans:numeroGuiaPrestador'))
        paciente = _tx(_find(guia, './/ans:dadosBeneficiario/ans:nomeBeneficiario'))
        medico   = _tx(_find(guia, './/ans:dadosProfissionaisResponsaveis/ans:nomeProfissional'))
        data_atd = _tx(_find(guia, './/ans:dataAtendimento'))

        for it in _itens_consulta(guia):
            it.update({
                'arquivo': arquivo_nome,
                'numero_lote': numero_lote,
                'tipo_guia': 'CONSULTA',
                'numeroGuiaPrestador': numero_guia_prest,
                'numeroGuiaOperadora': '',  # não presente na guia de consulta
                'paciente': paciente,
                'medico': medico,
                'data_atendimento': data_atd,
            })
            out.append(it)

    # SADT
    for guia in _findall(root, './/ans:guiaSP-SADT'):
        cab = _find(guia, 'ans:cabecalhoGuia')
        numero_guia_prest = _tx(_find(cab, 'ans:numeroGuiaPrestador')) if cab is not None else ''
        numero_guia_oper  = _tx(_find(cab, 'ans:numeroGuiaOperadora')) if cab is not None else ''
        paciente = _tx(_find(guia, './/ans:dadosBeneficiario/ans:nomeBeneficiario'))
        medico   = _tx(_find(guia, './/ans:dadosProfissionaisResponsaveis/ans:nomeProfissional'))
        data_atd = _tx(_find(guia, './/ans:dataAtendimento'))

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

def parse_itens_many(paths: List[Union[str, Path]]) -> pd.DataFrame:
    linhas: List[Dict] = []
    for p in paths:
        try:
            linhas.extend(parse_itens_tiss_xml(p))
        except Exception as e:
            linhas.append({'arquivo': Path(p).name if hasattr(p, 'name') else str(p),
                           'erro': str(e)})
    df = pd.DataFrame(linhas)
    # Normalizações
    if not df.empty:
        for c in ['quantidade', 'valor_unitario', 'valor_total']:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
        # Chaves de casamento
        df['chave_prest'] = (df.get('numeroGuiaPrestador', '').astype(str).str.strip()
                             + '__' + df.get('codigo_procedimento', '').astype(str).str.strip())
        df['chave_oper']  = (df.get('numeroGuiaOperadora', '').astype(str).str.strip()
                             + '__' + df.get('codigo_procedimento', '').astype(str).str.strip())
    return df

# ----------------------------
# Leitura Demonstrativo (.xlsx) — Itens
# ----------------------------
_COLMAPS = {
    'lote':       [r'^lote$'],
    'competencia':[r'^compet', r'^m[êe]s', r'^mes/?ano'],
    'guia_prest':[r'prestador', r'guia\s*prest'],
    'guia_oper': [r'operadora', r'guia\s*oper'],
    'cod_proc':  [r'c[oó]d.*proced', r'proced.*c[oó]d'],
    'desc_proc': [r'descri', r'proced.*descri'],
    'qtd_apres': [r'qtde|quant', r'apresent'],
    'qtd_paga':  [r'qtde|quant', r'(paga|autori)'],
    'val_apres': [r'valor', r'apresent'],
    'val_glosa': [r'glosa'],
    'val_pago':  [r'(valor.*(pago|apurado))|(pago$)|(apurado$)'],
    'motivo_cod':[r'(motivo.*c[oó]d)|(c[oó]d.*motivo)'],
    'motivo_desc':[r'(descri.*motivo)|(motivo.*descri)'],
}

def _match_col(cols: List[str], pats: List[str]) -> Optional[str]:
    for c in cols:
        s = str(c).strip()
        s_norm = re.sub(r'\s+', ' ', s.lower())
        ok = True
        for p in pats:
            if not re.search(p, s_norm):
                ok = False; break
        if ok:
            return s
    return None

def _find_header_row(df_raw: pd.DataFrame) -> int:
    # usa a mesma heurística anterior: procurar linha com "CPF/CNPJ" na 1ª coluna, senão assume a primeira linha é cabeçalho
    s0 = df_raw.iloc[:,0].astype(str).str.strip().str.lower()
    mask = s0.eq('cpf/cnpj')
    if mask.any():
        return int(mask.idxmax()) + 1  # header é a linha seguinte
    return 0

def ler_demo_itens_pagto_xlsx(source) -> pd.DataFrame:
    """Lê planilha itemizada do Demonstrativo; tenta detectar colunas por regex."""
    xls = pd.ExcelFile(source, engine='openpyxl')
    # tenta achar a(s) sheet(s) com “Analise” / “Item”
    sheet = None
    for s in xls.sheet_names:
        s_norm = s.strip().lower()
        if 'item' in s_norm or 'analise' in s_norm or 'análise' in s_norm:
            sheet = s; break
    if sheet is None:
        sheet = xls.sheet_names[0]

    df_raw = pd.read_excel(source, sheet_name=sheet, engine='openpyxl')
    hdr = _find_header_row(df_raw)
    df = df_raw.copy()
    if hdr > 0:
        df.columns = df_raw.iloc[hdr]
        df = df_raw.iloc[hdr+1:].reset_index(drop=True)

    cols = [str(c) for c in df.columns]
    pick = {k: _match_col(cols, v) for k, v in _COLMAPS.items()}

    need_any = ['val_apres', 'val_glosa', 'val_pago', 'cod_proc']
    if not any(pick.get(c) for c in need_any):
        raise ValueError("Não identifiquei colunas itemizadas no Demonstrativo. Envie um exemplo para mapearmos.")

    def col(c): return pick.get(c)

    out = pd.DataFrame({
        'numero_lote'        : df[col('lote')] if col('lote') else None,
        'competencia'        : df[col('competencia')] if col('competencia') else None,
        'numeroGuiaPrestador': (df[col('guia_prest')] if col('guia_prest') else None),
        'numeroGuiaOperadora': (df[col('guia_oper')]  if col('guia_oper')  else None),
        'codigo_procedimento': df[col('cod_proc')] if col('cod_proc') else None,
        'descricao_procedimento': df[col('desc_proc')] if col('desc_proc') else None,
        'quantidade_apresentada': pd.to_numeric(df[col('qtd_apres')], errors='coerce') if col('qtd_apres') else 0.0,
        'quantidade_paga'       : pd.to_numeric(df[col('qtd_paga')], errors='coerce')  if col('qtd_paga')  else 0.0,
        'valor_apresentado'     : pd.to_numeric(df[col('val_apres')], errors='coerce') if col('val_apres') else 0.0,
        'valor_glosa'           : pd.to_numeric(df[col('val_glosa')], errors='coerce') if col('val_glosa') else 0.0,
        'valor_pago'            : pd.to_numeric(df[col('val_pago')], errors='coerce')  if col('val_pago')  else 0.0,
        'motivo_glosa_codigo'   : df[col('motivo_cod')] if col('motivo_cod') else None,
        'motivo_glosa_descricao': df[col('motivo_desc')] if col('motivo_desc') else None,
    })

    # normalizações
    for c in ['numero_lote', 'numeroGuiaPrestador', 'numeroGuiaOperadora', 'codigo_procedimento']:
        if c in out.columns:
            out[c] = out[c].astype(str).str.strip()
    for c in ['valor_apresentado', 'valor_glosa', 'valor_pago', 'quantidade_apresentada', 'quantidade_paga']:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors='coerce').fillna(0.0)

    # chaves
    out['chave_prest'] = (out.get('numeroGuiaPrestador', '').astype(str).str.strip()
                          + '__' + out.get('codigo_procedimento', '').astype(str).str.strip())
    out['chave_oper']  = (out.get('numeroGuiaOperadora', '').astype(str).str.strip()
                          + '__' + out.get('codigo_procedimento', '').astype(str).str.strip())
    return out
