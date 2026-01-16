
# =========================================================
# app.py — TISS XML + Conciliação + Auditoria
# =========================================================
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import List, Dict, Optional, Union, IO
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
    page_title="TISS • Itens por Guia + Conciliação + Auditoria",
    layout="wide"
)
st.title("TISS — Itens por Guia (XML) + Conciliação com Demonstrativo + Auditoria")
st.caption("Lê XML TISS (Consulta / SADT), concilia com Demonstrativo itemizado, gera rankings e auditoria — sem editor de XML.")

# =========================================================
# Helpers gerais
# =========================================================

ANS_NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}
DEC_ZERO = Decimal('0')

def dec(txt: Optional[str]) -> Decimal:
    """Converte texto para Decimal."""
    if txt is None:
        return DEC_ZERO
    s = str(txt).strip().replace(',', '.')
    return Decimal(s) if s else DEC_ZERO

def tx(el: Optional[ET.Element]) -> str:
    """Extrai texto limpo de tag XML."""
    return (el.text or '').strip() if (el is not None and el.text) else ''

def f_currency(v: Union[int, float, Decimal, str]) -> str:
    """Formata valores monetários."""
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
    """Aplica formatação de moeda a colunas."""
    d = df.copy()
    for c in cols:
        if c in d.columns:
            d[c] = d[c].apply(f_currency)
    return d

def parse_date_flex(s: str) -> Optional[datetime]:
    """Tenta interpretar datas em formatos variados."""
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
    """Normaliza código TUSS removendo pontuação e zeros à esquerda."""
    if s is None:
        return ""
    s2 = re.sub(r'[\.\-_/ \t]', '', str(s)).strip()
    return s2.lstrip('0') if strip_zeros else s2

# =========================================================
# XML TISS → Itens por guia
# =========================================================

def _get_numero_lote(root: ET.Element) -> str:
    """Obtém número do lote do XML TISS."""
    el = root.find('.//ans:prestadorParaOperadora/ans:loteGuias/ans:numeroLote', ANS_NS)
    if el is not None and tx(el):
        return tx(el)
    # recurso de glosa
    el = root.find('.//ans:prestadorParaOperadora/ans:recursoGlosa/ans:guiaRecursoGlosa/ans:numeroLote', ANS_NS)
    if el is not None and tx(el):
        return tx(el)
    return ""

def _itens_consulta(guia: ET.Element) -> List[Dict]:
    """Extrai itens de guias de consulta (sempre 1 procedimento)."""
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
    """Extrai itens de guias SP/SADT."""
    out = []

    # Procedimentos executados
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

    # Outras despesas
    for desp in guia.findall('.//ans:outrasDespesas/ans:despesa', ANS_NS):
        ident = tx(desp.find('ans:identificadorDespesa', ANS_NS))
        sv = desp.find('ans:servicosExecutados', ANS_NS)

        codigo_tabela = tx(sv.find('ans:codigoTabela', ANS_NS)) if sv is not None else ''
        codigo_proc   = tx(sv.find('ans:codigoProcedimento', ANS_NS)) if sv is not None else ''
        descricao     = tx(sv.find('ans:descricaoProcedimento', ANS_NS)) if sv is not None else ''

        qtd  = dec(tx(sv.find('ans:quantidadeExecutada', ANS_NS)))
        vuni = dec(tx(sv.find('ans:valorUnitario', ANS_NS)))
        vtot = dec(tx(sv.find('ans:valorTotal', ANS_NS)))
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

def parse_itens_tiss_xml(source) -> List[Dict]:
    """Extrai itens por guia de um arquivo XML TISS."""
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
    out = []

    # CONSULTA
    for guia in root.findall('.//ans:guiaConsulta', ANS_NS):
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
# PARTE 2 — Demonstrativo (.xlsx) + Wizard + Tratamento Glosa
# =========================================================

# ---------- Normalização de texto p/ matching de colunas ----------
def _normtxt(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = s.lower().strip()
    return re.sub(r"\s+", " ", s)

# ---------- Padrões de reconhecimento de colunas (auto-detecção) ----------
# Observação: estes padrões cobrem seu layout (Lote, Competência, Guia, Cod. Procedimento etc.)
_COLMAPS = {
    'lote'       : [r'\blote\b'],
    'competencia': [r'compet|m[eê]s|period|refer'],
    'guia_prest' : [r'\bguia\b'],       # sua planilha usa apenas "Guia"
    'guia_oper'  : [r'\bguia\b'],       # usamos a mesma coluna para operadora (melhora casamento)
    'cod_proc'   : [r'cod.*proced|proced.*cod|tuss|cod\.\s*proced'],
    'desc_proc'  : [r'descri.*proced|proced.*descr|^descri|descri[cç][aã]o'],
    'qtd_apres'  : [r'qt|quant|qtd|quant\.\s*exec|quantidade.*exec'],
    'qtd_paga'   : [r'qt|quant|qtd|quant\.\s*paga|quantidade.*paga|autori|aprov'],
    'val_apres'  : [r'valor.*apres|apresent|cobrado'],
    'val_glosa'  : [r'\bglosa\b|glosado|glosada'],
    'val_pago'   : [r'valor.*(pago|apurado|liberado|aprovado)$|(pago|apurado|liberado)$'],
    # seu layout traz "Código Glosa" com "NNNN - DESCRIÇÃO"
    'motivo_cod' : [r'c[oó]d.*glosa|glosa'],
    'motivo_desc': [r'glosa|descri.*glosa'],
}

def _match_col(cols: List[str], pats: List[str]) -> Optional[str]:
    norm = {c: _normtxt(c) for c in cols}
    for c, cn in norm.items():
        ok = True
        for p in pats:
            if not re.search(p, cn):
                ok = False
                break
        if ok:
            return c
    return None

def _find_header_row(df_raw: pd.DataFrame) -> int:
    """
    Heurística igual ao seu relatório:
      - se a 1ª coluna em alguma linha tem 'CPF/CNPJ', o cabeçalho está na linha seguinte.
    """
    try:
        s0 = df_raw.iloc[:, 0].astype(str).str.strip().map(_normtxt)
        mask = s0.eq("cpf/cnpj")
        if mask.any():
            return int(mask.idxmax()) + 1
    except Exception:
        pass
    return 0

# ---------- Tratamento do campo "Código Glosa" → separa código e descrição ----------
def tratar_codigo_glosa(df_demo: pd.DataFrame) -> pd.DataFrame:
    """
    Se houver coluna 'Código Glosa' no formato 'NNNN - DESCRIÇÃO', separa em:
      - motivo_glosa_codigo
      - motivo_glosa_descricao
    """
    if 'Código Glosa' not in df_demo.columns:
        return df_demo

    gl = df_demo['Código Glosa'].astype(str).fillna('')

    df_demo['motivo_glosa_codigo']    = gl.str.extract(r'^(\d+)')
    df_demo['motivo_glosa_descricao'] = gl.str.extract(r'^\s*\d+\s*-\s*(.*)$')

    df_demo['motivo_glosa_codigo']    = df_demo['motivo_glosa_codigo'].fillna('').str.strip()
    df_demo['motivo_glosa_descricao'] = df_demo['motivo_glosa_descricao'].fillna('').str.strip()
    return df_demo

# ---------- Aplica mapeamento manual do wizard ----------
def _apply_manual_map(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    def pick(key: str):
        c = mapping.get(key)
        if not c or c == "(não usar)" or c not in df.columns:
            return None
        return df[c]

    out = pd.DataFrame({
        'numero_lote'           : pick('lote'),
        'competencia'           : pick('competencia'),
        'numeroGuiaPrestador'   : pick('guia_prest'),
        'numeroGuiaOperadora'   : pick('guia_oper'),
        'codigo_procedimento'   : pick('cod_proc'),
        'descricao_procedimento': pick('desc_proc'),
        'quantidade_apresentada': pd.to_numeric(pick('qtd_apres'), errors='coerce') if pick('qtd_apres') is not None else 0,
        'quantidade_paga'       : pd.to_numeric(pick('qtd_paga'), errors='coerce')  if pick('qtd_paga')  is not None else 0,
        'valor_apresentado'     : pd.to_numeric(pick('val_apres'), errors='coerce') if pick('val_apres') is not None else 0,
        'valor_glosa'           : pd.to_numeric(pick('val_glosa'), errors='coerce') if pick('val_glosa') is not None else 0,
        'valor_pago'            : pd.to_numeric(pick('val_pago'), errors='coerce')  if pick('val_pago')  is not None else 0,
        # se o wizard apontar diretamente para colunas já separadas
        'motivo_glosa_codigo'   : pick('motivo_cod'),
        'motivo_glosa_descricao': pick('motivo_desc'),
    })

    # normalizações
    for c in ['numero_lote', 'numeroGuiaPrestador', 'numeroGuiaOperadora', 'codigo_procedimento']:
        if c in out:
            out[c] = out[c].astype(str).str.strip()

    for c in ['valor_apresentado','valor_glosa','valor_pago','quantidade_apresentada','quantidade_paga']:
        if c in out:
            out[c] = pd.to_numeric(out[c], errors='coerce').fillna(0.0)

    # chaves
    out['codigo_procedimento_norm'] = out['codigo_procedimento'].astype(str).map(lambda s: normalize_code(s))
    out['chave_prest'] = out['numeroGuiaPrestador'].astype(str).str.strip() + "__" + out['codigo_procedimento_norm']
    out['chave_oper']  = out['numeroGuiaOperadora'].astype(str).str.strip() + "__" + out['codigo_procedimento_norm']
    return out

# ---------- Leitura do Demonstrativo (.xlsx) ----------
def ler_demo_itens_pagto_xlsx(
    source,
    strip_zeros_codes: bool = False,
    manual_map: Optional[Dict[str, Optional[str]]] = None,
    prefer_sheet: Optional[str] = None
) -> pd.DataFrame:
    """
    Lê o Demonstrativo itemizado e devolve colunas normalizadas para conciliação.
    - Se 'manual_map' for fornecido, usa o mapeamento do wizard.
    - Tenta achar o cabeçalho pela linha 'CPF/CNPJ' na 1ª coluna (como no seu arquivo).
    - Auto-detecção de colunas por regex, compatível com o layout enviado.
    """
    xls = pd.ExcelFile(source, engine='openpyxl')

    # Escolhe a aba
    if prefer_sheet and prefer_sheet in xls.sheet_names:
        sheet = prefer_sheet
    else:
        # tenta achar uma aba "DemonstrativoAnaliseDeContas" ou similar
        sheet = None
        for s in xls.sheet_names:
            s_norm = _normtxt(s)
            if 'analise' in s_norm or 'análise' in s_norm or 'demonstr' in s_norm or 'conta' in s_norm:
                sheet = s
                break
        sheet = sheet or xls.sheet_names[0]

    # Lê e encontra o cabeçalho
    df_raw = pd.read_excel(source, sheet_name=sheet, engine='openpyxl')
    hdr = _find_header_row(df_raw)
    df = df_raw.copy()
    if hdr > 0:
        df.columns = df_raw.iloc[hdr]
        df = df_raw.iloc[hdr + 1:].reset_index(drop=True)

    if manual_map:
        return _apply_manual_map(df, manual_map)

    # Auto-detecção de colunas
    cols = [str(c) for c in df.columns]
    pick = {k: _match_col(cols, v) for k, v in _COLMAPS.items()}

    # Campos mínimos: código procedimento + pelo menos um valor financeiro
    if not pick.get('cod_proc') or not (pick.get('val_apres') or pick.get('val_glosa') or pick.get('val_pago')):
        raise ValueError("Não identifiquei colunas itemizadas no Demonstrativo (auto). Use o mapeamento manual.")

    def col(c): return pick.get(c)

    out = pd.DataFrame({
        'numero_lote'           : df[col('lote')] if col('lote') else None,
        'competencia'           : df[col('competencia')] if col('competencia') else None,
        'numeroGuiaPrestador'   : df[col('guia_prest')] if col('guia_prest') else None,
        'numeroGuiaOperadora'   : df[col('guia_oper')] if col('guia_oper') else None,
        'codigo_procedimento'   : df[col('cod_proc')] if col('cod_proc') else None,
        'descricao_procedimento': df[col('desc_proc')] if col('desc_proc') else None,
        'quantidade_apresentada': pd.to_numeric(df[col('qtd_apres')], errors='coerce') if col('qtd_apres') else 0.0,
        'quantidade_paga'       : pd.to_numeric(df[col('qtd_paga')], errors='coerce')  if col('qtd_paga')  else 0.0,
        'valor_apresentado'     : pd.to_numeric(df[col('val_apres')], errors='coerce') if col('val_apres') else 0.0,
        'valor_glosa'           : pd.to_numeric(df[col('val_glosa')], errors='coerce') if col('val_glosa') else 0.0,
        'valor_pago'            : pd.to_numeric(df[col('val_pago')], errors='coerce')  if col('val_pago')  else 0.0,
        # Se o auto-map encontrar campos de motivo já separados
        'motivo_glosa_codigo'   : df[col('motivo_cod')] if col('motivo_cod') else None,
        'motivo_glosa_descricao': df[col('motivo_desc')] if col('motivo_desc') else None,
        # Mantém coluna bruta caso exista (para tratamento posterior)
        'Código Glosa'          : df['Código Glosa'] if 'Código Glosa' in df.columns else None,
    })

    # Normalizações
    for c in ['numero_lote', 'numeroGuiaPrestador', 'numeroGuiaOperadora', 'codigo_procedimento']:
        if c in out.columns and out[c] is not None:
            out[c] = out[c].astype(str).str.strip()

    for c in ['valor_apresentado', 'valor_glosa', 'valor_pago', 'quantidade_apresentada', 'quantidade_paga']:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors='coerce').fillna(0.0)

    # Códigos normalizados e chaves
    out['codigo_procedimento_norm'] = out['codigo_procedimento'].astype(str).map(
        lambda s: normalize_code(s, strip_zeros=strip_zeros_codes)
    )
    out['chave_prest'] = out['numeroGuiaPrestador'].astype(str).str.strip() + '__' + out['codigo_procedimento_norm']
    out['chave_oper']  = out['numeroGuiaOperadora'].astype(str).str.strip() + '__' + out['codigo_procedimento_norm']

    return out

# ---------- Wizard de mapeamento manual (UI) ----------
def _mapping_wizard_for_demo(uploaded_file) -> Optional[pd.DataFrame]:
    """
    Exibe UI para mapear colunas do demonstrativo manualmente e retorna o DataFrame mapeado.
    O resultado é salvo em st.session_state['demo_mappings'] para uso futuro.
    """
    st.warning(f"Precisamos mapear manualmente o demonstrativo: **{uploaded_file.name}**")

    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(f"Não consegui abrir o arquivo: {e}")
        return None

    # Escolhe a aba
    sheet = st.selectbox(
        f"Aba (sheet) do demonstrativo **{uploaded_file.name}**",
        xls.sheet_names,
        key=f"map_sheet_{uploaded_file.name}"
    )

    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, engine='openpyxl')
    st.markdown("Prévia da planilha (15 primeiras linhas):")
    st.dataframe(df_raw.head(15), use_container_width=True)

    cols = [str(c) for c in df_raw.columns]

    st.markdown("### Mapeie as colunas")
    fields = [
        ('lote', 'Lote (opcional)'),
        ('competencia', 'Competência (opcional)'),
        ('guia_prest', 'Número da Guia (Prestador)'),
        ('guia_oper', 'Número da Guia (Operadora) (opcional)'),
        ('cod_proc', 'Código do Procedimento (TUSS / interno)'),
        ('desc_proc', 'Descrição do Procedimento (opcional)'),
        ('qtd_apres', 'Quantidade Apresentada (opcional)'),
        ('qtd_paga',  'Quantidade Paga/Autorizada (opcional)'),
        ('val_apres', 'Valor Apresentado'),
        ('val_glosa', 'Valor Glosa'),
        ('val_pago',  'Valor Pago/Apurado'),
        ('motivo_cod', 'Motivo de Glosa (Código) (opcional)'),
        ('motivo_desc','Motivo de Glosa (Descrição) (opcional)'),
    ]

    def _default_for(key: str) -> int:
        for i, c in enumerate(cols):
            cn = _normtxt(c)
            pats = _COLMAPS.get(key, [])
            if any(re.search(p, cn) for p in pats):
                return i + 1  # +1 pela opção "(não usar)"
        return 0

    mapping: Dict[str, Optional[str]] = {}
    for key, label in fields:
        opt = ['(não usar)'] + cols
        idx = _default_for(key)
        sel = st.selectbox(label, opt, index=idx, key=f"map_{uploaded_file.name}_{key}")
        mapping[key] = None if sel == '(não usar)' else sel

    st.caption("Campos mínimos: **Código do Procedimento** e pelo menos um de **Valor Apresentado/Glosa/Pago**.")
    do_map = st.button(
        f"Usar este mapeamento para {uploaded_file.name}",
        type="primary",
        key=f"btn_map_{uploaded_file.name}"
    )
    if not do_map:
        return None

    # Salva mapeamento na sessão
    st.session_state.setdefault('demo_mappings', {})
    st.session_state['demo_mappings'][uploaded_file.name] = {'sheet': sheet, 'columns': mapping}

    # Aplica o mapeamento
    try:
        df_mapped = ler_demo_itens_pagto_xlsx(
            uploaded_file,
            manual_map=mapping,
            prefer_sheet=sheet
        )
        # separa código/descrição de glosa se vier em "Código Glosa"
        df_mapped = tratar_codigo_glosa(df_mapped)
        st.success("Mapeamento aplicado com sucesso!")
        return df_mapped
    except Exception as e:
        st.error(f"Falha aplicando mapeamento: {e}")
        return None

# ---------- Loader robusto do Demonstrativo (chamado FORA do botão) ----------
def build_demo_df(demo_files, strip_zeros_codes: bool = False) -> pd.DataFrame:
    """
    Lê todos os demonstrativos:
      - usa mapeamento salvo na sessão se existir
      - tenta auto-detecção
      - se falhar, abre wizard p/ mapear manualmente (sem resetar ao reexecutar)
    Retorna um único DataFrame concatenado.
    """
    parts: List[pd.DataFrame] = []
    if not demo_files:
        return pd.DataFrame()

    st.session_state.setdefault('demo_mappings', {})

    for f in demo_files:
        if hasattr(f, 'seek'):
            f.seek(0)
        fname = getattr(f, 'name', 'demo.xlsx')

        # 1) Usa mapeamento salvo, se houver
        mapping_info = st.session_state['demo_mappings'].get(fname)
        if mapping_info:
            try:
                df_demo = ler_demo_itens_pagto_xlsx(
                    f,
                    strip_zeros_codes=strip_zeros_codes,
                    manual_map=mapping_info.get('columns'),
                    prefer_sheet=mapping_info.get('sheet')
                )
                df_demo = tratar_codigo_glosa(df_demo)

                # Se operadora não estiver preenchida, replicamos Guia (melhora casamento por chave_oper)
                if 'numeroGuiaOperadora' in df_demo.columns and df_demo['numeroGuiaOperadora'].isna().all():
                    df_demo['numeroGuiaOperadora'] = df_demo['numeroGuiaPrestador']

                parts.append(df_demo)
                continue
            except Exception as e:
                st.warning(f"Mapeamento salvo para '{fname}' falhou ({e}). Tentando auto-detecção...")

        # 2) Tenta auto-detecção
        try:
            df_demo = ler_demo_itens_pagto_xlsx(f, strip_zeros_codes=strip_zeros_codes)
            df_demo = tratar_codigo_glosa(df_demo)

            if 'numeroGuiaOperadora' in df_demo.columns and df_demo['numeroGuiaOperadora'].isna().all():
                df_demo['numeroGuiaOperadora'] = df_demo['numeroGuiaPrestador']

            parts.append(df_demo)
            continue
        except Exception:
            # 3) Abre wizard manual
            with st.expander(f"⚙️ Configurar mapeamento para: {fname}", expanded=True):
                df_manual = _mapping_wizard_for_demo(f)
                if df_manual is not None:
                    if 'numeroGuiaOperadora' in df_manual.columns and df_manual['numeroGuiaOperadora'].isna().all():
                        df_manual['numeroGuiaOperadora'] = df_manual['numeroGuiaPrestador']
                    parts.append(df_manual)
                else:
                    st.error(f"Não foi possível mapear o demonstrativo '{fname}'. Preencha o mapeamento acima.")

    if parts:
        return pd.concat(parts, ignore_index=True)
    return pd.DataFrame()


# =========================================================
# PARTE 3 — Conciliação (XML × Demonstrativo)
# =========================================================

def build_xml_df(xml_files, strip_zeros_codes: bool = False) -> pd.DataFrame:
    """Lê e normaliza todos os XML TISS enviados."""
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
    df['chave_prest'] = (
        df['numeroGuiaPrestador'].fillna('').astype(str).str.strip()
        + '__'
        + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip()
    )
    df['chave_oper'] = (
        df['numeroGuiaOperadora'].fillna('').astype(str).str.strip()
        + '__'
        + df['codigo_procedimento_norm'].fillna('').astype(str).str.strip()
    )

    return df


# ---------------------------------------------------------
# Conciliação
# ---------------------------------------------------------

# --- helper para padronizar nomes "lado XML" após merges com sufixos ---
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

def _alias_xml_cols(df: pd.DataFrame, cols: list[str] = None, prefer_suffix: str = '_xml') -> pd.DataFrame:
    """
    Garante que as colunas de interesse (lado XML) existam SEM sufixo,
    copiando de 'colname_xml' caso necessário.
    """
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
    """
    PIPELINE DE CONCILIAÇÃO:
    1) Match XML × DEMO por chave_prest (número da guia do prestador + código)
    2) Match restante por chave_oper (número da guia da operadora + código)
    3) Fallback opcional: match por guia + descrição + tolerância valor
    4) Gera não-casados
    5) Calcula diffs e % de glosa
    """

    # ------------------------------------------------------
    # 1) Match por chave_prest
    # ------------------------------------------------------
    m1 = df_xml.merge(
        df_demo,
        on="chave_prest",
        how="left",
        suffixes=("_xml", "_demo"),
    )
    # padroniza visão XML: cria aliases sem sufixo a partir de *_xml
    m1 = _alias_xml_cols(m1)

    m1["matched_on"] = m1["valor_apresentado"].notna().map({True: "prestador", False: ""})

    # ------------------------------------------------------
    # 2) Match por chave_oper (nos que não casaram)
    # ------------------------------------------------------
    restante = m1[m1["matched_on"] == ""].copy()
    # garante que todas as colunas lado XML existam sem sufixo
    restante = _alias_xml_cols(restante)

    cols_for_second_join = [
        'arquivo', 'numero_lote', 'tipo_guia',
        'numeroGuiaPrestador', 'numeroGuiaOperadora',
        'paciente', 'medico', 'data_atendimento',
        'tipo_item', 'identificadorDespesa',
        'codigo_tabela', 'codigo_procedimento', 'codigo_procedimento_norm',
        'descricao_procedimento',
        'quantidade', 'valor_unitario', 'valor_total',
        'chave_oper', 'chave_prest',
    ]

    # filtra apenas as colunas que de fato existem
    cols_for_second_join = [c for c in cols_for_second_join if c in restante.columns]
    still_xml = restante[cols_for_second_join].copy()

    m2 = still_xml.merge(
        df_demo,
        on="chave_oper",
        how="left",
        suffixes=("_xml", "_demo"),
    )
    # novamente padroniza visão XML em m2
    m2 = _alias_xml_cols(m2)
    m2["matched_on"] = m2["valor_apresentado"].notna().map({True: "operadora", False: ""})

    # acumula conciliados
    conc = pd.concat(
        [
            m1[m1["matched_on"] != ""],
            m2[m2["matched_on"] != ""],
        ],
        ignore_index=True
    )

    # ------------------------------------------------------
    # 3) Fallback opcional — descrição + tolerância de valor
    # ------------------------------------------------------
    fallback_matches = pd.DataFrame()

    if fallback_por_descricao:
        rem1 = m1[m1["matched_on"] == ""].copy()
        rem2 = m2[m2["matched_on"] == ""].copy()

        # padroniza visão XML para ambas bases remanescentes
        rem1 = _alias_xml_cols(rem1)
        rem2 = _alias_xml_cols(rem2)

        rem_xml = pd.concat([rem1, rem2], ignore_index=True)

        if not rem_xml.empty:
            # chave de guia unificada: prestador se existir, senão operadora
            rem_xml["guia_join"] = rem_xml.apply(
                lambda r: r["numeroGuiaPrestador"] if str(r.get("numeroGuiaPrestador", "")).strip()
                else str(r.get("numeroGuiaOperadora", "")).strip(),
                axis=1
            )

            df_demo2 = df_demo.copy()
            df_demo2["guia_join"] = df_demo2.apply(
                lambda r: r["numeroGuiaPrestador"] if str(r.get("numeroGuiaPrestador", "")).strip()
                else str(r.get("numeroGuiaOperadora", "")).strip(),
                axis=1
            )

            # precisa garantir que 'descricao_procedimento' exista nos dois lados
            if "descricao_procedimento" in rem_xml.columns and "descricao_procedimento" in df_demo2.columns:
                tmp = rem_xml.merge(
                    df_demo2,
                    on=["guia_join", "descricao_procedimento"],
                    how="left",
                    suffixes=("_xml", "_demo")
                )

                tol = float(tolerance_valor)
                keep = (
                    tmp["valor_apresentado"].notna()
                    & ((tmp["valor_total"] - tmp["valor_apresentado"]).abs() <= tol)
                )

                fallback_matches = tmp[keep].copy()
                if not fallback_matches.empty:
                    fallback_matches["matched_on"] = "descricao+valor"
                    conc = pd.concat([conc, fallback_matches], ignore_index=True)

    # ------------------------------------------------------
    # 4) Não casados finais (com nomes sem sufixo)
    # ------------------------------------------------------
    unmatch = pd.concat(
        [
            m1[m1["matched_on"] == ""],
            m2[m2["matched_on"] == ""],
            fallback_matches[fallback_matches.get("matched_on", "") == ""]
            if not fallback_matches.empty else pd.DataFrame()
        ],
        ignore_index=True
    )
    # padroniza visão XML antes de deduplicar/usar colunas
    unmatch = _alias_xml_cols(unmatch)

    if not unmatch.empty:
        subset_cols = [
            "arquivo", "numero_lote", "tipo_guia",
            "numeroGuiaPrestador", "codigo_procedimento",
            "valor_total"
        ]
        subset_cols = [c for c in subset_cols if c in unmatch.columns]
        if subset_cols:
            unmatch = unmatch.drop_duplicates(subset=subset_cols)

    # ------------------------------------------------------
    # 5) Diffs e % glosa
    # ------------------------------------------------------
    if not conc.empty:
        # garante visão XML antes dos cálculos
        conc = _alias_xml_cols(conc)

        conc["apresentado_diff"] = conc["valor_total"] - conc["valor_apresentado"]
        conc["glosa_pct"] = conc.apply(
            lambda r: (
                (r["valor_glosa"] / r["valor_apresentado"])
                if r.get("valor_apresentado", 0) and r["valor_apresentado"] > 0
                else 0.0
            ),
            axis=1
        )

    return {
        "conciliacao": conc,
        "nao_casados": unmatch
    }

# =========================================================
# PARTE 4 — Auditoria de Guias
# =========================================================

def build_chave_guia(tipo: str, numeroGuiaPrestador: str, numeroGuiaOperadora: str) -> Optional[str]:
    """
    Define a chave única da guia para auditoria.
    Prioridade:
      1) número da guia do Prestador (se existir)
      2) número da guia da Operadora (fallback)
    Somente para guias assistenciais (CONSULTA / SADT).
    """
    t = (tipo or '').upper()
    if t in ('CONSULTA', 'SADT'):
        guia = str(numeroGuiaPrestador or '').strip() or str(numeroGuiaOperadora or '').strip()
        return guia if guia else None
    return None


def _parse_dt_series(s: pd.Series) -> pd.Series:
    """Converte série para datetime usando a função flexível de parsing."""
    return s.astype(str).map(lambda x: parse_date_flex(x))


def auditar_guias(df_xml_itens: pd.DataFrame, prazo_retorno: int = 30) -> pd.DataFrame:
    """
    Auditoria baseada nos itens do XML (nível guia).

    Regras:
    - Duplicidade: mesma 'chave_guia' aparecendo em mais de um arquivo e/ou lote.
    - Retorno: mesmo paciente volta ao mesmo médico dentro de 'prazo_retorno' dias.
    - Indicadores: quantidade de itens por guia e soma dos valores (valor_total) dos itens.

    Retorna um DataFrame no nível guia com:
      ['arquivo(s)', 'numero_lote(s)', 'tipo_guia', 'numeroGuiaPrestador', 'numeroGuiaOperadora',
       'paciente', 'medico', 'data_atendimento', 'itens_na_guia', 'valor_total_xml',
       'duplicada', 'arquivos_duplicados', 'lotes_duplicados', 'retorno_no_periodo', 'retorno_ref',
       'status_auditoria']
    """
    if df_xml_itens is None or df_xml_itens.empty:
        return pd.DataFrame()

    # Garante colunas necessárias
    req_cols = [
        'arquivo', 'numero_lote', 'tipo_guia',
        'numeroGuiaPrestador', 'numeroGuiaOperadora',
        'paciente', 'medico', 'data_atendimento',
        'valor_total'
    ]
    for c in req_cols:
        if c not in df_xml_itens.columns:
            df_xml_itens[c] = None

    # Parse de data
    df = df_xml_itens.copy()
    df['data_atendimento_dt'] = _parse_dt_series(df['data_atendimento'])

    # Indicadores por guia (agregando itens)
    agg = (df.groupby([
                'tipo_guia', 'numeroGuiaPrestador', 'numeroGuiaOperadora',
                'paciente', 'medico'
            ], dropna=False, as_index=False)
             .agg(
                 arquivo=('arquivo', lambda x: list(sorted(set(str(a) for a in x if str(a).strip())))),
                 numero_lote=('numero_lote', lambda x: list(sorted(set(str(a) for a in x if str(a).strip())))),
                 data_atendimento=('data_atendimento_dt', 'min'),  # primeira data do conjunto
                 itens_na_guia=('valor_total', 'count'),
                 valor_total_xml=('valor_total', 'sum'),
            ))

    # Reconstitui colunas string
    agg['arquivo(s)'] = agg['arquivo'].map(lambda L: ", ".join(L))
    agg['numero_lote(s)'] = agg['numero_lote'].map(lambda L: ", ".join(L))
    agg.drop(columns=['arquivo', 'numero_lote'], inplace=True)

    # Monta chave_guia
    agg['chave_guia'] = agg.apply(
        lambda r: build_chave_guia(r['tipo_guia'], r['numeroGuiaPrestador'], r['numeroGuiaOperadora']),
        axis=1
    )

    # Inicializa flags
    agg['duplicada'] = False
    agg['arquivos_duplicados'] = ''
    agg['lotes_duplicados'] = ''
    agg['retorno_no_periodo'] = False
    agg['retorno_ref'] = ''
    agg['status_auditoria'] = ''

    # ---------------------------
    # 1) DUPLICIDADE POR CHAVE
    # ---------------------------
    # Grupos com a mesma chave_guia
    dup_groups = (agg[agg['chave_guia'].notna()]
                  .groupby('chave_guia', as_index=False)
                  .agg(idx=('tipo_guia', lambda _: list(_.index))))

    # Mapeia duplicidade: se um mesmo 'chave_guia' tem mais de 1 registro, marca todos como duplicados
    dup_keys = set()
    for k, g in agg[agg['chave_guia'].notna()].groupby('chave_guia'):
        if len(g) > 1:
            dup_keys.add(k)
            # Para cada linha do grupo, lista lotes e arquivos dos demais
            indices = list(g.index)
            lotes_grupo = g['numero_lote(s)'].tolist()
            arqs_grupo = g['arquivo(s)'].tolist()
            for i_idx, i in enumerate(indices):
                outros_lotes = [l for j, l in enumerate(lotes_grupo) if j != i_idx and l]
                outros_arqs  = [a for j, a in enumerate(arqs_grupo)  if j != i_idx and a]
                lotes_dup = sorted(set(", ".join(outros_lotes).split(", "))) if outros_lotes else []
                arqs_dup  = sorted(set(", ".join(outros_arqs).split(", ")))  if outros_arqs  else []
                agg.loc[i, 'duplicada'] = True
                agg.loc[i, 'lotes_duplicados'] = ", ".join([x for x in lotes_dup if x])
                agg.loc[i, 'arquivos_duplicados'] = ", ".join([x for x in arqs_dup if x])

    # ---------------------------
    # 2) RETORNO (paciente volta ao mesmo médico em até X dias)
    # ---------------------------
    if prazo_retorno and prazo_retorno > 0:
        # para acelerar, index por paciente+medico
        agg['_pac'] = agg['paciente'].fillna('').astype(str).str.strip()
        agg['_med'] = agg['medico'].fillna('').astype(str).str.strip()

        # caminhamos item a item e vemos outros com mesma dupla paciente/médico
        for i, r in agg.iterrows():
            if not r['data_atendimento'] or pd.isna(r['data_atendimento']):
                continue
            pac, med = r['_pac'], r['_med']
            if not pac or not med:
                continue
            # candidatos com mesma dupla, exclui o próprio
            cand = agg[(agg.index != i) & (agg['_pac'] == pac) & (agg['_med'] == med)]
            # datas no raio
            refs = []
            for j, rr in cand.iterrows():
                d0 = r['data_atendimento']
                dj = rr['data_atendimento']
                if not dj or pd.isna(dj):
                    continue
                if abs((d0 - dj).days) <= int(prazo_retorno):
                    # monta referência (lote/arquivo/data)
                    lotes = rr['numero_lote(s)'] or ''
                    arqs  = rr['arquivo(s)'] or ''
                    data  = rr['data_atendimento'].strftime('%d/%m/%Y') if isinstance(rr['data_atendimento'], datetime) else str(rr['data_atendimento'])
                    refs.append(f"{lotes} @ {arqs} @ {data}")
            if refs:
                agg.loc[i, 'retorno_no_periodo'] = True
                agg.loc[i, 'retorno_ref'] = " | ".join(refs)

        agg.drop(columns=['_pac', '_med'], inplace=True)

    # ---------------------------
    # 3) STATUS CONSOLIDADO
    # ---------------------------
    def _status_row(r):
        flags = []
        if r.get('duplicada'):
            flags.append('Duplicidade')
        if r.get('retorno_no_periodo'):
            flags.append('Retorno')
        return " + ".join(flags) if flags else "OK"

    agg['status_auditoria'] = agg.apply(_status_row, axis=1)

    # Ordenação amigável
    cols_out = [
        'tipo_guia', 'numeroGuiaPrestador', 'numeroGuiaOperadora',
        'paciente', 'medico', 'data_atendimento',
        'itens_na_guia', 'valor_total_xml',
        'arquivo(s)', 'numero_lote(s)',
        'duplicada', 'arquivos_duplicados', 'lotes_duplicados',
        'retorno_no_periodo', 'retorno_ref',
        'status_auditoria'
    ]
    # Garante a existência de todas as colunas
    for c in cols_out:
        if c not in agg.columns:
            agg[c] = None

    # Conversões finais e retornos
    if 'valor_total_xml' in agg.columns:
        agg['valor_total_xml'] = pd.to_numeric(agg['valor_total_xml'], errors='coerce').fillna(0.0)

    # data em string para exibição estável
    if 'data_atendimento' in agg.columns:
        agg['data_atendimento'] = agg['data_atendimento'].apply(
            lambda d: d.strftime('%d/%m/%Y') if isinstance(d, datetime) else (d if d else '')
        )

    return agg[cols_out]


# =========================================================
# PARTE 5 — Interface UI (Uploads, Parâmetros, Processamento, Exportação)
# =========================================================

# -----------------------------
# Sidebar de parâmetros
# -----------------------------
with st.sidebar:
    st.header("Parâmetros")

    prazo_retorno = st.number_input(
        "Prazo de retorno (dias)",
        min_value=0, value=30, step=1
    )

    tolerance_valor = st.number_input(
        "Tolerância p/ fallback por descrição (R$)",
        min_value=0.00, value=0.02, step=0.01, format="%.2f"
    )

    fallback_desc = st.toggle(
        "Fallback por descrição + valor (quando código não casar)",
        value=False
    )

    strip_zeros_codes = st.toggle(
        "Normalizar códigos removendo zeros à esquerda",
        value=True
    )


# -----------------------------
# Upload dos arquivos
# -----------------------------
st.subheader("📤 Upload de arquivos")

xml_files = st.file_uploader(
    "XML TISS (um ou mais):",
    type=['xml'],
    accept_multiple_files=True
)

demo_files = st.file_uploader(
    "Demonstrativos de Pagamento (.xlsx) — itemizado:",
    type=['xlsx'],
    accept_multiple_files=True
)


# --------------------------------------------------------------
# PROCESSAMENTO DO DEMONSTRATIVO (SEMPRE) — para permitir wizard
# --------------------------------------------------------------
df_demo = build_demo_df(demo_files or [], strip_zeros_codes=strip_zeros_codes)

if not df_demo.empty:
    st.subheader("📘 Itens do Demonstrativo (Detectados)")
    st.dataframe(
        apply_currency(df_demo, ['valor_apresentado', 'valor_glosa', 'valor_pago']),
        use_container_width=True,
        height=380
    )
else:
    if demo_files:
        st.info("Carregue um Demonstrativo válido ou conclua o mapeamento manual.")


# --------------------------------------------------------------
# BOTÃO — Somente quando clicado processa XML + conciliação
# --------------------------------------------------------------
st.markdown("---")
if st.button("🚀 Processar Conciliação e Auditoria", type="primary"):

    # ---------------------------------------
    # 1) XML
    # ---------------------------------------
    df_xml = build_xml_df(xml_files or [], strip_zeros_codes=strip_zeros_codes)

    if df_xml.empty:
        st.warning("Nenhum item extraído do(s) XML(s). Verifique os arquivos.")
        st.stop()

    st.subheader("📄 Itens extraídos dos XML (Consulta / SADT)")
    st.dataframe(
        apply_currency(df_xml, ['valor_unitario', 'valor_total']),
        use_container_width=True,
        height=380
    )

    if df_demo.empty:
        st.warning("Nenhum demonstrativo válido para conciliar.")
        st.stop()

    # ---------------------------------------
    # 2) Conciliação
    # ---------------------------------------
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
        ['valor_unitario', 'valor_total', 'valor_apresentado', 'valor_glosa', 'valor_pago', 'apresentado_diff']
    )
    st.dataframe(conc_disp, use_container_width=True, height=480)

    c1, c2 = st.columns(2)
    c1.metric("Itens conciliados", len(conc))
    c2.metric("Itens não conciliados", len(unmatch))

    if not unmatch.empty:
        st.subheader("❗ Itens não conciliados")
        st.dataframe(
            apply_currency(unmatch.copy(), ['valor_unitario', 'valor_total']),
            use_container_width=True,
            height=320
        )

    # ---------------------------------------
    # 3) Rankings
    # ---------------------------------------
    st.markdown("---")
    st.subheader("📊 Rankings e Indicadores de Glosa")

    colA, colB = st.columns(2)

    # Motivos de glosa
    with colA:
        st.markdown("### Motivos de Glosa — TOP 50")
        if 'motivo_glosa_codigo' in conc.columns:
            mot = (conc.groupby(['motivo_glosa_codigo', 'motivo_glosa_descricao'], dropna=False, as_index=False)
                   .agg(valor_glosa=('valor_glosa', 'sum'),
                        valor_apresentado=('valor_apresentado', 'sum'),
                        itens=('codigo_procedimento', 'count')))
            mot['glosa_pct'] = mot.apply(
                lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
                axis=1
            )
            mot = mot.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
            st.dataframe(apply_currency(mot, ['valor_glosa', 'valor_apresentado']), use_container_width=True)
        else:
            st.info("Motivo de glosa não presente no demonstrativo.")

    # Procedimentos
    with colB:
        st.markdown("### Procedimentos com maior Glosa — TOP 50")
        proc = (conc.groupby(['codigo_procedimento', 'descricao_procedimento'], dropna=False, as_index=False)
                .agg(valor_apresentado=('valor_apresentado', 'sum'),
                     valor_glosa=('valor_glosa', 'sum'),
                     valor_pago=('valor_pago', 'sum'),
                     itens=('arquivo', 'count')))
        proc['glosa_pct'] = proc.apply(
            lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
            axis=1
        )
        proc = proc.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
        st.dataframe(apply_currency(proc, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True)

    # Rankings médicos/lotes
    colC, colD = st.columns(2)
    with colC:
        st.markdown("### Médicos com maior glosa")
        med = (conc.groupby(['medico'], dropna=False, as_index=False)
               .agg(valor_apresentado=('valor_apresentado', 'sum'),
                    valor_glosa=('valor_glosa', 'sum'),
                    valor_pago=('valor_pago', 'sum'),
                    itens=('arquivo', 'count')))
        med['glosa_pct'] = med.apply(
            lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
            axis=1
        )
        med = med.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
        st.dataframe(apply_currency(med, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True)

    with colD:
        st.markdown("### Lotes com maior glosa")
        lot = (conc.groupby(['numero_lote'], dropna=False, as_index=False)
               .agg(valor_apresentado=('valor_apresentado', 'sum'),
                    valor_glosa=('valor_glosa', 'sum'),
                    valor_pago=('valor_pago', 'sum'),
                    itens=('arquivo', 'count')))
        lot['glosa_pct'] = lot.apply(
            lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
            axis=1
        )
        lot = lot.sort_values(['glosa_pct', 'valor_glosa'], ascending=[False, False]).head(50)
        st.dataframe(apply_currency(lot, ['valor_apresentado', 'valor_glosa', 'valor_pago']), use_container_width=True)


    # ---------------------------------------
    # 4) Auditoria por guia
    # ---------------------------------------
    st.markdown("---")
    st.subheader("🔎 Auditoria por Guia (Duplicidade e Retorno)")

    df_aud = auditar_guias(df_xml, prazo_retorno=prazo_retorno)

    if df_aud.empty:
        st.info("Sem dados para auditoria.")
    else:
        st.dataframe(df_aud, use_container_width=True, height=420)


    # ---------------------------------------
    # 5) Exportação Excel consolidado
    # ---------------------------------------
    st.markdown("---")
    st.subheader("📥 Exportar Excel Consolidado")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as wr:

        df_xml.to_excel(wr, index=False, sheet_name='Itens_XML')
        df_demo.to_excel(wr, index=False, sheet_name='Itens_Demo')
        conc.to_excel(wr, index=False, sheet_name='Conciliação')
        unmatch.to_excel(wr, index=False, sheet_name='Nao_Casados')

        # Motivos, procedimentos, médicos, lotes
        if not conc.empty:
            mot_x = (conc.groupby(['motivo_glosa_codigo', 'motivo_glosa_descricao'], dropna=False, as_index=False)
                     .agg(valor_glosa=('valor_glosa', 'sum'),
                          valor_apresentado=('valor_apresentado', 'sum'),
                          itens=('codigo_procedimento', 'count')))
            mot_x['glosa_pct'] = mot_x.apply(
                lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
                axis=1
            )
            mot_x.to_excel(wr, index=False, sheet_name='Motivos_Glosa')

            proc_x = (conc.groupby(['codigo_procedimento', 'descricao_procedimento'], dropna=False, as_index=False)
                      .agg(valor_apresentado=('valor_apresentado', 'sum'),
                           valor_glosa=('valor_glosa', 'sum'),
                           valor_pago=('valor_pago', 'sum'),
                           itens=('arquivo', 'count')))
            proc_x['glosa_pct'] = proc_x.apply(
                lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
                axis=1
            )
            proc_x.to_excel(wr, index=False, sheet_name='Procedimentos_Glosa')

            med_x = (conc.groupby(['medico'], dropna=False, as_index=False)
                     .agg(valor_apresentado=('valor_apresentado', 'sum'),
                          valor_glosa=('valor_glosa', 'sum'),
                          valor_pago=('valor_pago', 'sum'),
                          itens=('arquivo', 'count')))
            med_x['glosa_pct'] = med_x.apply(
                lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
                axis=1
            )
            med_x.to_excel(wr, index=False, sheet_name='Medicos')

            lot_x = (conc.groupby(['numero_lote'], dropna=False, as_index=False)
                     .agg(valor_apresentado=('valor_apresentado', 'sum'),
                          valor_glosa=('valor_glosa', 'sum'),
                          valor_pago=('valor_pago', 'sum'),
                          itens=('arquivo', 'count')))
            lot_x['glosa_pct'] = lot_x.apply(
                lambda r: r['valor_glosa'] / r['valor_apresentado'] if r['valor_apresentado'] > 0 else 0,
                axis=1
            )
            lot_x.to_excel(wr, index=False, sheet_name='Lotes')

        df_aud.to_excel(wr, index=False, sheet_name='Auditoria_Guias')

        # Ajustes de largura e congelamento de cabeçalho
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
        file_name="tiss_itens_conciliacao_auditoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


