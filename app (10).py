
from __future__ import annotations

import io
import re
import math
from decimal import Decimal
from pathlib import Path
from typing import List, Dict

import pandas as pd
import streamlit as st

from tiss_parser import (
    parse_tiss_xml,
    parse_many_xmls,
    audit_por_guia,
    __version__ as PARSER_VERSION
)

# =========================================================
# Config & Header
# =========================================================
st.set_page_config(page_title="Leitor TISS XML (Consulta • SADT • Recurso)", layout="wide")
st.title("Leitor de XML TISS (Consulta, SP‑SADT e Recurso de Glosa)")
st.caption(f"Extrai nº do lote, protocolo (quando houver), quantidade de guias e valor total • Parser {PARSER_VERSION}")

tab1, tab2 = st.tabs(["Upload de XML(s)", "Ler de uma pasta local (clonada do GitHub)"])

# =========================================================
# FORMATAÇÃO DE MOEDA (BR)
# =========================================================
def format_currency_br(val) -> str:
    try:
        v = float(Decimal(str(val)))
    except Exception:
        v = 0.0
    if not math.isfinite(v):
        v = 0.0
    neg = v < 0
    v = abs(v)
    inteiro = int(v)
    centavos = int(round((v - inteiro) * 100))
    inteiro_fmt = f"{inteiro:,}".replace(",", ".")
    centavos_fmt = f"{centavos:02d}"
    s = f"R$ {inteiro_fmt},{centavos_fmt}"
    return f"-{s}" if neg else s

def _df_display_currency(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    dfd = df.copy()
    for c in cols:
        if c in dfd.columns:
            dfd[c] = dfd[c].apply(format_currency_br)
    return dfd

# =========================================================
# Regex para lote
# =========================================================
_LOTE_REGEX = re.compile(r'(?i)lote\s*[-_]*\s*(\d+)')

def extract_lote_from_filename(name: str) -> str | None:
    if not isinstance(name, str):
        return None
    m = _LOTE_REGEX.search(name)
    if m:
        return m.group(1)
    return None

# =========================================================
# Funções auxiliares
# =========================================================
def _to_float(val) -> float:
    try:
        return float(Decimal(str(val)))
    except Exception:
        return 0.0

def _df_format(df: pd.DataFrame) -> pd.DataFrame:
    if 'valor_total' in df.columns:
        df['valor_total'] = df['valor_total'].apply(_to_float)
    if 'qtde_guias' in df.columns and 'valor_total' in df.columns:
        df['suspeito'] = (df['qtde_guias'] > 0) & (df['valor_total'] == 0)
    else:
        df['suspeito'] = False
    if 'protocolo' not in df.columns:
        df['protocolo'] = None
    df['lote_arquivo'] = df['arquivo'].apply(extract_lote_from_filename)
    df['lote_arquivo_int'] = pd.to_numeric(df['lote_arquivo'], errors='coerce').astype('Int64')
    if 'numero_lote' in df.columns:
        df['lote_confere'] = (df['lote_arquivo'].fillna('') == df['numero_lote'].fillna(''))
    else:
        df['lote_confere'] = pd.NA
    if 'erro' not in df.columns:
        df['erro'] = None
    for c in ('valor_glosado', 'valor_liberado'):
        if c not in df.columns:
            df[c] = 0.0
    ordenar = [
        'numero_lote', 'protocolo', 'tipo', 'qtde_guias',
        'valor_total', 'valor_glosado', 'valor_liberado', 'estrategia_total',
        'arquivo', 'lote_arquivo', 'lote_arquivo_int', 'lote_confere',
        'suspeito', 'erro', 'parser_version'
    ]
    cols = [c for c in ordenar if c in df.columns] + [c for c in df.columns if c not in ordenar]
    df = df[cols].sort_values(['numero_lote', 'tipo', 'arquivo'], ignore_index=True)
    return df

def _make_agg(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=['numero_lote', 'tipo', 'qtde_arquivos', 'qtde_guias_total', 'valor_total'])
    agg = df.groupby(['numero_lote', 'tipo'], dropna=False, as_index=False).agg(
        qtde_arquivos=('arquivo', 'count'),
        qtde_guias_total=('qtde_guias', 'sum'),
        valor_total=('valor_total', 'sum')
    ).sort_values(['numero_lote', 'tipo'], ignore_index=True)
    return agg

# =========================================================
# Upload
# =========================================================
with tab1:
    files = st.file_uploader("Selecione um ou mais arquivos XML TISS", type=['xml'], accept_multiple_files=True)
    demo_files = st.file_uploader("Opcional: Demonstrativos (.xlsx)", type=['xlsx'], accept_multiple_files=True, key="demo_upload_tab1")

    st.markdown("### Banco de Demonstrativos (acumulado)")
    bcol1, bcol2, bcol3 = st.columns([1,1,2])
    with bcol1:
        add_disabled = not bool(demo_files)
        if st.button("➕ Adicionar demonstrativo(s) ao banco", disabled=add_disabled, use_container_width=True):
            try:
                demos = []
                for f in demo_files:
                    if hasattr(f, 'seek'):
                        f.seek(0)
                    demos.append(ler_demonstrativo_pagto_xlsx(f))
                if demos:
                    _add_to_demo_bank(pd.concat(demos, ignore_index=True))
                    st.success(f"{len(demos)} demonstrativo(s) adicionado(s). Lotes únicos: {st.session_state.demo_bank['numero_lote'].nunique()}")
            except Exception as e:
                st.error(f"Erro ao processar demonstrativo(s): {e}")
    with bcol2:
        if st.button("🗑️ Limpar banco", use_container_width=True):
            _clear_demo_bank()
            st.info("Banco limpo.")
    with bcol3:
        if not st.session_state.demo_bank.empty:
            lotes = st.session_state.demo_bank['numero_lote'].nunique()
            st.caption(f"**{lotes}** lote(s) no banco.")

    demo_agg_in_use = st.session_state.demo_bank.copy()

    if files:
        resultados: List[Dict] = []
        for f in files:
            try:
                if hasattr(f, "seek"):
                    f.seek(0)
                res = parse_tiss_xml(f)
                res['arquivo'] = f.name
                if 'erro' not in res:
                    res['erro'] = None
            except Exception as e:
                res = {'arquivo': f.name, 'numero_lote': '', 'tipo': 'DESCONHECIDO', 'qtde_guias': 0, 'valor_total': Decimal('0'), 'estrategia_total': 'erro', 'parser_version': PARSER_VERSION, 'erro': str(e)}
            resultados.append(res)

        if resultados:
            df = pd.DataFrame(resultados)
            df = _df_format(df)

            st.subheader("Resumo por arquivo (XML)")
            st.dataframe(_df_display_currency(df, ['valor_total', 'valor_glosado', 'valor_liberado']), use_container_width=True)

            st.subheader("Agregado por nº do lote e tipo (XML)")
            agg = _make_agg(df)
            st.dataframe(_df_display_currency(agg, ['valor_total']), use_container_width=True)

            # =========================================================
            # 🔎 Auditoria por guia e 🧩 Comparar/remover duplicadas
            # =========================================================
            with st.expander("🔎 Auditoria por guia (opcional)"):
                arquivo_escolhido = st.selectbox("Selecione um arquivo enviado", options=[r['arquivo'] for r in resultados])
                if st.button("Gerar auditoria do arquivo selecionado", type="primary"):
                    escolhido = next((f for f in files if f.name == arquivo_escolhido), None)
                    if escolhido is not None:
                        if hasattr(escolhido, "seek"):
                            escolhido.seek(0)
                        linhas = audit_por_guia(escolhido)
                        df_a = pd.DataFrame(linhas)
                        df_a_disp = df_a.copy()
                        for c in ('total_tag', 'subtotal_itens_proc', 'subtotal_itens_outras', 'subtotal_itens'):
                            if c in df_a_disp.columns:
                                df_a_disp[c] = df_a_disp[c].apply(format_currency_br)
                        st.dataframe(df_a_disp, use_container_width=True)
                        st.download_button("Baixar auditoria (CSV)", df_a.to_csv(index=False).encode('utf-8'), file_name=f"auditoria_{arquivo_escolhido}.csv", mime="text/csv")

            with st.expander("🧩 Comparar XML e remover guias duplicadas"):
                arquivo_base = st.selectbox("Selecione o arquivo base", options=[r['arquivo'] for r in resultados])
                if st.button("Remover guias duplicadas do arquivo base", type="primary"):
                    base_file = next((f for f in files if f.name == arquivo_base), None)
                    outros_files = [f for f in files if f.name != arquivo_base]

                    if base_file is None or not outros_files:
                        st.warning("É necessário selecionar um arquivo base e ter outros arquivos para comparar.")
                    else:
                        if hasattr(base_file, "seek"):
                            base_file.seek(0)
                        guias_base = audit_por_guia(base_file)

                        guias_outros = []
                        for f in outros_files:
                            if hasattr(f, "seek"):
                                f.seek(0)
                            guias_outros.extend(audit_por_guia(f))

                        duplicadas = []
                        for g in guias_base:
                            chave = None
                            if g['tipo'] in ('CONSULTA', 'SADT'):
                                chave = g.get('numeroGuiaPrestador')
                            elif g['tipo'] == 'RECURSO':
                                chave = g.get('numeroGuiaOrigem') or g.get('numeroGuiaOperadora')
                            if chave:
                                for o in guias_outros:
                                    chave_outro = None
                                    if o['tipo'] in ('CONSULTA', 'SADT'):
                                        chave_outro = o.get('numeroGuiaPrestador')
                                    elif o['tipo'] == 'RECURSO':
                                        chave_outro = o.get('numeroGuiaOrigem') or o.get('numeroGuiaOperadora')
                                    if chave_outro == chave:
                                        duplicadas.append(g)
                                        break

                        if not duplicadas:
                            st.success("Nenhuma guia duplicada encontrada.")
                        else:
                            st.warning(f"{len(duplicadas)} guia(s) duplicada(s) encontrada(s).")
                            df_dup = pd.DataFrame(duplicadas)
                            st.dataframe(df_dup, use_container_width=True)

                            from lxml import etree
                            base_file.seek(0)
                            parser = etree.XMLParser(remove_blank_text=True)
                            tree = etree.parse(base_file, parser)
                            root = tree.getroot()

                            def remover_guias(root, duplicadas):
                                for dup in duplicadas:
                                    tipo = dup['tipo']
                                    chave = dup.get('numeroGuiaPrestador') or dup.get('numeroGuiaOrigem') or dup.get('numeroGuiaOperadora')
                                    if not chave:
                                        continue
                                    if tipo == 'CONSULTA':
                                        for guia in root.xpath('.//ans:guiaConsulta', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}):
                                            num = guia.find('.//ans:numeroGuiaPrestador', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'})
                                            if num is not None and (num.text or '').strip() == chave:
                                                guia.getparent().remove(guia)
                                    elif tipo == 'SADT':
                                        for guia in root.xpath('.//ans:guiaSP-SADT', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}):
                                            num = guia.find('.//ans:cabecalhoGuia/ans:numeroGuiaPrestador', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'})
                                            if num is not None and (num.text or '').strip() == chave:
                                                guia.getparent().remove(guia)
                                    elif tipo == 'RECURSO':
                                        for guia in root.xpath('.//ans:recursoGuia', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}):
                                            num = guia.find('.//ans:numeroGuiaOrigem', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'})
                                            num2 = guia.find('.//ans:numeroGuiaOperadora', namespaces={'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'})
                                            if ((num is not None and (num.text or '').strip() == chave) or (num2 is not None and (num2.text or '').strip() == chave)):
                                                guia.getparent().remove(guia)
                                return root

                            root = remover_guias(root, duplicadas)
                            buffer_xml = io.BytesIO()
                            tree.write(buffer_xml, encoding="utf-8", xml_declaration=True, pretty_print=True)

                            st.download_button("Baixar XML sem duplicadas", data=buffer_xml.getvalue(), file_name=f"{arquivo_base.replace('.xml','')}_sem_duplicadas.xml", mime="application/xml")
