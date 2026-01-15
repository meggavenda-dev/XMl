
# file: app.py
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
# FORMATAÇÃO DE MOEDA (BR) — HOTFIX ROBUSTO
# =========================================================
def format_currency_br(val) -> str:
    """
    Converte número em string 'R$ 1.234,56'.
    - Valores None/NaN/Inf/Inválidos -> 'R$ 0,00'
    - Mantém sinal negativo com prefixo '-'
    """
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
# Extração "lote" a partir do nome do arquivo
# =========================================================
_LOTE_REGEX = re.compile(r'(?i)lote\s*[-_]*\s*(\d+)')

def extract_lote_from_filename(name: str) -> str | None:
    if not isinstance(name, str):
        return None
    m = _LOTE_REGEX.search(name)
    if m:
        return m.group(1)
    # Fallback opcional: descomente se quiser capturar primeiro número com >=5 dígitos
    # m2 = re.search(r'(\d{5,})', name)
    # return m2.group(1) if m2 else None
    return None

# =========================================================
# Utils de dataframe/saída
# =========================================================
def _to_float(val) -> float:
    try:
        return float(Decimal(str(val)))
    except Exception:
        return 0.0

def _df_format(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garantia de tipos, colunas auxiliares e ordenação para exibição/CSV.
    Adiciona:
      - suspeito: qtde_guias>0 e valor_total==0
      - lote_arquivo (+ lote_arquivo_int): extraídos do nome
      - lote_confere: confere lote_arquivo == numero_lote
    """
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
# Leitura do Demonstrativo de Pagamento (.xlsx)
# =========================================================
def _norm_lote(v) -> str | None:
    """Normaliza 'Lote' para string compatível com numero_lote do XML (remove '.0', pega só dígitos)."""
    if pd.isna(v):
        return None
    s = str(v).strip()
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    digits = ''.join(ch for ch in s if ch.isdigit())
    return digits if digits else s

def ler_demonstrativo_pagto_xlsx(source) -> pd.DataFrame:
    """
    Lê a planilha 'DemonstrativoAnaliseDeContas' e agrega por (numero_lote, competencia):
      - valor_apresentado, valor_apurado (liberado), valor_glosa, linhas
    Retorna colunas: numero_lote | competencia | valor_apresentado | valor_apurado | valor_glosa | linhas
    """
    df_raw = pd.read_excel(source, sheet_name='DemonstrativoAnaliseDeContas', engine='openpyxl')

    mask = df_raw.iloc[:, 0].astype(str).str.strip().eq('CPF/CNPJ')
    if not mask.any():
        raise ValueError("Cabeçalho 'CPF/CNPJ' não encontrado no Demonstrativo.")
    header_idx = mask.idxmax()

    df = df_raw.iloc[header_idx:]
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)

    need = ['Lote', 'Competência', 'Valor Apresentado', 'Valor Apurado', 'Valor Glosa']
    faltando = [c for c in need if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas ausentes no Demonstrativo: {faltando}")

    df['numero_lote'] = df['Lote'].apply(_norm_lote)
    for c in ['Valor Apresentado', 'Valor Apurado', 'Valor Glosa']:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

    demo_agg = (
        df.groupby(['numero_lote', df['Competência'].astype(str).str.strip()], dropna=False)
          .agg(valor_apresentado=('Valor Apresentado', 'sum'),
               valor_apurado=('Valor Apurado', 'sum'),
               valor_glosa=('Valor Glosa', 'sum'),
               linhas=('numero_lote', 'count'))
          .reset_index()
          .rename(columns={'Competência': 'competencia'})
    )
    return demo_agg

# =========================================================
# Conciliação — chave inteligente (prefere o que EXISTE no Demonstrativo)
# =========================================================
def _build_chave_concil(df_xml: pd.DataFrame, demo_agg: pd.DataFrame) -> pd.DataFrame:
    """
    Cria as colunas numero_lote_norm, lote_arquivo_norm e chave_concil.
    Regra:
      - Se numero_lote_norm estiver NOS LOTES do Demonstrativo -> usa numero_lote_norm
      - Senão, se lote_arquivo_norm estiver NOS LOTES do Demonstrativo -> usa lote_arquivo_norm
      - Heurística extra: se numero_lote_norm começa com lote_arquivo_norm e este existir no demo -> usa lote_arquivo_norm
      - Caso nada exista no demo, mantém numero_lote_norm (ou lote_arquivo_norm, se numero_lote_norm ausente)
    """
    df = df_xml.copy()
    demo_keys = set(demo_agg['numero_lote'].dropna().astype(str).str.strip()) if not demo_agg.empty else set()

    df['numero_lote_norm'] = df['numero_lote'].apply(_norm_lote)
    df['lote_arquivo_norm'] = df['lote_arquivo'].apply(_norm_lote)

    def choose(row):
        num = (row.get('numero_lote_norm') or '').strip()
        arq = (row.get('lote_arquivo_norm') or '').strip()

        if num and num in demo_keys:
            return num
        if arq and arq in demo_keys:
            return arq
        if num and arq and num.startswith(arq) and arq in demo_keys:
            return arq
        return num or arq or None

    df['chave_concil'] = df.apply(choose, axis=1)
    return df

# =========================================================
# Baixa por lote — com chave de conciliação inteligente
# =========================================================
def _make_baixa_por_lote(df_xml: pd.DataFrame, demo_agg: pd.DataFrame) -> pd.DataFrame:
    """
    Produz tabela de baixa por lote usando a chave de conciliação:
      - Preferir numero_lote_norm se existir no demonstrativo
      - Senão, usar lote_arquivo_norm (se existir no demonstrativo)
    """
    if df_xml.empty:
        return pd.DataFrame()

    tmp = _build_chave_concil(df_xml, demo_agg)

    xml_lote = (
        tmp.groupby('chave_concil', dropna=False)
           .agg(qtde_arquivos=('arquivo', 'count'),
                qtde_guias_xml=('qtde_guias', 'sum'),
                valor_total_xml=('valor_total', 'sum'),
                numero_lote=('numero_lote', 'first'),
                lote_arquivo=('lote_arquivo', 'first'))
           .reset_index()
    )

    demo_key = demo_agg.copy().assign(chave_concil=demo_agg['numero_lote'])

    baixa = xml_lote.merge(
        demo_key[['chave_concil', 'competencia', 'valor_apresentado', 'valor_apurado', 'valor_glosa']],
        on='chave_concil', how='left'
    )

    baixa['apresentado_diff'] = (baixa['valor_total_xml'] - baixa['valor_apresentado']).fillna(0.0)
    baixa['apresentado_confere'] = baixa['apresentado_diff'].abs() <= 0.01

    baixa['liberado_plus_glosa'] = (baixa['valor_apurado'].fillna(0.0) + baixa['valor_glosa'].fillna(0.0))
    baixa['demonstrativo_confere'] = (baixa['valor_apresentado'].fillna(0.0) - baixa['liberado_plus_glosa']).abs() <= 0.01

    baixa = baixa.sort_values('chave_concil', ignore_index=True)
    return baixa

# =========================================================
# Export Excel
# =========================================================
def _download_excel_button(df_resumo: pd.DataFrame, df_agg: pd.DataFrame, df_terceira: pd.DataFrame, label: str):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        (df_resumo if not df_resumo.empty else pd.DataFrame()).to_excel(
            writer, index=False, sheet_name="Resumo por arquivo"
        )
        (df_agg if not df_agg.empty else pd.DataFrame()).to_excel(
            writer, index=False, sheet_name="Agregado por lote"
        )
        sheet_name_3 = "Baixa por lote" if ('valor_total_xml' in df_terceira.columns) else "Auditoria"
        (df_terceira if not df_terceira.empty else pd.DataFrame()).to_excel(
            writer, index=False, sheet_name=sheet_name_3
        )

        def format_currency_sheet(ws, header_row=1, currency_cols=()):
            headers = {ws.cell(row=header_row, column=c).value: c for c in range(1, ws.max_column + 1)}
            numfmt = 'R$ #,##0.00'
            for col_name in currency_cols:
                if col_name in headers:
                    col_idx = headers[col_name]
                    for r in range(header_row + 1, ws.max_row + 1):
                        ws.cell(row=r, column=col_idx).number_format = numfmt

        ws = writer.sheets.get("Resumo por arquivo")
        if ws is not None:
            format_currency_sheet(ws, currency_cols=("valor_total", "valor_glosado", "valor_liberado"))

        ws = writer.sheets.get("Agregado por lote")
        if ws is not None:
            format_currency_sheet(ws, currency_cols=("valor_total",))

        ws = writer.sheets.get("Baixa por lote") or writer.sheets.get("Auditoria")
        if ws is not None:
            if writer.sheets.get("Baixa por lote") is not None:
                format_currency_sheet(ws, currency_cols=(
                    "valor_total_xml", "valor_apresentado", "valor_apurado",
                    "valor_glosa", "liberado_plus_glosa", "apresentado_diff"
                ))
            else:
                format_currency_sheet(ws, currency_cols=("total_tag","subtotal_itens_proc","subtotal_itens_outras","subtotal_itens"))

    st.download_button(
        label,
        data=buffer.getvalue(),
        file_name="resumo_xml_tiss.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def _auditar_alertas(df: pd.DataFrame) -> None:
    if df.empty:
        return
    sus = df[df['suspeito']]
    err = df[df['erro'].notna()] if 'erro' in df.columns else pd.DataFrame()

    if not sus.empty:
        st.warning(
            f"⚠️ {len(sus)} arquivo(s) com valor_total=0 e qtde_guias>0. Verifique: "
            + ", ".join(sus['arquivo'].tolist())[:500]
        )
    if not err.empty:
        st.error(
            f"❌ {len(err)} arquivo(s) com erro no parsing. Exemplos: "
            + ", ".join(err['arquivo'].head(5).tolist())
        )

# =========================================================
# 🔒 Banco acumulado de Demonstrativos (session_state)
# =========================================================
if 'demo_bank' not in st.session_state:
    st.session_state.demo_bank = pd.DataFrame(
        columns=['numero_lote', 'competencia', 'valor_apresentado', 'valor_apurado', 'valor_glosa', 'linhas']
    )

def _agg_demo(df: pd.DataFrame) -> pd.DataFrame:
    """Garante agregação por (numero_lote, competencia) após concatenação de múltiplos demonstrativos."""
    if df.empty:
        return df
    df = df.copy()
    return (df.groupby(['numero_lote', 'competencia'], dropna=False, as_index=False)
              .agg(valor_apresentado=('valor_apresentado','sum'),
                   valor_apurado=('valor_apurado','sum'),
                   valor_glosa=('valor_glosa','sum'),
                   linhas=('linhas','sum')))

def _add_to_demo_bank(demo_new: pd.DataFrame):
    bank = st.session_state.demo_bank
    bank = pd.concat([bank, demo_new], ignore_index=True)
    st.session_state.demo_bank = _agg_demo(bank)

def _clear_demo_bank():
    st.session_state.demo_bank = st.session_state.demo_bank.iloc[0:0]

# =========================================================
# Upload
# =========================================================
with tab1:
    files = st.file_uploader(
        "Selecione um ou mais arquivos XML TISS (Consulta, SP‑SADT ou Recurso de Glosa)",
        type=['xml'],
        accept_multiple_files=True
    )
    demo_files = st.file_uploader(
        "Opcional: Anexe um ou mais Demonstrativos de Pagamento (.xlsx) e adicione-os ao banco acumulado",
        type=['xlsx'],
        accept_multiple_files=True,
        key="demo_upload_tab1"
    )

    # ---- Painel do Banco de Demonstrativos
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
                    st.success(f"{len(demos)} demonstrativo(s) adicionado(s). "
                               f"Lotes únicos no banco: {st.session_state.demo_bank['numero_lote'].nunique()}")
            except Exception as e:
                st.error(f"Erro ao processar demonstrativo(s): {e}")
    with bcol2:
        if st.button("🗑️ Limpar banco", use_container_width=True):
            _clear_demo_bank()
            st.info("Banco de demonstrativos limpo.")
    with bcol3:
        if not st.session_state.demo_bank.empty:
            lotes = st.session_state.demo_bank['numero_lote'].nunique()
            st.caption(f"**{lotes}** lote(s) no banco. Use a seção abaixo normalmente — a conciliação "
                       f"usará o banco acumulado automaticamente.")

    # ---- Rodar a leitura dos XMLs
    demo_agg_in_use = st.session_state.demo_bank.copy()  # sempre prioriza o banco, se existir

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
                res = {
                    'arquivo': f.name,
                    'numero_lote': '',
                    'tipo': 'DESCONHECIDO',
                    'qtde_guias': 0,
                    'valor_total': Decimal('0'),
                    'estrategia_total': 'erro',
                    'parser_version': PARSER_VERSION,
                    'erro': str(e),
                }
            resultados.append(res)

        if resultados:
            df = pd.DataFrame(resultados)
            df = _df_format(df)

            # >>> Merge com Demonstrativo (usando o banco se existir)
            if not demo_agg_in_use.empty:
                df_keys = _build_chave_concil(df, demo_agg_in_use)

                map_glosa   = dict(zip(demo_agg_in_use['numero_lote'], demo_agg_in_use['valor_glosa']))
                map_apurado = dict(zip(demo_agg_in_use['numero_lote'], demo_agg_in_use['valor_apurado']))
                map_comp    = dict(zip(demo_agg_in_use['numero_lote'], demo_agg_in_use['competencia']))

                df['numero_lote_norm'] = df_keys['numero_lote_norm']
                df['lote_arquivo_norm'] = df_keys['lote_arquivo_norm']
                df['chave_concil'] = df_keys['chave_concil']

                df['valor_glosado']  = df['chave_concil'].map(map_glosa).fillna(df['valor_glosado']).fillna(0.0)
                df['valor_liberado'] = df['chave_concil'].map(map_apurado).fillna(df['valor_liberado']).fillna(0.0)

                df['competencia'] = df.get('competencia', pd.Series([None] * len(df)))
                df['competencia'] = df['chave_concil'].map(map_comp).fillna(df['competencia'])
            # <<< Fim do merge

            df_disp = _df_display_currency(df, ['valor_total', 'valor_glosado', 'valor_liberado'])

            st.subheader("Resumo por arquivo (XML)")
            st.dataframe(df_disp, use_container_width=True)

            st.subheader("Agregado por nº do lote e tipo (XML)")
            agg = _make_agg(df)
            agg_disp = _df_display_currency(agg, ['valor_total'])
            st.dataframe(agg_disp, use_container_width=True)

            baixa = pd.DataFrame()
            if not demo_agg_in_use.empty:
                st.subheader("Baixa por nº do lote (XML × Demonstrativo)")
                baixa = _make_baixa_por_lote(df, demo_agg_in_use)
                baixa_disp = baixa.copy()
                for c in ['valor_total_xml', 'valor_apresentado', 'valor_apurado',
                          'valor_glosa', 'liberado_plus_glosa', 'apresentado_diff']:
                    if c in baixa_disp.columns:
                        baixa_disp[c] = baixa_disp[c].fillna(0.0).apply(format_currency_br)
                st.dataframe(baixa_disp, use_container_width=True)

            _auditar_alertas(df)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    "Baixar resumo (CSV)",
                    df.to_csv(index=False).encode('utf-8'),
                    file_name="resumo_xml_tiss.csv",
                    mime="text/csv"
                )
            with col2:
                _download_excel_button(df, agg, baixa if not baixa.empty else df, "Baixar resumo (Excel .xlsx)")
            with col3:
                st.caption("O Excel inclui as abas: Resumo, Agregado e Auditoria/Baixa (moeda BR).")

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
                        st.download_button(
                            "Baixar auditoria (CSV)",
                            df_a.to_csv(index=False).encode('utf-8'),
                            file_name=f"auditoria_{arquivo_escolhido}.csv",
                            mime="text/csv"
                        )

# =========================================================
# Pasta local (útil para rodar local/clonado)
# =========================================================
with tab2:
    pasta = st.text_input(
        "Caminho da pasta com XMLs (ex.: ./data ou C:\\repos\\tiss-xmls)",
        value="./data"
    )
    st.caption("Esta aba reutiliza o mesmo **banco de demonstrativos** da aba de Upload (acumulado).")

    demo_files_local = st.file_uploader(
        "Adicionar mais Demonstrativos (.xlsx) ao banco (opcional)",
        type=['xlsx'],
        accept_multiple_files=True,
        key="demo_upload_tab2"
    )
    lcol1, lcol2 = st.columns([1,1])
    with lcol1:
        add_disabled_local = not bool(demo_files_local)
        if st.button("➕ Adicionar demonstrativo(s) ao banco (aba Pasta)", disabled=add_disabled_local, use_container_width=True):
            try:
                demos = []
                for f in demo_files_local:
                    if hasattr(f, 'seek'):
                        f.seek(0)
                    demos.append(ler_demonstrativo_pagto_xlsx(f))
                if demos:
                    _add_to_demo_bank(pd.concat(demos, ignore_index=True))
                    st.success(f"{len(demos)} demonstrativo(s) adicionado(s). "
                               f"Lotes únicos no banco: {st.session_state.demo_bank['numero_lote'].nunique()}")
            except Exception as e:
                st.error(f"Erro ao processar demonstrativo(s): {e}")
    with lcol2:
        if st.button("🗑️ Limpar banco (aba Pasta)", use_container_width=True):
            _clear_demo_bank()
            st.info("Banco de demonstrativos limpo.")

    if st.button("Ler pasta"):
        p = Path(pasta)
        if not p.exists():
            st.error("Pasta não encontrada.")
        else:
            xmls = list(p.glob("*.xml"))
            if not xmls:
                st.warning("Nenhum .xml encontrado nessa pasta.")
            else:
                resultados = parse_many_xmls(xmls)
                df = pd.DataFrame(resultados)
                df = _df_format(df)

                demo_agg_in_use = st.session_state.demo_bank.copy()

                baixa_local = pd.DataFrame()
                if not demo_agg_in_use.empty:
                    df_keys = _build_chave_concil(df, demo_agg_in_use)

                    map_glosa   = dict(zip(demo_agg_in_use['numero_lote'], demo_agg_in_use['valor_glosa']))
                    map_apurado = dict(zip(demo_agg_in_use['numero_lote'], demo_agg_in_use['valor_apurado']))
                    map_comp    = dict(zip(demo_agg_in_use['numero_lote'], demo_agg_in_use['competencia']))

                    df['numero_lote_norm'] = df_keys['numero_lote_norm']
                    df['lote_arquivo_norm'] = df_keys['lote_arquivo_norm']
                    df['chave_concil'] = df_keys['chave_concil']

                    df['valor_glosado']  = df['chave_concil'].map(map_glosa).fillna(df['valor_glosado']).fillna(0.0)
                    df['valor_liberado'] = df['chave_concil'].map(map_apurado).fillna(df['valor_liberado']).fillna(0.0)
                    df['competencia']    = df['chave_concil'].map(map_comp).fillna(df.get('competencia', pd.Series([None]*len(df))))

                    baixa_local = _make_baixa_por_lote(df, demo_agg_in_use)

                df_disp = _df_display_currency(df, ['valor_total', 'valor_glosado', 'valor_liberado'])

                st.subheader("Resumo por arquivo")
                st.dataframe(df_disp, use_container_width=True)

                st.subheader("Agregado por nº do lote e tipo")
                agg = _make_agg(df)
                agg_disp = _df_display_currency(agg, ['valor_total'])
                st.dataframe(agg_disp, use_container_width=True)

                if not baixa_local.empty:
                    st.subheader("Baixa por nº do lote (XML × Demonstrativo)")
                    baixa_disp = baixa_local.copy()
                    for c in ['valor_total_xml', 'valor_apresentado', 'valor_apurado',
                              'valor_glosa', 'liberado_plus_glosa', 'apresentado_diff']:
                        if c in baixa_disp.columns:
                            baixa_disp[c] = baixa_disp[c].fillna(0.0).apply(format_currency_br)
                    st.dataframe(baixa_disp, use_container_width=True)

                _auditar_alertas(df)

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.download_button(
                        "Baixar resumo (CSV)",
                        df.to_csv(index=False).encode('utf-8'),
                        file_name="resumo_xml_tiss.csv",
                        mime="text/csv"
                    )
                with col2:
                    _download_excel_button(df, agg, baixa_local if not baixa_local.empty else df, "Baixar resumo (Excel .xlsx)")
                with col3:
                    st.caption("O Excel inclui as abas: Resumo, Agregado e Auditoria/Baixa (moeda BR).")
