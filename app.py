
# file: app.py
from __future__ import annotations

import io
from decimal import Decimal
from pathlib import Path
import pandas as pd
import streamlit as st

from tiss_parser import parse_tiss_xml, parse_many_xmls

st.set_page_config(page_title="Leitor TISS XML", layout="wide")
st.title("Leitor de XML TISS (Consulta e SP‑SADT)")
st.caption("Extrai nº do lote, quantidade de guias e valor total por arquivo.")

tab1, tab2 = st.tabs(["Upload de XML(s)", "Ler de uma pasta local (clonada do GitHub)"])

def _df_format(df: pd.DataFrame) -> pd.DataFrame:
    """Garante tipos e formatação mínima para exibição/CSV."""
    if 'valor_total' in df.columns:
        df['valor_total'] = df['valor_total'].apply(lambda x: float(Decimal(str(x))))
    ordenar = ['numero_lote', 'tipo', 'qtde_guias', 'valor_total', 'arquivo']
    cols = [c for c in ordenar if c in df.columns] + [c for c in df.columns if c not in ordenar]
    return df[cols].sort_values(['numero_lote', 'tipo', 'arquivo'], ignore_index=True)

def _make_agg(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=['numero_lote', 'tipo', 'qtde_arquivos', 'qtde_guias_total', 'valor_total'])
    return df.groupby(['numero_lote', 'tipo'], dropna=False, as_index=False).agg(
        qtde_arquivos=('arquivo', 'count'),
        qtde_guias_total=('qtde_guias', 'sum'),
        valor_total=('valor_total', 'sum')
    ).sort_values(['numero_lote', 'tipo'], ignore_index=True)

def _download_excel_button(df_resumo: pd.DataFrame, df_agg: pd.DataFrame, label: str):
    # Gera um Excel em memória com abas "Resumo por arquivo" e "Agregado por lote"
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        (df_resumo if not df_resumo.empty else pd.DataFrame()).to_excel(
            writer, index=False, sheet_name="Resumo por arquivo"
        )
        (df_agg if not df_agg.empty else pd.DataFrame()).to_excel(
            writer, index=False, sheet_name="Agregado por lote"
        )
    st.download_button(
        label,
        data=buffer.getvalue(),
        file_name="resumo_xml_tiss.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with tab1:
    files = st.file_uploader(
        "Selecione um ou mais arquivos XML TISS",
        type=['xml'],
        accept_multiple_files=True
    )
    if files:
        resultados = []
        for f in files:
            try:
                # Passa UploadedFile (file-like) direto para o parser
                res = parse_tiss_xml(f)
                res['arquivo'] = f.name  # garante nome exibido igual ao upload
            except Exception as e:
                res = {
                    'arquivo': f.name, 'numero_lote': '', 'tipo': 'DESCONHECIDO',
                    'qtde_guias': 0, 'valor_total': Decimal('0'), 'erro': str(e)
                }
            resultados.append(res)

        if resultados:
            df = pd.DataFrame(resultados)
            df = _df_format(df)
            st.subheader("Resumo por arquivo")
            st.dataframe(df, use_container_width=True)

            agg = _make_agg(df)
            st.subheader("Agregado por nº do lote")
            st.dataframe(agg, use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "Baixar resumo (CSV)",
                    df.to_csv(index=False).encode('utf-8'),
                    file_name="resumo_xml_tiss.csv",
                    mime="text/csv"
                )
            with col2:
                _download_excel_button(df, agg, "Baixar resumo (Excel .xlsx)")

with tab2:
    pasta = st.text_input(
        "Caminho da pasta com XMLs (ex.: ./data ou C:\\repos\\tiss-xmls)",
        value="./data"
    )
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

                st.subheader("Resumo por arquivo")
                st.dataframe(df, use_container_width=True)

                agg = _make_agg(df)
                st.subheader("Agregado por nº do lote")
                st.dataframe(agg, use_container_width=True)

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        "Baixar resumo (CSV)",
                        df.to_csv(index=False).encode('utf-8'),
                        file_name="resumo_xml_tiss.csv",
                        mime="text/csv"
                    )
                with col2:
                    _download_excel_button(df, agg, "Baixar resumo (Excel .xlsx)")
