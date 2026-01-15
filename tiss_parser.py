
# file: app.py
import io
from pathlib import Path
from decimal import Decimal
import pandas as pd
import streamlit as st
from tiss_parser import parse_tiss_xml, parse_many_xmls

st.set_page_config(page_title="Leitor TISS XML", layout="wide")

st.title("Leitor de XML TISS (Consulta e SP‑SADT)")
st.caption("Extrai nº do lote, quantidade de guias e valor total por arquivo.")

tab1, tab2 = st.tabs(["Upload de XML(s)", "Ler de uma pasta local (clonada do GitHub)"])

def _df_format(df: pd.DataFrame) -> pd.DataFrame:
    # Garante tipos e formatação
    if 'valor_total' in df.columns:
        df['valor_total'] = df['valor_total'].apply(lambda x: float(Decimal(str(x))))
    ordenar = ['numero_lote', 'tipo', 'qtde_guias', 'valor_total', 'arquivo']
    cols = [c for c in ordenar if c in df.columns] + [c for c in df.columns if c not in ordenar]
    return df[cols].sort_values(['numero_lote', 'tipo', 'arquivo'], ignore_index=True)

with tab1:
    files = st.file_uploader(
        "Selecione um ou mais arquivos XML TISS",
        type=['xml'],
        accept_multiple_files=True
    )
    if files:
        resultados = []
        for f in files:
            data = f.read()
            # salvar em buffer para o parser do ElementTree
            tmp_path = Path(f.name)
            # ET aceita file-like; faremos via bytesIO para não gravar em disco:
            buff = io.BytesIO(data)
            try:
                import xml.etree.ElementTree as ET
                root = ET.parse(buff).getroot()
                # reaproveita núcleo de parsing: precisamos de um arquivo físico?
                # Para não duplicar lógica, vamos serializar temporariamente em memória:
                # Truque: escrever num NamedTemporaryFile se desejar, mas aqui
                # vamos adaptar parse_tiss_xml para receber root? Para simplificar,
                # reusamos o código do parser diretamente (copie a função se preferir).
            except Exception as e:
                st.error(f"Falha ao ler {f.name}: {e}")
                continue

            # Chamar parse_tiss_xml direto requer caminho; então criamos um arquivo temporário:
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".xml", delete=True) as tmp:
                tmp.write(data)
                tmp.flush()
                try:
                    res = parse_tiss_xml(tmp.name)
                    res['arquivo'] = f.name  # exibe o nome original do upload
                except Exception as e:
                    res = {'arquivo': f.name, 'numero_lote': '', 'tipo': 'DESCONHECIDO',
                           'qtde_guias': 0, 'valor_total': Decimal('0'), 'erro': str(e)}
                resultados.append(res)

        if resultados:
            df = pd.DataFrame(resultados)
            df = _df_format(df)
            st.subheader("Resumo por arquivo")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "Baixar resumo (CSV)",
                df.to_csv(index=False).encode('utf-8'),
                file_name="resumo_xml_tiss.csv",
                mime="text/csv"
            )

with tab2:
    pasta = st.text_input(
        "Caminho da pasta com XMLs (ex.: ./data ou C:\\\\repos\\\\tiss-xmls)",
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

                # Agregação por nº do lote (útil se houver mais de um arquivo do mesmo lote)
                st.subheader("Agregado por nº do lote")
                agg = df.groupby(['numero_lote', 'tipo'], dropna=False, as_index=False).agg(
                    qtde_arquivos=('arquivo', 'count'),
                    qtde_guias_total=('qtde_guias', 'sum'),
                    valor_total=('valor_total', 'sum')
                )
                st.dataframe(agg, use_container_width=True)

                st.download_button(
                    "Baixar resumo (CSV)",
                    df.to_csv(index=False).encode('utf-8'),
                    file_name="resumo_xml_tiss.csv",
                    mime="text/csv"
                )

