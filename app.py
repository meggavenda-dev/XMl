
# =========================================================
# TISS — Conciliação XML + Demonstrativo (AMHP)
# Template reescrito do zero — inclui todas as funções originais
# Melhorias: 1) Guia SADT robusto / 2) Chave com lote
#            3) Normalização rígida / 4) Fallback aprimorado
# Auditoria mantida mas desativada (# auditoria desativado)
# =========================================================

from __future__ import annotations
import io, os, re, json, unicodedata
from pathlib import Path
from typing import List, Dict, Optional, Union, IO, Tuple
from decimal import Decimal
from datetime import datetime

import pandas as pd
import xml.etree.ElementTree as ET
import streamlit as st

# =======================
# Configuração do Streamlit
# =======================
st.set_page_config(page_title="TISS • Conciliação & Analytics (Auditoria desativada)",
                   layout="wide")

st.title("TISS — XML → Conciliação com Demonstrativo AMHP + Analytics")
st.caption("Versão reescrita com chave por lote+guia+código, normalização rígida, "
           "fallback inteligente e auditoria desativada.")

# =======================
# Helpers gerais
# =======================
ANS_NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}
DEC_ZERO = Decimal("0")

def dec(v: Optional[str]) -> Decimal:
    if v is None:
        return DEC_ZERO
    s = str(v).strip().replace(",", ".")
    try:
        return Decimal(s)
    except:
        return DEC_ZERO

def tx(el: Optional[ET.Element]) -> str:
    return (el.text or "").strip() if el is not None else ""

def normalize_code(raw: str) -> str:
    if raw is None:
        return ""
    s = re.sub(r"[.\-_/ \t]", "", str(raw)).strip()
    return s.lstrip("0")  # rígido: sempre remove zeros à esquerda

def clean_text(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.strip().lower())

def f_money(x):
    try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def apply_money(df, cols):
    df2 = df.copy()
    for c in cols:
        if c in df2.columns:
            df2[c] = df2[c].apply(f_money)
    return df2

# =========================================================
# PARTE 2 — Parser XML TISS
# =========================================================

def extract_sadt_guia_id(guia: ET.Element) -> str:
    """
    Melhoria 1: busca universal do numeroGuiaPrestador:
    - cabecalhoGuia/numeroGuiaPrestador
    - numeroGuiaPrestador direto no nó
    - identificacaoGuia/numeroGuiaPrestador
    """
    paths = [
        ".//ans:cabecalhoGuia/ans:numeroGuiaPrestador",
        ".//ans:numeroGuiaPrestador",
        ".//ans:identificacaoGuia/ans:numeroGuiaPrestador",
    ]
    for p in paths:
        found = guia.find(p, ANS_NS)
        if found is not None and tx(found):
            return tx(found)
    return ""

def extract_xml_items(source: Union[IO[bytes], str, Path]) -> List[Dict]:
    """Lê um XML TISS inteiro e devolve todos os itens."""
    if hasattr(source, "read"):
        source.seek(0)
        root = ET.parse(source).getroot()
        fname = getattr(source, "name", "upload.xml")
    else:
        p = Path(source)
        root = ET.parse(p).getroot()
        fname = p.name

    # Lote
    lote = ""
    for path in [
        ".//ans:prestadorParaOperadora/ans:loteGuias/ans:numeroLote",
        ".//ans:loteGuias/ans:numeroLote",
        ".//ans:guiaRecursoGlosa/ans:numeroLote"
    ]:
        el = root.find(path, ANS_NS)
        if el is not None and tx(el):
            lote = tx(el)
            break

    itens = []

    # CONSULTA
    for g in root.findall(".//ans:guiaConsulta", ANS_NS):
        guia_id = tx(g.find(".//ans:numeroGuiaPrestador", ANS_NS))
        if not guia_id:
            guia_id = extract_sadt_guia_id(g)
        paciente = tx(g.find(".//ans:dadosBeneficiario/ans:nomeBeneficiario", ANS_NS))
        medico   = tx(g.find(".//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional", ANS_NS))
        data_atd = tx(g.find(".//ans:dataAtendimento", ANS_NS))

        proc = g.find(".//ans:procedimento", ANS_NS)
        cod_tab = tx(proc.find("ans:codigoTabela", ANS_NS)) if proc is not None else ""
        cod_proc = tx(proc.find("ans:codigoProcedimento", ANS_NS)) if proc is not None else ""
        desc = tx(proc.find("ans:descricaoProcedimento", ANS_NS)) if proc is not None else ""
        valor = dec(tx(proc.find("ans:valorProcedimento", ANS_NS))) if proc is not None else DEC_ZERO

        itens.append(dict(
            arquivo=fname,
            numero_lote=lote,
            tipo_guia="CONSULTA",
            numeroGuiaPrestador=guia_id,
            numeroGuiaOperadora="",
            paciente=paciente,
            medico=medico,
            data_atendimento=data_atd,
            codigo_tabela=cod_tab,
            codigo_procedimento=cod_proc,
            descricao_procedimento=desc,
            quantidade=1,
            valor_unitario=valor,
            valor_total=valor,
        ))

    # SADT
    for g in root.findall(".//ans:guiaSP-SADT", ANS_NS):
        guia_id = extract_sadt_guia_id(g)
        paciente = tx(g.find(".//ans:dadosBeneficiario/ans:nomeBeneficiario", ANS_NS))
        medico   = tx(g.find(".//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional", ANS_NS))
        data_atd = tx(g.find(".//ans:dataAtendimento", ANS_NS))

        # procedimentos executados
        for pe in g.findall(".//ans:procedimentoExecutado", ANS_NS):
            pr = pe.find("ans:procedimento", ANS_NS)
            cod_tab = tx(pr.find("ans:codigoTabela", ANS_NS)) if pr is not None else ""
            cod_proc = tx(pr.find("ans:codigoProcedimento", ANS_NS)) if pr is not None else ""
            desc = tx(pr.find("ans:descricaoProcedimento", ANS_NS)) if pr is not None else ""
            qtd = dec(tx(pe.find("ans:quantidadeExecutada", ANS_NS)))
            vuni = dec(tx(pe.find("ans:valorUnitario", ANS_NS)))
            vtot = dec(tx(pe.find("ans:valorTotal", ANS_NS)))
            if not vtot and vuni and qtd:
                vtot = vuni * qtd

            itens.append(dict(
                arquivo=fname,
                numero_lote=lote,
                tipo_guia="SADT",
                numeroGuiaPrestador=guia_id,
                numeroGuiaOperadora=guia_id,
                paciente=paciente,
                medico=medico,
                data_atendimento=data_atd,
                codigo_tabela=cod_tab,
                codigo_procedimento=cod_proc,
                descricao_procedimento=desc,
                quantidade=float(qtd) if qtd else 1,
                valor_unitario=float(vuni),
                valor_total=float(vtot),
            ))

        # outras despesas
        for d in g.findall(".//ans:despesa", ANS_NS):
            svc = d.find("ans:servicosExecutados", ANS_NS)
            cod_tab = tx(svc.find("ans:codigoTabela", ANS_NS)) if svc is not None else ""
            cod_proc = tx(svc.find("ans:codigoProcedimento", ANS_NS)) if svc is not None else ""
            desc = tx(svc.find("ans:descricaoProcedimento", ANS_NS)) if svc is not None else ""
            qtd = dec(tx(svc.find("ans:quantidadeExecutada", ANS_NS))) if svc is not None else DEC_ZERO
            vuni = dec(tx(svc.find("ans:valorUnitario", ANS_NS))) if svc is not None else DEC_ZERO
            vtot = dec(tx(svc.find("ans:valorTotal", ANS_NS))) if svc is not None else DEC_ZERO
            if not vtot and vuni and qtd:
                vtot = vuni * qtd

            itens.append(dict(
                arquivo=fname,
                numero_lote=lote,
                tipo_guia="SADT",
                numeroGuiaPrestador=guia_id,
                numeroGuiaOperadora=guia_id,
                paciente=paciente,
                medico=medico,
                data_atendimento=data_atd,
                codigo_tabela=cod_tab,
                codigo_procedimento=cod_proc,
                descricao_procedimento=desc,
                quantidade=float(qtd) if qtd else 1,
                valor_unitario=float(vuni),
                valor_total=float(vtot),
            ))

    return itens


# =========================================================
# PARTE 2 — Parser XML TISS
# =========================================================

def extract_sadt_guia_id(guia: ET.Element) -> str:
    """
    Melhoria 1: busca universal do numeroGuiaPrestador:
    - cabecalhoGuia/numeroGuiaPrestador
    - numeroGuiaPrestador direto no nó
    - identificacaoGuia/numeroGuiaPrestador
    """
    paths = [
        ".//ans:cabecalhoGuia/ans:numeroGuiaPrestador",
        ".//ans:numeroGuiaPrestador",
        ".//ans:identificacaoGuia/ans:numeroGuiaPrestador",
    ]
    for p in paths:
        found = guia.find(p, ANS_NS)
        if found is not None and tx(found):
            return tx(found)
    return ""

def extract_xml_items(source: Union[IO[bytes], str, Path]) -> List[Dict]:
    """Lê um XML TISS inteiro e devolve todos os itens."""
    if hasattr(source, "read"):
        source.seek(0)
        root = ET.parse(source).getroot()
        fname = getattr(source, "name", "upload.xml")
    else:
        p = Path(source)
        root = ET.parse(p).getroot()
        fname = p.name

    # Lote
    lote = ""
    for path in [
        ".//ans:prestadorParaOperadora/ans:loteGuias/ans:numeroLote",
        ".//ans:loteGuias/ans:numeroLote",
        ".//ans:guiaRecursoGlosa/ans:numeroLote"
    ]:
        el = root.find(path, ANS_NS)
        if el is not None and tx(el):
            lote = tx(el)
            break

    itens = []

    # CONSULTA
    for g in root.findall(".//ans:guiaConsulta", ANS_NS):
        guia_id = tx(g.find(".//ans:numeroGuiaPrestador", ANS_NS))
        if not guia_id:
            guia_id = extract_sadt_guia_id(g)
        paciente = tx(g.find(".//ans:dadosBeneficiario/ans:nomeBeneficiario", ANS_NS))
        medico   = tx(g.find(".//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional", ANS_NS))
        data_atd = tx(g.find(".//ans:dataAtendimento", ANS_NS))

        proc = g.find(".//ans:procedimento", ANS_NS)
        cod_tab = tx(proc.find("ans:codigoTabela", ANS_NS)) if proc is not None else ""
        cod_proc = tx(proc.find("ans:codigoProcedimento", ANS_NS)) if proc is not None else ""
        desc = tx(proc.find("ans:descricaoProcedimento", ANS_NS)) if proc is not None else ""
        valor = dec(tx(proc.find("ans:valorProcedimento", ANS_NS))) if proc is not None else DEC_ZERO

        itens.append(dict(
            arquivo=fname,
            numero_lote=lote,
            tipo_guia="CONSULTA",
            numeroGuiaPrestador=guia_id,
            numeroGuiaOperadora="",
            paciente=paciente,
            medico=medico,
            data_atendimento=data_atd,
            codigo_tabela=cod_tab,
            codigo_procedimento=cod_proc,
            descricao_procedimento=desc,
            quantidade=1,
            valor_unitario=valor,
            valor_total=valor,
        ))

    # SADT
    for g in root.findall(".//ans:guiaSP-SADT", ANS_NS):
        guia_id = extract_sadt_guia_id(g)
        paciente = tx(g.find(".//ans:dadosBeneficiario/ans:nomeBeneficiario", ANS_NS))
        medico   = tx(g.find(".//ans:dadosProfissionaisResponsaveis/ans:nomeProfissional", ANS_NS))
        data_atd = tx(g.find(".//ans:dataAtendimento", ANS_NS))

        # procedimentos executados
        for pe in g.findall(".//ans:procedimentoExecutado", ANS_NS):
            pr = pe.find("ans:procedimento", ANS_NS)
            cod_tab = tx(pr.find("ans:codigoTabela", ANS_NS)) if pr is not None else ""
            cod_proc = tx(pr.find("ans:codigoProcedimento", ANS_NS)) if pr is not None else ""
            desc = tx(pr.find("ans:descricaoProcedimento", ANS_NS)) if pr is not None else ""
            qtd = dec(tx(pe.find("ans:quantidadeExecutada", ANS_NS)))
            vuni = dec(tx(pe.find("ans:valorUnitario", ANS_NS)))
            vtot = dec(tx(pe.find("ans:valorTotal", ANS_NS)))
            if not vtot and vuni and qtd:
                vtot = vuni * qtd

            itens.append(dict(
                arquivo=fname,
                numero_lote=lote,
                tipo_guia="SADT",
                numeroGuiaPrestador=guia_id,
                numeroGuiaOperadora=guia_id,
                paciente=paciente,
                medico=medico,
                data_atendimento=data_atd,
                codigo_tabela=cod_tab,
                codigo_procedimento=cod_proc,
                descricao_procedimento=desc,
                quantidade=float(qtd) if qtd else 1,
                valor_unitario=float(vuni),
                valor_total=float(vtot),
            ))

        # outras despesas
        for d in g.findall(".//ans:despesa", ANS_NS):
            svc = d.find("ans:servicosExecutados", ANS_NS)
            cod_tab = tx(svc.find("ans:codigoTabela", ANS_NS)) if svc is not None else ""
            cod_proc = tx(svc.find("ans:codigoProcedimento", ANS_NS)) if svc is not None else ""
            desc = tx(svc.find("ans:descricaoProcedimento", ANS_NS)) if svc is not None else ""
            qtd = dec(tx(svc.find("ans:quantidadeExecutada", ANS_NS))) if svc is not None else DEC_ZERO
            vuni = dec(tx(svc.find("ans:valorUnitario", ANS_NS))) if svc is not None else DEC_ZERO
            vtot = dec(tx(svc.find("ans:valorTotal", ANS_NS))) if svc is not None else DEC_ZERO
            if not vtot and vuni and qtd:
                vtot = vuni * qtd

            itens.append(dict(
                arquivo=fname,
                numero_lote=lote,
                tipo_guia="SADT",
                numeroGuiaPrestador=guia_id,
                numeroGuiaOperadora=guia_id,
                paciente=paciente,
                medico=medico,
                data_atendimento=data_atd,
                codigo_tabela=cod_tab,
                codigo_procedimento=cod_proc,
                descricao_procedimento=desc,
                quantidade=float(qtd) if qtd else 1,
                valor_unitario=float(vuni),
                valor_total=float(vtot),
            ))

    return itens


# =========================================================
# PARTE 4 — Conciliação
# =========================================================

def build_xml_df(files) -> pd.DataFrame:
    rows = []
    for f in files:
        rows.extend(extract_xml_items(f))
    df = pd.DataFrame(rows)

    if df.empty:
        return df

    # normalizar códigos rigidamente
    df["codigo_procedimento_norm"] = df["codigo_procedimento"].astype(str).map(normalize_code)

    # chave nova
    df["chave"] = (
        df["numero_lote"].astype(str) + "__" +
        df["numeroGuiaPrestador"].astype(str).str.strip() + "__" +
        df["codigo_procedimento_norm"]
    )
    return df


def conciliar(df_xml: pd.DataFrame, df_demo: pd.DataFrame,
              fallback: bool = False,
              tol: float = 0.02) -> Tuple[pd.DataFrame, pd.DataFrame]:

    if df_xml.empty or df_demo.empty:
        return pd.DataFrame(), df_xml

    # merge principal
    m = df_xml.merge(df_demo, on="chave", how="left", suffixes=("", "_demo"))
    m["matched"] = m["valor_apresentado"].notna()

    conciliados = m[m["matched"]].copy()
    nao = m[~m["matched"]].copy()  # apenas itens do XML (demonstrativo extra é ignorado)

    # fallback (descrição + valor)
    if fallback:
        df_fallback = nao.copy()
        df_fallback["k_desc"] = df_fallback["descricao_procedimento"].str.strip().str.lower()
        demo2 = df_demo.copy()
        demo2["k_desc"] = demo2["descricao_procedimento"].str.strip().str.lower()

        fb = df_fallback.merge(demo2,
                               on=["numero_lote", "numeroGuiaPrestador", "k_desc"],
                               suffixes=("", "_fb"),
                               how="left")

        keep = fb["valor_apresentado"].notna() & (
            (fb["valor_total"] - fb["valor_apresentado"]).abs() <= tol
        )
        fb_ok = fb[keep]

        if not fb_ok.empty:
            conciliados = pd.concat([conciliados, fb_ok], ignore_index=True)
            nao = fb[~keep]  # só o que ainda não casou

    return conciliados, nao


# =========================================================
# PARTE 5 — Analytics
# =========================================================

def kpis_conc(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    base = df.copy()
    if "competencia" not in base.columns:
        base["competencia"] = ""
    grp = base.groupby("competencia", as_index=False).agg(
        valor_apresentado=("valor_apresentado", "sum"),
        valor_glosa=("valor_glosa", "sum"),
        valor_pago=("valor_pago", "sum")
    )
    grp["glosa_pct"] = grp["valor_glosa"] / grp["valor_apresentado"].replace(0, 1)
    return grp

def ranking_glosa(df: pd.DataFrame):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    grp = df.groupby(["codigo_procedimento", "descricao_procedimento"], as_index=False).agg(
        valor_apresentado=("valor_apresentado", "sum"),
        valor_glosa=("valor_glosa", "sum"),
        valor_pago=("valor_pago", "sum"),
        itens=("arquivo", "count")
    )
    grp["glosa_pct"] = grp["valor_glosa"] / grp["valor_apresentado"].replace(0, 1)
    top_valor = grp.sort_values("valor_glosa", ascending=False).head(20)
    top_pct = grp[grp["valor_apresentado"] >= 500].sort_values("glosa_pct",
                                                               ascending=False).head(20)
    return top_valor, top_pct

def motivos(df):
    if df.empty:
        return df
    grp = df.groupby(["motivo_glosa_codigo", "motivo_glosa_descricao"], as_index=False).agg(
        valor_glosa=("valor_glosa", "sum"),
        valor_apresentado=("valor_apresentado", "sum"),
        itens=("arquivo", "count")
    )
    grp["glosa_pct"] = grp["valor_glosa"] / grp["valor_apresentado"].replace(0, 1)
    return grp.sort_values("valor_glosa", ascending=False)

# =========================================================
# PARTE 6 — Interface e Exportação
# =========================================================

st.sidebar.header("Parâmetros")
fallback_desc = st.sidebar.toggle("Fallback por descrição + valor", False)
tolerance = st.sidebar.number_input("Tolerância (R$)", 0.0, 1.0, 0.02, 0.01)

xml_files = st.file_uploader("XML TISS", type=["xml"], accept_multiple_files=True)
demo_files = st.file_uploader("Demonstrativos AMHP", type=["xlsx"], accept_multiple_files=True)

df_demo = pd.concat([parse_demo_file(f) for f in (demo_files or [])], ignore_index=True) \
           if demo_files else pd.DataFrame()

if st.button("Processar"):

    df_xml = build_xml_df(xml_files or [])
    if df_xml.empty:
        st.error("Nenhum item XML encontrado.")
        st.stop()

    conc, nao = conciliar(df_xml, df_demo, fallback=fallback_desc, tol=tolerance)

    st.subheader("Itens XML conciliados")
    st.dataframe(apply_money(conc, ["valor_total", "valor_apresentado",
                                    "valor_glosa", "valor_pago"]))

    st.subheader("Itens do XML NÃO conciliados")
    st.dataframe(apply_money(nao, ["valor_total"]))

    # Analytics
    st.markdown("## Analytics")
    st.markdown("### KPIs por Competência")
    st.dataframe(apply_money(kpis_conc(conc), ["valor_apresentado", "valor_glosa", "valor_pago"]))

    st.markdown("### Ranking de itens glosados")
    t_val, t_pct = ranking_glosa(conc)
    st.write("Top por valor de glosa")
    st.dataframe(apply_money(t_val, ["valor_apresentado", "valor_glosa", "valor_pago"]))
    st.write("Top por percentual de glosa")
    st.dataframe(apply_money(t_pct, ["valor_apresentado", "valor_glosa", "valor_pago"]))

    st.markdown("### Motivos de Glosa")
    st.dataframe(apply_money(motivos(conc), ["valor_glosa", "valor_apresentado"]))

    # Export
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        conc.to_excel(wr, index=False, sheet_name="Conciliação")
        nao.to_excel(wr, index=False, sheet_name="Nao_Conciliados")
    st.download_button("Baixar Excel", buf.getvalue(),
                       file_name="conciliacao_tiss.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


