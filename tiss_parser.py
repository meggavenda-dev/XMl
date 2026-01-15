
# file: tiss_parser.py
from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import IO, Union, List, Dict
import xml.etree.ElementTree as ET

# Namespace TISS
ANS_NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}

__version__ = "2026.01.15-ptbr-02"

class TissParsingError(Exception):
    pass

# ----------------------------
# Helpers
# ----------------------------
def _dec(txt: str | None) -> Decimal:
    """
    Converte string numérica do XML em Decimal; vazio/None => Decimal(0).
    Faz replace de ',' por '.' por segurança.
    """
    if not txt:
        return Decimal('0')
    return Decimal(txt.strip().replace(',', '.'))

def _is_consulta(root: ET.Element) -> bool:
    return root.find('.//ans:guiaConsulta', ANS_NS) is not None

def _get_numero_lote(root: ET.Element) -> str:
    el = root.find('.//ans:prestadorParaOperadora/ans:loteGuias/ans:numeroLote', ANS_NS)
    if el is None or not (txt := el.text):
        raise TissParsingError('numeroLote não encontrado no XML.')
    return txt.strip()

# ----------------------------
# CONSULTA
# ----------------------------
def _sum_consulta(root: ET.Element) -> tuple[int, Decimal, str]:
    """
    Para CONSULTA: soma ans:procedimento/ans:valorProcedimento por ans:guiaConsulta.
    Retorna (qtde_guias, total, estrategia_str).
    """
    total = Decimal('0')
    guias = root.findall('.//ans:prestadorParaOperadora/ans:loteGuias/ans:guiasTISS/ans:guiaConsulta', ANS_NS)
    for g in guias:
        val_el = g.find('.//ans:procedimento/ans:valorProcedimento', ANS_NS)
        total += _dec(val_el.text if val_el is not None else None)
    estrategia = "consulta_valorProcedimento"
    return len(guias), total, estrategia

# ----------------------------
# SADT - Estratégias em cascata
# ----------------------------
def _sum_sadt_by_total(guia: ET.Element) -> tuple[Decimal, str]:
    """
    (1) Tenta usar valorTotal/valorTotalGeral por guia SP-SADT.
    """
    vt = guia.find('.//ans:valorTotal', ANS_NS)
    if vt is None:
        return Decimal('0'), ''
    vtg = vt.find('ans:valorTotalGeral', ANS_NS)
    val = _dec(vtg.text if vtg is not None else None)
    if val > 0:
        return val, 'valorTotalGeral'
    return Decimal('0'), ''

def _sum_sadt_components(guia: ET.Element) -> tuple[Decimal, str]:
    """
    (2) Soma componentes do bloco valorTotal:
        valorProcedimentos, valorDiarias, valorTaxasAlugueis,
        valorMateriais, valorMedicamentos, valorGasesMedicinais.
    Útil quando valorTotalGeral vem vazio, mas os componentes foram preenchidos.
    """
    total = Decimal('0')
    vt = guia.find('.//ans:valorTotal', ANS_NS)
    if vt is not None:
        for tag in (
            'valorProcedimentos',
            'valorDiarias',
            'valorTaxasAlugueis',
            'valorMateriais',
            'valorMedicamentos',
            'valorGasesMedicinais',
        ):
            el = vt.find(f'ans:{tag}', ANS_NS)
            total += _dec(el.text if el is not None else None)
    if total > 0:
        return total, 'componentes_valorTotal'
    return Decimal('0'), ''

def _sum_sadt_outras_despesas_itens(guia: ET.Element) -> tuple[Decimal, str]:
    """
    (3) Soma item a item em outrasDespesas/despesa/servicosExecutados/valorTotal.
    """
    total = Decimal('0')
    for desp in guia.findall('.//ans:outrasDespesas/ans:despesa', ANS_NS):
        sv = desp.find('ans:servicosExecutados', ANS_NS)
        if sv is None:
            continue
        el_val = sv.find('ans:valorTotal', ANS_NS)
        total += _dec(el_val.text if el_val is not None else None)
    if total > 0:
        return total, 'outrasDespesas_itens'
    return Decimal('0'), ''

def _sum_sadt_procedimentos_itens(guia: ET.Element) -> tuple[Decimal, str]:
    """
    (4) Soma item a item em procedimentosExecutados/procedimentoExecutado/valorTotal.
        Se valorTotal do item não existir, usa valorUnitario * quantidadeExecutada.
    """
    total = Decimal('0')
    for it in guia.findall('.//ans:procedimentosExecutados/ans:procedimentoExecutado', ANS_NS):
        vtot = it.find('ans:valorTotal', ANS_NS)
        if vtot is not None and vtot.text and vtot.text.strip():
            total += _dec(vtot.text)
        else:
            vuni = it.find('ans:valorUnitario', ANS_NS)
            qtd  = it.find('ans:quantidadeExecutada', ANS_NS)
            if (vuni is not None and vuni.text) and (qtd is not None and qtd.text):
                total += _dec(vuni.text) * _dec(qtd.text)
    if total > 0:
        return total, 'procedimentos_itens'
    return Decimal('0'), ''

def _sum_sadt(root: ET.Element) -> tuple[int, Decimal, str]:
    """
    Calcula total por guia SP-SADT seguindo a ordem de tentativa:
      (1) valorTotalGeral
      (2) componentes do bloco valorTotal
      (3) outrasDespesas (itens)
      (4) procedimentosExecutados (itens)
    Retorna (qtde_guias, total, estrategia_total_arquivo).

    A estratégia reportada no arquivo será:
      - a estratégia única, se todas as guias usarem a mesma;
      - "misto: <estratégia1=x>, <estratégia2=y>, ..." quando houver combinação.
    """
    total = Decimal('0')
    estrategias: Dict[str, int] = {}
    guias = root.findall('.//ans:prestadorParaOperadora/ans:loteGuias/ans:guiasTISS/ans:guiaSP-SADT', ANS_NS)

    for g in guias:
        val, strat = _sum_sadt_by_total(g)
        if val == 0:
            val, strat = _sum_sadt_components(g)
        if val == 0:
            val, strat = _sum_sadt_outras_despesas_itens(g)
        if val == 0:
            val, strat = _sum_sadt_procedimentos_itens(g)
        if val == 0:
            strat = 'zero'

        total += val
        estrategias[strat] = estrategias.get(strat, 0) + 1

    # Monta descrição de estratégia para o arquivo
    if len(estrategias) == 1:
        estrategia_arquivo = next(iter(estrategias.keys()))
    else:
        parts = [f"{k}={v}" for k, v in sorted(estrategias.items(), key=lambda x: (-x[1], x[0]))]
        estrategia_arquivo = "misto: " + ", ".join(parts)

    return len(guias), total, estrategia_arquivo

# ----------------------------
# Parser público
# ----------------------------
def _parse_root(root: ET.Element, arquivo_nome: str) -> Dict:
    numero_lote = _get_numero_lote(root)
    if _is_consulta(root):
        tipo = 'CONSULTA'
        n_guias, total, estrategia = _sum_consulta(root)
    else:
        tipo = 'SADT'
        n_guias, total, estrategia = _sum_sadt(root)

    return {
        'arquivo': arquivo_nome,
        'numero_lote': numero_lote,
        'tipo': tipo,
        'qtde_guias': n_guias,
        'valor_total': total,
        'estrategia_total': estrategia,
        'parser_version': __version__,
    }

def parse_tiss_xml(source: Union[str, Path, IO[bytes]]) -> Dict:
    """
    Lê um XML TISS a partir de caminho (str/Path) OU arquivo (IO[bytes]/BytesIO).
    Retorna um dicionário com: arquivo, numero_lote, tipo, qtde_guias,
    valor_total (Decimal), estrategia_total (str), parser_version (str).
    """
    # UploadedFile/BytesIO
    if hasattr(source, 'read'):
        root = ET.parse(source).getroot()
        arquivo_nome = getattr(source, 'name', 'upload.xml')
        return _parse_root(root, Path(arquivo_nome).name)

    # Caminho de arquivo
    path = Path(source)
    root = ET.parse(path).getroot()
    return _parse_root(root, path.name)

def parse_many_xmls(paths: List[Union[str, Path]]) -> List[Dict]:
    """
    Processa vários XMLs (caminhos) e retorna lista de dicionários do parse_tiss_xml().
    Em caso de erro por arquivo, inclui 'erro' no dict e mantém valor_total=0.
    """
    resultados: List[Dict] = []
    for p in paths:
        try:
            resultados.append(parse_tiss_xml(p))
        except Exception as e:
            resultados.append({
                'arquivo': Path(p).name if hasattr(p, 'name') else str(p),
                'numero_lote': '',
                'tipo': 'DESCONHECIDO',
                'qtde_guias': 0,
                'valor_total': Decimal('0'),
                'estrategia_total': 'erro',
                'parser_version': __version__,
                'erro': str(e),
            })
    return resultados
