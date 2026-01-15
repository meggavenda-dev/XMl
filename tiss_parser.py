
# file: tiss_parser.py
from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import IO, Union, List, Dict
import xml.etree.ElementTree as ET

ANS_NS = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}

class TissParsingError(Exception):
    pass

def _dec(txt: str | None) -> Decimal:
    """
    Converte string numérica do XML em Decimal; vazio/None => Decimal(0).
    Também troca ',' por '.' por segurança.
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

def _sum_consulta(root: ET.Element) -> tuple[int, Decimal]:
    """
    Para CONSULTA: soma ans:procedimento/ans:valorProcedimento por ans:guiaConsulta.
    """
    total = Decimal('0')
    guias = root.findall('.//ans:prestadorParaOperadora/ans:loteGuias/ans:guiasTISS/ans:guiaConsulta', ANS_NS)
    for g in guias:
        val_el = g.find('.//ans:procedimento/ans:valorProcedimento', ANS_NS)
        total += _dec(val_el.text if val_el is not None else None)
    return len(guias), total

def _sum_sadt_by_total(guia: ET.Element) -> Decimal:
    """
    Primeiro tenta usar valorTotal/valorTotalGeral por guia SP-SADT.
    """
    vt = guia.find('.//ans:valorTotal', ANS_NS)
    if vt is None:
        return Decimal('0')
    vtg = vt.find('ans:valorTotalGeral', ANS_NS)
    return _dec(vtg.text if vtg is not None else None)

def _sum_sadt_by_items(guia: ET.Element) -> Decimal:
    """
    Fallback: reconstrói total somando:
      - valorProcedimentos, valorDiarias, valorTaxasAlugueis, valorMateriais, valorMedicamentos, valorGasesMedicinais
      - e, se houver, outrasDespesas/despesa/servicosExecutados/valorTotal (item a item)
    """
    total = Decimal('0')

    vt = guia.find('.//ans:valorTotal', ANS_NS)
    if vt is not None:
        for tag in ('valorProcedimentos', 'valorDiarias', 'valorTaxasAlugueis',
                    'valorMateriais', 'valorMedicamentos', 'valorGasesMedicinais'):
            el = vt.find(f'ans:{tag}', ANS_NS)
            total += _dec(el.text if el is not None else None)

    for desp in guia.findall('.//ans:outrasDespesas/ans:despesa', ANS_NS):
        sv = desp.find('ans:servicosExecutados', ANS_NS)
        if sv is None:
            continue
        el_val = sv.find('ans:valorTotal', ANS_NS)
        total += _dec(el_val.text if el_val is not None else None)

    return total

def _sum_sadt(root: ET.Element) -> tuple[int, Decimal]:
    """
    Para SP-SADT: tenta valorTotalGeral; se não houver, soma itens/outrasDespesas.
    """
    total = Decimal('0')
    guias = root.findall('.//ans:prestadorParaOperadora/ans:loteGuias/ans:guiasTISS/ans:guiaSP-SADT', ANS_NS)
    for g in guias:
        v = _sum_sadt_by_total(g)
        if v == 0:
            v = _sum_sadt_by_items(g)
        total += v
    return len(guias), total

def _parse_root(root: ET.Element, arquivo_nome: str) -> Dict:
    numero_lote = _get_numero_lote(root)
    if _is_consulta(root):
        tipo = 'CONSULTA'
        n_guias, total = _sum_consulta(root)
    else:
        tipo = 'SADT'
        n_guias, total = _sum_sadt(root)

    return {
        'arquivo': arquivo_nome,
        'numero_lote': numero_lote,
        'tipo': tipo,
        'qtde_guias': n_guias,
        'valor_total': total
    }

def parse_tiss_xml(source: Union[str, Path, IO[bytes]]) -> Dict:
    """
    Lê um XML TISS a partir de caminho (str/Path) OU arquivo (IO[bytes]/BytesIO).
    Retorna:
    {
      'arquivo': '...xml',
      'numero_lote': '9148401',
      'tipo': 'CONSULTA'|'SADT',
      'qtde_guias': 19,
      'valor_total': Decimal('12198.38')
    }
    """
    # Se for file-like (tem .read), usamos ET.parse direto
    if hasattr(source, 'read'):
        root = ET.parse(source).getroot()
        arquivo_nome = getattr(source, 'name', 'upload.xml')
        return _parse_root(root, Path(arquivo_nome).name)

    # Caso contrário, tratamos como caminho
    path = Path(source)
    root = ET.parse(path).getroot()
    return _parse_root(root, path.name)

def parse_many_xmls(paths: List[Union[str, Path]]) -> List[Dict]:
    """
    Processa vários XMLs (caminhos) e retorna lista de dicionários (parse_tiss_xml()).
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
                'erro': str(e)
            })
    return resultados
