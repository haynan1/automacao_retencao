# -*- coding: utf-8 -*-
"""Parser merged-aware: leitura correta apesar de mesclagens, e anti-ruído."""
from decimal import Decimal

from services.parser_listagem import extrair_lancamentos, limpar_valor


def test_competencia_e_lotacoes(listagem):
    r = extrair_lancamentos(listagem)
    assert r["competencia"] == "06/2026"
    assert "5.0043.0000 - FMS, AGENTE COMUNITÁRIO DE SAÚDE - ACS" in r["lotacoes"]
    assert "5.9999.0000 - FMS, LOTAÇÃO NOVA SEM SETOR" in r["lotacoes"]


def test_descricao_deslocada_e_lida(listagem):
    """A Descrição (mesclada em R:W, deslocada do cabeçalho Q:V) deve ser lida."""
    r = extrair_lancamentos(listagem)
    descricoes = {l["descricao_original"] for l in r["lancamentos"]}
    assert "PREVIBELOS" in descricoes
    assert "IRRF" in descricoes
    assert "12/120 - EMPRESTIMO CEF 1" in descricoes


def test_banners_e_resumo_nao_viram_lancamento(listagem):
    r = extrair_lancamentos(listagem)
    descricoes = {l["descricao_original"] for l in r["lancamentos"]}
    assert "MUNICÍPIO DE SÃO LUÍS DE MONTES BELOS" not in descricoes
    assert "LISTAGEM DE EVENTOS" not in descricoes
    assert "06/2026" not in descricoes
    # O número solto do resumo (99999.99) não pode ter entrado.
    assert all(l["valor"] != Decimal("99999.99") for l in r["lancamentos"])


def test_banners_capturados_para_deteccao(listagem):
    r = extrair_lancamentos(listagem)
    juntos = " ".join(r["banners"]).upper()
    assert "FUNDO MUNICIPAL DE SAUDE" in juntos


def test_valor_e_folha(listagem):
    r = extrair_lancamentos(listagem)
    prev = next(l for l in r["lancamentos"] if l["descricao_original"] == "PREVIBELOS"
                and l["folha"] == "MENSAL")
    assert prev["valor"] == Decimal("698.39")
    ferias = next(l for l in r["lancamentos"] if l["folha"] == "FÉRIAS")
    assert ferias["tipo_destino"] == "Mensal"  # FÉRIAS -> Mensal


def test_limpar_valor_formatos():
    assert limpar_valor("1.234,56") == Decimal("1234.56")
    assert limpar_valor("88.0") == Decimal("88.00")
    assert limpar_valor("(42,10)") == Decimal("-42.10")
    assert limpar_valor("") == Decimal("0.00")
    assert limpar_valor(None) == Decimal("0.00")
