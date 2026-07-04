# -*- coding: utf-8 -*-
"""Histórico local: gravação append-only, leitura mais-recente-primeiro, trim."""
from decimal import Decimal

import pytest

from services import historico


@pytest.fixture
def hist_tmp(tmp_path, monkeypatch):
    monkeypatch.setattr(historico, "_ARQ", tmp_path / "operacoes.jsonl")
    return tmp_path


def test_vazio(hist_tmp):
    assert historico.listar_operacoes() == []
    assert historico.total_operacoes() == 0


def test_registra_e_serializa_decimal(hist_tmp):
    historico.registrar_operacao({
        "perfil": "saude", "competencia": "06/2026",
        "total_lido": Decimal("685765.57"), "confere": True,
    })
    ops = historico.listar_operacoes()
    assert len(ops) == 1
    assert ops[0]["total_lido"] == "685765.57"   # Decimal -> string
    assert ops[0]["confere"] is True
    assert "datahora" in ops[0]                   # preenchido automaticamente


def test_ordem_mais_recente_primeiro(hist_tmp):
    for i in range(3):
        historico.registrar_operacao({"competencia": f"0{i+1}/2026"})
    ops = historico.listar_operacoes()
    assert [o["competencia"] for o in ops] == ["03/2026", "02/2026", "01/2026"]


def test_trim_limita_tamanho(hist_tmp, monkeypatch):
    monkeypatch.setattr(historico, "_MAX_ENTRADAS", 5)
    for i in range(12):
        historico.registrar_operacao({"n": i})
    assert historico.total_operacoes() == 5
    # os 5 mais recentes são n=7..11
    ns = [o["n"] for o in historico.listar_operacoes()]
    assert ns == [11, 10, 9, 8, 7]


def test_linha_corrompida_e_ignorada(hist_tmp):
    historico.registrar_operacao({"ok": 1})
    historico._ARQ.open("a", encoding="utf-8").write("{lixo não-json}\n")
    ops = historico.listar_operacoes()
    assert len(ops) == 1 and ops[0]["ok"] == 1
