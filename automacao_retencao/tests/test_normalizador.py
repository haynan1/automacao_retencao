# -*- coding: utf-8 -*-
from services.normalizador import normalizar_folha, normalizar_texto
from services import mapeador


def test_normalizar_texto():
    assert normalizar_texto("Vigilância  Sanitária") == "VIGILANCIA SANITARIA"
    assert normalizar_texto("  ITAÚ ") == "ITAU"


def test_normalizar_folha():
    assert normalizar_folha("MENSAL")["tipo_destino"] == "Mensal"
    assert normalizar_folha("FÉRIAS")["tipo_destino"] == "Mensal"
    r = normalizar_folha("RESCISÃO")
    assert r["tipo_destino"] == "Mensal" and r["observacao"]
    c = normalizar_folha("COMPLEMENTAR")
    assert c["tipo_destino"] == "Mensal" and c["observacao"]
    assert normalizar_folha("13º SALÁRIO")["tipo_destino"] == "13º salário"
    d = normalizar_folha("QUALQUER COISA")
    assert d["tipo_destino"] is None and d["reconhecida"] is False


def test_mapear_rubrica():
    regras = mapeador.carregar_regras_rubricas()
    assert mapeador.mapear_rubrica("IRRF", regras) == "IR"
    assert mapeador.mapear_rubrica("12/120 - EMPRESTIMO CEF 1", regras) == "CEF"
    assert mapeador.mapear_rubrica("IPASGO DEPENDENTES..", regras) == "IPASGO DEPENDENTES"
    assert mapeador.mapear_rubrica("IPASGO - BASICO", regras) == "IPASGO"


def test_fora_de_escopo_configurado():
    cfg = mapeador.carregar_config_rubricas()
    fora = {normalizar_texto(x) for x in cfg["fora_de_escopo"]}
    assert normalizar_texto("INSS") in fora
    assert normalizar_texto("PENSAO ALIMENTICIA") in fora
