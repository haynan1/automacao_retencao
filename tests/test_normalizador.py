# -*- coding: utf-8 -*-
from services.normalizador import explicar_rubrica, familia_folha, normalizar_texto
from services import mapeador


def test_normalizar_texto():
    assert normalizar_texto("Vigilância  Sanitária") == "VIGILANCIA SANITARIA"
    assert normalizar_texto("  ITAÚ ") == "ITAU"


def test_familia_folha_classifica_sem_decidir_destino():
    """A família NOMEIA a folha; quem escolhe a linha é o vínculo do perfil.

    Antes, tudo que não fosse 13º virava 'Mensal' aqui dentro — férias e
    rescisão perdiam a identidade antes de qualquer tela poder mostrá-la.
    """
    assert familia_folha("MENSAL") == "Mensal"
    assert familia_folha("FÉRIAS") == "Férias"
    assert familia_folha("RESCISÃO") == "Rescisão"
    assert familia_folha("COMPLEMENTAR") == "Complementar"
    assert familia_folha("13º SALÁRIO") == "13º salário"
    assert familia_folha("QUALQUER COISA") is None
    assert familia_folha("") is None


def test_familia_13_nao_casa_numero_qualquer():
    """'13' precisa ser o número 13, não um pedaço de '130' ou '2013'."""
    assert familia_folha("MENSAL 130") == "Mensal"
    assert familia_folha("13") == "13º salário"
    assert familia_folha("DECIMO TERCEIRO") == "13º salário"
    # 13º tem prioridade sobre complementar: "13º COMPLEMENTAR" é 13º.
    assert familia_folha("13º COMPLEMENTAR") == "13º salário"


def test_mapear_rubrica():
    regras = mapeador.carregar_regras_rubricas()
    assert mapeador.mapear_rubrica("IRRF", regras) == "IR"
    assert mapeador.mapear_rubrica("12/120 - EMPRESTIMO CEF 1", regras) == "CEF"
    assert mapeador.mapear_rubrica("IPASGO DEPENDENTES..", regras) == "IPASGO DEPENDENTES"
    assert mapeador.mapear_rubrica("IPASGO - BASICO", regras) == "IPASGO"


def test_explicar_rubrica_devolve_o_termo_que_bateu():
    """Sem o termo, ninguém consegue ver POR QUE dois eventos viraram um."""
    regras = mapeador.carregar_regras_rubricas()
    assert explicar_rubrica("IRRF", regras) == ("IR", "IRRF")
    rubrica, termo = explicar_rubrica("INSS DO 13º SALÁRIO", regras)
    assert (rubrica, termo) == ("INSS", "INSS")
    assert explicar_rubrica("COISA NENHUMA", regras) == (None, "")


def test_fora_de_escopo_configurado():
    cfg = mapeador.carregar_config_rubricas()
    fora = {normalizar_texto(x) for x in cfg["fora_de_escopo"]}
    assert normalizar_texto("INSS") in fora
    assert normalizar_texto("PENSAO ALIMENTICIA") in fora
