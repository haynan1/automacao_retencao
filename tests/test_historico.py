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


def test_remover_operacao_por_id(hist_tmp):
    for i in range(3):
        historico.registrar_operacao({"competencia": f"0{i+1}/2026"})
    ops = historico.listar_operacoes()
    alvo = ops[1]["id"]                      # remove o do meio
    assert historico.remover_operacao(alvo) is True
    restantes = historico.listar_operacoes()
    assert len(restantes) == 2
    assert alvo not in {o["id"] for o in restantes}
    # id inexistente não remove nada
    assert historico.remover_operacao("naoexiste") is False
    assert historico.total_operacoes() == 2


def test_limpar_historico(hist_tmp):
    for i in range(4):
        historico.registrar_operacao({"n": i})
    assert historico.limpar_historico() == 4
    assert historico.listar_operacoes() == []
    assert historico.limpar_historico() == 0   # já vazio


def test_linha_corrompida_e_ignorada(hist_tmp):
    historico.registrar_operacao({"ok": 1})
    historico._ARQ.open("a", encoding="utf-8").write("{lixo não-json}\n")
    ops = historico.listar_operacoes()
    assert len(ops) == 1 and ops[0]["ok"] == 1


# ---------------------------------------------------------------------------
# Retrato da operação — o que permite reemitir o PDF meses depois
# ---------------------------------------------------------------------------

_RETRATO = {
    "aba_destino": "SEC CULTURA",
    "reconciliacao": {"total_lido": Decimal("100.00"), "confere": True},
    "por_setor": {"ACS": Decimal("100.00")},
}


def test_retrato_vai_para_arquivo_proprio_e_nao_para_a_linha(hist_tmp):
    """A linha é lida INTEIRA a cada abertura da tela; o retrato, só ao clicar.

    Guardar a conferência dentro do JSONL engordaria em uma ordem de grandeza
    o arquivo que a listagem percorre — por um dado que a listagem não usa.
    """
    historico.registrar_operacao({"competencia": "07/2026", "retrato": _RETRATO})
    op = historico.listar_operacoes()[0]

    assert "retrato" not in op
    assert op["tem_retrato"] is True
    linha = historico._ARQ.read_text(encoding="utf-8")
    assert "por_setor" not in linha

    guardado = historico.carregar_retrato(op["id"])
    assert guardado["por_setor"] == {"ACS": "100.00"}   # Decimal -> string
    assert guardado["reconciliacao"]["confere"] is True


def test_operacao_sem_retrato_e_marcada_como_tal(hist_tmp):
    """Registro anterior a este recurso: a tela não pode oferecer o botão."""
    historico.registrar_operacao({"competencia": "07/2026"})
    op = historico.listar_operacoes()[0]
    assert op["tem_retrato"] is False
    assert historico.carregar_retrato(op["id"]) is None


def test_remover_operacao_leva_o_retrato_junto(hist_tmp):
    """Apagar do histórico tem de apagar o que ele guardava — não deixar no disco."""
    historico.registrar_operacao({"n": 1, "retrato": _RETRATO})
    op_id = historico.listar_operacoes()[0]["id"]
    caminho = historico._caminho_retrato(op_id)
    assert caminho.exists()

    assert historico.remover_operacao(op_id) is True
    assert not caminho.exists()
    assert historico.carregar_retrato(op_id) is None


def test_limpar_tudo_apaga_tambem_os_retratos(hist_tmp):
    for i in range(3):
        historico.registrar_operacao({"n": i, "retrato": _RETRATO})
    pasta = historico._pasta_retratos()
    assert len(list(pasta.iterdir())) == 3

    historico.limpar_historico()
    assert not pasta.exists() or not list(pasta.iterdir())


def test_trim_nao_deixa_retrato_orfao(hist_tmp, monkeypatch):
    """O JSONL é limitado; sem isto a pasta de retratos cresceria para sempre."""
    monkeypatch.setattr(historico, "_MAX_ENTRADAS", 3)
    for i in range(8):
        historico.registrar_operacao({"n": i, "retrato": _RETRATO})

    vivos = {o["id"] for o in historico.listar_operacoes()}
    em_disco = {p.stem for p in historico._pasta_retratos().iterdir()}
    assert em_disco == vivos and len(vivos) == 3


@pytest.mark.parametrize("op_id", [
    "../fora", "..\fora", "a/b", "C:evil", "", None, "x" * 65, ".", "..",
])
def test_id_que_nao_vira_nome_de_arquivo_e_recusado(hist_tmp, op_id):
    """O id entra num caminho: não pode escapar da pasta de retratos."""
    assert historico._caminho_retrato(op_id) is None
    assert historico.salvar_retrato(op_id, _RETRATO) is False
    assert historico.carregar_retrato(op_id) is None
    historico.remover_retrato(op_id)   # não pode levantar nem apagar nada de fora


def test_retrato_corrompido_no_disco_nao_derruba_a_leitura(hist_tmp):
    historico.registrar_operacao({"n": 1, "retrato": _RETRATO})
    op_id = historico.listar_operacoes()[0]["id"]
    historico._caminho_retrato(op_id).write_text("{não é json", encoding="utf-8")
    assert historico.carregar_retrato(op_id) is None


def test_retrato_guardado_nao_tem_dado_pessoal(hist_tmp):
    """O histórico não guarda PII; o que fica ao lado dele também não."""
    import re

    historico.registrar_operacao({"n": 1, "retrato": _RETRATO})
    op_id = historico.listar_operacoes()[0]["id"]
    conteudo = historico._caminho_retrato(op_id).read_text(encoding="utf-8")
    assert not re.search(r"\d{3}\.\d{3}\.\d{3}-\d{2}", conteudo)
    for campo in ("funcionario", "matricula", "cpf"):
        assert campo not in conteudo.lower()


def test_retrato_grande_demais_e_recusado_inteiro(hist_tmp, monkeypatch):
    """Recusar, nunca cortar.

    Um retrato pela metade renderizaria um TOTAL menor que o verdadeiro — um
    documento de conferência que mente é pior que documento nenhum. A tela
    apenas deixa de oferecer o PDF; a planilha continua com a aba completa.
    """
    monkeypatch.setattr(historico, "MAX_RETRATO_KB", 1)
    gigante = dict(_RETRATO, por_setor={f"SETOR {i:05d}": Decimal("1.00") for i in range(500)})

    historico.registrar_operacao({"n": 1, "retrato": gigante})
    op = historico.listar_operacoes()[0]

    assert op["tem_retrato"] is False
    assert historico.carregar_retrato(op["id"]) is None
    assert not historico._caminho_retrato(op["id"]).exists()


def test_retrato_nao_serializavel_nao_derruba_o_processamento(hist_tmp):
    """O histórico é best-effort: gravar não pode ser o que quebra o mês."""
    class Exotico:
        pass

    historico.registrar_operacao({"n": 1, "retrato": {"x": Exotico()}})
    ops = historico.listar_operacoes()
    assert len(ops) == 1 and ops[0]["tem_retrato"] is False


def test_linha_corrompida_nao_interrompe_a_faxina_de_retratos(hist_tmp, monkeypatch):
    """Uma linha ilegível não pode impedir o trim de limpar os retratos das outras."""
    monkeypatch.setattr(historico, "_MAX_ENTRADAS", 2)
    for i in range(3):
        historico.registrar_operacao({"n": i, "retrato": _RETRATO})
    with historico._ARQ.open("a", encoding="utf-8") as fh:
        fh.write("{linha corrompida}\n")

    historico.registrar_operacao({"n": 99, "retrato": _RETRATO})   # dispara o trim

    vivos = {o["id"] for o in historico.listar_operacoes()}
    em_disco = {p.stem for p in historico._pasta_retratos().iterdir()}
    assert em_disco == vivos


# ---------------------------------------------------------------------------
# Durabilidade — o histórico virou o registro permanente
# ---------------------------------------------------------------------------

def test_falha_ao_reescrever_nao_destroi_o_historico(hist_tmp, monkeypatch):
    """Reescrita atômica: ou o arquivo é o novo, ou continua o antigo.

    `write_text` trunca antes de escrever — uma queda no meio deixaria o
    registro permanente pela metade, e ele guarda meses de operações.
    """
    import os as os_mod

    for i in range(3):
        historico.registrar_operacao({"n": i})
    antes = historico._ARQ.read_text(encoding="utf-8")
    alvo = historico.listar_operacoes()[0]["id"]

    monkeypatch.setattr(os_mod, "replace", lambda *a, **k: (_ for _ in ()).throw(OSError("disco cheio")))
    assert historico.remover_operacao(alvo) is False

    assert historico._ARQ.read_text(encoding="utf-8") == antes
    assert historico.total_operacoes() == 3


def test_retrato_nao_fica_orfao_se_a_linha_nao_for_gravada(hist_tmp, monkeypatch):
    """O retrato é gravado antes da linha; sem a linha, ninguém o apagaria."""
    def falha_ao_abrir(*_a, **_k):
        raise OSError("disco cheio")

    monkeypatch.setattr(type(historico._ARQ), "open", falha_ao_abrir)
    historico.registrar_operacao({"n": 1, "retrato": _RETRATO})

    pasta = historico._pasta_retratos()
    assert not pasta.exists() or not list(pasta.iterdir())
