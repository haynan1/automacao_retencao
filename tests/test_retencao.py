# -*- coding: utf-8 -*-
"""Política de retenção: o que expira e o que fica.

A assimetria é a regra de negócio, não um detalhe de faxina:

  * o que foi ENVIADO carrega nome, matrícula e CPF de cada servidor e sai
    em 24h — guardar depois de terminado o trabalho é risco sem contrapartida;
  * o que foi GERADO é o produto, vai anexado a processo e conferido meses
    depois. Some por idade e o app destrói aquilo que existe para produzir.
"""
import time

import pytest

from services import utils


@pytest.fixture
def pastas(tmp_path, monkeypatch):
    """Redireciona as três pastas de retenção para lugares descartáveis."""
    caminhos = {}
    for nome, attr in (("uploads", "UPLOADS_DIR"), ("outputs", "OUTPUTS_DIR"),
                       ("sessions", "SESSIONS_DIR")):
        pasta = tmp_path / nome
        pasta.mkdir(parents=True, exist_ok=True)
        monkeypatch.setattr(utils, attr, pasta)
        caminhos[nome] = pasta
    return caminhos


def _envelhecer(caminho, dias: float):
    antigo = time.time() - dias * 24 * 3600
    caminho.write_text("x", encoding="utf-8")
    import os
    os.utime(caminho, (antigo, antigo))
    return caminho


def test_planilha_gerada_nao_expira_nunca(pastas):
    """O padrão do app: saídas ficam. Um ano depois ainda estão lá."""
    antiga = _envelhecer(pastas["outputs"] / "RETENCAO_PREENCHIDA_20250101_120000.xlsx", 365)
    utils.limpar_temporarios()
    assert antiga.exists(), "a planilha gerada não pode ser apagada por idade"


def test_arquivo_enviado_com_pii_continua_expirando(pastas):
    """A contrapartida da permanência: o que tem PII sai em 24h."""
    enviado = _envelhecer(pastas["uploads"] / "ListagemEventos.xlsx", 2)
    sessao = _envelhecer(pastas["sessions"] / "abc.json", 2)
    recente = _envelhecer(pastas["uploads"] / "de_hoje.xlsx", 0.1)

    utils.limpar_temporarios()

    assert not enviado.exists()
    assert not sessao.exists()
    assert recente.exists()   # dentro da janela, o trabalho pode estar em curso


def test_varredura_de_saidas_ainda_e_possivel_sob_pedido(pastas):
    """`outputs_horas=None` é o padrão, não uma amarra: quem pedir, varre."""
    antiga = _envelhecer(pastas["outputs"] / "velha.xlsx", 10)
    nova = _envelhecer(pastas["outputs"] / "nova.xlsx", 1)

    utils.limpar_temporarios(outputs_horas=24 * 7)

    assert not antiga.exists()
    assert nova.exists()


def test_gitkeep_nunca_e_removido(pastas):
    """Apagar o .gitkeep quebraria a estrutura de pastas do repositório."""
    marca = _envelhecer(pastas["uploads"] / ".gitkeep", 400)
    utils.limpar_temporarios()
    assert marca.exists()


# ---------------------------------------------------------------------------
# Permanência tem uma contrapartida: nada pode sobrescrever o que ficou
# ---------------------------------------------------------------------------

def test_duas_operacoes_no_mesmo_segundo_nao_se_sobrescrevem(pastas):
    """O carimbo tem precisão de segundo — e agora o arquivo é para sempre.

    Duas secretarias processadas em sequência caíam no mesmo nome, e a
    segunda apagava a primeira em silêncio: o histórico ficava com duas
    operações apontando para um arquivo só.
    """
    primeira = utils.gerar_nome_saida()
    (pastas["outputs"] / primeira).write_text("planilha entregue", encoding="utf-8")

    segunda = utils.gerar_nome_saida()
    assert segunda != primeira

    (pastas["outputs"] / segunda).write_text("segunda planilha", encoding="utf-8")
    assert (pastas["outputs"] / primeira).read_text(encoding="utf-8") == "planilha entregue"


def test_nomes_seguem_distintos_sob_repeticao(pastas):
    nomes = set()
    for _ in range(25):
        nome = utils.gerar_nome_saida()
        (pastas["outputs"] / nome).write_text("x", encoding="utf-8")
        nomes.add(nome)
    assert len(nomes) == 25


def test_o_pdf_continua_pareado_com_a_planilha_que_ganhou_sufixo(pastas):
    """O nome do PDF sai do carimbo da planilha; o sufixo não pode quebrar isso."""
    from services import conferencia_pdf

    (pastas["outputs"] / utils.gerar_nome_saida()).write_text("x", encoding="utf-8")
    com_sufixo = utils.gerar_nome_saida()
    assert "_" in com_sufixo.replace("RETENCAO_PREENCHIDA_", "")  # ganhou sufixo

    nome_pdf = conferencia_pdf.nome_arquivo(com_sufixo)
    assert nome_pdf.startswith("CONFERENCIA_") and nome_pdf.endswith(".pdf")
    # O carimbo do PDF é o mesmo da planilha, sufixo incluído.
    assert nome_pdf[len("CONFERENCIA_"):-len(".pdf")] in com_sufixo
