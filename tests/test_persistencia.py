# -*- coding: utf-8 -*-
"""Dados sagrados: o aprendizado dos três eixos sobrevive a quedas.

Os vínculos são meses de decisões de quem opera — valem mais que o código,
porque refazê-los custa a conferência inteira. Dois riscos concretos:

  * gravar com `open("w")` trunca o arquivo ANTES de escrever: uma queda no
    meio deixa um JSON pela metade;
  * ler sem proteção transforma esse arquivo pela metade num erro 500 que
    bloqueia o processamento do mês até alguém apagar o arquivo na mão.

Este arquivo prende as duas pontas.
"""
import json

import pytest

from services import mapeador, molde, perfis


@pytest.fixture
def perfil(perfis_tmp):
    perfis.registrar_perfil("Secretaria de Teste", slug="saude")
    return "saude"


# ---------------------------------------------------------------------------
# Escrita atômica
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("salvar,carregar,caminho", [
    (mapeador.salvar_mapeamento_lotacoes, mapeador.carregar_mapeamento_lotacoes,
     perfis.caminho_mapa_lotacoes),
    (mapeador.salvar_vinculos, mapeador.carregar_vinculos, perfis.caminho_vinculos),
    (mapeador.salvar_vinculos_folhas, mapeador.carregar_vinculos_folhas,
     perfis.caminho_vinculos_folhas),
])
def test_vinculos_gravam_e_releem(perfil, salvar, carregar, caminho):
    salvar(perfil, {"CHAVE": "DESTINO", "": "descartado", "VAZIO": ""})
    assert carregar(perfil) == {"CHAVE": "DESTINO"}   # entradas vazias saem
    assert list(caminho(perfil).parent.glob("*.tmp")) == []  # sem resíduo


def test_falha_no_meio_da_escrita_nao_destroi_o_arquivo_anterior(perfil):
    """O ponto da escrita atômica: o arquivo antigo fica intacto.

    A falha é provocada de verdade — um valor que o json não serializa faz
    `json.dump` levantar DEPOIS de já ter escrito parte do conteúdo. Sem a
    escrita atômica, `open("w")` teria truncado o arquivo real nesse ponto e o
    aprendizado de meses viraria um arquivo pela metade.
    """
    mapeador.salvar_vinculos_folhas(perfil, {"FERIAS": "Férias"})
    caminho = perfis.caminho_vinculos_folhas(perfil)
    antes = caminho.read_text(encoding="utf-8")

    with pytest.raises(TypeError):
        mapeador.salvar_vinculos_folhas(perfil, {"FERIAS": object()})

    assert caminho.read_text(encoding="utf-8") == antes
    assert mapeador.carregar_vinculos_folhas(perfil) == {"FERIAS": "Férias"}


# ---------------------------------------------------------------------------
# Leitura tolerante
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("carregar,caminho", [
    (mapeador.carregar_mapeamento_lotacoes, perfis.caminho_mapa_lotacoes),
    (mapeador.carregar_vinculos, perfis.caminho_vinculos),
    (mapeador.carregar_vinculos_folhas, perfis.caminho_vinculos_folhas),
])
@pytest.mark.parametrize("lixo", ['{"a": ', "", "não é json", "[1, 2, 3]"])
def test_vinculo_corrompido_nao_derruba_o_processamento(perfil, carregar, caminho, lixo):
    """Um JSON quebrado degrada para 'sem aprendizado', nunca para erro 500."""
    caminho(perfil).write_text(lixo, encoding="utf-8")
    assert carregar(perfil) == {}


def test_config_de_rubricas_corrompida_degrada_para_vazio(monkeypatch, tmp_path):
    arq = tmp_path / "mapeamento_rubricas.json"
    arq.write_text("{quebrado", encoding="utf-8")
    monkeypatch.setattr(mapeador, "_ARQ_RUBRICAS", arq)
    assert mapeador.carregar_config_rubricas() == {"regras": [], "fora_de_escopo": []}


def test_estrutura_de_molde_gravada_atomicamente(perfil):
    spec = molde.validar_spec({
        "titulo": "T", "subtitulo": "S", "abas": ["MOLDE"], "colunas": ["ARSEM"],
        "setores": [{"nome": "ACS", "apelido": ""}], "tipos": ["Mensal", "Férias"],
        "opcoes": {"coluna_total_evento": True, "linha_total_bloco": True,
                   "linha_total_geral": True, "rodape_apelido": True},
    })
    molde.salvar_estrutura(perfil, spec)
    caminho = molde.caminho_estrutura(perfil)
    assert list(caminho.parent.glob("*.tmp")) == []
    assert molde.carregar_estrutura(perfil) == spec
    # As linhas de tipo desenhadas voltam inteiras — inclusive as novas.
    assert json.loads(caminho.read_text(encoding="utf-8"))["tipos"] == ["Mensal", "Férias"]


def test_registro_de_perfis_gravado_atomicamente(perfis_tmp):
    perfis.registrar_perfil("Educação", slug="educacao")
    assert list(perfis.REGISTRO.parent.glob("*.tmp")) == []
    assert any(p["slug"] == "educacao" for p in perfis.listar_perfis())
