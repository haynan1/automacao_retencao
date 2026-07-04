# -*- coding: utf-8 -*-
"""Perfis por secretaria — isolados em diretório temporário (não toca no repo)."""
import pytest

from services import mapeador, perfis


@pytest.fixture
def perfis_tmp(tmp_path, monkeypatch):
    monkeypatch.setattr(perfis, "PERFIS_CONFIG_DIR", tmp_path / "config" / "perfis")
    monkeypatch.setattr(perfis, "PERFIS_MODELOS_DIR", tmp_path / "modelos" / "perfis")
    monkeypatch.setattr(perfis, "REGISTRO", tmp_path / "config" / "perfis.json")
    (tmp_path / "config").mkdir(parents=True, exist_ok=True)
    return tmp_path


def test_registro_padrao_quando_vazio(perfis_tmp):
    assert perfis.slug_padrao() == "saude"
    assert perfis.detectar_perfil(["FMS FUNDO MUNICIPAL DE SAUDE SLMB"]) == "saude"
    assert perfis.detectar_perfil(["OUTRA COISA"]) is None


def test_registrar_e_editar_nome(perfis_tmp):
    slug = perfis.registrar_perfil("Secretaria de Educação", deteccao=["EDUCACAO"])
    assert slug == "secretaria-de-educacao"
    assert perfis.nome_perfil(slug) == "Secretaria de Educação"
    # editar o nome mantém o slug
    perfis.registrar_perfil("Educação (SME)", slug=slug)
    assert perfis.nome_perfil(slug) == "Educação (SME)"
    assert perfis.detectar_perfil(["... EDUCACAO ..."]) == slug


def test_definir_padrao(perfis_tmp):
    slug = perfis.registrar_perfil("Secretaria de Educacao", deteccao=["EDUCACAO"])
    perfis.definir_padrao(slug)
    assert perfis.slug_padrao() == slug

    perfis.definir_padrao("inexistente")
    assert perfis.slug_padrao() == slug


def test_remover_perfil_bloqueios_e_sucesso(perfis_tmp):
    assert perfis.remover_perfil("saude") is False

    educacao = perfis.registrar_perfil("Secretaria de Educacao")
    cultura = perfis.registrar_perfil("Secretaria de Cultura")

    assert perfis.remover_perfil("saude") is False
    assert perfis.remover_perfil("inexistente") is False
    assert perfis.remover_perfil(educacao) is True
    assert perfis.perfil_valido(educacao) is False
    assert perfis.perfil_valido(cultura) is True

    assert perfis.remover_perfil(cultura) is True
    assert perfis.remover_perfil("saude") is False


def test_slug_guard_rejeita_traversal(perfis_tmp):
    import pytest
    for ruim in ("../../etc", "a/b", "..", "MAIUSC", "", "com espaco"):
        with pytest.raises(ValueError):
            perfis.caminho_molde(ruim)
    # slug válido não levanta
    assert perfis.caminho_vinculos("educacao-2").name == "vinculo_rubrica_coluna.json"


def test_vinculos_isolados_por_perfil(perfis_tmp):
    mapeador.salvar_vinculos("saude", {"evento x": "CEF"})
    mapeador.salvar_vinculos("educacao", {"evento x": mapeador.IGNORAR})
    assert mapeador.carregar_vinculos("saude") == {"evento x": "CEF"}
    assert mapeador.carregar_vinculos("educacao") == {"evento x": "__IGNORAR__"}


def test_molde_por_perfil_e_backup(perfis_tmp, tmp_path):
    from openpyxl import Workbook
    src = tmp_path / "m.xlsx"
    Workbook().save(src)
    assert perfis.existe_molde("saude") is False
    perfis.definir_molde("saude", src, "RETENCAO.xlsx")
    assert perfis.existe_molde("saude") is True
    assert perfis.info_molde("saude")["nome_original"] == "RETENCAO.xlsx"
    # redefinir gera um backup .bak.xlsx
    perfis.definir_molde("saude", src, "RETENCAO2.xlsx")
    backups = list(perfis.caminho_molde("saude").parent.glob("*.bak.xlsx"))
    assert len(backups) == 1


def test_remover_molde_limpa_atual_e_metadados(perfis_tmp, tmp_path):
    from openpyxl import Workbook
    src = tmp_path / "m.xlsx"
    Workbook().save(src)

    perfis.definir_molde("saude", src, "RETENCAO.xlsx")
    meta = perfis.caminho_molde("saude").with_suffix(".json")
    assert perfis.existe_molde("saude") is True
    assert meta.exists() is True

    assert perfis.remover_molde("saude") is True
    assert perfis.existe_molde("saude") is False
    assert meta.exists() is False

    assert perfis.remover_molde("saude") is False
    assert perfis.remover_molde("inexistente") is False
