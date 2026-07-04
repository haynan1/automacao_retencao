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
