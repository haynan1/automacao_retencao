# -*- coding: utf-8 -*-
"""Camada web do construtor de molde.

Nenhum teste aqui pode gravar no repositório: a rota de salvar é exercida
com os diretórios de perfil redirecionados para um tmp_path.
"""
import pytest
from openpyxl import load_workbook

import app as app_mod
from services import molde, perfis

CLIENT = app_mod.app.test_client()

SPEC = {
    "titulo": "PREFEITURA DE SÃO LUÍS DE MONTES BELOS",
    "subtitulo": "MODELO DE RETENÇÕES — BASE ZERADA",
    "abas": ["MOLDE"],
    "colunas": ["ARSEM", "CEF", "IR"],
    "setores": [{"nome": "ADMINISTRATIVO", "apelido": "ADM"}, {"nome": "ACS", "apelido": "ACS"}],
    "tipos": ["Mensal", "13º salário"],
    "opcoes": {
        "coluna_total_evento": True,
        "linha_total_bloco": True,
        "linha_total_geral": True,
        "rodape_apelido": True,
    },
}


@pytest.fixture
def perfis_tmp(tmp_path, monkeypatch):
    """Isola molde, estrutura e registro do perfil em disco temporário."""
    monkeypatch.setattr(perfis, "PERFIS_CONFIG_DIR", tmp_path / "config" / "perfis")
    monkeypatch.setattr(perfis, "PERFIS_MODELOS_DIR", tmp_path / "modelos" / "perfis")
    monkeypatch.setattr(perfis, "REGISTRO", tmp_path / "config" / "perfis.json")
    (tmp_path / "config").mkdir(parents=True, exist_ok=True)
    return tmp_path


# ---------------------------------------------------------------------------
# Tela
# ---------------------------------------------------------------------------

def test_editor_abre_para_perfil_valido():
    resposta = CLIENT.get("/molde/saude")
    assert resposta.status_code == 200
    assert "Construtor de molde" in resposta.get_data(as_text=True)


def test_editor_404_para_perfil_inexistente():
    assert CLIENT.get("/molde/naoexiste").status_code == 404
    assert CLIENT.get("/molde/nao-existe/importar").status_code == 404


def test_rotas_de_escrita_recusam_origem_externa():
    for rota in ("previa", "baixar", "salvar"):
        resposta = CLIENT.post(
            f"/molde/saude/{rota}",
            headers={"Origin": "http://evil.example"},
            json={"spec": SPEC},
        )
        assert resposta.status_code == 403, rota


# ---------------------------------------------------------------------------
# Prévia
# ---------------------------------------------------------------------------

def test_previa_de_spec_valida():
    dados = CLIENT.post("/molde/saude/previa", json={"spec": SPEC}).get_json()
    assert dados["ok"] is True
    assert dados["problemas"] == [] and dados["divergencias"] == []
    assert dados["resumo"] == {"abas": 1, "setores": 2, "colunas": 3, "tipos": 2,
                               "celulas_preenchiveis": 12}
    papeis = [linha["papel"] for linha in dados["grade"]["linhas"]]
    assert papeis[:4] == ["titulo", "subtitulo", "secao", "setor"]
    assert "total_geral" in papeis


def test_previa_de_spec_invalida_devolve_problemas_sem_estourar():
    ruim = dict(SPEC, colunas=["ARSEM", "arsem", "TOTAL X"])
    resposta = CLIENT.post("/molde/saude/previa", json={"spec": ruim})
    assert resposta.status_code == 200
    dados = resposta.get_json()
    assert dados["ok"] is False
    assert len(dados["problemas"]) == 2


def test_previa_sem_corpo_json_nao_quebra():
    dados = CLIENT.post("/molde/saude/previa", json={}).get_json()
    assert dados["ok"] is False
    assert dados["problemas"] == ["Estrutura do molde não recebida."]


def test_previa_recusa_corpo_grande_sem_desserializar():
    """Uma spec é texto: 512 KB já é folga. O teto de 30 MB é para planilhas."""
    gordo = {"spec": dict(SPEC, colunas=[f"RUBRICA {'X' * 60} {i}" for i in range(8000)])}
    resposta = CLIENT.post("/molde/saude/previa", json=gordo)
    dados = resposta.get_json()
    assert dados["ok"] is False
    assert "grande demais" in dados["problemas"][0]


def test_previa_ignora_corpo_que_nao_e_json():
    resposta = CLIENT.post("/molde/saude/previa", data="setores=1")
    assert resposta.status_code == 200
    assert resposta.get_json()["ok"] is False


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def test_baixar_devolve_xlsx_valido(tmp_path):
    resposta = CLIENT.post("/molde/saude/baixar", json={"spec": SPEC})
    assert resposta.status_code == 200
    assert "spreadsheet" in resposta.headers["Content-Type"]

    caminho = tmp_path / "baixado.xlsx"
    caminho.write_bytes(resposta.data)
    assert molde.extrair_spec(caminho) == molde.validar_spec(SPEC)


def test_baixar_recusa_spec_invalida():
    resposta = CLIENT.post("/molde/saude/baixar", json={"spec": dict(SPEC, setores=[])})
    assert resposta.status_code == 422
    assert "ao menos um setor" in " ".join(resposta.get_json()["problemas"])


# ---------------------------------------------------------------------------
# Gravação (isolada em tmp_path)
# ---------------------------------------------------------------------------

def test_salvar_fixa_o_molde_e_guarda_a_estrutura(perfis_tmp):
    assert perfis.existe_molde("saude") is False

    dados = CLIENT.post("/molde/saude/salvar", json={"spec": SPEC}).get_json()
    assert dados["ok"] is True
    assert dados["redirect"].endswith("/secretarias")

    assert perfis.existe_molde("saude") is True
    assert perfis.info_molde("saude")["nome_original"] == "molde_construido_saude.xlsx"
    # A estrutura volta idêntica ao desenho — inclusive os apelidos.
    assert molde.carregar_estrutura("saude") == molde.validar_spec(SPEC)

    wb = load_workbook(perfis.caminho_molde("saude"))
    from services import preenchimento
    assert [b["setor"] for b in preenchimento.localizar_blocos_setores(wb["MOLDE"])] == \
        ["ADMINISTRATIVO", "ACS"]
    wb.close()


def test_salvar_recusa_spec_invalida_e_nao_grava(perfis_tmp):
    resposta = CLIENT.post("/molde/saude/salvar", json={"spec": dict(SPEC, colunas=[])})
    assert resposta.status_code == 422
    assert perfis.existe_molde("saude") is False
    assert molde.carregar_estrutura("saude") is None


def test_salvar_de_novo_guarda_backup_do_anterior(perfis_tmp):
    CLIENT.post("/molde/saude/salvar", json={"spec": SPEC})
    CLIENT.post("/molde/saude/salvar", json={"spec": dict(SPEC, colunas=["ARSEM", "CEF"])})

    backups = list(perfis.caminho_molde("saude").parent.glob("*.bak.xlsx"))
    assert len(backups) == 1
    assert molde.carregar_estrutura("saude")["colunas"] == ["ARSEM", "CEF"]


def test_importar_le_o_molde_fixo_atual(perfis_tmp):
    CLIENT.post("/molde/saude/salvar", json={"spec": SPEC})
    dados = CLIENT.get("/molde/saude/importar").get_json()
    assert dados["ok"] is True
    assert [s["nome"] for s in dados["spec"]["setores"]] == ["ADMINISTRATIVO", "ACS"]


def test_importar_sem_molde_fixo_avisa(perfis_tmp):
    resposta = CLIENT.get("/molde/saude/importar")
    assert resposta.status_code == 404
    assert resposta.get_json()["ok"] is False


def test_editor_reabre_o_desenho_salvo(perfis_tmp):
    CLIENT.post("/molde/saude/salvar", json={"spec": SPEC})
    corpo = CLIENT.get("/molde/saude").get_data(as_text=True)
    assert "Desenho salvo" in corpo
