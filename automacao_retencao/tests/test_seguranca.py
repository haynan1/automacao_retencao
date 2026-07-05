# -*- coding: utf-8 -*-
"""Barreira de origem (anti drive-by): POST de outra origem é bloqueado."""
import app as app_mod

CLIENT = app_mod.app.test_client()


def test_post_de_origem_externa_bloqueado():
    r = CLIENT.post("/analisar", headers={"Origin": "http://evil.example"}, data={})
    assert r.status_code == 403


def test_post_sem_origem_passa_o_hook():
    # Sem Origin (curl / form same-origin) — não é 403 (será 400 por falta de listagem).
    r = CLIENT.post("/analisar", data={})
    assert r.status_code != 403


def test_post_mesma_origem_passa_o_hook():
    r = CLIENT.post("/analisar", headers={"Origin": "http://localhost"}, data={})
    assert r.status_code != 403


def test_get_nao_e_afetado():
    assert CLIENT.get("/secretarias").status_code == 200


def test_baixar_molde_saude():
    r = CLIENT.get("/secretarias/saude/molde/baixar")
    assert r.status_code == 200
    assert "spreadsheet" in r.headers.get("Content-Type", "")


def test_baixar_molde_slug_inexistente():
    assert CLIENT.get("/secretarias/naoexiste/molde/baixar").status_code == 404
