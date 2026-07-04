"""Perfis por secretaria.

Cada perfil (secretaria) e uma 'gaveta' isolada com o proprio molde fixo e
os proprios vinculos aprendidos (lotacao->setor e evento/rubrica->coluna).
O MOTOR e 100% reaproveitado entre secretarias — o que muda sao apenas os
DADOS de cada perfil, evitando colisao de vinculos entre secretarias.

Layout em disco:
    config/perfis.json                                   (registro de perfis)
    config/perfis/<slug>/mapeamento_lotacoes.json
    config/perfis/<slug>/vinculo_rubrica_coluna.json
    modelos/perfis/<slug>/molde_padrao.xlsx
    modelos/perfis/<slug>/molde_padrao.json
    config/mapeamento_rubricas.json                      (regras genericas, compartilhadas)
"""

from __future__ import annotations

import json
import logging
import re
import shutil
from datetime import datetime
from pathlib import Path

from .normalizador import normalizar_texto
from .utils import CONFIG_DIR, MODELOS_DIR, criar_pastas

log = logging.getLogger("automacao_retencao")

PERFIS_CONFIG_DIR = CONFIG_DIR / "perfis"
PERFIS_MODELOS_DIR = MODELOS_DIR / "perfis"
REGISTRO = CONFIG_DIR / "perfis.json"

_REGISTRO_PADRAO = {
    "padrao": "saude",
    "perfis": [
        {"slug": "saude", "nome": "Secretaria da Saúde (FMS)",
         "deteccao": ["FUNDO MUNICIPAL DE SAUDE", "FMS"]},
    ],
}

_RE_SLUG = re.compile(r"[^a-z0-9]+")
# Forma canonica de um slug seguro (usado como nome de pasta).
_RE_SLUG_OK = re.compile(r"^[a-z0-9][a-z0-9-]*$")


def _slug_seguro(slug: str) -> str:
    """Garante que o slug e um nome de pasta seguro (sem path traversal).

    Defesa em profundidade: os callers atuais ja sanitizam, mas nenhum
    caminho de perfil e montado com um slug de formato invalido.
    """
    if not isinstance(slug, str) or not _RE_SLUG_OK.fullmatch(slug):
        raise ValueError(f"slug de perfil inválido: {slug!r}")
    return slug


# ---------------------------------------------------------------------------
# Registro
# ---------------------------------------------------------------------------

def _ler_registro() -> dict:
    if REGISTRO.exists():
        try:
            with REGISTRO.open("r", encoding="utf-8") as fh:
                dados = json.load(fh)
            dados.setdefault("perfis", [])
            dados.setdefault("padrao", dados["perfis"][0]["slug"] if dados["perfis"] else "saude")
            return dados
        except (OSError, json.JSONDecodeError, KeyError, IndexError) as exc:
            log.warning("perfis.json inválido (%s) — usando registro padrão.", exc)
    return {k: (v[:] if isinstance(v, list) else v) for k, v in _REGISTRO_PADRAO.items()}


def _salvar_registro(reg: dict) -> None:
    criar_pastas()
    with REGISTRO.open("w", encoding="utf-8") as fh:
        json.dump(reg, fh, ensure_ascii=False, indent=2)


def listar_perfis() -> list[dict]:
    return _ler_registro().get("perfis", [])


def slug_padrao() -> str:
    reg = _ler_registro()
    return reg.get("padrao") or (reg["perfis"][0]["slug"] if reg.get("perfis") else "saude")


def perfil_valido(slug: str) -> bool:
    return any(p["slug"] == slug for p in listar_perfis())


def info_perfil(slug: str) -> dict | None:
    for p in listar_perfis():
        if p["slug"] == slug:
            return p
    return None


def nome_perfil(slug: str) -> str:
    p = info_perfil(slug)
    return p["nome"] if p else slug


def gerar_slug(nome: str) -> str:
    base = _RE_SLUG.sub("-", normalizar_texto(nome).lower()).strip("-")
    return base or "perfil"


def registrar_perfil(nome: str, deteccao: list[str] | None = None, slug: str | None = None) -> str:
    """Cria (ou atualiza) um perfil no registro e devolve o slug."""
    reg = _ler_registro()
    slug = slug or gerar_slug(nome)
    existente = next((p for p in reg["perfis"] if p["slug"] == slug), None)
    if existente:
        existente["nome"] = nome
        if deteccao is not None:
            existente["deteccao"] = deteccao
    else:
        reg["perfis"].append({"slug": slug, "nome": nome, "deteccao": deteccao or []})
    _salvar_registro(reg)
    return slug


def definir_padrao(slug: str) -> None:
    """Define o perfil padrao. Ignora slugs inexistentes."""
    if not perfil_valido(slug):
        return
    reg = _ler_registro()
    reg["padrao"] = slug
    _salvar_registro(reg)


def remover_perfil(slug: str) -> bool:
    """Remove o perfil do registro sem apagar dados em disco."""
    reg = _ler_registro()
    if slug == reg.get("padrao") or len(reg.get("perfis", [])) <= 1:
        return False

    antes = len(reg["perfis"])
    reg["perfis"] = [p for p in reg["perfis"] if p["slug"] != slug]
    if len(reg["perfis"]) == antes:
        return False

    _salvar_registro(reg)
    log.info("Perfil '%s' removido do registro (dados preservados em disco).", slug)
    return True


def detectar_perfil(banners: list[str]) -> str | None:
    """Sugere um perfil a partir dos banners/cabecalho da Listagem."""
    if not banners:
        return None
    texto = normalizar_texto(" ".join(banners))
    for p in listar_perfis():
        for termo in p.get("deteccao", []):
            if termo and normalizar_texto(termo) in texto:
                return p["slug"]
    return None


# ---------------------------------------------------------------------------
# Caminhos por perfil
# ---------------------------------------------------------------------------

def _dir_config(slug: str) -> Path:
    d = PERFIS_CONFIG_DIR / _slug_seguro(slug)
    d.mkdir(parents=True, exist_ok=True)
    return d


def _dir_modelo(slug: str) -> Path:
    d = PERFIS_MODELOS_DIR / _slug_seguro(slug)
    d.mkdir(parents=True, exist_ok=True)
    return d


def caminho_mapa_lotacoes(slug: str) -> Path:
    return _dir_config(slug) / "mapeamento_lotacoes.json"


def caminho_vinculos(slug: str) -> Path:
    return _dir_config(slug) / "vinculo_rubrica_coluna.json"


def caminho_molde(slug: str) -> Path:
    return _dir_modelo(slug) / "molde_padrao.xlsx"


def _caminho_molde_meta(slug: str) -> Path:
    return _dir_modelo(slug) / "molde_padrao.json"


# ---------------------------------------------------------------------------
# Molde fixo por perfil
# ---------------------------------------------------------------------------

def existe_molde(slug: str) -> bool:
    return caminho_molde(slug).exists()


def info_molde(slug: str) -> dict | None:
    if not existe_molde(slug):
        return None
    meta = _caminho_molde_meta(slug)
    if meta.exists():
        try:
            with meta.open("r", encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, json.JSONDecodeError) as exc:
            log.warning("Metadados do molde do perfil '%s' inválidos: %s", slug, exc)
    return {"nome_original": "molde_padrao.xlsx", "definido_em": None}


def definir_molde(slug: str, origem: str | Path, nome_original: str) -> None:
    """Fixa um .xlsx (ja validado) como molde padrao do perfil.

    Se ja existir um molde fixo, guarda um backup timestampado antes de
    sobrescrever (evita perder o molde anterior por engano).
    """
    destino = caminho_molde(slug)
    if destino.exists():
        carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = destino.with_name(f"molde_padrao.{carimbo}.bak.xlsx")
        try:
            shutil.copyfile(destino, backup)
        except OSError as exc:
            log.warning("Não foi possível fazer backup do molde do perfil '%s': %s", slug, exc)
    shutil.copyfile(origem, destino)
    meta = {
        "nome_original": nome_original or "molde_padrao.xlsx",
        "definido_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    with _caminho_molde_meta(slug).open("w", encoding="utf-8") as fh:
        json.dump(meta, fh, ensure_ascii=False, indent=2)


def remover_molde(slug: str) -> bool:
    """Remove o molde fixo atual e seus metadados. Mantem backups no disco."""
    if not perfil_valido(slug):
        return False

    removido = False
    for caminho in (caminho_molde(slug), _caminho_molde_meta(slug)):
        if not caminho.exists():
            continue
        try:
            caminho.unlink()
            removido = True
        except OSError as exc:
            log.warning("Nao foi possivel remover '%s': %s", caminho, exc)
    if removido:
        log.info("Molde fixo do perfil '%s' removido.", slug)
    return removido
