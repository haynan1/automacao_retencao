"""Mapeamento de lotacoes -> setores e descricoes -> rubricas.

Os mapeamentos vivem em arquivos JSON (config/), permitindo manutencao
sem tocar no codigo. Este modulo tambem sugere associacoes automaticas
e detecta pendencias antes do preenchimento.
"""

from __future__ import annotations

import json
from pathlib import Path

from .normalizador import normalizar_rubrica, normalizar_texto
from .utils import CONFIG_DIR

_ARQ_LOTACOES = CONFIG_DIR / "mapeamento_lotacoes.json"
_ARQ_RUBRICAS = CONFIG_DIR / "mapeamento_rubricas.json"


# ---------------------------------------------------------------------------
# Lotacoes
# ---------------------------------------------------------------------------

def carregar_mapeamento_lotacoes() -> dict:
    """Le o JSON de lotacoes -> setor. {} se o arquivo nao existir."""
    if not _ARQ_LOTACOES.exists():
        return {}
    with _ARQ_LOTACOES.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def salvar_mapeamento_lotacoes(mapa: dict) -> None:
    """Persiste o JSON de lotacoes -> setor (ordenado por chave)."""
    _ARQ_LOTACOES.parent.mkdir(parents=True, exist_ok=True)
    limpo = {k: v for k, v in mapa.items() if k and v}
    with _ARQ_LOTACOES.open("w", encoding="utf-8") as fh:
        json.dump(dict(sorted(limpo.items())), fh, ensure_ascii=False, indent=2)


def mapear_lotacao(lotacao_original: str, mapa: dict | None = None) -> str | None:
    """Retorna o setor destino de uma lotacao, ou None se nao mapeada."""
    mapa = mapa if mapa is not None else carregar_mapeamento_lotacoes()
    if lotacao_original in mapa:
        return mapa[lotacao_original]

    # Tolerancia a diferencas de acento/caixa/espaco na chave.
    alvo = normalizar_texto(lotacao_original)
    for chave, setor in mapa.items():
        if normalizar_texto(chave) == alvo:
            return setor
    return None


def sugerir_setor(lotacao_original: str, setores_planilha: list[str]) -> str | None:
    """Sugestao heuristica de setor quando a lotacao nao esta mapeada.

    Procura o nome de um setor da planilha dentro do texto da lotacao.
    """
    alvo = normalizar_texto(lotacao_original)
    melhor = None
    melhor_tam = 0
    for setor in setores_planilha:
        setor_norm = normalizar_texto(setor)
        if setor_norm and setor_norm in alvo and len(setor_norm) > melhor_tam:
            melhor = setor
            melhor_tam = len(setor_norm)
    return melhor


# ---------------------------------------------------------------------------
# Rubricas
# ---------------------------------------------------------------------------

def carregar_regras_rubricas() -> list[dict]:
    """Le a lista ordenada de regras de rubrica."""
    if not _ARQ_RUBRICAS.exists():
        return []
    with _ARQ_RUBRICAS.open("r", encoding="utf-8") as fh:
        dados = json.load(fh)
    return dados.get("regras", [])


def mapear_rubrica(descricao_original: str, regras: list[dict] | None = None) -> str | None:
    """Resolve a rubrica destino a partir da descricao do lancamento."""
    regras = regras if regras is not None else carregar_regras_rubricas()
    return normalizar_rubrica(descricao_original, regras)


# ---------------------------------------------------------------------------
# Aplicacao do mapeamento aos lancamentos + deteccao de pendencias
# ---------------------------------------------------------------------------

def aplicar_mapeamentos(
    lancamentos: list[dict],
    mapa_lotacoes: dict,
    regras_rubricas: list[dict],
) -> list[dict]:
    """Preenche setor_destino e rubrica_destino em cada lancamento,
    modificando e devolvendo a propria lista."""
    for reg in lancamentos:
        reg["setor_destino"] = mapear_lotacao(reg["lotacao_original"], mapa_lotacoes)
        reg["rubrica_destino"] = mapear_rubrica(reg["descricao_original"], regras_rubricas)
    return lancamentos


def detectar_pendencias(lancamentos: list[dict]) -> dict:
    """Classifica o que nao pode ser preenchido automaticamente.

    Retorna listas distintas de pendencias (sem duplicar), prontas para
    exibir na tela e para a aba de conferencia.
    """
    lotacoes_nao_mapeadas: dict[str, int] = {}
    rubricas_nao_mapeadas: dict[str, int] = {}
    folhas_desconhecidas: dict[str, int] = {}

    for reg in lancamentos:
        if not reg.get("setor_destino"):
            chave = reg["lotacao_original"] or "(sem lotacao)"
            lotacoes_nao_mapeadas[chave] = lotacoes_nao_mapeadas.get(chave, 0) + 1
        if not reg.get("rubrica_destino"):
            chave = reg["descricao_original"] or "(sem descricao)"
            rubricas_nao_mapeadas[chave] = rubricas_nao_mapeadas.get(chave, 0) + 1
        if not reg.get("folha_reconhecida"):
            chave = reg.get("folha", "(sem folha)")
            folhas_desconhecidas[chave] = folhas_desconhecidas.get(chave, 0) + 1

    return {
        "lotacoes_nao_mapeadas": lotacoes_nao_mapeadas,
        "rubricas_nao_mapeadas": rubricas_nao_mapeadas,
        "folhas_desconhecidas": folhas_desconhecidas,
    }
