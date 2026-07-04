"""Mapeamento em dois eixos:
  A) Lotacao  -> Setor  (config/mapeamento_lotacoes.json)
  B) Rubrica  -> Coluna do modelo (regras + vinculos aprendidos)

As colunas do modelo sao a FONTE DA VERDADE: nunca se escreve numa coluna
que nao exista. Rubricas 'fora de escopo' (ex.: INSS, faltas, pensao) sao
ignoradas de proposito — aparecem como fora de escopo, nao como erro.
Vinculos definidos manualmente sao aprendidos e salvos para os proximos meses.
"""

from __future__ import annotations

import json
import re
from decimal import Decimal

from . import perfis
from .normalizador import normalizar_rubrica, normalizar_texto
from .utils import CONFIG_DIR

# Prefixo de parcela em emprestimos: "12/120 - EMPRESTIMO CEF 1" -> "EMPRESTIMO CEF 1".
_RE_PARCELA = re.compile(r"^\s*\d+\s*/\s*\d+\s*-\s*")

# Regras de rubrica: genericas, compartilhadas entre secretarias.
_ARQ_RUBRICAS = CONFIG_DIR / "mapeamento_rubricas.json"

IGNORAR = "__IGNORAR__"  # marcador de vinculo "fora de escopo / nao preencher"
_ZERO = Decimal("0.00")


# ===========================================================================
# Eixo A — Lotacoes -> Setor (por perfil)
# ===========================================================================

def carregar_mapeamento_lotacoes(perfil: str) -> dict:
    arq = perfis.caminho_mapa_lotacoes(perfil)
    if not arq.exists():
        return {}
    with arq.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def salvar_mapeamento_lotacoes(perfil: str, mapa: dict) -> None:
    arq = perfis.caminho_mapa_lotacoes(perfil)
    limpo = {k: v for k, v in mapa.items() if k and v}
    with arq.open("w", encoding="utf-8") as fh:
        json.dump(dict(sorted(limpo.items())), fh, ensure_ascii=False, indent=2)


def mapear_lotacao(lotacao_original: str, mapa: dict) -> str | None:
    if lotacao_original in mapa:
        return mapa[lotacao_original]
    alvo = normalizar_texto(lotacao_original)
    for chave, setor in mapa.items():
        if normalizar_texto(chave) == alvo:
            return setor
    return None


def sugerir_setor(lotacao_original: str, setores_planilha: list[str]) -> str | None:
    """Sugestao heuristica: nome de um setor contido no texto da lotacao."""
    alvo = normalizar_texto(lotacao_original)
    melhor, melhor_tam = None, 0
    for setor in setores_planilha:
        s = normalizar_texto(setor)
        if s and s in alvo and len(s) > melhor_tam:
            melhor, melhor_tam = setor, len(s)
    return melhor


# ===========================================================================
# Eixo B — Rubricas -> Coluna do modelo
# ===========================================================================

def carregar_config_rubricas() -> dict:
    if not _ARQ_RUBRICAS.exists():
        return {"regras": [], "fora_de_escopo": []}
    with _ARQ_RUBRICAS.open("r", encoding="utf-8") as fh:
        dados = json.load(fh)
    dados.setdefault("regras", [])
    dados.setdefault("fora_de_escopo", [])
    return dados


def carregar_regras_rubricas() -> list[dict]:
    return carregar_config_rubricas()["regras"]


def mapear_rubrica(descricao_original: str, regras: list[dict] | None = None) -> str | None:
    """Descricao crua -> rubrica canonica (via regras 'contains')."""
    regras = regras if regras is not None else carregar_regras_rubricas()
    return normalizar_rubrica(descricao_original, regras)


def carregar_vinculos(perfil: str) -> dict:
    """Vinculos aprendidos do perfil: rubrica/descricao (norm) -> coluna | IGNORAR."""
    arq = perfis.caminho_vinculos(perfil)
    if not arq.exists():
        return {}
    with arq.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def salvar_vinculos(perfil: str, vinculos: dict) -> None:
    arq = perfis.caminho_vinculos(perfil)
    limpo = {k: v for k, v in vinculos.items() if k and v}
    with arq.open("w", encoding="utf-8") as fh:
        json.dump(dict(sorted(limpo.items())), fh, ensure_ascii=False, indent=2)


def _chave_grupo(reg: dict) -> tuple[str, str]:
    """Chave estavel do grupo de rubrica de um lancamento.

    Se a rubrica canonica foi reconhecida, agrupa por ela; senao, agrupa pela
    descricao crua. Retorna (rotulo_exibicao, chave_normalizada).
    """
    if reg.get("rubrica_destino"):
        rotulo = reg["rubrica_destino"]
    else:
        # Evento cru sem regra: remove o prefixo de parcela para agrupar o evento.
        rotulo = _RE_PARCELA.sub("", reg.get("descricao_original", "")).strip()
    return rotulo, normalizar_texto(rotulo)


def resolver_coluna(
    rotulo: str,
    chave_norm: str,
    colunas_norm: dict,
    fora_escopo_norm: set,
    vinculos: dict,
) -> tuple[str | None, str]:
    """Resolve o destino de um grupo de rubrica.

    Retorna (coluna_modelo | None, status) com status em:
      'ok'          -> tem coluna no modelo (coluna preenchida)
      'fora_escopo' -> ignorar de proposito
      'sem_vinculo' -> precisa de decisao manual (dropdown)
    """
    # 1) vinculo aprendido tem prioridade
    if chave_norm in vinculos:
        destino = vinculos[chave_norm]
        if destino == IGNORAR:
            return None, "fora_escopo"
        if normalizar_texto(destino) in colunas_norm:
            return colunas_norm[normalizar_texto(destino)], "ok"
        return None, "sem_vinculo"

    # 2) fora de escopo por configuracao
    if chave_norm in fora_escopo_norm:
        return None, "fora_escopo"

    # 3) casamento direto com uma coluna do modelo
    if chave_norm in colunas_norm:
        return colunas_norm[chave_norm], "ok"

    # 4) precisa de decisao manual
    return None, "sem_vinculo"


def aplicar_mapeamentos(
    lancamentos: list[dict],
    mapa_lotacoes: dict,
    regras_rubricas: list[dict],
) -> list[dict]:
    """Preenche setor_destino e rubrica_destino (canonica) em cada lancamento."""
    for reg in lancamentos:
        reg["setor_destino"] = mapear_lotacao(reg["lotacao_original"], mapa_lotacoes)
        reg["rubrica_destino"] = mapear_rubrica(reg["descricao_original"], regras_rubricas)
    return lancamentos


def resolver_colunas(
    lancamentos: list[dict],
    colunas_modelo: list[str],
    fora_de_escopo: list[str],
    vinculos: dict,
) -> list[dict]:
    """Preenche coluna_destino e rubrica_status em cada lancamento."""
    colunas_norm = {normalizar_texto(c): c for c in colunas_modelo}
    fora_norm = {normalizar_texto(x) for x in fora_de_escopo}
    for reg in lancamentos:
        rotulo, chave = _chave_grupo(reg)
        coluna, status = resolver_coluna(rotulo, chave, colunas_norm, fora_norm, vinculos)
        reg["coluna_destino"] = coluna
        reg["rubrica_status"] = status
    return lancamentos


def construir_grupos_rubricas(
    lancamentos: list[dict],
    colunas_modelo: list[str],
    fora_de_escopo: list[str],
    vinculos: dict,
) -> list[dict]:
    """Agrupa as rubricas para a tela de vinculo (dropdown).

    Cada grupo: {rotulo, chave, n, total, coluna, status}. Ordena por status
    (pendentes primeiro) e depois por valor.
    """
    colunas_norm = {normalizar_texto(c): c for c in colunas_modelo}
    fora_norm = {normalizar_texto(x) for x in fora_de_escopo}

    grupos: dict[str, dict] = {}
    for reg in lancamentos:
        rotulo, chave = _chave_grupo(reg)
        if chave not in grupos:
            coluna, status = resolver_coluna(rotulo, chave, colunas_norm, fora_norm, vinculos)
            grupos[chave] = {
                "rotulo": rotulo, "chave": chave, "n": 0, "total": _ZERO,
                "coluna": coluna, "status": status,
            }
        grupos[chave]["n"] += 1
        grupos[chave]["total"] += reg["valor"]

    ordem = {"sem_vinculo": 0, "ok": 1, "fora_escopo": 2}
    return sorted(grupos.values(), key=lambda g: (ordem.get(g["status"], 9), -g["total"]))


def detectar_pendencias(lancamentos: list[dict]) -> dict:
    """Consolida pendencias para telas e conferencia (contagem por item)."""
    lotacoes_nao_mapeadas: dict[str, int] = {}
    rubricas_sem_vinculo: dict[str, int] = {}
    rubricas_fora_escopo: dict[str, int] = {}
    folhas_desconhecidas: dict[str, int] = {}

    for reg in lancamentos:
        if not reg.get("setor_destino"):
            k = reg["lotacao_original"] or "(sem lotação)"
            lotacoes_nao_mapeadas[k] = lotacoes_nao_mapeadas.get(k, 0) + 1
        status = reg.get("rubrica_status")
        rotulo = reg.get("rubrica_destino") or reg.get("descricao_original") or "(sem descrição)"
        if status == "sem_vinculo":
            rubricas_sem_vinculo[rotulo] = rubricas_sem_vinculo.get(rotulo, 0) + 1
        elif status == "fora_escopo":
            rubricas_fora_escopo[rotulo] = rubricas_fora_escopo.get(rotulo, 0) + 1
        if not reg.get("folha_reconhecida"):
            k = reg.get("folha", "(sem folha)")
            folhas_desconhecidas[k] = folhas_desconhecidas.get(k, 0) + 1

    return {
        "lotacoes_nao_mapeadas": lotacoes_nao_mapeadas,
        "rubricas_sem_vinculo": rubricas_sem_vinculo,
        "rubricas_fora_escopo": rubricas_fora_escopo,
        "folhas_desconhecidas": folhas_desconhecidas,
    }
