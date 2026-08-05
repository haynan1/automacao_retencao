"""Mapeamento em tres eixos:
  A) Lotacao        -> Setor          (bloco da planilha)
  B) Evento/Rubrica -> Coluna         (regras + vinculos aprendidos)
  C) Folha          -> Linha de tipo  (vinculos aprendidos)

A planilha e uma GRADE. Cada lancamento precisa de tres coordenadas para
achar sua celula: o bloco (setor), a coluna (rubrica) e a linha (tipo de
folha). Enquanto o eixo C nao existia, a linha era escolhida por uma regra
fixa que mandava tudo que nao fosse 13o para 'Mensal' — ferias, rescisao e
complementar somavam na mesma linha e nao havia onde discordar.

O modelo e a FONTE DA VERDADE nos dois eixos da grade: nunca se escreve
numa coluna nem numa linha que nao exista. O que nao encontra destino vira
pendencia visivel — nunca preenchimento no lugar errado.

Granularidade do eixo B: o vinculo e do EVENTO, nao da rubrica. 'INSS' e
'INSS DO 13º SALÁRIO' batem na mesma regra ('contem: INSS') e, quando o
grupo era a rubrica, viravam uma linha so — impossivel mandar um para uma
coluna e outro para outra. Agora cada evento e um grupo, e a regra vira
apenas o PADRAO herdado: quem quiser separa; quem nao quiser nao faz nada.

Toda decisao carrega um `motivo` em texto. Poder discordar do sistema
comeca por conseguir ler o que ele fez.
"""

from __future__ import annotations

import re
from decimal import Decimal

from . import perfis
from .normalizador import (
    explicar_rubrica,
    familia_folha,
    normalizar_rubrica,
    normalizar_texto,
)
from .utils import CONFIG_DIR, gravar_json, ler_json

# Prefixo de parcela em emprestimos: "12/120 - EMPRESTIMO CEF 1" -> "EMPRESTIMO CEF 1".
_RE_PARCELA = re.compile(r"^\s*\d+\s*/\s*\d+\s*-\s*")

# Regras de rubrica: genericas, compartilhadas entre secretarias.
_ARQ_RUBRICAS = CONFIG_DIR / "mapeamento_rubricas.json"

IGNORAR = "__IGNORAR__"  # marcador de vinculo "fora de escopo / nao preencher"
_ZERO = Decimal("0.00")


# ===========================================================================
# Eixo A — Lotacoes -> Setor (por perfil)
# ===========================================================================

def _ler_vinculos(caminho, descricao: str) -> dict:
    """Le um arquivo de vinculos. Corrompido ou ausente -> dicionario vazio."""
    dados = ler_json(caminho, {}, descricao)
    return dados if isinstance(dados, dict) else {}


def _gravar_vinculos(caminho, vinculos: dict) -> None:
    """Grava vinculos ordenados, sem entradas vazias, de forma atomica."""
    limpo = {k: v for k, v in vinculos.items() if k and v}
    gravar_json(caminho, dict(sorted(limpo.items())))


def carregar_mapeamento_lotacoes(perfil: str) -> dict:
    return _ler_vinculos(perfis.caminho_mapa_lotacoes(perfil), "mapeamento_lotacoes.json")


def salvar_mapeamento_lotacoes(perfil: str, mapa: dict) -> None:
    _gravar_vinculos(perfis.caminho_mapa_lotacoes(perfil), mapa)


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
    vazio = {"regras": [], "fora_de_escopo": []}
    dados = ler_json(_ARQ_RUBRICAS, vazio, "mapeamento_rubricas.json")
    if not isinstance(dados, dict):
        return vazio
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
    """Vinculos aprendidos do perfil: evento/rubrica (norm) -> coluna | IGNORAR."""
    return _ler_vinculos(perfis.caminho_vinculos(perfil), "vinculo_rubrica_coluna.json")


def salvar_vinculos(perfil: str, vinculos: dict) -> None:
    _gravar_vinculos(perfis.caminho_vinculos(perfil), vinculos)


def _chave_evento(reg: dict) -> tuple[str, str]:
    """Identidade do EVENTO de um lancamento — nao da sua rubrica.

    O prefixo de parcela sai ("12/120 - EMPRESTIMO CEF 1" -> "EMPRESTIMO
    CEF 1") porque ele muda a cada servidor e a cada mes: mante-lo faria de
    cada funcionario um grupo. O resto do texto fica INTEIRO, e e essa a
    diferenca que permite separar 'INSS' de 'INSS DO 13º SALÁRIO'.

    Retorna (rotulo_exibicao, chave_normalizada).
    """
    rotulo = _RE_PARCELA.sub("", reg.get("descricao_original", "") or "").strip()
    return rotulo, normalizar_texto(rotulo)


def _destino_do_vinculo(destino: str, colunas_norm: dict, motivo: str):
    """Traduz um vinculo salvo em (coluna, status, motivo)."""
    if destino == IGNORAR:
        return None, "fora_escopo", f"{motivo}: fora de escopo"
    alvo = normalizar_texto(destino)
    if alvo in colunas_norm:
        return colunas_norm[alvo], "ok", motivo
    # O vinculo existe mas aponta para uma coluna que esta planilha nao tem
    # (molde trocado, coluna renomeada). Silenciar isso seria perder dinheiro
    # sem aviso: vira pendencia com o motivo explicito.
    return None, "sem_vinculo", f"{motivo}, mas “{destino}” não é coluna desta planilha"


def resolver_coluna(
    chave_evento: str,
    rubrica: str | None,
    termo_regra: str,
    colunas_norm: dict,
    fora_escopo_norm: set,
    vinculos: dict,
) -> tuple[str | None, str, str]:
    """Resolve o destino de um evento. Retorna (coluna | None, status, motivo).

    Precedencia — do mais especifico (voce) para o mais generico (a regra):

      1. vinculo salvo para ESTE evento        -> 'ok'
      2. vinculo salvo para a RUBRICA do evento -> 'ok'   (vale para o grupo)
      3. fora de escopo por configuracao        -> 'fora_escopo'
      4. coluna com o nome exato do evento      -> 'ok'
      5. regra de rubrica aponta uma coluna     -> 'regra' (preenche, inferido)
      6. nada                                   -> 'sem_vinculo'

    'ok' e 'regra' preenchem igual; a diferenca e de autoria. 'ok' foi
    decidido por alguem, 'regra' foi deduzido pelo sistema — e por isso a
    tela mostra os 'regra' em cima, para serem olhados antes de virarem
    rotina.
    """
    chave_rubrica = normalizar_texto(rubrica) if rubrica else ""

    if chave_evento in vinculos:
        return _destino_do_vinculo(
            vinculos[chave_evento], colunas_norm, "vínculo salvo deste evento"
        )

    if chave_rubrica and chave_rubrica in vinculos:
        return _destino_do_vinculo(
            vinculos[chave_rubrica], colunas_norm, f"vínculo salvo da rubrica “{rubrica}”"
        )

    if chave_evento in fora_escopo_norm:
        return None, "fora_escopo", "evento marcado como fora de escopo"

    if chave_rubrica and chave_rubrica in fora_escopo_norm:
        return None, "fora_escopo", f"rubrica “{rubrica}” é fora de escopo"

    if chave_evento in colunas_norm:
        return colunas_norm[chave_evento], "ok", "o evento tem uma coluna de mesmo nome"

    if chave_rubrica and chave_rubrica in colunas_norm:
        detalhe = f" (contém “{termo_regra}”)" if termo_regra else ""
        return colunas_norm[chave_rubrica], "regra", f"regra “{rubrica}”{detalhe}"

    return None, "sem_vinculo", "nenhuma coluna corresponde a este evento"


def aplicar_mapeamentos(
    lancamentos: list[dict],
    mapa_lotacoes: dict,
    regras_rubricas: list[dict],
) -> list[dict]:
    """Preenche setor_destino, rubrica_destino (canonica) e o termo que bateu."""
    for reg in lancamentos:
        reg["setor_destino"] = mapear_lotacao(reg["lotacao_original"], mapa_lotacoes)
        rubrica, termo = explicar_rubrica(reg["descricao_original"], regras_rubricas)
        reg["rubrica_destino"] = rubrica
        reg["rubrica_termo"] = termo
    return lancamentos


def resolver_colunas(
    lancamentos: list[dict],
    colunas_modelo: list[str],
    fora_de_escopo: list[str],
    vinculos: dict,
) -> list[dict]:
    """Preenche coluna_destino, rubrica_status e rubrica_motivo."""
    colunas_norm = {normalizar_texto(c): c for c in colunas_modelo}
    fora_norm = {normalizar_texto(x) for x in fora_de_escopo}
    for reg in lancamentos:
        rotulo, chave = _chave_evento(reg)
        coluna, status, motivo = resolver_coluna(
            chave, reg.get("rubrica_destino"), reg.get("rubrica_termo", ""),
            colunas_norm, fora_norm, vinculos,
        )
        # O rótulo do evento fica no lançamento para que quem soma não
        # precise reimplementar o agrupamento (nem importar este módulo).
        reg["evento"] = rotulo
        reg["coluna_destino"] = coluna
        reg["rubrica_status"] = status
        reg["rubrica_motivo"] = motivo
    return lancamentos


# Pendentes primeiro; depois o que o sistema deduziu sozinho (para ser
# conferido); so entao o que ja foi decidido e o que se ignora de proposito.
_ORDEM_RUBRICA = {"sem_vinculo": 0, "regra": 1, "ok": 2, "fora_escopo": 3}


def construir_grupos_rubricas(
    lancamentos: list[dict],
    colunas_modelo: list[str],
    fora_de_escopo: list[str],
    vinculos: dict,
) -> list[dict]:
    """Agrupa os EVENTOS para a tela de vinculo (um menu por evento).

    Cada grupo: {rotulo, chave, rubrica, n, total, coluna, status, motivo,
    folhas}. `folhas` lista as folhas em que aquele evento apareceu — e o
    que mostra, sem abrir o relatorio, que 'INSS' veio da mensal e 'INSS DO
    13º' veio do decimo terceiro.
    """
    colunas_norm = {normalizar_texto(c): c for c in colunas_modelo}
    fora_norm = {normalizar_texto(x) for x in fora_de_escopo}

    grupos: dict[str, dict] = {}
    for reg in lancamentos:
        rotulo, chave = _chave_evento(reg)
        grupo = grupos.get(chave)
        if grupo is None:
            coluna, status, motivo = resolver_coluna(
                chave, reg.get("rubrica_destino"), reg.get("rubrica_termo", ""),
                colunas_norm, fora_norm, vinculos,
            )
            grupo = grupos[chave] = {
                "rotulo": rotulo, "chave": chave, "rubrica": reg.get("rubrica_destino"),
                "n": 0, "total": _ZERO, "coluna": coluna, "status": status,
                "motivo": motivo, "folhas": [],
            }
        grupo["n"] += 1
        grupo["total"] += reg["valor"]
        folha = (reg.get("folha") or "").strip()
        if folha and folha not in grupo["folhas"]:
            grupo["folhas"].append(folha)

    return sorted(
        grupos.values(),
        key=lambda g: (_ORDEM_RUBRICA.get(g["status"], 9), -g["total"]),
    )


# ===========================================================================
# Eixo C — Folha -> Linha de tipo do modelo
# ===========================================================================

def carregar_vinculos_folhas(perfil: str) -> dict:
    """Vinculos aprendidos do perfil: folha (norm) -> tipo | IGNORAR."""
    return _ler_vinculos(perfis.caminho_vinculos_folhas(perfil), "vinculo_folha_tipo.json")


def salvar_vinculos_folhas(perfil: str, vinculos: dict) -> None:
    _gravar_vinculos(perfis.caminho_vinculos_folhas(perfil), vinculos)


def _chave_folha(reg: dict) -> tuple[str, str]:
    """Identidade da folha: o texto do relatorio, inteiro e normalizado."""
    rotulo = (reg.get("folha") or "").strip()
    return rotulo, normalizar_texto(rotulo)


def resolver_tipo(
    folha: str,
    chave_folha: str,
    familia: str | None,
    tipos_norm: dict,
    vinculos: dict,
) -> tuple[str | None, str, str]:
    """Resolve a linha de destino de uma folha. (tipo | None, status, motivo).

    Precedencia:

      1. vinculo salvo para esta folha            -> 'ok'
      2. linha com o nome exato da folha          -> 'ok'
      3. linha cujo nome esta contido na folha    -> 'ok'   (o mais longo vence)
      4. linha com o nome da FAMILIA da folha     -> 'ok'
      5. familia conhecida sem linha propria      -> 'sugerido' (cai em Mensal)
      6. nada                                     -> 'sem_vinculo'

    O passo 5 e a regra antiga — ferias e rescisao somando na mensal —
    rebaixada de lei a sugestao: continua preenchendo (nada regride), mas
    aparece marcada na tela e pode ser trocada num clique. Folha que o
    sistema nao reconhece NAO recebe destino automatico: preencher no
    escuro seria pior que segurar o valor e perguntar.
    """
    if chave_folha in vinculos:
        destino = vinculos[chave_folha]
        if destino == IGNORAR:
            return None, "fora_escopo", "vínculo salvo desta folha: não preencher"
        alvo = normalizar_texto(destino)
        if alvo in tipos_norm:
            return tipos_norm[alvo], "ok", "vínculo salvo desta folha"
        return None, "sem_vinculo", (
            f"vínculo salvo desta folha, mas “{destino}” não é linha desta planilha"
        )

    if chave_folha in tipos_norm:
        return tipos_norm[chave_folha], "ok", "a planilha tem uma linha com este nome"

    # Linha cujo rotulo aparece dentro do texto da folha ("FÉRIAS 06/2026"
    # -> linha "Férias"). O rotulo mais longo vence, para que "Mensal" nao
    # roube uma folha que casaria melhor com "Mensal complementar".
    melhor, melhor_tam = None, 0
    for norm, rotulo in tipos_norm.items():
        if norm and norm in chave_folha and len(norm) > melhor_tam:
            melhor, melhor_tam = rotulo, len(norm)
    if melhor:
        return melhor, "ok", f"“{folha}” contém o nome da linha “{melhor}”"

    if familia:
        chave_familia = normalizar_texto(familia)
        if chave_familia in tipos_norm:
            return tipos_norm[chave_familia], "ok", f"folha do tipo {familia}"
        mensal = tipos_norm.get(normalizar_texto("Mensal"))
        if mensal:
            return mensal, "sugerido", (
                f"{familia} não tem linha própria nesta planilha — somando em “{mensal}”"
            )

    return None, "sem_vinculo", "nenhuma linha corresponde a esta folha"


def resolver_tipos(
    lancamentos: list[dict],
    tipos_modelo: list[str],
    vinculos_folhas: dict,
) -> list[dict]:
    """Preenche tipo_destino, folha_status e folha_motivo em cada lancamento."""
    tipos_norm = {normalizar_texto(t): t for t in tipos_modelo}
    for reg in lancamentos:
        rotulo, chave = _chave_folha(reg)
        familia = reg.get("folha_familia") or familia_folha(rotulo)
        reg["folha_familia"] = familia
        tipo, status, motivo = resolver_tipo(rotulo, chave, familia, tipos_norm, vinculos_folhas)
        reg["tipo_destino"] = tipo
        reg["folha_status"] = status
        reg["folha_motivo"] = motivo
    return lancamentos


_ORDEM_FOLHA = {"sem_vinculo": 0, "sugerido": 1, "ok": 2, "fora_escopo": 3}


def construir_grupos_folhas(
    lancamentos: list[dict],
    tipos_modelo: list[str],
    vinculos_folhas: dict,
) -> list[dict]:
    """Agrupa as folhas para a tela de vinculo (um menu por folha).

    Cada grupo: {rotulo, chave, familia, n, total, tipo, status, motivo}.
    """
    tipos_norm = {normalizar_texto(t): t for t in tipos_modelo}

    grupos: dict[str, dict] = {}
    for reg in lancamentos:
        rotulo, chave = _chave_folha(reg)
        if not chave:
            continue
        grupo = grupos.get(chave)
        if grupo is None:
            familia = reg.get("folha_familia") or familia_folha(rotulo)
            tipo, status, motivo = resolver_tipo(
                rotulo, chave, familia, tipos_norm, vinculos_folhas
            )
            grupo = grupos[chave] = {
                "rotulo": rotulo, "chave": chave, "familia": familia,
                "n": 0, "total": _ZERO, "tipo": tipo, "status": status, "motivo": motivo,
            }
        grupo["n"] += 1
        grupo["total"] += reg["valor"]

    return sorted(
        grupos.values(),
        key=lambda g: (_ORDEM_FOLHA.get(g["status"], 9), -g["total"]),
    )


# ===========================================================================
# Pendencias consolidadas
# ===========================================================================

def detectar_pendencias(lancamentos: list[dict]) -> dict:
    """Consolida pendencias para telas e conferencia (contagem por item)."""
    lotacoes_nao_mapeadas: dict[str, int] = {}
    rubricas_sem_vinculo: dict[str, int] = {}
    rubricas_fora_escopo: dict[str, int] = {}
    rubricas_por_regra: dict[str, int] = {}
    folhas_sem_vinculo: dict[str, int] = {}
    folhas_sugeridas: dict[str, int] = {}
    folhas_fora_escopo: dict[str, int] = {}

    def _conta(alvo: dict, chave: str) -> None:
        alvo[chave] = alvo.get(chave, 0) + 1

    for reg in lancamentos:
        if not reg.get("setor_destino"):
            _conta(lotacoes_nao_mapeadas, reg["lotacao_original"] or "(sem lotação)")

        evento = _chave_evento(reg)[0] or "(sem descrição)"
        status_rubrica = reg.get("rubrica_status")
        if status_rubrica == "sem_vinculo":
            _conta(rubricas_sem_vinculo, evento)
        elif status_rubrica == "fora_escopo":
            _conta(rubricas_fora_escopo, evento)
        elif status_rubrica == "regra":
            _conta(rubricas_por_regra, f"{evento} → {reg.get('coluna_destino')}")

        folha = (reg.get("folha") or "").strip() or "(sem folha)"
        status_folha = reg.get("folha_status")
        if status_folha == "sem_vinculo":
            _conta(folhas_sem_vinculo, folha)
        elif status_folha == "sugerido":
            _conta(folhas_sugeridas, f"{folha} → {reg.get('tipo_destino')}")
        elif status_folha == "fora_escopo":
            _conta(folhas_fora_escopo, folha)

    return {
        "lotacoes_nao_mapeadas": lotacoes_nao_mapeadas,
        "rubricas_sem_vinculo": rubricas_sem_vinculo,
        "rubricas_fora_escopo": rubricas_fora_escopo,
        "rubricas_por_regra": rubricas_por_regra,
        "folhas_sem_vinculo": folhas_sem_vinculo,
        "folhas_sugeridas": folhas_sugeridas,
        "folhas_fora_escopo": folhas_fora_escopo,
    }
