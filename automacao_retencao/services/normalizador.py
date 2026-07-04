"""Normalizacao de texto e de dominios (folha, rubrica).

A comparacao de rubricas e feita por 'contains' sobre um texto
normalizado: sem acentos, em caixa alta e com espacos colapsados.
"""

from __future__ import annotations

import unicodedata


def remover_acentos(texto: str) -> str:
    """Remove diacriticos preservando as letras base (ç -> c, á -> a)."""
    if texto is None:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar_texto(texto: str) -> str:
    """Texto canonico para comparacao: sem acento, caixa alta, sem
    espacos duplicados nas bordas ou no meio."""
    if texto is None:
        return ""
    limpo = remover_acentos(texto).upper()
    return " ".join(limpo.split())


def normalizar_folha(folha: str) -> dict:
    """Classifica o tipo de folha conforme as regras de negocio.

    Retorna um dicionario:
        {
          "tipo_destino": "Mensal" | "13º salário" | None,
          "reconhecida": bool,
          "observacao": str,   # nota interna para conferencia
        }

    - Contem "13" ou "13º"      -> "13º salário"
    - "MENSAL"                  -> "Mensal"
    - "FERIAS"                  -> "Mensal"
    - "RESCISAO"                -> "Mensal" (marcado para conferencia)
    - Desconhecida              -> tipo_destino None e reconhecida False
    """
    bruto = (folha or "").strip()
    norm = normalizar_texto(bruto)

    if not norm:
        return {"tipo_destino": None, "reconhecida": False, "observacao": "Folha vazia"}

    # 13o salario tem prioridade (pode vir como "13", "13º", "DECIMO TERCEIRO").
    if "13" in norm or "DECIMO TERCEIRO" in norm:
        return {"tipo_destino": "13º salário", "reconhecida": True, "observacao": ""}

    if "MENSAL" in norm:
        return {"tipo_destino": "Mensal", "reconhecida": True, "observacao": ""}

    if "FERIAS" in norm:
        return {"tipo_destino": "Mensal", "reconhecida": True, "observacao": ""}

    if "RESCISAO" in norm:
        return {
            "tipo_destino": "Mensal",
            "reconhecida": True,
            "observacao": "Origem: RESCISÃO — conferir manualmente",
        }

    if "COMPLEMENTAR" in norm:
        return {
            "tipo_destino": "Mensal",
            "reconhecida": True,
            "observacao": "Origem: COMPLEMENTAR — conferir manualmente",
        }

    return {
        "tipo_destino": None,
        "reconhecida": False,
        "observacao": f"Tipo de folha desconhecido: {bruto}",
    }


def normalizar_rubrica(descricao: str, regras: list[dict]) -> str | None:
    """Aplica as regras de rubrica (lista ordenada de {rubrica, contem}).

    Retorna o nome canonico da rubrica destino ou None se nenhuma regra
    bater. A avaliacao respeita a ordem das regras (especifica antes de
    generica).
    """
    alvo = normalizar_texto(descricao)
    if not alvo:
        return None

    for regra in regras:
        for termo in regra.get("contem", []):
            if normalizar_texto(termo) in alvo:
                return regra["rubrica"]
    return None
