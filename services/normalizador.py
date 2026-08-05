"""Normalizacao de texto e de dominios (folha, rubrica).

A comparacao de rubricas e feita por 'contains' sobre um texto
normalizado: sem acentos, em caixa alta e com espacos colapsados.

Nada aqui DECIDE destino. Este modulo classifica e explica; quem escolhe
onde o dinheiro cai e o vinculo do perfil (services/mapeador.py), com a
tela na frente. A diferenca importa: enquanto a classificacao de folha
decidia sozinha, toda folha que nao fosse 13o virava 'Mensal' em silencio —
ferias, rescisao e complementar iam para a mesma linha sem ninguem ver.
"""

from __future__ import annotations

import re
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


# Familias de folha, da mais especifica para a mais generica. Servem para
# SUGERIR uma linha quando o molde nao tem uma linha com o nome exato da
# folha — e para dar nome ao que o relatorio trouxe. "13" vem antes porque
# "13º COMPLEMENTAR" e decimo terceiro, nao complementar.
#
# `\b13O?\b` casa "13" e "13O" (de "13º" apos a normalizacao) sem casar
# "130" — um numero solto na descricao nao vira decimo terceiro.
FAMILIAS_FOLHA: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("13º salário", (r"\b13O?\b", r"DECIMO TERCEIRO")),
    ("Férias", (r"FERIAS",)),
    ("Rescisão", (r"RESCISAO",)),
    ("Complementar", (r"COMPLEMENTAR",)),
    ("Mensal", (r"MENSAL",)),
)


def familia_folha(folha: str) -> str | None:
    """Familia canonica da folha bruta do relatorio, ou None.

    'FÉRIAS 06/2026' -> 'Férias'. Nao e um destino: e o nome da familia,
    usado para sugerir uma linha e para agrupar o que a tela mostra. A
    folha nunca e reescrita — o texto original do relatorio segue inteiro
    no lancamento, porque e por ele que o vinculo e aprendido.
    """
    norm = normalizar_texto(folha)
    if not norm:
        return None
    for nome, padroes in FAMILIAS_FOLHA:
        for padrao in padroes:
            if re.search(padrao, norm):
                return nome
    return None


def normalizar_rubrica(descricao: str, regras: list[dict]) -> str | None:
    """Aplica as regras de rubrica (lista ordenada de {rubrica, contem}).

    Retorna o nome canonico da rubrica destino ou None se nenhuma regra
    bater. A avaliacao respeita a ordem das regras (especifica antes de
    generica).
    """
    return explicar_rubrica(descricao, regras)[0]


def explicar_rubrica(descricao: str, regras: list[dict]) -> tuple[str | None, str]:
    """Como `normalizar_rubrica`, mas devolve tambem o termo que bateu.

    Retorna (rubrica_canonica | None, termo_que_bateu). O termo e o que a
    tela mostra para explicar a decisao: sem ele, 'INSS' e 'INSS DO 13º'
    caem na mesma rubrica e ninguem consegue ver POR QUE. Ver o criterio e
    o primeiro passo para poder discordar dele.
    """
    alvo = normalizar_texto(descricao)
    if not alvo:
        return None, ""

    for regra in regras:
        for termo in regra.get("contem", []):
            if normalizar_texto(termo) in alvo:
                return regra["rubrica"], termo
    return None, ""
