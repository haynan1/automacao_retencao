"""Leitura e interpretacao do relatorio de origem (Listagem de Eventos).

O relatorio e um XLSX paginado: cabecalhos repetidos, rodapes, blocos
por lotacao e linhas de lancamento. Este modulo transforma esse texto
bruto em uma lista de lancamentos estruturados, sem depender cegamente
de letras de coluna fixas.
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook

from .normalizador import normalizar_folha, normalizar_texto

# Lotacao no formato "5.0037.0000 - FMS, ...".
_RE_LOTACAO = re.compile(r"^\s*\d+\.\d+\.\d+\s+-\s+")
# Competencia MM/AAAA.
_RE_COMPETENCIA = re.compile(r"\b(0[1-9]|1[0-2])/(\d{4})\b")

# Rotulos possiveis de cada coluna (normalizados) para localizar o cabecalho.
_ROTULOS_COLUNAS = {
    "matricula": ("MAT.", "MAT", "MATRICULA"),
    "funcionario": ("FUNCIONARIO", "NOME", "SERVIDOR"),
    "cpf": ("CPF",),
    "folha": ("FOLHA",),
    "descricao": ("DESCRICAO", "EVENTO", "DESCR"),
    "base_calculo": ("BASE DE CALC.", "BASE DE CALCULO", "BASE CALC", "BASE"),
    "referencia": ("REFER.", "REFERENCIA", "REFER"),
    "valor": ("VALOR",),
}

# Linhas de ruido que devem ser ignoradas na leitura de lancamentos.
_MARCADORES_IGNORAR = ("FP038", "EMITIDO EM", "PAGINA", "PÁGINA")


def _texto(celula) -> str:
    """Valor da celula como string limpa (nunca None)."""
    if celula is None:
        return ""
    return str(celula).strip()


def limpar_valor(valor) -> Decimal:
    """Converte um valor de celula (str brasileira ou numero) em Decimal.

    Aceita "28.391,15", "28391.15", 28391.15, "" -> Decimal('0.00').
    Valores negativos e parenteses contabeis sao respeitados.
    """
    if valor is None or valor == "":
        return Decimal("0.00")

    if isinstance(valor, (int, float)):
        try:
            return Decimal(str(valor)).quantize(Decimal("0.01"))
        except InvalidOperation:
            return Decimal("0.00")

    texto = str(valor).strip()
    if not texto:
        return Decimal("0.00")

    negativo = False
    if texto.startswith("(") and texto.endswith(")"):
        negativo = True
        texto = texto[1:-1]
    if texto.startswith("-"):
        negativo = True
        texto = texto[1:]

    # Remove qualquer coisa que nao seja digito, ponto ou virgula.
    texto = re.sub(r"[^\d.,]", "", texto)
    if not texto:
        return Decimal("0.00")

    # Formato brasileiro: ponto como milhar, virgula como decimal.
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    # Caso so tenha pontos, o ultimo ponto e o decimal (ja compativel).

    try:
        numero = Decimal(texto).quantize(Decimal("0.01"))
    except InvalidOperation:
        return Decimal("0.00")

    return -numero if negativo else numero


def detectar_competencia(ws) -> str | None:
    """Procura a competencia MM/AAAA nas primeiras linhas do relatorio."""
    for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
        for celula in row:
            if celula is None:
                continue
            m = _RE_COMPETENCIA.search(str(celula))
            if m:
                return f"{m.group(1)}/{m.group(2)}"
    return None


def _eh_linha_lotacao(texto_a: str) -> bool:
    return bool(_RE_LOTACAO.match(texto_a))


def detectar_lotacoes(ws) -> list[str]:
    """Lista, na ordem de aparicao, todas as lotacoes do relatorio."""
    lotacoes: list[str] = []
    for row in ws.iter_rows(values_only=True):
        texto_a = _texto(row[0]) if row else ""
        if _eh_linha_lotacao(texto_a) and texto_a not in lotacoes:
            lotacoes.append(texto_a)
    return lotacoes


def detectar_cabecalhos(ws) -> dict | None:
    """Localiza a primeira linha de cabecalho de lancamentos e devolve o
    mapa {campo: indice_coluna} (0-based). Retorna None se nao achar."""
    for row in ws.iter_rows(values_only=True):
        mapa = _mapear_cabecalho_da_linha(row)
        if mapa:
            return mapa
    return None


def _mapear_cabecalho_da_linha(row) -> dict | None:
    """Se a linha for um cabecalho de lancamento, retorna {campo: idx}."""
    if not row:
        return None

    encontrados: dict[str, int] = {}
    for idx, celula in enumerate(row):
        rotulo = normalizar_texto(_texto(celula))
        if not rotulo:
            continue
        for campo, opcoes in _ROTULOS_COLUNAS.items():
            if campo in encontrados:
                continue
            if rotulo in opcoes or any(rotulo == normalizar_texto(o) for o in opcoes):
                encontrados[campo] = idx

    # Consideramos cabecalho valido se localizamos descricao + valor,
    # que sao as colunas indispensaveis para o lancamento.
    if "descricao" in encontrados and "valor" in encontrados:
        return encontrados
    return None


def linha_eh_lancamento(row, cabecalho: dict) -> bool:
    """Heuristica: a linha e um lancamento se tiver descricao textual e
    um valor numerico nao nulo, e nao for ruido/cabecalho/lotacao."""
    if not row or not cabecalho:
        return False

    texto_a = _texto(row[0]) if len(row) > 0 else ""
    if _eh_linha_lotacao(texto_a):
        return False

    # Ruido conhecido em qualquer celula.
    linha_norm = normalizar_texto(" ".join(_texto(c) for c in row))
    if not linha_norm:
        return False
    if any(marca in linha_norm for marca in _MARCADORES_IGNORAR):
        return False
    # Cabecalho repetido.
    if _mapear_cabecalho_da_linha(row):
        return False

    descricao = _celula(row, cabecalho.get("descricao"))
    if not _texto(descricao):
        return False

    valor = limpar_valor(_celula(row, cabecalho.get("valor")))
    return valor != Decimal("0.00")


def _celula(row, idx):
    """Acesso seguro a uma celula por indice (pode estar fora do range)."""
    if idx is None or idx < 0 or idx >= len(row):
        return None
    return row[idx]


def extrair_lancamentos(caminho_xlsx: str | Path, regras_rubrica: list[dict] | None = None) -> dict:
    """Le o relatorio e devolve a estrutura completa de trabalho.

    Retorna:
        {
          "competencia": "06/2026" | None,
          "aba": "Page1",
          "lotacoes": [...],            # ordem de aparicao
          "lancamentos": [ {registro}, ... ],
          "cabecalho": {campo: idx},
        }

    A rubrica destino so e resolvida se 'regras_rubrica' for informada;
    caso contrario fica None (a camada de mapeamento resolve depois).
    """
    caminho = Path(caminho_xlsx)
    # data_only=True: o relatorio exportado ja traz valores, nao formulas.
    wb = load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]  # primeira aba, independente do nome ("Page1").

    competencia = detectar_competencia(ws)
    cabecalho = detectar_cabecalhos(ws)

    lancamentos: list[dict] = []
    lotacao_atual: str | None = None

    if cabecalho is None:
        wb.close()
        return {
            "competencia": competencia,
            "aba": ws.title,
            "lotacoes": [],
            "lancamentos": [],
            "cabecalho": None,
        }

    for numero_linha, row in enumerate(ws.iter_rows(values_only=True), start=1):
        texto_a = _texto(row[0]) if row else ""

        if _eh_linha_lotacao(texto_a):
            lotacao_atual = texto_a
            continue

        if not linha_eh_lancamento(row, cabecalho):
            continue

        folha_bruta = _texto(_celula(row, cabecalho.get("folha")))
        info_folha = normalizar_folha(folha_bruta)
        descricao = _texto(_celula(row, cabecalho.get("descricao")))

        registro = {
            "lotacao_original": lotacao_atual or "",
            "setor_destino": None,
            "matricula": _texto(_celula(row, cabecalho.get("matricula"))),
            "funcionario": _texto(_celula(row, cabecalho.get("funcionario"))),
            "cpf": _texto(_celula(row, cabecalho.get("cpf"))),
            "folha": folha_bruta or "MENSAL",
            "tipo_destino": info_folha["tipo_destino"],
            "folha_reconhecida": info_folha["reconhecida"],
            "observacao": info_folha["observacao"],
            "descricao_original": descricao,
            "rubrica_destino": None,
            "base_calculo": limpar_valor(_celula(row, cabecalho.get("base_calculo"))),
            "referencia": limpar_valor(_celula(row, cabecalho.get("referencia"))),
            "valor": limpar_valor(_celula(row, cabecalho.get("valor"))),
            "linha_origem": numero_linha,
        }
        lancamentos.append(registro)

    wb.close()

    lotacoes = []
    for reg in lancamentos:
        if reg["lotacao_original"] and reg["lotacao_original"] not in lotacoes:
            lotacoes.append(reg["lotacao_original"])

    return {
        "competencia": competencia,
        "aba": ws.title,
        "lotacoes": lotacoes,
        "lancamentos": lancamentos,
        "cabecalho": cabecalho,
    }
