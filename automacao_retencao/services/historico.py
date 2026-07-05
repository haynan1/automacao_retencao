"""Histórico local das operações de preenchimento.

Append-only em JSONL (uma operação por linha) — robusto e sem reescrever o
arquivo a cada gravação. Guardado apenas em disco, em `historico/`, que está
no .gitignore: **nunca vai para o GitHub**.

Não guarda PII: só metadados e totais agregados (nada de nome/CPF/servidor).
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from decimal import Decimal

from .utils import HISTORICO_DIR, criar_pastas

log = logging.getLogger("automacao_retencao")

_ARQ = HISTORICO_DIR / "operacoes.jsonl"
_MAX_ENTRADAS = 2000  # limite defensivo de crescimento do arquivo


def _serializavel(obj):
    """Converte Decimal -> string com 2 casas, recursivamente."""
    if isinstance(obj, Decimal):
        return f"{obj:.2f}"
    if isinstance(obj, dict):
        return {k: _serializavel(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_serializavel(v) for v in obj]
    return obj


def registrar_operacao(entry: dict) -> None:
    """Acrescenta uma operação ao histórico. Tolerante a falhas (não quebra o app)."""
    try:
        criar_pastas()
        registro = _serializavel(dict(entry))
        registro.setdefault("datahora", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        with _ARQ.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(registro, ensure_ascii=False) + "\n")
        _trim()
    except (OSError, TypeError, ValueError) as exc:
        # Best-effort: o histórico nunca deve quebrar o processamento.
        log.warning("Não foi possível gravar o histórico: %s", exc)


def _trim() -> None:
    """Mantém no máximo _MAX_ENTRADAS linhas (remove as mais antigas)."""
    try:
        linhas = _ARQ.read_text(encoding="utf-8").splitlines()
    except OSError:
        return
    if len(linhas) > _MAX_ENTRADAS:
        _ARQ.write_text("\n".join(linhas[-_MAX_ENTRADAS:]) + "\n", encoding="utf-8")


def listar_operacoes(limite: int = 200) -> list[dict]:
    """Lê o histórico, mais recentes primeiro. Ignora linhas corrompidas."""
    if not _ARQ.exists():
        return []
    entradas: list[dict] = []
    try:
        for linha in _ARQ.read_text(encoding="utf-8").splitlines():
            linha = linha.strip()
            if not linha:
                continue
            try:
                entradas.append(json.loads(linha))
            except json.JSONDecodeError:
                continue
    except OSError as exc:
        log.warning("Não foi possível ler o histórico: %s", exc)
        return []
    return list(reversed(entradas))[:limite]


def total_operacoes() -> int:
    if not _ARQ.exists():
        return 0
    try:
        return sum(1 for ln in _ARQ.read_text(encoding="utf-8").splitlines() if ln.strip())
    except OSError:
        return 0
