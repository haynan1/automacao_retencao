"""Utilitarios de infraestrutura: pastas, logs, nomes de saida e
persistencia de sessao de trabalho (sem banco de dados).

Todo o estado intermediario entre as telas (analise -> preview ->
mapeamento -> processamento) e serializado em JSON dentro de uma pasta
de sessao isolada. Valores monetarios sao guardados como string para
preservar a precisao do Decimal.
"""

from __future__ import annotations

import json
import logging
import uuid
from datetime import datetime
from decimal import Decimal
from logging.handlers import RotatingFileHandler
from pathlib import Path

# Raiz do projeto (pasta que contem app.py).
BASE_DIR = Path(__file__).resolve().parent.parent

UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
LOGS_DIR = BASE_DIR / "logs"
CONFIG_DIR = BASE_DIR / "config"
MODELOS_DIR = BASE_DIR / "modelos"      # molde(s) versionavel(is) no repositorio
SESSIONS_DIR = OUTPUTS_DIR / "_sessions"

# O molde fixo e os vinculos aprendidos vivem por perfil (services/perfis.py).

_logger: logging.Logger | None = None


def criar_pastas() -> None:
    """Garante que todas as pastas de trabalho existam."""
    for pasta in (UPLOADS_DIR, OUTPUTS_DIR, LOGS_DIR, CONFIG_DIR, MODELOS_DIR, SESSIONS_DIR):
        pasta.mkdir(parents=True, exist_ok=True)


def configurar_logs() -> logging.Logger:
    """Configura o log tecnico rotativo em logs/app.log."""
    global _logger
    if _logger is not None:
        return _logger

    criar_pastas()
    logger = logging.getLogger("automacao_retencao")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        handler = RotatingFileHandler(
            LOGS_DIR / "app.log", maxBytes=2_000_000, backupCount=5, encoding="utf-8"
        )
        formato = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
        )
        handler.setFormatter(formato)
        logger.addHandler(handler)

    _logger = logger
    return logger


def gerar_nome_saida(prefixo: str = "RETENCAO_PREENCHIDA") -> str:
    """Nome padronizado do arquivo final: PREFIXO_YYYYMMDD_HHMMSS.xlsx."""
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefixo}_{carimbo}.xlsx"


# ---------------------------------------------------------------------------
# Serializacao segura de estruturas com Decimal
# ---------------------------------------------------------------------------

def _encode(obj):
    """Converte Decimal -> {'__decimal__': '...'} de forma recursiva."""
    if isinstance(obj, Decimal):
        return {"__decimal__": str(obj)}
    if isinstance(obj, dict):
        return {k: _encode(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_encode(v) for v in obj]
    return obj


def _decode(obj):
    """Reconstroi Decimal a partir do marcador serializado."""
    if isinstance(obj, dict):
        if "__decimal__" in obj and len(obj) == 1:
            return Decimal(obj["__decimal__"])
        return {k: _decode(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_decode(v) for v in obj]
    return obj


# ---------------------------------------------------------------------------
# Sessao de trabalho persistida em disco
# ---------------------------------------------------------------------------

def novo_id_sessao() -> str:
    """Gera um identificador opaco (sem dado do usuario no nome do arquivo)."""
    return uuid.uuid4().hex


def salvar_sessao(session_id: str, dados: dict) -> None:
    """Persiste o estado de trabalho da sessao em JSON."""
    criar_pastas()
    caminho = SESSIONS_DIR / f"{session_id}.json"
    with caminho.open("w", encoding="utf-8") as fh:
        json.dump(_encode(dados), fh, ensure_ascii=False, indent=2)


def carregar_sessao(session_id: str) -> dict | None:
    """Le o estado de trabalho da sessao; None se nao existir."""
    caminho = SESSIONS_DIR / f"{session_id}.json"
    if not caminho.exists():
        return None
    with caminho.open("r", encoding="utf-8") as fh:
        return _decode(json.load(fh))
