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
import os
import time
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation
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
HISTORICO_DIR = BASE_DIR / "historico"  # historico local (nunca vai para o git)

# O molde fixo e os vinculos aprendidos vivem por perfil (services/perfis.py).

_logger: logging.Logger | None = None


def criar_pastas() -> None:
    """Garante que todas as pastas de trabalho existam."""
    for pasta in (UPLOADS_DIR, OUTPUTS_DIR, LOGS_DIR, CONFIG_DIR, MODELOS_DIR,
                  SESSIONS_DIR, HISTORICO_DIR):
        pasta.mkdir(parents=True, exist_ok=True)


def _limpar_antigos(pasta: Path, max_horas: float, manter: set[str] = frozenset({".gitkeep"})) -> int:
    """Remove arquivos da pasta com mais de 'max_horas'. Retorna quantos removeu."""
    if not pasta.exists():
        return 0
    limite = time.time() - max_horas * 3600
    removidos = 0
    for arq in pasta.iterdir():
        if not arq.is_file() or arq.name in manter:
            continue
        try:
            if arq.stat().st_mtime < limite:
                arq.unlink()
                removidos += 1
        except OSError:
            continue  # arquivo em uso/inacessível — ignora nesta passada
    return removidos


def limpar_temporarios(uploads_horas: float = 24, sessoes_horas: float = 24,
                       outputs_horas: float = 24 * 7) -> None:
    """Faz uma limpeza de retenção: uploads (PII) e sessões em 24h, saídas em 7 dias.

    Idempotente e tolerante a falhas — pensada para rodar no startup do app.
    """
    n = (_limpar_antigos(UPLOADS_DIR, uploads_horas)
         + _limpar_antigos(SESSIONS_DIR, sessoes_horas)
         + _limpar_antigos(OUTPUTS_DIR, outputs_horas))
    if n:
        configurar_logs().info("Limpeza de temporários: %d arquivo(s) removido(s).", n)


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


def formatar_moeda(valor) -> str:
    """Decimal/float -> "R$ 1.234.567,89". Sem depender de locale do sistema.

    Uma implementacao so para as duas superficies que exibem dinheiro como
    TEXTO: o filtro das telas e o PDF. Duas versoes do mesmo agrupamento
    acabariam divergindo num caso de borda — e o numero na tela deixaria de
    bater com o numero no relatorio assinado.

    (A planilha nao usa isto: la o valor vai como numero e quem formata e o
    Excel, via `estilo_xlsx.MOEDA`.)
    """
    try:
        numero = Decimal(str(valor))
    except (InvalidOperation, ValueError, TypeError):
        return "R$ 0,00"

    inteiro, _, decimais = f"{abs(numero):.2f}".partition(".")
    grupos = []
    while len(inteiro) > 3:
        grupos.insert(0, inteiro[-3:])
        inteiro = inteiro[:-3]
    grupos.insert(0, inteiro)
    sinal = "-" if numero < 0 else ""
    return f"R$ {sinal}{'.'.join(grupos)},{decimais}"


def gerar_nome_saida(prefixo: str = "RETENCAO_PREENCHIDA") -> str:
    """Nome padronizado do arquivo final: PREFIXO_YYYYMMDD_HHMMSS.xlsx."""
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefixo}_{carimbo}.xlsx"


# ---------------------------------------------------------------------------
# JSON em disco: escrita atomica e leitura tolerante
# ---------------------------------------------------------------------------
#
# Todo estado aprendido do app (vinculos dos tres eixos, estrutura do molde,
# registro de perfis) vive em JSON. Sao meses de decisoes de quem opera, e
# valem mais que o codigo: refaze-las custa a conferencia inteira.
#
# Por isso as duas metades andam juntas:
#   * gravar sempre num temporario e trocar de uma vez — `open("w")` trunca
#     antes de escrever, e uma queda no meio deixaria o arquivo pela metade;
#   * ler tolerando ausencia e corrupcao — um JSON quebrado nao pode derrubar
#     o processamento do mes; ele volta ao padrao e o log registra o motivo.

def gravar_json(caminho: Path, dados) -> None:
    """Grava JSON de forma atomica (temporario + troca)."""
    caminho.parent.mkdir(parents=True, exist_ok=True)
    temporario = caminho.with_suffix(caminho.suffix + ".tmp")
    with temporario.open("w", encoding="utf-8") as fh:
        json.dump(dados, fh, ensure_ascii=False, indent=2)
    os.replace(temporario, caminho)


def ler_json(caminho: Path, padrao, descricao: str = ""):
    """Le JSON; devolve `padrao` se o arquivo nao existir ou estiver corrompido.

    Arquivo corrompido e registrado no log: perder o aprendizado em silencio
    seria pior que perde-lo — ninguem investigaria o mes que voltou a pedir
    todos os vinculos de novo.
    """
    if not caminho.exists():
        return padrao
    try:
        with caminho.open("r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        configurar_logs().warning(
            "%s inválido (%s) — seguindo com o padrão.", descricao or caminho.name, exc
        )
        return padrao


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
