"""Histórico local das operações de preenchimento.

Append-only em JSONL (uma operação por linha) — robusto e sem reescrever o
arquivo a cada gravação. Guardado apenas em disco, em `historico/`, que está
no .gitignore: **nunca vai para o GitHub**.

Não guarda PII: só metadados e totais agregados (nada de nome/CPF/servidor).
Vale para o JSONL e para os retratos ao lado dele.

**Duas gravações, de propósito.** A linha do JSONL é o resumo que a tela
lista; o *retrato* da operação — a conferência inteira, que vira o PDF — vai
num arquivo próprio em `historico/retratos/<id>.json`. Guardar o retrato
dentro da linha engordaria o JSONL em uma ordem de grandeza, e ele é lido
INTEIRO a cada abertura da tela de histórico: a lista ficaria lenta por causa
de um dado que só interessa a quem clica em "PDF".
"""

from __future__ import annotations

import json
import logging
import re
import shutil
import uuid
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from .utils import HISTORICO_DIR, criar_pastas, gravar_texto, ler_json

log = logging.getLogger("automacao_retencao")

_ARQ = HISTORICO_DIR / "operacoes.jsonl"
_MAX_ENTRADAS = 2000  # limite defensivo de crescimento do arquivo

# Teto de um retrato em disco. Um retrato real dá poucas dezenas de KB; 2 MB é
# ~100x isso e ainda limita o pior caso a 4 GB no histórico cheio. Ver
# `salvar_retrato`: acima do teto ele recusa, nunca corta pela metade.
MAX_RETRATO_KB = 2048

# O id vira nome de arquivo. Os ids reais são uuid4 em hexa; este filtro
# aceita também os ids curtos dos testes e recusa qualquer coisa que possa
# escapar da pasta (`..`, barra, dois-pontos de drive no Windows).
_ID_VALIDO = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


def _pasta_retratos() -> Path:
    """Ao lado do JSONL — assim redirecionar `_ARQ` leva os retratos junto, e
    um histórico que mudou de lugar não deixa retrato órfão para trás."""
    return _ARQ.parent / "retratos"


def _caminho_retrato(op_id) -> Path | None:
    """Caminho do retrato, ou None se o id não puder virar nome de arquivo."""
    if not op_id or not _ID_VALIDO.match(str(op_id)):
        return None
    pasta = _pasta_retratos().resolve()
    caminho = (pasta / f"{op_id}.json").resolve()
    if pasta not in caminho.parents:  # defesa em profundidade
        return None
    return caminho


def salvar_retrato(op_id: str, retrato: dict) -> bool:
    """Grava o retrato da operação. Best-effort: nunca quebra o processamento."""
    caminho = _caminho_retrato(op_id)
    if caminho is None or not isinstance(retrato, dict):
        return False
    try:
        # Mede antes de gravar. O retrato de uma operação real dá dezenas de
        # KB; um modelo enviado com milhares de blocos de setor levaria isso
        # a megabytes, e o histórico guarda 2.000 operações.
        corpo = json.dumps(retrato, ensure_ascii=False)
    except (TypeError, ValueError) as exc:
        log.warning("Retrato da operação %s não é serializável: %s", op_id, exc)
        return False

    if len(corpo.encode("utf-8")) > MAX_RETRATO_KB * 1024:
        # Recusa em vez de cortar: um retrato pela metade renderizaria um
        # TOTAL menor que o verdadeiro, e um documento de conferência que
        # mente é pior que documento nenhum. A tela deixa de oferecer o PDF
        # desta operação; a planilha, com a aba CONFERÊNCIA_AUTOMAÇÃO
        # completa, continua no lugar.
        log.warning("Retrato da operação %s grande demais (%d KB) — PDF não será oferecido.",
                    op_id, len(corpo) // 1024)
        return False

    try:
        gravar_texto(caminho, corpo)
        return True
    except (OSError, TypeError, ValueError) as exc:
        log.warning("Não foi possível gravar o retrato da operação %s: %s", op_id, exc)
        return False


def carregar_retrato(op_id: str) -> dict | None:
    """Retrato da operação, ou None se ela é anterior a este recurso."""
    caminho = _caminho_retrato(op_id)
    if caminho is None or not caminho.exists():
        return None
    dados = ler_json(caminho, None, f"retrato da operação {op_id}")
    return dados if isinstance(dados, dict) else None


def remover_retrato(op_id: str) -> None:
    """Apaga o retrato. Apagar a operação tem de apagar o que ela guardava —
    senão 'remover do histórico' deixaria o dado no disco."""
    caminho = _caminho_retrato(op_id)
    if caminho is None:
        return
    try:
        caminho.unlink(missing_ok=True)
    except OSError as exc:
        log.warning("Não foi possível remover o retrato da operação %s: %s", op_id, exc)


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
    """Acrescenta uma operação ao histórico. Tolerante a falhas (não quebra o app).

    `entry["retrato"]` — a conferência inteira, se vier — não entra na linha:
    sai para `historico/retratos/<id>.json`. O id só existe aqui dentro (é
    gerado logo abaixo), então é aqui que os dois têm de ser amarrados.
    """
    gravado = None
    try:
        criar_pastas()
        registro = _serializavel(dict(entry))
        retrato = registro.pop("retrato", None)
        registro.setdefault("id", uuid.uuid4().hex)  # id estável para remoção
        registro.setdefault("datahora", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        # A flag evita um acesso a disco por linha ao montar a tela; o
        # download confere o arquivo de novo, porque a flag pode envelhecer.
        registro["tem_retrato"] = bool(retrato) and salvar_retrato(registro["id"], retrato)
        if registro["tem_retrato"]:
            gravado = registro["id"]
        with _ARQ.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(registro, ensure_ascii=False) + "\n")
        _trim()
    except (OSError, TypeError, ValueError) as exc:
        # Best-effort: o histórico nunca deve quebrar o processamento.
        log.warning("Não foi possível gravar o histórico: %s", exc)
        # O retrato é gravado antes da linha (a flag entra nela). Se a linha
        # não foi, o retrato ficaria no disco sem nenhum registro que o
        # mencione — e ninguém o apagaria: o trim só limpa o que ele descarta.
        if gravado:
            remover_retrato(gravado)


def _trim() -> None:
    """Mantém no máximo _MAX_ENTRADAS linhas (remove as mais antigas).

    Leva junto o retrato de cada linha descartada: sem isso a pasta de
    retratos cresceria para sempre enquanto o JSONL fica limitado.
    """
    try:
        linhas = _ARQ.read_text(encoding="utf-8").splitlines()
    except OSError:
        return
    if len(linhas) <= _MAX_ENTRADAS:
        return
    for linha in linhas[:-_MAX_ENTRADAS]:
        try:
            remover_retrato(json.loads(linha).get("id"))
        except json.JSONDecodeError:
            continue
    gravar_texto(_ARQ, "\n".join(linhas[-_MAX_ENTRADAS:]) + "\n")


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


def remover_operacao(op_id: str) -> bool:
    """Remove uma operação do histórico pelo id. Retorna True se removeu."""
    if not op_id or not _ARQ.exists():
        return False
    try:
        linhas = _ARQ.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        log.warning("Não foi possível ler o histórico para remover: %s", exc)
        return False

    mantidas, removeu = [], False
    for linha in linhas:
        if not linha.strip():
            continue
        try:
            if json.loads(linha).get("id") == op_id:
                removeu = True
                continue
        except json.JSONDecodeError:
            continue  # descarta linha corrompida junto
        mantidas.append(linha)

    if not removeu:
        return False
    try:
        conteudo = ("\n".join(mantidas) + "\n") if mantidas else ""
        gravar_texto(_ARQ, conteudo)
    except OSError as exc:
        log.warning("Não foi possível gravar o histórico após remoção: %s", exc)
        return False
    remover_retrato(op_id)
    return True


def limpar_historico() -> int:
    """Apaga todo o histórico — inclusive os retratos.

    "Limpar tudo" tem de limpar tudo mesmo: deixar as conferências no disco
    depois de o usuário mandar apagar seria guardar o que ele pediu para
    esquecer. Retorna quantas operações foram removidas.
    """
    total = total_operacoes()
    try:
        if _ARQ.exists():
            _ARQ.unlink()
    except OSError as exc:
        log.warning("Não foi possível limpar o histórico: %s", exc)
        return 0
    try:
        shutil.rmtree(_pasta_retratos(), ignore_errors=True)
    except OSError as exc:  # pragma: no cover — rmtree já engole o que sabe tratar
        log.warning("Não foi possível remover os retratos: %s", exc)
    return total


def buscar_operacao(op_id: str) -> dict | None:
    """Retorna a operação com o id informado, ou None."""
    if not op_id or not _ARQ.exists():
        return None
    try:
        for linha in _ARQ.read_text(encoding="utf-8").splitlines():
            linha = linha.strip()
            if not linha:
                continue
            try:
                op = json.loads(linha)
            except json.JSONDecodeError:
                continue
            if op.get("id") == op_id:
                return op
    except OSError as exc:
        log.warning("Não foi possível ler o histórico para buscar: %s", exc)
    return None


def buscar_operacoes(ids) -> list[dict]:
    """Operações cujos ids foram informados, na ordem da listagem (recentes primeiro).

    Uma leitura só do arquivo, em vez de uma por id — e ids desconhecidos são
    simplesmente ignorados: uma tela aberta há tempo pode citar uma operação
    que já foi apagada, e isso não é motivo para negar o resto da seleção.
    """
    alvo = {i for i in (ids or []) if i}
    if not alvo:
        return []
    return [op for op in listar_operacoes(limite=_MAX_ENTRADAS) if op.get("id") in alvo]


def total_operacoes() -> int:
    if not _ARQ.exists():
        return 0
    try:
        return sum(1 for ln in _ARQ.read_text(encoding="utf-8").splitlines() if ln.strip())
    except OSError:
        return 0
