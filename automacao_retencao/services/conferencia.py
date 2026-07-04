"""Agregacao e conferencia de totais.

Todos os totais sao calculados na propria aplicacao com Decimal, sem
depender do Excel recalcular no momento da geracao. Ao final, uma aba
'CONFERÊNCIA_AUTOMAÇÃO' e escrita no arquivo de saida.
"""

from __future__ import annotations

from datetime import datetime
from decimal import Decimal

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

_ZERO = Decimal("0.00")


def agregar_lancamentos(lancamentos: list[dict]) -> tuple[dict, list[dict]]:
    """Agrupa lancamentos por (setor, tipo, rubrica) somando os valores.

    Apenas lancamentos totalmente mapeados (setor + tipo + rubrica) entram
    na agregacao para preenchimento. Os demais sao devolvidos como
    'ignorados' para a conferencia.

    Retorna (agregados, ignorados).
    """
    agregados: dict[tuple, Decimal] = {}
    ignorados: list[dict] = []

    for reg in lancamentos:
        setor = reg.get("setor_destino")
        tipo = reg.get("tipo_destino")
        rubrica = reg.get("rubrica_destino")

        if setor and tipo and rubrica:
            chave = (setor, tipo, rubrica)
            agregados[chave] = agregados.get(chave, _ZERO) + reg["valor"]
        else:
            ignorados.append(reg)

    return agregados, ignorados


def calcular_totais_lidos(lancamentos: list[dict]) -> Decimal:
    """Total bruto: soma de todos os lancamentos lidos do relatorio."""
    total = _ZERO
    for reg in lancamentos:
        total += reg["valor"]
    return total


def calcular_totais_agregados(agregados: dict) -> dict:
    """Deriva totais por setor, por rubrica e por tipo a partir dos agregados."""
    por_setor: dict[str, Decimal] = {}
    por_rubrica: dict[str, Decimal] = {}
    por_tipo: dict[str, Decimal] = {}
    total = _ZERO

    for (setor, tipo, rubrica), valor in agregados.items():
        por_setor[setor] = por_setor.get(setor, _ZERO) + valor
        por_rubrica[rubrica] = por_rubrica.get(rubrica, _ZERO) + valor
        por_tipo[tipo] = por_tipo.get(tipo, _ZERO) + valor
        total += valor

    return {
        "por_setor": dict(sorted(por_setor.items())),
        "por_rubrica": dict(sorted(por_rubrica.items())),
        "por_tipo": dict(sorted(por_tipo.items())),
        "total": total,
    }


def calcular_pendencias(ignorados: list[dict], pendencias_estrutura: list[dict]) -> dict:
    """Consolida o valor pendente e as listas de motivos.

    'ignorados' = lancamentos sem mapeamento completo.
    'pendencias_estrutura' = itens que existiam mas nao acharam lugar na planilha.
    """
    total_pendente = _ZERO
    for reg in ignorados:
        total_pendente += reg["valor"]
    for item in pendencias_estrutura:
        total_pendente += item.get("valor", _ZERO)

    return {"total_pendente": total_pendente}


# ---------------------------------------------------------------------------
# Escrita da aba de conferencia
# ---------------------------------------------------------------------------

_FONTE_TITULO = Font(bold=True, size=14, color="FFFFFF")
_FONTE_SECAO = Font(bold=True, size=11, color="FFFFFF")
_FONTE_CABECALHO = Font(bold=True, color="FFFFFF")
_FILL_TITULO = PatternFill("solid", fgColor="0F172A")
_FILL_SECAO = PatternFill("solid", fgColor="1E293B")
_FILL_CABECALHO = PatternFill("solid", fgColor="334155")
_FILL_ALERTA = PatternFill("solid", fgColor="7F1D1D")
_BORDA = Border(*(Side(style="thin", color="CBD5E1"),) * 4)
_MOEDA = 'R$ #,##0.00'


def criar_aba_conferencia(wb, dados: dict) -> None:
    """Cria (ou recria) a aba 'CONFERÊNCIA_AUTOMAÇÃO' com todo o resumo.

    'dados' deve conter:
        competencia, arquivo_origem, arquivo_modelo, aba_destino,
        total_lido, total_preenchido, total_pendente,
        por_setor, por_rubrica, por_tipo,
        lotacoes_nao_mapeadas, rubricas_nao_mapeadas, folhas_desconhecidas,
        pendencias_estrutura
    """
    nome = "CONFERÊNCIA_AUTOMAÇÃO"
    if nome in wb.sheetnames:
        del wb[nome]
    ws = wb.create_sheet(title=nome)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 16

    linha = 1
    linha = _titulo(ws, linha, "CONFERÊNCIA DA AUTOMAÇÃO DE RETENÇÕES")
    linha += 1

    # Metadados -----------------------------------------------------------
    meta = [
        ("Processado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("Competência", dados.get("competencia") or "—"),
        ("Arquivo de origem", dados.get("arquivo_origem") or "—"),
        ("Planilha modelo", dados.get("arquivo_modelo") or "—"),
        ("Aba preenchida", dados.get("aba_destino") or "—"),
    ]
    for rotulo, valor in meta:
        ws.cell(row=linha, column=1, value=rotulo).font = Font(bold=True)
        ws.cell(row=linha, column=2, value=valor)
        linha += 1
    linha += 1

    # Totais gerais -------------------------------------------------------
    linha = _secao(ws, linha, "TOTAIS GERAIS")
    totais = [
        ("Total lido (bruto)", dados.get("total_lido", _ZERO)),
        ("Total preenchido", dados.get("total_preenchido", _ZERO)),
        ("Total pendente", dados.get("total_pendente", _ZERO)),
    ]
    for rotulo, valor in totais:
        ws.cell(row=linha, column=1, value=rotulo).font = Font(bold=True)
        c = ws.cell(row=linha, column=2, value=float(valor))
        c.number_format = _MOEDA
        linha += 1
    linha += 1

    # Tabelas -------------------------------------------------------------
    linha = _tabela_valor(ws, linha, "TOTAL POR SETOR", dados.get("por_setor", {}))
    linha = _tabela_valor(ws, linha, "TOTAL POR RUBRICA", dados.get("por_rubrica", {}))
    linha = _tabela_valor(ws, linha, "TOTAL POR TIPO", dados.get("por_tipo", {}))

    # Pendencias ----------------------------------------------------------
    linha = _secao(ws, linha, "PENDÊNCIAS", alerta=True)
    linha = _tabela_contagem(
        ws, linha, "Lotações não mapeadas", dados.get("lotacoes_nao_mapeadas", {})
    )
    linha = _tabela_contagem(
        ws, linha, "Rubricas não mapeadas", dados.get("rubricas_nao_mapeadas", {})
    )
    linha = _tabela_contagem(
        ws, linha, "Tipos de folha desconhecidos", dados.get("folhas_desconhecidas", {})
    )

    estrutura = dados.get("pendencias_estrutura", [])
    if estrutura:
        ws.cell(row=linha, column=1, value="Itens sem lugar na planilha").font = Font(bold=True, italic=True)
        linha += 1
        for cab, col in (("Motivo", 1), ("Setor / Rubrica", 2), ("Valor", 3)):
            cel = ws.cell(row=linha, column=col, value=cab)
            cel.font = _FONTE_CABECALHO
            cel.fill = _FILL_CABECALHO
        linha += 1
        for item in estrutura:
            ws.cell(row=linha, column=1, value=item.get("motivo"))
            ws.cell(row=linha, column=2, value=f"{item.get('setor')} / {item.get('rubrica')}")
            c = ws.cell(row=linha, column=3, value=float(item.get("valor", _ZERO)))
            c.number_format = _MOEDA
            linha += 1


def _titulo(ws, linha, texto):
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
    c = ws.cell(row=linha, column=1, value=texto)
    c.font = _FONTE_TITULO
    c.fill = _FILL_TITULO
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[linha].height = 26
    return linha + 1


def _secao(ws, linha, texto, alerta=False):
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
    c = ws.cell(row=linha, column=1, value=texto)
    c.font = _FONTE_SECAO
    c.fill = _FILL_ALERTA if alerta else _FILL_SECAO
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[linha].height = 20
    return linha + 1


def _tabela_valor(ws, linha, titulo, dados: dict):
    linha = _secao(ws, linha, titulo)
    if not dados:
        ws.cell(row=linha, column=1, value="— nada a exibir —").font = Font(italic=True, color="94A3B8")
        return linha + 2

    for cab, col in (("Descrição", 1), ("Valor", 2)):
        cel = ws.cell(row=linha, column=col, value=cab)
        cel.font = _FONTE_CABECALHO
        cel.fill = _FILL_CABECALHO
    linha += 1

    total = _ZERO
    for nome, valor in dados.items():
        ws.cell(row=linha, column=1, value=nome).border = _BORDA
        c = ws.cell(row=linha, column=2, value=float(valor))
        c.number_format = _MOEDA
        c.border = _BORDA
        total += valor
        linha += 1

    ws.cell(row=linha, column=1, value="TOTAL").font = Font(bold=True)
    c = ws.cell(row=linha, column=2, value=float(total))
    c.number_format = _MOEDA
    c.font = Font(bold=True)
    return linha + 2


def _tabela_contagem(ws, linha, titulo, dados: dict):
    ws.cell(row=linha, column=1, value=titulo).font = Font(bold=True, italic=True)
    linha += 1
    if not dados:
        ws.cell(row=linha, column=1, value="Nenhuma").font = Font(italic=True, color="16A34A")
        return linha + 2

    for cab, col in (("Item", 1), ("Ocorrências", 2)):
        cel = ws.cell(row=linha, column=col, value=cab)
        cel.font = _FONTE_CABECALHO
        cel.fill = _FILL_CABECALHO
    linha += 1
    for nome, qtd in dados.items():
        ws.cell(row=linha, column=1, value=nome).border = _BORDA
        ws.cell(row=linha, column=2, value=qtd).border = _BORDA
        linha += 1
    return linha + 2
