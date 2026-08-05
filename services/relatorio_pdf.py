"""O mesmo compilado do historico, em PDF.

Este modulo NAO soma nada. Ele recebe `relatorio.compilar(...)` e desenha —
exatamente a mesma compilacao que vira a planilha. Se o PDF e o .xlsx
pudessem divergir num centavo, um dos dois estaria mentindo e ninguem
saberia qual; por isso a conta acontece uma vez so, em services/relatorio.py.

Por que reportlab: e puro Python (nada de GTK, wkhtmltopdf ou Excel
instalado), desenha tabelas paginadas com cabecalho repetido e tem graficos
nativos — o PDF sai com o mesmo desenho da planilha, sem virar imagem.

Diferencas deliberadas em relacao ao .xlsx:

* **Sem formulas.** Um PDF nao recalcula; os totais sao os numeros ja
  somados pela compilacao. E a mesma soma que a planilha manda o Excel
  fazer — so que congelada, porque e isso que um PDF e.
* **Retrato, nao paisagem.** O compilado e uma coluna de nome e uma de
  valor; paisagem so criaria vazio.
* **Listas longas paginam.** Uma dimensao com 300 eventos atravessa as
  paginas repetindo o cabecalho, em vez de ser cortada.
"""

from __future__ import annotations

import io
from decimal import Decimal
from xml.sax.saxutils import escape

from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from . import relatorio
from .utils import formatar_moeda

# A mesma paleta slate da interface e das planilhas geradas.
_ESCURO = colors.HexColor("#0F172A")
_SECAO = colors.HexColor("#1E293B")
_CABECALHO = colors.HexColor("#334155")
_BORDA = colors.HexColor("#CBD5E1")
_ZEBRA = colors.HexColor("#F1F5F9")
_NEGATIVO = colors.HexColor("#B91C1C")
_APAGADO = colors.HexColor("#64748B")

# Paleta categorica dos graficos: matizes distintos em luminancia parecida,
# para que as barras se diferenciem tambem impressas em preto e branco.
_SERIES = [colors.HexColor(c) for c in (
    "#2563EB", "#0D9488", "#D97706", "#7C3AED", "#DC2626", "#0891B2",
    "#65A30D", "#DB2777", "#4F46E5", "#CA8A04", "#059669", "#9333EA",
    "#E11D48",
)]

_MARGEM = 16 * mm
_LARGURA_UTIL = A4[0] - 2 * _MARGEM


def _texto(valor) -> str:
    """Neutraliza a marcacao que o `Paragraph` do reportlab interpreta.

    `Paragraph` nao recebe texto puro: recebe um mini-XML com `<b>`, `<font>`
    e afins. Um nome de setor lido de um .xlsx enviado, ou uma secretaria
    digitada na tela, chegam aqui como texto — e bastava um "<font size=99>"
    solto para derrubar a geracao inteira com erro de parse, ou um "<b>" para
    reformatar o documento a revelia de quem o assina.

    Contraparte em PDF do que `estilo_xlsx.texto_seguro` faz na planilha.
    """
    return escape("" if valor is None else str(valor))


def _moeda(valor) -> str:
    """R$ 1.234.567,89 — o MESMO formatador que as telas usam.

    Se o PDF agrupasse milhar por conta propria, um caso de borda faria o
    numero do relatorio assinado divergir do que se leu na tela.
    """
    return formatar_moeda(valor)


def _estilos() -> dict:
    base = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            "titulo", parent=base["Title"], fontName="Helvetica-Bold",
            fontSize=16, leading=20, textColor=colors.white, alignment=TA_CENTER,
        ),
        "secao": ParagraphStyle(
            "secao", parent=base["Heading2"], fontName="Helvetica-Bold",
            fontSize=10.5, leading=14, textColor=colors.white, spaceBefore=0, spaceAfter=0,
        ),
        "corpo": ParagraphStyle(
            "corpo", parent=base["Normal"], fontName="Helvetica", fontSize=8.5, leading=11,
        ),
        "nota": ParagraphStyle(
            "nota", parent=base["Normal"], fontName="Helvetica-Oblique",
            fontSize=7.5, leading=10, textColor=_APAGADO,
        ),
    }


def _faixa(texto: str, estilos: dict, cor, estilo_texto: str, altura: float) -> Table:
    """Faixa colorida de largura total — título ou cabeçalho de seção."""
    tabela = Table([[Paragraph(texto, estilos[estilo_texto])]],
                   colWidths=[_LARGURA_UTIL], rowHeights=[altura])
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), cor),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return tabela


def _tabela(linhas: list[list], larguras: list[float], alinhamentos: list[str],
            negrito_ultima: bool = False, repetir_cabecalho: bool = True) -> Table:
    """Tabela com cabeçalho escuro, zebra e a última linha em destaque."""
    tabela = Table(linhas, colWidths=larguras,
                   repeatRows=1 if repetir_cabecalho else 0)
    comandos = [
        ("BACKGROUND", (0, 0), (-1, 0), _CABECALHO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, _BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _ZEBRA]),
    ]
    for coluna, alinhamento in enumerate(alinhamentos):
        comandos.append(("ALIGN", (coluna, 0), (coluna, -1), alinhamento))
    if negrito_ultima:
        comandos += [
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, -1), (-1, -1), _ZEBRA),
            ("LINEABOVE", (0, -1), (-1, -1), 0.9, _CABECALHO),
        ]
    tabela.setStyle(TableStyle(comandos))
    return tabela


# ---------------------------------------------------------------------------
# Seções
# ---------------------------------------------------------------------------

def _cabecalho(dados: dict, estilos: dict) -> list:
    contexto = [
        ["Gerado em", dados["gerado_em"]],
        ["Operações compiladas", str(dados["resumo"]["operacoes"])],
        ["Competências", ", ".join(dados["competencias"]) or "—"],
        ["Secretarias", ", ".join(dados["secretarias"]) or "—"],
    ]
    tabela = Table(
        [[Paragraph(f"<b>{r}</b>", estilos["corpo"]), Paragraph(_texto(v), estilos["corpo"])]
         for r, v in contexto],
        colWidths=[45 * mm, _LARGURA_UTIL - 45 * mm],
    )
    tabela.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return [
        _faixa("RELATÓRIO CONSOLIDADO DE RETENÇÕES", estilos, _ESCURO, "titulo", 13 * mm),
        Spacer(1, 5 * mm), tabela, Spacer(1, 6 * mm),
    ]


def _resumo(dados: dict, estilos: dict) -> list:
    resumo = dados["resumo"]
    linhas = [["Descrição", "Valor"]]
    linhas += [[rotulo, _moeda(resumo[chave])] for rotulo, chave, _d in relatorio.LINHAS_RESUMO]
    linhas.append(["Diferença (deve ser zero)", _moeda(resumo["diferenca"])])

    tabela = _tabela(linhas, [_LARGURA_UTIL - 45 * mm, 45 * mm], ["LEFT", "RIGHT"])
    comandos = [("FONTNAME", (0, 1), (-1, 2), "Helvetica-Bold")]  # lido e preenchido
    if abs(resumo["diferenca"]) >= Decimal("0.01"):
        comandos += [("TEXTCOLOR", (0, -1), (-1, -1), _NEGATIVO),
                     ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]
    tabela.setStyle(TableStyle(comandos))

    volume = Paragraph(
        f"{resumo['lancamentos']:,} lançamentos lidos · "
        f"{resumo['celulas']:,} células preenchidas".replace(",", "."),
        estilos["nota"],
    )
    return [
        _faixa("RESUMO GERAL DO PERÍODO", estilos, _SECAO, "secao", 7 * mm),
        Spacer(1, 2 * mm), tabela, Spacer(1, 1.5 * mm), volume, Spacer(1, 6 * mm),
    ]


# Espaco reservado a esquerda para os nomes das categorias.
_FAIXA_NOMES = 132
_ALTURA_BARRA = 15


def _grafico_barras(itens) -> Drawing:
    """Barras horizontais: o nome da categoria lê na horizontal, sem corte.

    Barras verticais forcam rotacionar o rotulo, e nome de setor ou de
    evento em portugues ("12/120 - EMPRESTIMO CEF 1") vira um rabisco
    ilegivel de 6pt. Na horizontal o texto ocupa a faixa da esquerda e a
    barra ocupa o resto — que e como a leitura acontece de qualquer forma.
    """
    altura = 26 + len(itens) * _ALTURA_BARRA
    desenho = Drawing(_LARGURA_UTIL, altura)

    grafico = HorizontalBarChart()
    grafico.x, grafico.y = _FAIXA_NOMES, 16
    grafico.height = altura - 24
    grafico.width = _LARGURA_UTIL - _FAIXA_NOMES - 8

    # reportlab desenha o primeiro item embaixo; a lista vem da maior para a
    # menor, e invertê-la coloca a maior no topo — como se lê a tabela.
    invertidos = list(reversed(itens))
    grafico.data = [[float(v) for _k, v in invertidos]]
    grafico.categoryAxis.categoryNames = [k[:34] for k, _v in invertidos]
    grafico.categoryAxis.labels.fontSize = 7
    grafico.categoryAxis.labels.boxAnchor = "e"
    grafico.categoryAxis.labels.dx = -3
    grafico.categoryAxis.strokeWidth = 0.4
    grafico.categoryAxis.strokeColor = _BORDA
    # Ancora em zero só quando tudo é positivo. Com um estorno na lista, um
    # mínimo fixo em zero esconderia a barra negativa — e um gráfico que
    # omite dinheiro é pior que gráfico nenhum.
    valores = [float(v) for _k, v in invertidos]
    grafico.valueAxis.valueMin = 0 if min(valores, default=0) >= 0 else None
    grafico.valueAxis.labels.fontSize = 6
    grafico.valueAxis.strokeColor = _BORDA
    grafico.valueAxis.gridStrokeColor = _BORDA
    grafico.valueAxis.gridStrokeWidth = 0.25
    grafico.valueAxis.visibleGrid = True
    grafico.barSpacing = 2
    grafico.groupSpacing = 4
    grafico.bars.strokeWidth = 0
    for indice in range(len(invertidos)):
        grafico.bars[(0, indice)].fillColor = _SERIES[
            (len(invertidos) - 1 - indice) % len(_SERIES)
        ]
    desenho.add(grafico)
    return desenho


def _grafico_pizza(itens, estilos: dict) -> Drawing:
    desenho = Drawing(_LARGURA_UTIL, 78)
    pizza = Pie()
    pizza.x, pizza.y = 8, 4
    pizza.width = pizza.height = 70
    pizza.data = [float(v) for _k, v in itens]
    pizza.labels = [k[:18] for k, _v in itens]
    pizza.slices.fontSize = 6
    pizza.slices.strokeWidth = 0.4
    pizza.slices.strokeColor = colors.white
    for indice in range(len(itens)):
        pizza.slices[indice].fillColor = _SERIES[indice % len(_SERIES)]
    desenho.add(pizza)

    total = sum(Decimal(str(v)) for _k, v in itens)
    for indice, (rotulo, valor) in enumerate(itens[:8]):
        y = 66 - indice * 8
        desenho.add(String(96, y, "■", fontSize=7,
                           fillColor=_SERIES[indice % len(_SERIES)]))
        parte = (Decimal(str(valor)) / total * 100) if total else Decimal(0)
        desenho.add(String(106, y, f"{rotulo[:38]} — {_moeda(valor)} ({parte:.1f}%)",
                           fontName="Helvetica", fontSize=7))
    return desenho


def _dimensao(dim: dict, estilos: dict) -> list:
    """Uma seção do compilado: faixa, gráfico e a tabela inteira."""
    itens = dim["itens"]
    if not itens:
        return []

    linhas = [[dim["rotulo_coluna"], "Total preenchido"]]
    linhas += [[Paragraph(_texto(rotulo), estilos["corpo"]), _moeda(valor)]
               for rotulo, valor in itens]
    linhas.append(["TOTAL", _moeda(dim["total"])])
    tabela = _tabela(linhas, [_LARGURA_UTIL - 45 * mm, 45 * mm],
                     ["LEFT", "RIGHT"], negrito_ultima=True)

    # Pizza só faz sentido como parte de um todo: com valor negativo em
    # cena, "parte do todo" não existe e as fatias mentiriam. Cai para
    # barras, que representam negativo sem distorcer.
    recorte = relatorio.para_grafico(dim)
    cabe_pizza = dim["chave"] == "Tipo" and all(valor >= 0 for _rotulo, valor in recorte)
    desenho = _grafico_pizza(recorte, estilos) if cabe_pizza else _grafico_barras(recorte)

    # A faixa e o gráfico nunca se separam da tabela que explicam; a tabela
    # em si pode paginar (e repete o cabeçalho ao virar a página).
    abertura = KeepTogether([
        _faixa(dim["titulo"], estilos, _SECAO, "secao", 7 * mm),
        Spacer(1, 2 * mm), desenho, Spacer(1, 2 * mm),
    ])
    partes = [abertura, tabela]
    if dim["sem_detalhe"]:
        partes.append(Paragraph(
            f"{dim['sem_detalhe']} operação(ões) não têm este detalhamento "
            f"(registradas antes de a dimensão existir): "
            f"{_moeda(dim['valor_sem_detalhe'])} preenchido não aparece acima.",
            estilos["nota"],
        ))
    partes.append(Spacer(1, 6 * mm))
    return partes


def _operacoes(dados: dict, estilos: dict) -> list:
    linhas = [list(relatorio.COLUNAS_OPERACOES)]
    for op in dados["operacoes"]:
        linhas.append([
            _texto(op["datahora"]), Paragraph(_texto(op["secretaria"]), estilos["corpo"]),
            _texto(op["competencia"]) or "—", _texto(op["aba"]),
            _moeda(op["lido"]), _moeda(op["preenchido"]),
        ])
    linhas.append(["TOTAL", "", "", "",
                   _moeda(dados["resumo"]["lido"]), _moeda(dados["resumo"]["preenchido"])])

    larguras = [26 * mm, _LARGURA_UTIL - 128 * mm, 20 * mm, 24 * mm, 29 * mm, 29 * mm]
    tabela = _tabela(linhas, larguras,
                     ["LEFT", "LEFT", "CENTER", "LEFT", "RIGHT", "RIGHT"],
                     negrito_ultima=True)
    return [
        _faixa("OPERAÇÕES COMPILADAS", estilos, _SECAO, "secao", 7 * mm),
        Spacer(1, 2 * mm), tabela, Spacer(1, 4 * mm),
    ]


def _rodape(canvas, documento) -> None:
    """Numeração e origem em toda página — um relatório sem procedência é
    um número que ninguém consegue conferir."""
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(_APAGADO)
    canvas.drawString(_MARGEM, 10 * mm, "Automação de Retenções · histórico local (sem dados pessoais)")
    canvas.drawRightString(A4[0] - _MARGEM, 10 * mm, f"Página {documento.page}")
    canvas.setStrokeColor(_BORDA)
    canvas.setLineWidth(0.4)
    canvas.line(_MARGEM, 13 * mm, A4[0] - _MARGEM, 13 * mm)
    canvas.restoreState()


# ---------------------------------------------------------------------------
# Montagem
# ---------------------------------------------------------------------------

def montar_pdf(operacoes: list[dict]) -> bytes:
    """Compila e desenha o PDF. Levanta `relatorio.SemOperacoes` se vazio."""
    dados = relatorio.compilar(operacoes)
    estilos = _estilos()

    elementos = _cabecalho(dados, estilos) + _resumo(dados, estilos)
    for dimensao in dados["dimensoes"]:
        elementos += _dimensao(dimensao, estilos)
    elementos.append(PageBreak())
    elementos += _operacoes(dados, estilos)
    elementos.append(Paragraph(
        "Os totais por dimensão somam o que foi destinado à planilha. "
        "Este relatório vem do histórico local, que não guarda dados pessoais: "
        "nenhum nome, matrícula ou CPF.",
        estilos["nota"],
    ))

    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=_MARGEM, rightMargin=_MARGEM,
        topMargin=14 * mm, bottomMargin=18 * mm,
        title="Relatório consolidado de retenções",
        author="Automação de Retenções",
    )
    documento.build(elementos, onFirstPage=_rodape, onLaterPages=_rodape)
    return buffer.getvalue()
