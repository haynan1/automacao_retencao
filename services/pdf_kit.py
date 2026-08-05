"""Vocabulario visual dos PDFs que o app gera.

Contraparte em PDF de services/estilo_xlsx. Duas superficies desenham PDF: o
relatorio consolidado do historico (services/relatorio_pdf) e a conferencia
de uma operacao (services/conferencia_pdf). Elas saem da mesma mao — mesma
paleta, mesma hierarquia, mesmo formato de moeda — porque quem recebe as
duas nao deveria conseguir dizer que foram escritas por codigos diferentes.

Aqui ficam SO os atomos: paleta, estilos de texto, as faixas de hierarquia,
a tabela zebrada, os dois graficos e o rodape. O CONTEUDO de cada documento
continua no modulo de quem o escreve — este arquivo nao sabe o que e uma
reconciliacao nem uma operacao compilada.

Por que reportlab: e puro Python (nada de GTK, wkhtmltopdf ou Excel
instalado), desenha tabelas paginadas com cabecalho repetido e tem graficos
nativos — o PDF sai desenhado, sem virar imagem.
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
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

from .utils import formatar_moeda, recortar_serie

# A mesma paleta slate da interface e das planilhas geradas.
ESCURO = colors.HexColor("#0F172A")
SECAO = colors.HexColor("#1E293B")
CABECALHO = colors.HexColor("#334155")
BORDA = colors.HexColor("#CBD5E1")
ZEBRA = colors.HexColor("#F1F5F9")
NEGATIVO = colors.HexColor("#B91C1C")
APAGADO = colors.HexColor("#64748B")
# Faixas de veredito, iguais as da aba de conferencia (estilo_xlsx).
OK = colors.HexColor("#14532D")
ALERTA = colors.HexColor("#7F1D1D")

# Paleta categorica dos graficos: matizes distintos em luminancia parecida,
# para que as barras se diferenciem tambem impressas em preto e branco.
SERIES = [colors.HexColor(c) for c in (
    "#2563EB", "#0D9488", "#D97706", "#7C3AED", "#DC2626", "#0891B2",
    "#65A30D", "#DB2777", "#4F46E5", "#CA8A04", "#059669", "#9333EA",
    "#E11D48",
)]

MARGEM = 16 * mm
LARGURA_UTIL = A4[0] - 2 * MARGEM

# Largura padrao da coluna de dinheiro. Uma so, para que as tabelas de
# documentos diferentes alinhem os valores na mesma posicao da folha.
COL_VALOR = 45 * mm

# Altura das duas faixas de hierarquia.
ALTURA_TITULO = 13 * mm
ALTURA_SECAO = 7 * mm


def texto(valor) -> str:
    """Neutraliza a marcacao que o `Paragraph` do reportlab interpreta.

    `Paragraph` nao recebe texto puro: recebe um mini-XML com `<b>`, `<font>`
    e afins. Um nome de setor lido de um .xlsx enviado, ou uma secretaria
    digitada na tela, chegam aqui como texto — e bastava um "<font size=99>"
    solto para derrubar a geracao inteira com erro de parse, ou um "<b>" para
    reformatar o documento a revelia de quem o assina.

    Contraparte em PDF do que `estilo_xlsx.texto_seguro` faz na planilha.
    """
    return escape("" if valor is None else str(valor))


def moeda(valor) -> str:
    """R$ 1.234.567,89 — o MESMO formatador que as telas usam.

    Se o PDF agrupasse milhar por conta propria, um caso de borda faria o
    numero do relatorio assinado divergir do que se leu na tela.
    """
    return formatar_moeda(valor)


def estilos() -> dict:
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
        "veredito": ParagraphStyle(
            "veredito", parent=base["Normal"], fontName="Helvetica-Bold",
            fontSize=11, leading=14, textColor=colors.white, alignment=TA_CENTER,
        ),
        "corpo": ParagraphStyle(
            "corpo", parent=base["Normal"], fontName="Helvetica", fontSize=8.5, leading=11,
        ),
        # O texto em vermelho vive num estilo proprio porque `Paragraph`
        # carrega a propria cor e IGNORA o TEXTCOLOR da tabela: pintar a
        # celula pelo TableStyle nao teria efeito nenhum.
        "corpo_alerta": ParagraphStyle(
            "corpo_alerta", parent=base["Normal"], fontName="Helvetica-Bold",
            fontSize=8.5, leading=11, textColor=NEGATIVO,
        ),
        "nota": ParagraphStyle(
            "nota", parent=base["Normal"], fontName="Helvetica-Oblique",
            fontSize=7.5, leading=10, textColor=APAGADO,
        ),
    }


# ---------------------------------------------------------------------------
# Faixas de hierarquia
# ---------------------------------------------------------------------------

def _faixa(rotulo: str, estilo: dict, cor, estilo_texto: str, altura: float) -> Table:
    """Faixa colorida de largura total — título, seção ou veredito."""
    tabela_faixa = Table([[Paragraph(rotulo, estilo[estilo_texto])]],
                         colWidths=[LARGURA_UTIL], rowHeights=[altura])
    tabela_faixa.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), cor),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return tabela_faixa


def titulo(rotulo: str, estilo: dict) -> Table:
    return _faixa(rotulo, estilo, ESCURO, "titulo", ALTURA_TITULO)


def secao(rotulo: str, estilo: dict) -> Table:
    return _faixa(rotulo, estilo, SECAO, "secao", ALTURA_SECAO)


def veredito(rotulo: str, estilo: dict, positivo: bool) -> Table:
    """Faixa verde ou vermelha — o veredito se lê antes de qualquer número."""
    return _faixa(rotulo, estilo, OK if positivo else ALERTA, "veredito", ALTURA_SECAO)


def tabela_contexto(pares, estilo: dict) -> Table:
    """Rótulo em negrito à esquerda, valor à direita, sem grade.

    A procedência do documento — quem gerou, quando, de qual arquivo. Abre os
    dois PDFs, e por isso mora aqui: dois blocos de abertura com respiros
    diferentes fariam os documentos parecerem de origens diferentes.
    """
    tabela_pronta = Table(
        [[Paragraph(f"<b>{texto(rotulo)}</b>", estilo["corpo"]),
          Paragraph(texto(valor), estilo["corpo"])] for rotulo, valor in pares],
        colWidths=[45 * mm, LARGURA_UTIL - 45 * mm],
    )
    tabela_pronta.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return tabela_pronta


# ---------------------------------------------------------------------------
# Tabelas
# ---------------------------------------------------------------------------

def tabela(linhas: list[list], larguras: list[float], alinhamentos: list[str],
           negrito_ultima: bool = False, repetir_cabecalho: bool = True) -> Table:
    """Tabela com cabeçalho escuro, zebra e a última linha em destaque."""
    tabela_pronta = Table(linhas, colWidths=larguras,
                          repeatRows=1 if repetir_cabecalho else 0)
    comandos = [
        ("BACKGROUND", (0, 0), (-1, 0), CABECALHO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ZEBRA]),
    ]
    for coluna, alinhamento in enumerate(alinhamentos):
        comandos.append(("ALIGN", (coluna, 0), (coluna, -1), alinhamento))
    if negrito_ultima:
        comandos += [
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, -1), (-1, -1), ZEBRA),
            ("LINEABOVE", (0, -1), (-1, -1), 0.9, CABECALHO),
        ]
    tabela_pronta.setStyle(TableStyle(comandos))
    return tabela_pronta


def tabela_valor(itens, estilo: dict, rotulo_coluna: str,
                 rotulo_valor: str = "Total preenchido", total=None) -> Table:
    """A tabela mais comum dos dois documentos: rótulo à esquerda, dinheiro à direita.

    O rótulo vai em `Paragraph` para quebrar linha — nome de setor não cabe
    numa linha só e cortá-lo esconderia de qual setor é o número. A linha de
    TOTAL vai em texto puro, porque é o estilo da tabela que a põe em negrito
    e um `Paragraph` carrega a própria fonte, ignorando esse comando.
    """
    corpo = [[rotulo_coluna, rotulo_valor]]
    corpo += [[Paragraph(texto(rotulo), estilo["corpo"]), moeda(valor)]
              for rotulo, valor in itens]
    if total is not None:
        corpo.append(["TOTAL", moeda(total)])
    return tabela(corpo, [LARGURA_UTIL - COL_VALOR, COL_VALOR],
                  ["LEFT", "RIGHT"], negrito_ultima=total is not None)


# ---------------------------------------------------------------------------
# Gráficos
# ---------------------------------------------------------------------------

# Quantas categorias os GRAFICOS mostram. Nao afeta as tabelas: elas listam
# tudo. Acima disto as legendas se sobrepoem e o desenho para de comunicar,
# entao o excedente vira uma unica categoria "outras" — e a soma continua
# igual a da tabela.
TOP_GRAFICO = 12

# Espaco reservado a esquerda para os nomes das categorias.
_FAIXA_NOMES = 132
_ALTURA_BARRA = 15


def recorte(itens, top: int = TOP_GRAFICO) -> list[tuple[str, Decimal]]:
    """Recorte legível de uma série para desenhar — ver `utils.recortar_serie`,
    que é a mesma regra que o gráfico do .xlsx aplica."""
    return recortar_serie(itens, top)


def grafico_barras(itens) -> Drawing:
    """Barras horizontais: o nome da categoria lê na horizontal, sem corte.

    Barras verticais forcam rotacionar o rotulo, e nome de setor ou de
    evento em portugues ("12/120 - EMPRESTIMO CEF 1") vira um rabisco
    ilegivel de 6pt. Na horizontal o texto ocupa a faixa da esquerda e a
    barra ocupa o resto — que e como a leitura acontece de qualquer forma.
    """
    altura = 26 + len(itens) * _ALTURA_BARRA
    desenho = Drawing(LARGURA_UTIL, altura)

    grafico = HorizontalBarChart()
    grafico.x, grafico.y = _FAIXA_NOMES, 16
    grafico.height = altura - 24
    grafico.width = LARGURA_UTIL - _FAIXA_NOMES - 8

    # reportlab desenha o primeiro item embaixo; a lista vem da maior para a
    # menor, e invertê-la coloca a maior no topo — como se lê a tabela.
    invertidos = list(reversed(itens))
    grafico.data = [[float(v) for _k, v in invertidos]]
    grafico.categoryAxis.categoryNames = [k[:34] for k, _v in invertidos]
    grafico.categoryAxis.labels.fontSize = 7
    grafico.categoryAxis.labels.boxAnchor = "e"
    grafico.categoryAxis.labels.dx = -3
    grafico.categoryAxis.strokeWidth = 0.4
    grafico.categoryAxis.strokeColor = BORDA
    # Ancora em zero só quando tudo é positivo. Com um estorno na lista, um
    # mínimo fixo em zero esconderia a barra negativa — e um gráfico que
    # omite dinheiro é pior que gráfico nenhum.
    valores = [float(v) for _k, v in invertidos]
    grafico.valueAxis.valueMin = 0 if min(valores, default=0) >= 0 else None
    grafico.valueAxis.labels.fontSize = 6
    grafico.valueAxis.strokeColor = BORDA
    grafico.valueAxis.gridStrokeColor = BORDA
    grafico.valueAxis.gridStrokeWidth = 0.25
    grafico.valueAxis.visibleGrid = True
    grafico.barSpacing = 2
    grafico.groupSpacing = 4
    grafico.bars.strokeWidth = 0
    for indice in range(len(invertidos)):
        grafico.bars[(0, indice)].fillColor = SERIES[
            (len(invertidos) - 1 - indice) % len(SERIES)
        ]
    desenho.add(grafico)
    return desenho


def grafico_pizza(itens) -> Drawing:
    """Pizza com legenda em dinheiro e percentual, à direita das fatias."""
    desenho = Drawing(LARGURA_UTIL, 78)
    pizza = Pie()
    pizza.x, pizza.y = 8, 4
    pizza.width = pizza.height = 70
    pizza.data = [float(v) for _k, v in itens]
    pizza.labels = [k[:18] for k, _v in itens]
    pizza.slices.fontSize = 6
    pizza.slices.strokeWidth = 0.4
    pizza.slices.strokeColor = colors.white
    for indice in range(len(itens)):
        pizza.slices[indice].fillColor = SERIES[indice % len(SERIES)]
    desenho.add(pizza)

    total = sum(Decimal(str(v)) for _k, v in itens)
    for indice, (rotulo, valor) in enumerate(itens[:8]):
        y = 66 - indice * 8
        desenho.add(String(96, y, "■", fontSize=7,
                           fillColor=SERIES[indice % len(SERIES)]))
        parte = (Decimal(str(valor)) / total * 100) if total else Decimal(0)
        desenho.add(String(106, y, f"{rotulo[:38]} — {moeda(valor)} ({parte:.1f}%)",
                           fontName="Helvetica", fontSize=7))
    return desenho


def cabe_pizza(itens) -> bool:
    """Pizza só faz sentido como parte de um todo.

    Com valor negativo em cena, "parte do todo" não existe e as fatias
    mentiriam; quem chama cai para barras, que representam negativo sem
    distorcer. Sem itens, não há o que fatiar.
    """
    return bool(itens) and all(valor >= 0 for _rotulo, valor in itens)


# ---------------------------------------------------------------------------
# Montagem do documento
# ---------------------------------------------------------------------------

def _fazer_rodape(origem: str):
    """Numeração e procedência em toda página — um relatório sem origem é um
    número que ninguém consegue conferir."""
    def desenhar(canvas, documento) -> None:
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(APAGADO)
        canvas.drawString(MARGEM, 10 * mm, origem)
        canvas.drawRightString(A4[0] - MARGEM, 10 * mm, f"Página {documento.page}")
        canvas.setStrokeColor(BORDA)
        canvas.setLineWidth(0.4)
        canvas.line(MARGEM, 13 * mm, A4[0] - MARGEM, 13 * mm)
        canvas.restoreState()
    return desenhar


def montar_documento(elementos: list, titulo_documento: str, origem: str) -> bytes:
    """Desenha os elementos em A4 retrato e devolve os bytes do PDF.

    Retrato, nao paisagem: os dois documentos sao uma coluna de nome e uma de
    valor; paisagem so criaria vazio.
    """
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=MARGEM, rightMargin=MARGEM,
        topMargin=14 * mm, bottomMargin=18 * mm,
        title=titulo_documento,
        author="Automação de Retenções",
    )
    rodape = _fazer_rodape(origem)
    documento.build(elementos, onFirstPage=rodape, onLaterPages=rodape)
    return buffer.getvalue()
