"""O mesmo compilado do historico, em PDF.

Este modulo NAO soma nada. Ele recebe `relatorio.compilar(...)` e desenha —
exatamente a mesma compilacao que vira a planilha. Se o PDF e o .xlsx
pudessem divergir num centavo, um dos dois estaria mentindo e ninguem
saberia qual; por isso a conta acontece uma vez so, em services/relatorio.py.

O desenho (paleta, faixas, tabela zebrada, graficos, rodape) vem de
services/pdf_kit, compartilhado com o PDF da conferencia de uma operacao:
os dois documentos saem da mesma mao.

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

from decimal import Decimal

from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, Spacer, TableStyle

from . import pdf_kit as kit
from . import relatorio

_RODAPE = "Automação de Retenções · histórico local (sem dados pessoais)"


# ---------------------------------------------------------------------------
# Seções
# ---------------------------------------------------------------------------

def _cabecalho(dados: dict, estilos: dict) -> list:
    contexto = kit.tabela_contexto([
        ("Gerado em", dados["gerado_em"]),
        ("Operações compiladas", str(dados["resumo"]["operacoes"])),
        ("Competências", ", ".join(dados["competencias"]) or "—"),
        ("Secretarias", ", ".join(dados["secretarias"]) or "—"),
    ], estilos)
    return [
        kit.titulo("RELATÓRIO CONSOLIDADO DE RETENÇÕES", estilos),
        Spacer(1, 5 * mm), contexto, Spacer(1, 6 * mm),
    ]


def _resumo(dados: dict, estilos: dict) -> list:
    resumo = dados["resumo"]
    linhas = [["Descrição", "Valor"]]
    linhas += [[rotulo, kit.moeda(resumo[chave])] for rotulo, chave, _d in relatorio.LINHAS_RESUMO]
    linhas.append(["Diferença (deve ser zero)", kit.moeda(resumo["diferenca"])])

    tabela = kit.tabela(linhas, [kit.LARGURA_UTIL - kit.COL_VALOR, kit.COL_VALOR],
                        ["LEFT", "RIGHT"])
    comandos = [("FONTNAME", (0, 1), (-1, 2), "Helvetica-Bold")]  # lido e preenchido
    if abs(resumo["diferenca"]) >= Decimal("0.01"):
        comandos += [("TEXTCOLOR", (0, -1), (-1, -1), kit.NEGATIVO),
                     ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]
    tabela.setStyle(TableStyle(comandos))

    volume = Paragraph(
        f"{resumo['lancamentos']:,} lançamentos lidos · "
        f"{resumo['celulas']:,} células preenchidas".replace(",", "."),
        estilos["nota"],
    )
    return [
        kit.secao("RESUMO GERAL DO PERÍODO", estilos),
        Spacer(1, 2 * mm), tabela, Spacer(1, 1.5 * mm), volume, Spacer(1, 6 * mm),
    ]


def _dimensao(dim: dict, estilos: dict) -> list:
    """Uma seção do compilado: faixa, gráfico e a tabela inteira."""
    itens = dim["itens"]
    if not itens:
        return []

    tabela = kit.tabela_valor(itens, estilos, dim["rotulo_coluna"], total=dim["total"])

    recorte = relatorio.para_grafico(dim)
    desenho = (kit.grafico_pizza(recorte) if dim["chave"] == "Tipo" and kit.cabe_pizza(recorte)
               else kit.grafico_barras(recorte))

    # A faixa e o gráfico nunca se separam da tabela que explicam; a tabela
    # em si pode paginar (e repete o cabeçalho ao virar a página).
    abertura = KeepTogether([
        kit.secao(dim["titulo"], estilos),
        Spacer(1, 2 * mm), desenho, Spacer(1, 2 * mm),
    ])
    partes = [abertura, tabela]
    if dim["sem_detalhe"]:
        partes.append(Paragraph(
            f"{dim['sem_detalhe']} operação(ões) não têm este detalhamento "
            f"(registradas antes de a dimensão existir): "
            f"{kit.moeda(dim['valor_sem_detalhe'])} preenchido não aparece acima.",
            estilos["nota"],
        ))
    partes.append(Spacer(1, 6 * mm))
    return partes


def _operacoes(dados: dict, estilos: dict) -> list:
    linhas = [list(relatorio.COLUNAS_OPERACOES)]
    for op in dados["operacoes"]:
        linhas.append([
            kit.texto(op["datahora"]), Paragraph(kit.texto(op["secretaria"]), estilos["corpo"]),
            kit.texto(op["competencia"]) or "—", kit.texto(op["aba"]),
            kit.moeda(op["lido"]), kit.moeda(op["preenchido"]),
        ])
    linhas.append(["TOTAL", "", "", "",
                   kit.moeda(dados["resumo"]["lido"]), kit.moeda(dados["resumo"]["preenchido"])])

    larguras = [26 * mm, kit.LARGURA_UTIL - 128 * mm, 20 * mm, 24 * mm, 29 * mm, 29 * mm]
    tabela = kit.tabela(linhas, larguras,
                        ["LEFT", "LEFT", "CENTER", "LEFT", "RIGHT", "RIGHT"],
                        negrito_ultima=True)
    return [
        kit.secao("OPERAÇÕES COMPILADAS", estilos),
        Spacer(1, 2 * mm), tabela, Spacer(1, 4 * mm),
    ]


# ---------------------------------------------------------------------------
# Montagem
# ---------------------------------------------------------------------------

def montar_pdf(operacoes: list[dict]) -> bytes:
    """Compila e desenha o PDF. Levanta `relatorio.SemOperacoes` se vazio."""
    dados = relatorio.compilar(operacoes)
    estilos = kit.estilos()

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

    return kit.montar_documento(elementos, "Relatório consolidado de retenções", _RODAPE)
