"""A conferencia de UMA operacao, em PDF — a tela de resultado num documento.

O que a tela mostra depois de preencher a planilha (reconciliacao, totais por
setor, rubrica e linha de folha, como cada folha e cada evento foram
destinados, e o que ficou pendente) e o que precisa ser anexado ao processo,
impresso e arquivado. Este modulo desenha exatamente isso.

**Este modulo NAO soma nada.** Ele recebe o retrato da operacao — o mesmo
dicionario que a tela renderiza e que virou a aba `CONFERÊNCIA_AUTOMAÇÃO`
dentro da planilha (ver services/conferencia) — e desenha. A conta acontece
uma vez so, em `conferencia.reconciliar` e `conferencia.calcular_totais_*`,
no momento do preenchimento. Tres superficies mostrando a mesma operacao com
tres somas diferentes seria pior que nao mostrar nenhuma.

Duas coisas que este PDF tem e a aba nao:

* **Graficos.** Impresso, um bloco de barras diz em um segundo qual setor
  concentra a retencao — a aba depende de quem abre construir o grafico.
* **Veredito no topo.** A faixa verde ou vermelha da reconciliacao e a
  primeira coisa da folha: quem assina precisa ver se bate ao centavo antes
  de qualquer numero.

E uma coisa que a aba tem e este PDF nao: a lista integral. Uma tabela aqui
para em `MAX_LINHAS` e diz quantas linhas ficaram de fora — o registro
exaustivo e a planilha, que acompanha o mesmo processamento.

O desenho (paleta, faixas, tabelas, graficos, rodape) vem de
services/pdf_kit, compartilhado com o relatorio consolidado do historico.
"""

from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal, InvalidOperation

from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, KeepTogether, Paragraph, Spacer, TableStyle

from . import conferencia
from . import pdf_kit as kit

_ZERO = Decimal("0.00")
_CENTAVO = Decimal("0.01")

_RODAPE = "Automação de Retenções · conferência da operação (totais agregados)"

# Teto de linhas por tabela. Vale para TODAS elas — inclusive os totais por
# setor e por rubrica, que saem da grade de um .xlsx ENVIADO: o teto de 200
# setores de `molde.MAX_SETORES` so vale para o molde desenhado na interface,
# e um modelo enviado com milhares de blocos renderizaria centenas de paginas
# que ninguem le. A aba CONFERÊNCIA_AUTOMAÇÃO da planilha guarda a lista
# inteira; o PDF corta e DIZ quanto cortou, com o valor do que ficou de fora.
MAX_LINHAS = 500

_COL_ROTULO = kit.LARGURA_UTIL - kit.COL_VALOR

# Larguras da tabela de decisao: item | destino | valor | por quê.
_COLS_DECISAO = [46 * mm, 32 * mm, 26 * mm, kit.LARGURA_UTIL - 104 * mm]


def _decimal(valor) -> Decimal:
    """Valor monetario do retrato. Ausente, nulo ou ilegivel -> zero.

    A sessao atravessa versoes do app: um retrato gravado antes de um campo
    existir nao pode derrubar a geracao do documento.
    """
    if valor is None or valor == "":
        return _ZERO
    try:
        return Decimal(str(valor))
    except (InvalidOperation, ValueError, ArithmeticError):
        return _ZERO


def _itens(mapa) -> list[tuple[str, Decimal]]:
    """{rótulo: valor} -> lista de pares, na ordem em que a tela mostra."""
    if not isinstance(mapa, dict):
        return []
    return [(str(rotulo), _decimal(valor)) for rotulo, valor in mapa.items()]


def _cortar(itens: list, estilos: dict, unidade: str) -> tuple[list, list]:
    """Aplica `MAX_LINHAS`. Devolve (linhas exibidas, aviso do que sobrou)."""
    if len(itens) <= MAX_LINHAS:
        return itens, []
    sobra = len(itens) - MAX_LINHAS
    return itens[:MAX_LINHAS], [Paragraph(
        f"Mais {sobra} {unidade} não cabem neste documento. A lista integral "
        f"está na aba CONFERÊNCIA_AUTOMAÇÃO da planilha preenchida.",
        estilos["nota"],
    )]


# ---------------------------------------------------------------------------
# Seções
# ---------------------------------------------------------------------------

def _cabecalho(dados: dict, estilos: dict) -> list:
    contexto = [
        ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("Secretaria", dados.get("perfil_nome") or "—"),
        ("Competência", dados.get("competencia") or "—"),
        ("Aba preenchida", dados.get("aba_destino") or "—"),
        ("Arquivo de origem", dados.get("arquivo_origem") or "—"),
        ("Planilha modelo", dados.get("arquivo_modelo") or "—"),
        ("Planilha gerada", dados.get("output_nome") or "—"),
        ("Células lançadas", str(dados.get("qtd_celulas") or 0)),
    ]
    return [
        kit.titulo("CONFERÊNCIA DA AUTOMAÇÃO DE RETENÇÕES", estilos),
        Spacer(1, 5 * mm), kit.tabela_contexto(contexto, estilos), Spacer(1, 6 * mm),
    ]


def _reconciliacao(dados: dict, estilos: dict) -> list:
    """O veredito primeiro, os números depois — é nessa ordem que se confere."""
    rec = dados.get("reconciliacao") or {}
    diferenca = _decimal(rec.get("diferenca"))
    confere = bool(rec.get("confere"))

    faixa = kit.veredito(
        "✔ CONFERE — a soma bate ao centavo" if confere
        else f"✘ DIVERGÊNCIA de {kit.moeda(diferenca)} — confira antes de enviar",
        estilos, confere,
    )

    linhas = [["Descrição", "Valor"]]
    linhas += [[rotulo, kit.moeda(_decimal(rec.get(chave)))]
               for rotulo, chave, _destaque in conferencia.LINHAS_RECONCILIACAO]
    linhas.append(["Diferença (deve ser zero)", kit.moeda(diferenca)])

    tabela = kit.tabela(linhas, [_COL_ROTULO, kit.COL_VALOR], ["LEFT", "RIGHT"])
    comandos = [
        ("FONTNAME", (0, 1), (-1, 2), "Helvetica-Bold"),  # lido e preenchido
        # A diferença não é mais uma parcela: é a conferência das anteriores,
        # e o traço acima dela diz isso sem precisar de legenda.
        ("LINEABOVE", (0, -1), (-1, -1), 0.9, kit.CABECALHO),
    ]
    if abs(diferenca) >= _CENTAVO:
        comandos += [("TEXTCOLOR", (0, -1), (-1, -1), kit.NEGATIVO),
                     ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]
    tabela.setStyle(TableStyle(comandos))

    return [
        kit.secao("RECONCILIAÇÃO", estilos), Spacer(1, 2 * mm),
        faixa, Spacer(1, 2 * mm), tabela,
        Spacer(1, 1.5 * mm),
        Paragraph(
            "Lido = preenchido + sem lugar na planilha + fora de escopo + pendente. "
            "Toda diferença aparece acima: nada é arredondado nem descartado em silêncio.",
            estilos["nota"],
        ),
        Spacer(1, 6 * mm),
    ]


def _dimensao(titulo: str, rotulo_coluna: str, itens: list, estilos: dict,
              pizza: bool = False) -> list:
    """Faixa, gráfico e tabela de um total — a seção que se repete três vezes."""
    if not itens:
        return []

    # O TOTAL soma TUDO, mesmo o que não coube na folha: um total que só
    # somasse as linhas visíveis divergiria da tela e da planilha. O aviso
    # abaixo diz quanto ficou de fora, para a soma visível fechar na conta.
    total = sum((v for _k, v in itens), _ZERO)
    exibidos, sobra = itens[:MAX_LINHAS], itens[MAX_LINHAS:]
    tabela = kit.tabela_valor(exibidos, estilos, rotulo_coluna, total=total)
    aviso = [] if not sobra else [Paragraph(
        f"Outras {len(sobra)} linha(s), somando {kit.moeda(sum((v for _k, v in sobra), _ZERO))}, "
        f"não cabem neste documento — já estão no TOTAL acima. A lista integral está "
        f"na aba CONFERÊNCIA_AUTOMAÇÃO da planilha preenchida.",
        estilos["nota"],
    )]
    # O gráfico ordena por valor; a tabela mantém a ordem da tela. Quem
    # desenha quer a hierarquia; quem confere quer achar o nome na lista.
    ordenados = sorted(itens, key=lambda kv: kv[1], reverse=True)
    recorte = kit.recorte(ordenados)
    desenho = (kit.grafico_pizza(recorte) if pizza and kit.cabe_pizza(recorte)
               else kit.grafico_barras(recorte))

    # A faixa e o gráfico nunca se separam da tabela que explicam; a tabela
    # em si pode paginar, repetindo o cabeçalho ao virar a página.
    return [
        KeepTogether([kit.secao(titulo, estilos), Spacer(1, 2 * mm),
                      desenho, Spacer(1, 2 * mm)]),
        tabela,
    ] + aviso + [Spacer(1, 6 * mm)]


def _totais(dados: dict, estilos: dict) -> list:
    elementos = _dimensao("TOTAL POR SETOR", "Setor", _itens(dados.get("por_setor")), estilos)
    elementos += _dimensao("TOTAL POR RUBRICA (COLUNA DA PLANILHA)", "Rubrica",
                           _itens(dados.get("por_coluna")), estilos)
    elementos += _dimensao("TOTAL POR LINHA DE TIPO DE FOLHA", "Linha",
                           _itens(dados.get("por_tipo")), estilos, pizza=True)
    return elementos


def _tabela_decisao(titulo: str, grupos: list, cab_item: str, cab_destino: str,
                    campo: str, estilos: dict) -> list:
    """Item → destino → por quê: o rastro de auditoria de cada vínculo."""
    if not isinstance(grupos, list) or not grupos:
        return [kit.secao(titulo, estilos), Spacer(1, 2 * mm),
                Paragraph("Nada a exibir.", estilos["nota"]), Spacer(1, 5 * mm)]

    exibidos, aviso = _cortar(grupos, estilos, "linha(s)")
    linhas = [[cab_item, cab_destino, "Valor", "Por quê"]]
    for grupo in exibidos:
        destino = grupo.get(campo)
        # O que não achou destino sai em vermelho e em negrito: numa folha
        # impressa é o único jeito de a linha problemática se distinguir das
        # outras duzentas.
        estilo_destino = estilos["corpo"] if destino else estilos["corpo_alerta"]
        status = conferencia.ROTULO_STATUS.get(grupo.get("status"), grupo.get("status") or "")
        motivo = grupo.get("motivo") or ""
        linhas.append([
            Paragraph(kit.texto(grupo.get("rotulo")), estilos["corpo"]),
            Paragraph(kit.texto(destino or "não preenchida"), estilo_destino),
            kit.moeda(_decimal(grupo.get("total"))),
            Paragraph(kit.texto(f"{status} — {motivo}" if motivo else status), estilos["corpo"]),
        ])

    tabela = kit.tabela(linhas, _COLS_DECISAO, ["LEFT", "LEFT", "RIGHT", "LEFT"])
    return [kit.secao(titulo, estilos), Spacer(1, 2 * mm), tabela] + aviso + [Spacer(1, 5 * mm)]


def _decisoes(dados: dict, estilos: dict) -> list:
    return (
        _tabela_decisao("COMO CADA FOLHA FOI DESTINADA", dados.get("decisoes_folhas"),
                        "Folha (relatório)", "Linha", "tipo", estilos)
        + _tabela_decisao("COMO CADA EVENTO FOI DESTINADO", dados.get("decisoes_rubricas"),
                          "Evento (relatório)", "Coluna", "coluna", estilos)
    )


def _tabela_contagem(titulo: str, mapa, estilos: dict) -> list:
    """Item + ocorrências. Lista vazia vira uma linha de 'Nenhuma' — o
    silêncio de uma seção ausente pareceria esquecimento, não ausência."""
    itens = [(str(nome), qtd) for nome, qtd in mapa.items()] if isinstance(mapa, dict) else []
    if not itens:
        return [Paragraph(f"<b>{kit.texto(titulo)}:</b> nada a registrar.", estilos["corpo"]),
                Spacer(1, 2 * mm)]

    exibidos, aviso = _cortar(itens, estilos, "item(ns)")
    linhas = [[titulo, "Ocorrências"]]
    linhas += [[Paragraph(kit.texto(nome), estilos["corpo"]), str(qtd)]
               for nome, qtd in exibidos]
    tabela = kit.tabela(linhas, [kit.LARGURA_UTIL - 30 * mm, 30 * mm], ["LEFT", "RIGHT"])
    return [tabela] + aviso + [Spacer(1, 4 * mm)]


def _tabela_estrutura(itens, estilos: dict) -> list:
    if not isinstance(itens, list) or not itens:
        return []
    exibidos, aviso = _cortar(itens, estilos, "item(ns)")
    linhas = [["Motivo", "Setor", "Linha", "Rubrica", "Valor"]]
    for item in exibidos:
        linhas.append([
            Paragraph(kit.texto(item.get("motivo")), estilos["corpo"]),
            Paragraph(kit.texto(item.get("setor")), estilos["corpo"]),
            Paragraph(kit.texto(item.get("tipo")), estilos["corpo"]),
            Paragraph(kit.texto(item.get("rubrica")), estilos["corpo"]),
            kit.moeda(_decimal(item.get("valor"))),
        ])
    larguras = [48 * mm, 32 * mm, 24 * mm, kit.LARGURA_UTIL - 134 * mm, 30 * mm]
    tabela = kit.tabela(linhas, larguras, ["LEFT", "LEFT", "LEFT", "LEFT", "RIGHT"])
    return [Paragraph("<b>Itens sem lugar na planilha</b>", estilos["corpo"]),
            Spacer(1, 1.5 * mm), tabela] + aviso + [Spacer(1, 4 * mm)]


def _pendencias(dados: dict, estilos: dict) -> list:
    tem = any(dados.get(chave) for chave in (
        "lotacoes_nao_mapeadas", "rubricas_sem_vinculo", "folhas_sem_vinculo",
        "pendencias_estrutura", "folhas_sugeridas", "rubricas_fora_escopo",
        "folhas_fora_escopo",
    ))
    if not tem:
        return [kit.secao("PENDÊNCIAS E OBSERVAÇÕES", estilos), Spacer(1, 2 * mm),
                Paragraph("Sem pendências acionáveis: tudo o que foi reconhecido "
                          "entrou na planilha ou foi marcado como fora de escopo.",
                          estilos["corpo"]), Spacer(1, 4 * mm)]

    elementos = [kit.secao("PENDÊNCIAS E OBSERVAÇÕES", estilos), Spacer(1, 2 * mm)]
    for titulo, chave in (
        ("Lotações não mapeadas", "lotacoes_nao_mapeadas"),
        ("Eventos sem coluna de destino", "rubricas_sem_vinculo"),
        ("Folhas sem linha de destino", "folhas_sem_vinculo"),
        ("Folhas somadas em outra linha (sugestão do sistema)", "folhas_sugeridas"),
        ("Eventos fora de escopo (ignorados de propósito)", "rubricas_fora_escopo"),
        ("Folhas fora de escopo (ignoradas de propósito)", "folhas_fora_escopo"),
    ):
        elementos += _tabela_contagem(titulo, dados.get(chave), estilos)
    elementos += _tabela_estrutura(dados.get("pendencias_estrutura"), estilos)
    return elementos


# ---------------------------------------------------------------------------
# Montagem
# ---------------------------------------------------------------------------

def nome_arquivo(output_nome: str | None = None) -> str:
    """`CONFERENCIA_<carimbo>.pdf`, com o carimbo da planilha que ele confere.

    O nome sai de um `\\d{8}_\\d{6}` extraido por expressao regular — nunca do
    texto recebido. Assim os dois arquivos da mesma operacao ficam lado a
    lado na pasta de downloads, e um nome de saida adulterado nao tem como
    virar caminho ou extensao no `Content-Disposition`.
    """
    achado = re.search(r"\d{8}_\d{6}(?:_[0-9a-f]{6})?", output_nome or "")
    carimbo = achado.group(0) if achado else datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"CONFERENCIA_{carimbo}.pdf"


def montar_pdf(dados: dict) -> bytes:
    """Desenha o retrato de uma operação. Não recalcula nada — ver o módulo."""
    estilos = kit.estilos()

    elementos = _cabecalho(dados, estilos)
    elementos += _reconciliacao(dados, estilos)
    elementos += _totais(dados, estilos)
    # Vira a página só se o rastro de auditoria não tiver onde começar.
    # Uma quebra incondicional deixaria meia folha em branco em todo
    # documento — num papel que vai para o processo, isso é desperdício.
    elementos.append(CondPageBreak(60 * mm))
    elementos += _decisoes(dados, estilos)
    elementos += _pendencias(dados, estilos)
    elementos.append(Paragraph(
        "Documento gerado a partir da mesma conferência que acompanha a planilha "
        "preenchida (aba CONFERÊNCIA_AUTOMAÇÃO). Traz apenas totais agregados: "
        "nenhum nome, matrícula ou CPF.",
        estilos["nota"],
    ))

    return kit.montar_documento(elementos, "Conferência da automação de retenções", _RODAPE)
