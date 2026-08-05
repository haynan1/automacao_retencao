# Automação de Retenções — Fundo Municipal de Saúde (SLMB)

Aplicação web em Flask que transforma o **relatório bruto de eventos** (Listagem
de Eventos, XLSX) em uma **planilha de Retenção preenchida** por lotação, rubrica
e tipo de folha — preservando fórmulas, estilos, mesclagens e dashboard do modelo,
e gerando uma aba de conferência completa.

> A ferramenta não preenche às cegas: ela mostra o que **identificou**, o que
> **preencheu** e o que ficou **pendente**. Confira sempre antes de enviar.

---

## 1. Como rodar (Windows)

Dê um duplo clique em **`iniciar.bat`**. Ele cria o ambiente virtual se não
existir, instala as dependências dentro dele, sobe o servidor e abre o
navegador. Nas próximas vezes ele reinstala apenas se o `requirements.txt`
tiver mudado — então a partida é imediata.

Requer **Python 3.11+** instalado (com *Add python.exe to PATH* marcado).

## 2. Instalação manual (ou macOS / Linux)

```powershell
python -m venv venv
venv\Scripts\activate          # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
python app.py
```

Acesse: **http://127.0.0.1:5000**

O servidor sobe em `127.0.0.1` com **debug desligado**. Para desenvolvimento (auto-reload
e tracebacks), use `FLASK_DEBUG=1 python app.py`. Não exponha o app para fora do localhost
com debug ligado.

### Testes

```bash
pip install -r requirements-dev.txt
python -m pytest
```

## 3. Como usar

### 3.1. O molde (uma vez por secretaria)

O molde é a planilha de destino: os blocos de setor, as colunas de rubrica e as
linhas de tipo. Ele **não precisa mais ser editado no Excel** — em
**Secretarias → Construir molde** (`/molde/<secretaria>`) a estrutura é
desenhada na tela:

- **Folha em branco** ou **partir do molde atual** (a estrutura do `.xlsx`
  existente é lida de volta e vira formulário).
- Setores, rubricas, abas e **linhas de tipo de folha** sem limite prático —
  um por linha, com colar de lista em lote, reordenação e remoção.
- **Linhas de tipo livres**: `Mensal`, `13º salário`, `Férias`, `Rescisão`,
  `Complementar` ou o nome que a secretaria usar. É o que decide se férias
  ganham uma linha própria ou continuam somando na mensal.
- Parâmetros do bloco: linha `TOTAL` por setor, coluna `TOTAL DO EVENTO`,
  linha `TOTAL GERAL` e rodapé com apelido.
- **Prévia ao vivo** com o mesmo layout do arquivo, célula por célula.

Nada é salvo às cegas: a cada alteração o sistema **gera o `.xlsx` e o relê com
o mesmo motor que preenche a planilha todo mês** (`localizar_blocos_setores` e
`listar_colunas_modelo`). Se o motor não enxergar exatamente os setores, as
colunas e os tipos desenhados, a gravação é recusada e a divergência aparece na
tela. O molde anterior vira cópia de segurança timestampada.

O construtor recusa por escrito o que o motor descartaria em silêncio: coluna
repetida a menos de acento ou caixa, rótulo iniciado por `TOTAL`, aba com
`CONFER` no nome, texto que o Excel leria como fórmula (`=`, `+`, `@`).

### 3.2. O processamento (todo mês)

1. **Exporte** a *Listagem de Eventos* do sistema em **XLSX** (não PDF).
2. Na tela inicial, envie a *Listagem de Eventos* e a planilha *RETENÇÃO*.
3. Clique em **Analisar arquivos**.
4. Na **pré-visualização**, confira competência, lançamentos, lotações e rubricas
   detectadas — e os avisos de itens não mapeados.
5. Na tela de **mapeamento**, confirme os **três vínculos** (com sugestão
   automática, menus suspensos e o **motivo de cada decisão à vista**):
   - **Eixo 1 — Lotação → Setor**: cada lotação do relatório para um bloco da planilha.
   - **Eixo 2 — Evento → Coluna**: cada **evento** do relatório para uma coluna
     do modelo, ou **“Fora de escopo”** para não preencher. É por evento, não
     por rubrica: `INSS` e `INSS do 13º salário` aparecem separados e podem ir
     para colunas diferentes.
   - **Eixo 3 — Folha → Linha**: cada folha do relatório para uma linha do bloco
     (`Mensal`, `13º salário`, `Férias`…), ou **“Não preencher”**.
   Escolha a **aba de destino** e, se quiser, marque para **aprender/salvar os
   vínculos** (ficam prontos para os próximos meses).
6. Clique em **Processar e preencher**.
7. Na tela final, confira a **reconciliação** (batendo ao centavo), os totais,
   o registro de **como cada folha foi destinada** e as **pendências** — e
   **baixe** o XLSX.

### 3.3. Relatório consolidado (quando quiser)

Na tela **Histórico**, marque as operações desejadas e escolha **Excel** ou
**PDF**. Não é uma lista do que foi processado — é o **compilado**: quanto foi
retido no período inteiro, somado por dimensão.

- **Resumo geral** — total lido, preenchido, fora de escopo, sem lugar,
  pendente e a **diferença** (que deve ser zero).
- **Por secretaria**, **por setor**, **por evento do relatório**, **por
  rubrica (coluna)** e **por linha de tipo de folha** — cada seção com a
  lista completa, seu total e um gráfico.
- **Operações compiladas** — a procedência, para que qualquer número do
  relatório possa ser rastreado até o processamento que o gerou.

O `.xlsx` traz os totais em **fórmula** e uma segunda aba com os gráficos
nativos; o `.pdf` traz o mesmo compilado paginado, pronto para imprimir ou
anexar. Os dois leem a **mesma compilação** — não têm como divergir.

Seções cujo detalhamento não existe em alguma operação (registro anterior à
criação daquela dimensão) trazem, por escrito, quanto não está representado.
Os valores são **números**, nunca texto: a planilha soma, ordena e recalcula.
O relatório sai do histórico local, que **não guarda dados pessoais**.

## 4. Observações importantes

- Use preferencialmente **XLSX**, nunca PDF.
- **Confira as pendências** antes de enviar o arquivo final. Eventos, folhas e
  lotações sem vínculo **não são preenchidos** — aparecem listados na tela e na
  aba `CONFERÊNCIA_AUTOMAÇÃO`.
- **Abra o arquivo final no Excel** para que as fórmulas e o dashboard recalculem
  automaticamente (o openpyxl não recalcula; o app marca o workbook para
  recálculo na abertura).
- São preenchidas as **linhas de tipo que existirem no molde**. A linha **TOTAL**
  nunca é tocada (mantém a fórmula).
- Folha sem linha correspondente é **sugerida** para `Mensal` e marcada na tela;
  folha que o sistema não reconhece **não recebe destino automático** — segura o
  valor e pergunta, em vez de preencher no escuro.

---

## Arquitetura

```
rentencoes/
├── iniciar.bat            # Sobe o sistema (venv + dependências + servidor)
├── app.py                 # Camada web (Flask) — apenas orquestração
├── config/
│   ├── mapeamento_rubricas.json        # regras de rubrica + "fora de escopo" (compartilhado)
│   ├── perfis.json                     # registro de secretarias (perfis)
│   └── perfis/<secretaria>/            # dados isolados por secretaria (aprendidos)
│       ├── mapeamento_lotacoes.json    #   eixo 1: lotação → setor
│       ├── vinculo_rubrica_coluna.json #   eixo 2: evento → coluna
│       └── vinculo_folha_tipo.json     #   eixo 3: folha → linha de tipo
├── modelos/
│   └── perfis/<secretaria>/
│       ├── molde_padrao.xlsx      # molde fixo versionado (blank)
│       └── molde_estrutura.json   # o desenho do molde (reabrível no construtor)
├── services/              # Regra de negócio, isolada do Flask
│   ├── parser_listagem.py # Lê o relatório de origem (merged-aware, anti-ruído)
│   ├── normalizador.py    # Texto, família de folha e rubrica normalizados
│   ├── mapeador.py        # Três eixos de vínculo + aprendizado + pendências
│   ├── molde.py           # Construtor: spec validada -> xlsx -> reverificado
│   ├── perfis.py          # Perfis por secretaria: molde fixo + vínculos isolados
│   ├── preenchimento.py   # Localiza blocos/colunas e preenche o modelo
│   ├── conferencia.py     # Agregação, reconciliação e aba de conferência
│   ├── relatorio.py       # Compila o histórico e desenha o .xlsx
│   ├── relatorio_pdf.py   # O MESMO compilado, desenhado em .pdf
│   ├── estilo_xlsx.py     # Vocabulário visual das planilhas que o app gera
│   ├── historico.py       # Registro local das operações (JSONL, sem PII)
│   └── utils.py           # Pastas, logs, JSON atômico e sessão em disco
├── templates/             # Jinja2 + Bootstrap 5 (dark-first)
├── static/style.css
├── uploads/  outputs/  logs/
```

### Decisões de projeto

- **Leitura *merged-aware***: o relatório real usa células mescladas e a coluna
  Descrição vem deslocada em relação ao cabeçalho. Cada campo é lido pelo segmento
  mesclado com maior **sobreposição** à faixa do cabeçalho — imune a deslocamentos.
  Banners/rodapés/resumo são descartados por estrutura (o mesmo segmento cobre
  Descrição e Folha ⇒ não é linha tabular).
- **Detecção dinâmica**: setores, colunas de rubrica e linhas de tipo são
  localizados pela *estrutura* da planilha, não por coordenadas fixas.
- **Perfis por secretaria**: o motor é 100% reaproveitado entre secretarias; o que
  muda é só o *dado* de cada uma (molde fixo + vínculos aprendidos), isolado em
  `config/perfis/<secretaria>/` e `modelos/perfis/<secretaria>/`. A secretaria é
  escolhida na tela inicial e **auto-sugerida** pelo cabeçalho da Listagem.
- **Molde fixo versionado**: cada secretaria pode ter um molde padrão no repositório
  (blank, sem dados pessoais) — quem clona já processa enviando só a Listagem.
- **Molde como dado, não como arquivo**: a estrutura vive numa *spec* validada
  (`molde_estrutura.json`) e o `.xlsx` é um artefato gerado. Isso tira o Excel do
  caminho crítico e transforma erro silencioso (coluna repetida, rótulo reservado)
  em recusa explícita. O layout é descrito **uma única vez**, em
  `molde._linhas_layout` — o gerador e a prévia da tela consomem a mesma sequência,
  então a prévia não é uma aproximação da planilha: é a planilha.
- **Verificação de mão dupla**: todo molde gerado é relido pelo motor de
  preenchimento antes de ser aceito. O construtor não confia na própria geração.
- **Três eixos de vínculo com aprendizado**: a planilha é uma grade, e todo
  lançamento precisa de três coordenadas — bloco (lotação→setor), coluna
  (evento→coluna) e linha (folha→tipo). Os três com sugestão automática, menu
  suspenso, motivo explícito e persistência em JSON. Enquanto o terceiro eixo
  não existia, a linha era escolhida por uma regra fixa que mandava tudo que
  não fosse 13º para `Mensal`: férias, rescisão e complementar somavam no mesmo
  lugar e não havia onde discordar.
- **Vínculo por evento, com a rubrica como padrão herdado**: `INSS` e `INSS do
  13º salário` batem na mesma regra `contém: INSS`. Agrupar por rubrica os
  fundia num item só — impossível separar. Agora cada evento é um grupo e a
  regra vira apenas o padrão: quem quiser separa, quem não quiser não faz nada.
- **Toda decisão carrega um motivo em texto** (“regra INSS (contém INSS)”,
  “vínculo salvo deste evento”, “Férias não tem linha própria — somando em
  Mensal”), na tela e na aba de conferência. Poder discordar do sistema começa
  por conseguir ler o que ele fez.
- **Linhas de tipo abertas**: o motor lê da planilha os rótulos que existem, em
  vez de carregar uma lista fechada. O bloco é delimitado por *estrutura* (a
  linha `Tipo` sob o nome do setor, o `TOTAL` que o fecha), não por vocabulário.
- **Interface dark-first com modo claro** em um clique (tema persistido no navegador).
- **Colunas do modelo = fonte da verdade**: nunca se escreve numa coluna que não
  exista; rubricas sem coluna viram pendência ou “fora de escopo”, nunca erro mudo.
- **Uma compilação, dois formatos**: `relatorio.compilar()` faz a conta uma vez
  e devolve só números; o `.xlsx` e o `.pdf` apenas desenham. Se os dois
  pudessem divergir num centavo, um deles estaria mentindo e ninguém saberia
  qual — um teste confere valor por valor entre os formatos.
- **Reconciliação exata**: `total lido = preenchido + fora de escopo (rubrica) +
  fora de escopo (folha) + evento sem coluna + folha sem linha + setor não
  mapeado`. Cada balde é um motivo diferente para o dinheiro não ter entrado na
  planilha, e a tela mostra todos separados — “ignorei de propósito” e “não soube
  onde pôr” não podem parecer a mesma coisa. Se não bater ao centavo, a tela avisa.
- **`Decimal`** para todo valor monetário; totais calculados na aplicação.
- **Mapeamentos em JSON**: manutenção sem tocar no código.
- **Segurança**: `secure_filename`, limite de 30 MB, apenas `.xlsx`, uploads e
  saídas restritos às suas pastas, sem execução de macros.

### Secretarias (perfis)

Cada secretaria é um **perfil** com molde fixo e vínculos aprendidos isolados. O registro
fica em `config/perfis.json`:

```json
{ "slug": "saude", "nome": "Secretaria da Saúde (FMS)", "deteccao": ["FUNDO MUNICIPAL DE SAUDE", "FMS"] }
```

- **`slug`** — id interno (nome da pasta de dados). Não altere depois de criado.
- **`nome`** — texto exibido no seletor. Edite à vontade.
- **`deteccao`** — termos que auto-sugerem a secretaria pelo cabeçalho da Listagem.

Para **adicionar** uma secretaria: crie a entrada pela tela **Secretarias** (ou use
`services.perfis.registrar_perfil(nome, deteccao, slug)`) e defina o molde fixo dela
pelo **construtor de molde** — ou enviando um `.xlsx` na tela inicial com
“Definir como molde fixo”. Os dados ficam em `config/perfis/<slug>/` e
`modelos/perfis/<slug>/`.

### Conferência

Cada arquivo gerado recebe a aba **`CONFERÊNCIA_AUTOMAÇÃO`** com data/hora,
arquivos de origem e modelo, competência, total lido/preenchido/pendente, totais
por setor, por rubrica e por linha de tipo, e todas as pendências.

Ela traz também a seção **“COMO O SISTEMA DECIDIU”**: cada folha → linha e cada
evento → coluna, com valor e **motivo por extenso**. Quem abrir a planilha daqui
a seis meses consegue reconstruir por que cada valor foi parar naquela célula,
sem ter o app na frente.
