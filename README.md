# Automação de Retenções — Fundo Municipal de Saúde (SLMB)

Aplicação web em Flask que transforma o **relatório bruto de eventos** (Listagem
de Eventos, XLSX) em uma **planilha de Retenção preenchida** por lotação, rubrica
e tipo de folha — preservando fórmulas, estilos, mesclagens e dashboard do modelo,
e gerando uma aba de conferência completa.

> A ferramenta não preenche às cegas: ela mostra o que **identificou**, o que
> **preencheu** e o que ficou **pendente**. Confira sempre antes de enviar.

---

## 1. Instalação

Windows (PowerShell):

```powershell
cd automacao_retencao
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

macOS / Linux:

```bash
cd automacao_retencao
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Requer **Python 3.11+**.

## 2. Como rodar

```bash
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
- Setores, rubricas e abas **sem limite prático** — um por linha, com colar de
  lista em lote, reordenação e remoção.
- Parâmetros do bloco: tipos de folha, linha `TOTAL` por setor, coluna
  `TOTAL DO EVENTO`, linha `TOTAL GERAL` e rodapé com apelido.
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
5. Na tela de **mapeamento**, confirme os **dois vínculos** (com sugestão
   automática e menus suspensos):
   - **Eixo 1 — Lotação → Setor**: cada lotação do relatório para um setor da planilha.
   - **Eixo 2 — Evento/Rubrica → Coluna**: cada evento do relatório para uma coluna
     do modelo, ou marque **“Fora de escopo”** para não preencher.
   Escolha a **aba de destino** e, se quiser, marque para **aprender/salvar os
   vínculos** (ficam prontos para os próximos meses).
6. Clique em **Processar e preencher**.
7. Na tela final, confira a **reconciliação** (`total lido = preenchido + fora de
   escopo + sem vínculo + setor não mapeado`, batendo ao centavo), os totais e as
   **pendências**, e **baixe** o XLSX.

## 4. Observações importantes

- Use preferencialmente **XLSX**, nunca PDF.
- **Confira as pendências** antes de enviar o arquivo final. Rubricas/lotações não
  mapeadas **não são preenchidas** — aparecem listadas na tela e na aba
  `CONFERÊNCIA_AUTOMAÇÃO`.
- **Abra o arquivo final no Excel** para que as fórmulas e o dashboard recalculem
  automaticamente (o openpyxl não recalcula; o app marca o workbook para
  recálculo na abertura).
- Apenas as linhas **Mensal** e **13º salário** são preenchidas. A linha **TOTAL**
  nunca é tocada (mantém a fórmula).

---

## Arquitetura

```
automacao_retencao/
├── app.py                 # Camada web (Flask) — apenas orquestração
├── config/
│   ├── mapeamento_rubricas.json        # regras de rubrica + "fora de escopo" (compartilhado)
│   ├── perfis.json                     # registro de secretarias (perfis)
│   └── perfis/<secretaria>/            # dados isolados por secretaria (aprendidos)
│       ├── mapeamento_lotacoes.json    #   lotação → setor
│       └── vinculo_rubrica_coluna.json #   evento/rubrica → coluna
├── modelos/
│   └── perfis/<secretaria>/
│       ├── molde_padrao.xlsx      # molde fixo versionado (blank)
│       └── molde_estrutura.json   # o desenho do molde (reabrível no construtor)
├── services/              # Regra de negócio, isolada do Flask
│   ├── parser_listagem.py # Lê o relatório de origem (merged-aware, anti-ruído)
│   ├── normalizador.py    # Texto, folha e rubrica normalizados
│   ├── mapeador.py        # Dois eixos de vínculo + aprendizado + pendências
│   ├── molde.py           # Construtor: spec validada -> xlsx -> reverificado
│   ├── perfis.py          # Perfis por secretaria: molde fixo + vínculos isolados
│   ├── preenchimento.py   # Localiza blocos/colunas e preenche o modelo
│   ├── conferencia.py     # Agregação, reconciliação e aba de conferência
│   └── utils.py           # Pastas, logs, nomes e sessão em disco
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
- **Dois eixos de vínculo com aprendizado**: lotação→setor e evento/rubrica→coluna,
  ambos com sugestão automática, menu suspenso e persistência em JSON.
- **Interface dark-first com modo claro** em um clique (tema persistido no navegador).
- **Colunas do modelo = fonte da verdade**: nunca se escreve numa coluna que não
  exista; rubricas sem coluna viram pendência ou “fora de escopo”, nunca erro mudo.
- **Reconciliação exata**: `total lido = preenchido + fora de escopo + sem vínculo
  + setor não mapeado`. Se não bater ao centavo, a tela avisa.
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
por setor, por rubrica e por tipo, e todas as pendências.
