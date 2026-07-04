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

## 3. Como usar

1. **Exporte** a *Listagem de Eventos* do sistema em **XLSX** (não PDF).
2. Na tela inicial, envie a *Listagem de Eventos* e a planilha *RETENÇÃO*.
3. Clique em **Analisar arquivos**.
4. Na **pré-visualização**, confira competência, lançamentos, lotações e rubricas
   detectadas — e os avisos de itens não mapeados.
5. Na tela de **mapeamento**, associe cada lotação a um setor da planilha
   (há sugestão automática). Escolha a **aba de destino** (sugerida pela
   competência) e, se quiser, marque para **salvar o mapeamento**.
6. Clique em **Processar e preencher**.
7. Na tela final, revise os totais e as **pendências**, e **baixe** o XLSX.

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
├── config/                # Mapeamentos editáveis (JSON, sem código)
│   ├── mapeamento_lotacoes.json
│   └── mapeamento_rubricas.json
├── services/              # Regra de negócio, isolada do Flask
│   ├── parser_listagem.py # Lê o relatório de origem (robusto a colunas)
│   ├── normalizador.py    # Texto, folha e rubrica normalizados
│   ├── mapeador.py        # Lotação→setor, descrição→rubrica, pendências
│   ├── preenchimento.py   # Localiza blocos/colunas e preenche o modelo
│   ├── conferencia.py     # Agregação, totais e aba de conferência
│   └── utils.py           # Pastas, logs, nomes e sessão em disco
├── templates/             # Jinja2 + Bootstrap 5 (dark-first)
├── static/style.css
├── uploads/  outputs/  logs/
```

### Decisões de projeto

- **Detecção dinâmica**: setores, colunas de rubrica e linhas de tipo são
  localizados pela *estrutura* da planilha, não por coordenadas fixas — o layout
  pode mudar de mês para mês.
- **Colunas por cabeçalho**: no relatório de origem, as posições das colunas são
  descobertas lendo os rótulos (`Mat.`, `Descrição`, `Valor`, …), não pela letra.
- **`Decimal`** para todo valor monetário; totais são calculados na aplicação,
  sem depender do recálculo do Excel.
- **Mapeamentos em JSON**: manutenção sem tocar no código.
- **Segurança**: `secure_filename`, limite de 30 MB, apenas `.xlsx`, uploads e
  saídas restritos às suas pastas, sem execução de macros.

### Conferência

Cada arquivo gerado recebe a aba **`CONFERÊNCIA_AUTOMAÇÃO`** com data/hora,
arquivos de origem e modelo, competência, total lido/preenchido/pendente, totais
por setor, por rubrica e por tipo, e todas as pendências.
