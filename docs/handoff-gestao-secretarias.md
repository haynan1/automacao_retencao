# Handoff — Tela de gestão de secretarias (perfis)

**Objetivo:** construir uma interface para **criar, renomear, editar termos de detecção,
definir padrão e remover** secretarias (perfis), sem editar `config/perfis.json` na mão.

O motor de perfis **já existe** (`automacao_retencao/services/perfis.py`). Esta tarefa é
quase toda camada web (rotas Flask + 1 template) + 2 funções novas no serviço.

Contexto do projeto: app Flask em `automacao_retencao/`, roda em `127.0.0.1` (single-user,
sem auth). Regra de negócio isolada em `services/`. Templates Jinja2 + Bootstrap 5 (dark-first,
com modo claro via `data-bs-theme`). Estilos utilitários em `static/style.css`
(`.panel`, `.panel-pad`, `.section-title`, `.form-select`, `.form-control`, `.btn-primary`,
`.chip`, `.table-wrap`, `.opt-modelo`, `.alert-soft-danger/warn/success`).

---

## O que é um "perfil" hoje

Registro em `config/perfis.json`:
```json
{
  "padrao": "saude",
  "perfis": [
    { "slug": "saude", "nome": "Secretaria da Saúde (FMS)", "deteccao": ["FUNDO MUNICIPAL DE SAUDE", "FMS"] }
  ]
}
```
- `slug` — id interno; **imutável** (é o nome das pastas `config/perfis/<slug>/` e
  `modelos/perfis/<slug>/`). Renomear a secretaria muda só `nome`, nunca `slug`.
- `nome` — texto exibido.
- `deteccao` — termos usados para auto-sugerir a secretaria pelo cabeçalho da Listagem.

Dados por perfil (não mexer, já funcionam):
- `config/perfis/<slug>/mapeamento_lotacoes.json`
- `config/perfis/<slug>/vinculo_rubrica_coluna.json`
- `modelos/perfis/<slug>/molde_padrao.xlsx` (+ `.json` de metadados)

## API existente em `services/perfis.py` (reutilizar)

```python
listar_perfis() -> list[dict]                 # [{slug, nome, deteccao}, ...]
slug_padrao() -> str
perfil_valido(slug) -> bool
info_perfil(slug) -> dict | None
nome_perfil(slug) -> str
gerar_slug(nome) -> str                        # "Educação" -> "educacao"
registrar_perfil(nome, deteccao=None, slug=None) -> slug   # cria OU atualiza (se slug existe)
detectar_perfil(banners) -> slug | None
existe_molde(slug) -> bool
info_molde(slug) -> dict | None
# privados: _ler_registro(), _salvar_registro(reg)
# constantes: PERFIS_CONFIG_DIR, PERFIS_MODELOS_DIR, REGISTRO
```
`registrar_perfil` já cobre **criar** e **editar (nome/deteccao)** — se `slug` for passado e
já existir, atualiza `nome` e `deteccao` mantendo o slug.

---

## Passo 1 — 2 funções novas em `services/perfis.py`

```python
def definir_padrao(slug: str) -> None:
    """Define o perfil padrão. Ignora se o slug não existir."""
    if not perfil_valido(slug):
        return
    reg = _ler_registro()
    reg["padrao"] = slug
    _salvar_registro(reg)


def remover_perfil(slug: str) -> bool:
    """Remove o perfil do registro. NÃO apaga os dados em disco (PII/histórico).

    Bloqueia se for o padrão ou o último perfil. Retorna True se removeu.
    """
    reg = _ler_registro()
    if slug == reg.get("padrao") or len(reg.get("perfis", [])) <= 1:
        return False
    antes = len(reg["perfis"])
    reg["perfis"] = [p for p in reg["perfis"] if p["slug"] != slug]
    if len(reg["perfis"]) == antes:
        return False
    _salvar_registro(reg)
    log.info("Perfil '%s' removido do registro (dados preservados em disco).", slug)
    return True
```
> Decisão de segurança: **não deletar** `config/perfis/<slug>/` nem `modelos/perfis/<slug>/`
> (contêm dados aprendidos e molde). Remover é só "desregistrar". Se quiser expurgo real,
> fazer numa ação separada com confirmação dupla — fora do escopo deste handoff.

## Passo 2 — Rotas em `app.py`

`perfis` já está importado. Adicionar:

```python
@app.route("/secretarias")
def secretarias():
    return render_template(
        "secretarias.html",
        perfis=perfis.listar_perfis(),
        padrao=perfis.slug_padrao(),
        moldes={p["slug"]: perfis.info_molde(p["slug"]) for p in perfis.listar_perfis()},
    )

@app.route("/secretarias/salvar", methods=["POST"])
def secretarias_salvar():
    nome = (request.form.get("nome") or "").strip()
    slug = (request.form.get("slug") or "").strip()   # vazio = criar; preenchido = editar
    deteccao = [t.strip() for t in re.split(r"[,\n;]", request.form.get("deteccao") or "") if t.strip()]
    if not nome:
        return render_template("erro.html", titulo="Nome obrigatório",
                               mensagem="Informe o nome da secretaria."), 400
    # slug só é usado se já existir (edição); nunca aceitar slug arbitrário para criar.
    slug_final = slug if (slug and perfis.perfil_valido(slug)) else None
    perfis.registrar_perfil(nome, deteccao=deteccao, slug=slug_final)
    return redirect(url_for("secretarias"))

@app.route("/secretarias/<slug>/padrao", methods=["POST"])
def secretarias_padrao(slug):
    perfis.definir_padrao(slug)
    return redirect(url_for("secretarias"))

@app.route("/secretarias/<slug>/remover", methods=["POST"])
def secretarias_remover(slug):
    if not perfis.remover_perfil(slug):
        return render_template("erro.html", titulo="Não foi possível remover",
                               mensagem="Não dá para remover o perfil padrão nem o último."), 400
    return redirect(url_for("secretarias"))
```
`re` já é importado? Se não, adicionar `import re` no topo de `app.py`.

> **Restrição de segurança:** ao **criar**, gerar o slug via `gerar_slug` dentro de
> `registrar_perfil` (já é o comportamento). Nunca deixar o usuário definir o `slug` de um
> perfil novo (evita path/colisão). Na **edição**, o slug vem de um campo hidden e é validado
> com `perfil_valido`.

## Passo 3 — Template `templates/secretarias.html`

Seguir o padrão dos outros templates (`{% extends "base.html" %}`). Sem stepper (não faz
parte do fluxo de processamento). Estrutura:

- Título + botão "Nova secretaria" (abre o mesmo form com campos vazios).
- Tabela/cards das secretarias: `nome`, `slug` (badge), chips de `deteccao`, indicador de
  molde fixo (`moldes[slug]`), e ações:
  - **Editar** (preenche o form com `nome`, `slug` em hidden, `deteccao` juntos por vírgula).
  - **Definir padrão** (POST `/secretarias/<slug>/padrao`) — desabilitado se já é o padrão.
  - **Remover** (POST `/secretarias/<slug>/remover`, com `confirm()` no submit) — desabilitado
    se for o padrão ou o único.
- Form (criar/editar): inputs `nome` (`form-control`), `deteccao` (`textarea`/`form-control`,
  ajuda: "termos separados por vírgula"), `slug` (hidden). `action="{{ url_for('secretarias_salvar') }}"`.

Reaproveitar classes: `.panel-pad`, `.section-title`, `.table-wrap`, `.chip`, `.btn-primary`,
`.btn-outline-light`. Manter dark-first + funcionar no modo claro (as variáveis do CSS já
cobrem os dois temas).

## Passo 4 — Link de acesso em `templates/base.html`

No navbar (dentro do `<div class="d-flex align-items-center gap-2">`, junto ao botão de tema e
"Novo processamento"), adicionar:
```html
<a href="{{ url_for('secretarias') }}" class="btn btn-sm btn-outline-light" title="Secretarias">
  <i class="bi bi-buildings"></i>
</a>
```

## Passo 5 — Testes em `tests/test_perfis.py`

Usar o fixture `perfis_tmp` já existente (monkeypatcha os diretórios). Adicionar:
- `definir_padrao` muda o padrão; `slug_padrao()` reflete.
- `remover_perfil` bloqueia remover o padrão e o último; remove um não-padrão com sucesso.
- Editar via `registrar_perfil(nome_novo, slug=slug)` mantém o slug e troca o nome
  (já coberto parcialmente por `test_registrar_e_editar_nome`).

Rodar: `python -m pytest` (deve seguir verde, 17 casos atuais + os novos).

---

## Critérios de aceite

1. `/secretarias` lista as secretarias com nome, slug, termos de detecção e se têm molde fixo.
2. Criar uma secretaria pela UI → aparece no seletor da tela inicial (`index`).
3. Renomear troca só o `nome`; **slug, molde e vínculos permanecem** (mesma pasta).
4. Definir padrão muda o selecionado por default no `index`.
5. Remover desregistra sem apagar dados; bloqueado para padrão/último.
6. `slug` de um perfil **nunca** muda após criado.
7. `python -m pytest` verde.

## Não fazer
- Não deletar diretórios de dados de perfil ao remover (só desregistrar).
- Não permitir o usuário definir/alterar `slug` manualmente.
- Não quebrar o fluxo atual (`index → analisar → preview → mapeamento → processar → resultado`).
- Não commitar nada de `arquivos_base_temp/`, `uploads/`, `outputs/` (têm PII; já no `.gitignore`).
