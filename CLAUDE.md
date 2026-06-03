# CLAUDE.md

Orientações para o Claude Code (e para humanos) trabalharem neste projeto.

## O que é

Projeto de **um único script** que monitora se a Anatel já publicou a competência
mensal nova de **acessos móveis (SMP – Serviço Móvel Pessoal)** no Portal de Dados
Abertos. Roda 1x/dia (manualmente) no Windows com o Python do Anaconda e diz se a
competência-alvo já apareceu, com totais por operadora e por tecnologia.

Arquivo principal: [monitor_anatel_smp.py](monitor_anatel_smp.py).

## Ambiente

- **SO:** Windows 11. **Python:** Anaconda 3.11 em `C:\ProgramData\Anaconda3\python.exe`.
- **Dependências:** `requests` + `pandas` (já no Anaconda) + stdlib. Sem Selenium.
- **Rodar:**
  ```powershell
  C:\ProgramData\Anaconda3\python.exe c:\dev\Acessos_Moveis\monitor_anatel_smp.py --verbose
  ```

## Como executar

- Sem flags: silencioso/idempotente — só destaca quando a competência **muda**.
- `--verbose`: mostra qual camada respondeu, colunas do CSV, competência detectada e
  força a saída completa.
- Saídas esperadas: `✅ DADO NOVO DISPONÍVEL: <mês>/<ano>...` ou
  `⏳ AINDA NÃO: a competência mais recente é <mês/ano>...`.

## Arquivos gerados

- `anatel_smp_estado.json` — última competência vista. **Versionado de propósito**:
  é assim que o GitHub Actions persiste o estado entre execuções (e decide se houve
  novidade para mandar e-mail). Não editar à mão.
- `anatel_smp_log.txt` — histórico append-only de cada checagem (ignorado pelo git).
- `email_config.json` — credenciais SMTP. **Nunca versionar** (no Actions vem do
  secret `EMAIL_CONFIG_JSON`); está no `.gitignore`.
- `dados_anatel/`, `__pycache__/` — baixados/cache; ignorados.

## Automação (GitHub Actions) e e-mail

- Workflow: [.github/workflows/monitor.yml](.github/workflows/monitor.yml) — cron a
  cada 30 min (`13,43 * * * *`) + `workflow_dispatch`. Escreve `email_config.json` a
  partir do secret, roda o script e commita `anatel_smp_estado.json` de volta.
- Padrão copiado de `github.com/danxhp/cvm-fatos-relevantes`.
- **Secrets necessários no repositório** (Settings → Secrets and variables → Actions):
  - `EMAIL_CONFIG_JSON` — JSON com `enabled:true`, `smtp_*` (Gmail: senha de APP,
    porta 465 ou 587) e `to_addrs`.
  - `CHAVE_DADOS_GOV` — chave gratuita do dados.gov.br. **Necessária nos runners do
    GitHub (EUA)**: sem chave a API responde HTTP 401.
- E-mail (`enviar_email` + `_render_email_html`): disparado **só quando a competência
  muda** (nova publicação) e nunca na primeira execução (evita alerta no marco zero).
  Idempotente — não reenvia se nada mudou.

## Fonte de dados (confirmado)

- Dataset: **"Acessos - Telefonia Móvel"** (acessos do SMP).
- slug/id no dados.gov.br: **`acessos-autorizadas-smp`**.
  Página: https://dados.gov.br/dataset/acessos-autorizadas-smp
- Painel (JS, só referência em avisos): https://informacoes.anatel.gov.br/paineis/acessos/telefonia-movel
- **Encoding dos arquivos da Anatel: latin-1 (ISO-8859-1) com separador `;`** — nunca
  assumir utf-8/vírgula.

## Decisões de arquitetura (por que é assim)

1. **Descoberta de recursos em CAMADAS** (`descobrir_recursos`), porque a API do
   dados.gov.br passou a exigir chave (retorna **HTTP 401** sem ela):
   - A) CKAN legado `package_show` → B) API nova `/dados/api/publico/...`
     (manda header `chave-api-dados-abertos` se a env var `CHAVE_DADOS_GOV` existir)
     → C) fallback raspando a página HTML por links `.csv`/`.zip`.
   - Usa a primeira camada que retornar recursos.
2. **Leitura em chunks** (`competencia_mais_recente`): o consolidado de SMP é gigante
   (série desde 2007). 1º passe acha o `max(Ano, Mês)`; 2º passe carrega só as linhas
   da competência mais recente. Preferir sempre o CSV do ano corrente, se existir.
3. **Cabeçalho confirmado antes de assumir** (`detectar_formato`/`_mapear_colunas`):
   nomes de coluna são casados com normalização de acento/caixa (`Ano`, `Mês`,
   `Acessos`, `Grupo Econômico`, `Tecnologia`).
4. **Nunca afirmar "o dado não saiu"** quando a API pode estar defasada vs. o painel:
   nesses casos o script avisa e sugere conferir o painel manualmente.
5. **Saída em UTF-8 forçada** (`sys.stdout.reconfigure`) porque o console do Windows
   é cp1252 e quebra com emojis/acentos.

## Configuração

Constantes no topo de [monitor_anatel_smp.py](monitor_anatel_smp.py):
`COMPETENCIA_ALVO` (ex.: `"2026-04"`), `DATASET_SLUG`, `PASTA_DOWNLOAD`, `TIMEOUT`,
`MAX_TENTATIVAS`. A chave da API vem da env var `CHAVE_DADOS_GOV`
(`setx CHAVE_DADOS_GOV "sua-chave"`), nunca hardcoded.

## Convenções para alterações

- Código e mensagens ao usuário **em português**.
- Manter o tratamento de rede com retry + backoff (`http_get`) e mensagens de erro
  claras; a Anatel/dados.gov.br derruba conexão e muda de API com frequência.
- Não introduzir Selenium nem renderização de JS — ficar em API/CSV.
- Ao mudar parsing, validar com um CSV sintético latin-1/`;` antes de confiar
  (foi assim que a lógica de competência/agregados foi verificada).
