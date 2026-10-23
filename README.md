# monitor-precos — demo de monitoramento de preços

Script em Python que **raspa o catálogo de um concorrente, guarda um histórico diário
e compara com a sua tabela de preços**, entregando uma planilha Excel pronta para
análise e uma página HTML com KPIs e gráfico.

Demo de portfólio de **Marcos Galani** (automação de processos). A fonte raspada é o
[books.toscrape.com](https://books.toscrape.com), um site público criado justamente
para treino de web scraping; os "meus preços" são fictícios.

## O que o script faz

1. Baixa as 3 primeiras páginas do catálogo (60 produtos) com delay de 1 s entre
   páginas, User-Agent identificado e 3 tentativas com backoff em caso de erro.
2. Extrai título, preço (convertido para número), avaliação em estrelas (1–5),
   disponibilidade e URL do produto.
3. Grava um snapshot em `historico/precos_AAAA-MM-DD.csv` (uma linha por produto, com a data).
4. Cruza com `dados_exemplo/meus_precos.csv` pelo título e calcula diferença em moeda
   e em %, classificando cada item em **estou mais caro / estou mais barato / igual**.
5. Gera:
   - `saida/comparativo_precos.xlsx` — abas **Comparativo** e **Catalogo_raspado**
     (cabeçalho em negrito, filtros, painel congelado; linha vermelha quando estou
     mais de 5 % acima do concorrente, verde quando estou abaixo). Itens da minha
     tabela que não existem no concorrente vão para a aba **Sem_par_no_concorrente**.
   - `docs/index.html` — página estática autocontida (sem CDN) com KPIs, gráfico de
     barras em SVG da diferença % por produto e tabela ordenável. Abre em qualquer
     navegador, inclusive no celular, e pode ser publicada no GitHub Pages.

Exemplo de saída no terminal:

```
Produtos raspados ........ 60
Produtos monitorados ..... 20  (sem par no concorrente: 2)
Estou mais caro .......... 5  (alertas > 5%: 4)
Estou mais barato ........ 13
Iguais ................... 2
Diferença média .......... -1.95%
```

## Como rodar

Requisitos: Python 3.10+ e as bibliotecas de `requirements.txt`.

```bash
pip install -r requirements.txt
python monitor.py
```

Na primeira execução, se `dados_exemplo/meus_precos.csv` não existir, o script gera
uma tabela sintética com 20 títulos do catálogo e preços fictícios (seed fixa, logo
reproduzível). Para usar a sua tabela real, basta substituir o arquivo mantendo as
colunas `codigo_interno;titulo;meu_preco`.

Os CSVs usam o padrão brasileiro (separador `;` e vírgula decimal) para abrirem
direto no Excel. Isso é configurável em `CSV_OPCOES` no `monitor.py`.

## Estrutura

```
monitor-precos/
├── monitor.py           # fluxo principal: raspa → histórico → compara → xlsx + html
├── scraper_base.py      # sessão HTTP, retry/backoff, delay, robots.txt, parser HTML (stdlib)
├── requirements.txt
├── dados_exemplo/
│   └── meus_precos.csv  # a sua tabela (aqui, fictícia)
├── historico/           # um CSV por dia de coleta
├── saida/
│   └── comparativo_precos.xlsx
└── docs/
    └── index.html       # relatório estático
```

## Como trocar a fonte (outro site ou marketplace)

Tudo que é específico do site raspado está num único bloco do `monitor.py`,
marcado com `FONTE`:

- `URL_PAGINA` — modelo da URL de listagem com `{n}` no lugar do número da página,
  e `PAGINAS` — quantas páginas percorrer.
- Os seletores `SEL_CARD`, `SEL_TITULO`, `SEL_PRECO`, `SEL_DISPONIBILIDADE`,
  `SEL_ESTRELAS` — pares `(tag, classe CSS)` que localizam o card do produto e
  cada campo dentro dele.
- A função `extrair_produtos(html, url_pagina)` — recebe o HTML de uma página e
  devolve uma lista de dicionários com as chaves `titulo`, `preco_concorrente`,
  `avaliacao_estrelas`, `disponibilidade` e `url`. Mantendo essas chaves, o restante
  do pipeline (histórico, comparação, planilha, página) continua funcionando sem
  alteração.

O parser (`scraper_base.py`) usa apenas a biblioteca padrão: `parsear(html)` monta
uma árvore de nós e `buscar(no, tag, classe)` / `primeiro(...)` / `texto(...)`
localizam os elementos — equivalente a um `find_all` simplificado. Para sites que
só carregam os preços via JavaScript, a abordagem muda: geralmente existe uma API
JSON interna que pode ser consumida diretamente com `requests`, o que é mais
leve e estável do que automatizar um navegador.

A chave de cruzamento com a sua tabela é o título normalizado. Em um caso real, o
ideal é cruzar por EAN/código do fabricante ou manter uma tabela "de-para"
entre o seu SKU e a URL do concorrente.

## Agendar a execução diária

### Windows (Agendador de Tarefas, via `schtasks`)

Abra o Prompt de Comando e ajuste o caminho do Python e da pasta:

```bat
schtasks /Create /SC DAILY /ST 08:00 /TN "MonitorPrecos" ^
  /TR "cmd /c cd /d C:\caminho\monitor-precos && python monitor.py >> log.txt 2>&1" /F
```

Para testar na hora: `schtasks /Run /TN "MonitorPrecos"`. Para remover:
`schtasks /Delete /TN "MonitorPrecos" /F`.

### Linux (cron)

```bash
crontab -e
# todo dia às 08:00
0 8 * * * cd /caminho/monitor-precos && /usr/bin/python3 monitor.py >> log.txt 2>&1
```

Em qualquer um dos casos, o histórico acumula um arquivo por dia em `historico/`,
o que permite analisar depois a evolução de preço de cada produto.

## Aviso sobre scraping responsável

- Antes de raspar qualquer site, leia o `robots.txt` e os termos de uso. O script
  consulta o `robots.txt` automaticamente e se recusa a baixar URLs bloqueadas, mas
  isso não substitui a leitura dos termos.
- Mantenha o delay entre requisições, identifique-se no User-Agent e colete apenas
  o necessário. Muitos marketplaces oferecem APIs oficiais para parceiros, que são
  sempre a opção preferível quando existem.
- Dados pessoais não devem ser coletados. Este demo trata apenas de informações
  públicas de produtos.

## Limitações do demo

- Comparação por título exato (normalizado). Variações de grafia exigem de-para.
- Não trata paginação dinâmica nem conteúdo carregado por JavaScript.
- Os preços do site de treino estão em libras (£); o script trata apenas o número.
- Sem envio de alertas (e-mail/WhatsApp) — é a extensão natural seguinte.

---
Marcos Galani · automação de processos
