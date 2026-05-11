# Documentação Técnica — Gerador de Apresentações · Matriz Educação

> **Última atualização:** 2026-05-11
> **Responsável técnico:** Desenvolvido com Claude Code (Anthropic)
> **Repositório:** https://github.com/rafaelxavier-cmyk/reuniao-diretores-
> **Deploy (produção):** Streamlit Community Cloud — share.streamlit.io

---

## 1. Visão Geral

O projeto automatiza a geração de apresentações PowerPoint profissionais para as **Reuniões de Diretores da Matriz Educação**. O usuário sobe um documento de pauta (`.docx` ou `.pdf`), o sistema extrai a estrutura, gera os slides com identidade visual padronizada, passa por um loop de qualidade automático e entrega o arquivo pronto para download.

### Problema que resolve
Antes, cada apresentação era montada manualmente ou gerada via prompt individual para o Claude Code. O processo era repetitivo, demorado e sujeito a inconsistências visuais. O app elimina essa etapa e garante padrão visual em toda reunião.

### Dois modos de uso
1. **Upload de Pauta** — sobe o arquivo DOCX/PDF da pauta e gera automaticamente
2. **Criar Pauta Guiada** — assistente passo a passo que constrói a pauta do zero dentro do app

---

## 2. Estrutura de Arquivos

```
Reunião de diretores/
├── DOCUMENTACAO.md          ← este arquivo
├── Modelo.pptx              ← template de referência visual (não usado em runtime)
├── Procfile                 ← instrução de start para Railway/Heroku (legado)
├── requirements.txt         ← dependências Python (nível raiz, para plataformas cloud)
├── runtime.txt              ← versão Python para plataformas cloud (python-3.11)
├── .gitignore               ← exclui saida/, entrada/, PPTs gerados, __pycache__
├── entrada/                 ← pasta local para arrastar pautas (não enviada ao repo)
├── saida/                   ← PPTs gerados localmente (não enviada ao repo)
└── app/
    ├── app.py               ← interface Streamlit (ponto de entrada do app)
    ├── parser.py            ← extrai estrutura de DOCX/PDF para dict Python
    ├── generator.py         ← converte dict de pauta em arquivo .pptx
    ├── quality_loop.py      ← loop iterativo: gera → audita → ajusta → repete
    ├── validator.py         ← validador básico de qualidade visual (auto-fix de fontes)
    ├── layout_audit.py      ← auditor de precisão: overflow, overlap, fora dos limites
    ├── requirements.txt     ← dependências Python (cópia para uso local)
    └── assets/
        └── logo.png         ← logo oficial da Matriz Educação (68 KB, extraída
                                manualmente do PPT corrigido; md5=298b377c)
```

### Arquivos que NÃO estão no repositório (por design)
- `saida/` e `entrada/` — conteúdo gerado/temporário
- `📌 PAUTA – REUNIÃO DE DIRETORES 07.05.pdf` — pauta de exemplo (dado interno)
- `Iniciar App.bat` — script Windows para uso local
- `app/criar_capa_modelo.py` — script de uso único que adicionou a capa ao Modelo.pptx (já executado)

---

## 3. Dependências

```
streamlit>=1.35.0      # interface web
python-pptx>=0.6.23   # criação e manipulação de arquivos .pptx
python-docx>=1.1.2    # leitura de arquivos .docx
pdfplumber>=0.11.0    # extração de texto de PDFs
Pillow>=10.0.0        # manipulação de imagens (logo)
```

Python recomendado: **3.11+**

---

## 4. Como Rodar Localmente

```bash
cd "Reunião de diretores/app"
pip install -r requirements.txt
streamlit run app.py
```

O app abre em `http://localhost:8501`.

Ou use o arquivo `Iniciar App.bat` com duplo-clique (Windows).

---

## 5. Identidade Visual

### Paleta de Cores

| Nome       | Hex       | Uso                                      |
|------------|-----------|------------------------------------------|
| NAVY       | `#0E3D52` | Fundo principal, textos escuros          |
| TEAL       | `#5DD4B4` | Acentos, supertítulos, bordas de destaque|
| DARK_NAVY  | `#0A2D3D` | Blocos de cards, painéis laterais        |
| LIGHT_BG   | `#F4F7F8` | Fundo de cards claros                    |
| AMBER      | `#D97706` | Alertas, acentos alternativos            |
| GREEN      | `#16A34A` | Indicadores positivos                    |
| GRAY_TXT   | `#6B7280` | Textos secundários, footers              |
| WHITE      | `#FFFFFF` | Textos sobre fundos escuros              |

### Tipografia (em pontos)

| Elemento         | Tamanho | Estilo     |
|------------------|---------|------------|
| Título principal | 30pt    | Bold       |
| Título de card   | 16pt    | Bold       |
| Corpo / bullets  | 13pt    | Normal     |
| Labels / badges  | 10pt    | Bold       |
| Supertítulo      | 11pt    | Normal     |
| Footer           | 10pt    | Normal     |
| KPI valor        | 24pt    | Bold       |

**Regra absoluta:** nenhum texto de conteúdo visível abaixo de 10pt.

### Dimensões do Slide

Formato 16:9 em EMUs (English Metric Units — unidade interna do python-pptx):

```
Largura:   12.192.000 EMU  (≈ 33,87 cm)
Altura:     6.858.000 EMU  (≈ 19,05 cm)
1 pt      =    12.700 EMU
1 polegada =  914.400 EMU
```

### Elementos Estruturais de Cada Slide

```
┌─────────────────────────────────────────────────────────────┐
│ [sidebar NAVY 164.592 EMU]  [logo topo esq, top=-116.672]  │
│                              [supertítulo 11pt TEAL]        │
│                              [título 30pt WHITE]            │
│                              [empresa/data 10pt GRAY]       │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│          ÁREA DE CONTEÚDO (começa em Y=2.150.000)           │
│          Cards, bullets, tabelas, KPIs                       │
│                                                              │
├─────────────────────────────────────────────────────────────┤
│ Footer:  "Matriz Educação  ·  Reunião de Diretores"  10pt  │
│ Y = 6.537.960  (margem segura até Y = 6.450.000)           │
└─────────────────────────────────────────────────────────────┘
```

### Posição da Logo

```python
left=465055, top=-116672, width=2864869, height=1053344
```
A logo ultrapassa levemente o topo (top negativo) — comportamento intencional para alinhar visualmente com a barra lateral.

---

## 6. Fluxo de Geração

```
Arquivo (DOCX/PDF)
       │
       ▼
   parser.py
   extrai dict "pauta"
       │
       ▼
quality_loop.py
  ┌──────────────────────────────┐
  │  generator.py  → .pptx      │
  │       ↓                     │
  │  layout_audit.py → score    │
  │       ↓                     │
  │  score ≥ 85? → sai          │
  │  senão: ajusta params        │
  │  (até 4 iterações)          │
  └──────────────────────────────┘
       │
       ▼
  validator.py (auto-fix final de fontes)
       │
       ▼
   .pptx entregue ao usuário
```

---

## 7. Parser (`parser.py`)

### Responsabilidade
Converte um arquivo `.docx` ou `.pdf` em um dicionário Python com a estrutura:

```python
{
    "titulo":   "REUNIÃO DE DIRETORES",
    "data":     "07/05/2026",
    "objetivo": "Alinhamento estratégico...",
    "agenda": [
        {"num": 1, "nome": "Cultura",    "tempo": "15 min"},
        {"num": 0, "nome": "Intervalo",  "tempo": "20 min"},  # num=0 = intervalo
        {"num": 2, "nome": "Operacional","tempo": "30 min"},
    ],
    "secoes": [
        {
            "num": 1, "titulo": "Cultura", "tempo": "15 min",
            "subtitulo": "Regras de Ouro",
            "items": ["Bullet 1", "Bullet 2"],
            "subsecoes": [
                {"titulo": "Sub-bloco", "items": ["Item a", "Item b"]}
            ]
        }
    ]
}
```

### Heurísticas de Parsing (PDF)

PDFs não têm estrutura semântica como DOCX. O parser usa fases:

**Fase 1 — `header`**: captura título, data, objetivo até encontrar a agenda.

**Fase 2 — `agenda`**: captura apenas linhas com padrão `"N. Nome – XX min"`. Se um item numerado aparece **sem** sufixo de tempo, a fase muda para `content`.

**Fase 3 — `content`**: captura seções e seus bullets.

### Detecção de Seção vs Sub-item (problema resolvido)

O principal desafio foi distinguir `"3. Estaremos no Museal"` (sub-item da seção Pedagógica) de `"3. Pauta Pedagógica"` (nova seção). A solução usa duas condições:

```python
# É seção principal se:
name_match = _matches_agenda(num, title, agenda)   # nome coincide com agenda
num_match  = num in agenda_nums and num not in created_nums  # número ainda não criado

# É sub-item se:
is_sequential = (num == last_sub_num + 1) and num <= 4  # sequencial E número pequeno
```

A trava `num <= 4` resolveu o caso específico onde a seção 5 ("Meta Final de Ciclo") seria confundida com sub-item 5 da seção anterior (cujos sub-itens iam até 4).

### Matching Fuzzy de Nomes

Nomes na agenda podem diferir dos títulos no corpo do documento ("Pauta Pedagógica" vs "Discussão Pedagógica"). O matching usa normalização Unicode e contagem de palavras em comum:

```python
def _matches_agenda(num, title, agenda):
    # normaliza, remove acentos, compara palavras
    # retorna True se ≥2 palavras coincidem ou num está na agenda
```

---

## 8. Generator (`generator.py`)

### Responsabilidade
Recebe o dict de pauta e cria um arquivo `.pptx` do zero usando `python-pptx`, sem depender de template externo.

### Slides Gerados

| Slide | Função | Condição |
|-------|--------|----------|
| 1 | Capa | Sempre |
| 2 | Agenda (sumário visual) | Sempre |
| 3..N | Um slide por seção | Uma seção = um slide |
| Último | Encerramento | Sempre |

### Parâmetros do Quality Loop

O generator aceita parâmetros de controle via `pauta['_gen_params']`:

```python
{
    'max_bullets_per_card': 6,      # máximo de bullets por card
    'card_title_lines':     2,      # linhas do título do card
    'extra_spacing_factor': 1.0,    # multiplicador de espaçamento
    'truncate_long_items':  False,  # truncar bullets longos
}
```

O `quality_loop.py` injeta esses parâmetros e os torna mais conservadores a cada iteração com problemas.

### Constantes de Layout Críticas

```python
CONTENT_TOP = 2150000   # Y onde o conteúdo começa (abaixo do cabeçalho)
SAFE_BOTTOM = 6450000   # Y máximo para conteúdo (acima do footer)
FOOTER_T    = 6537960   # Y do footer (intencionalmente abaixo do SAFE_BOTTOM)
CONTENT_L   = 548640    # X mínimo do conteúdo (direita da sidebar)
```

**Por que `FOOTER_T > SAFE_BOTTOM`?**
O footer existe em `Y=6.537.960` mas o conteúdo dinâmico não deve ultrapassar `Y=6.450.000`. O footer é um elemento estático fixo, não conteúdo gerado, por isso tem posição própria fora da margem segura.

---

## 9. Quality Loop (`quality_loop.py`)

### Por que existe
Após a geração inicial, caixas de texto podiam se sobrepor, texto podia transbordar ou fontes ficavam pequenas demais dependendo do volume de conteúdo. Em vez de corrigir manualmente, o sistema regenera com parâmetros mais conservadores até atingir qualidade aceitável.

### Algoritmo

```
para iteração em 1..4:
    gera o PPT com os parâmetros atuais
    audita o layout (layout_audit.py)
    
    se score ≥ 85 e sem erros:
        para o loop (sucesso)
    senão:
        tighten_params():
            overflow/tiny_box → reduz max_bullets, aumenta espaçamento
            overlap           → reduz card_title_lines, aumenta espaçamento

aplica auto-fix final de fontes (validator.py)
retorna (bytes, AuditReport, log_de_iterações)
```

### Resultado típico
Na grande maioria dos casos (pautas com até 7 seções e ~15 bullets por seção), o score é 100/100 na primeira iteração.

---

## 10. Layout Auditor (`layout_audit.py`)

### Checks realizados

| Tipo | Severidade | Descrição |
|------|-----------|-----------|
| `tiny_box` | error | Caixa menor que 1 linha de texto — conteúdo invisível |
| `overflow` | warning | Texto estimado >140% da altura da caixa |
| `out_of_bounds` | error | Shape ultrapassa o limite inferior do slide |
| `overlap` | error | Sobreposição >30% entre dois shapes de conteúdo |
| `overlap` | warning | Sobreposição 15–30% entre dois shapes de conteúdo |

### Zonas excluídas dos checks
- **Sidebar**: `left < 164.592 + 50.000 EMU` — barra lateral decorativa
- **Topo/logo**: `top < 50.000 EMU` com largura >30% do slide
- **Cabeçalho estrutural**: `top < 900.000 EMU`
- **Footer**: `top ≥ 6.430.000 EMU` — footer intencional

### Score e aprovação

```
score = 100 - (erros × 15) - (avisos × 5)
passed = score ≥ 85 AND sem erros
```

---

## 11. Validator (`validator.py`)

Camada mais simples, roda após o quality loop como auto-fix final:

- Verifica tamanho mínimo de fonte (conteúdo ≥ 12pt, títulos ≥ 18pt)
- Verifica densidade de texto (máx. 600 caracteres/slide)
- Detecta sobreposições básicas
- **Auto-fix**: fontes abaixo de 10pt em zona de conteúdo são bump para 11–12pt

Gera nota A/B/C/D baseada em score 0–100.

---

## 12. Interface (`app.py`)

### Tecnologia
Streamlit — framework Python que cria interfaces web sem HTML/JS.

### Estrutura da UI

```
┌─────────────────────────────────────────────────────┐
│  HERO BANNER (gradiente navy→teal)                  │
├──────────────────────┬──────────────────────────────┤
│  Tab: Upload de Pauta│  Tab: Criar Pauta Guiada      │
└──────────────────────┴──────────────────────────────┘
```

**Tab Upload:**
- Upload de arquivo DOCX/PDF
- Edição de data e objetivo extraídos
- Preview da estrutura detectada
- Botão "Gerar" → chama `_run_generation(pauta)`

**Tab Criar Pauta Guiada (Wizard 4 etapas):**
1. **Identificação** — data, formato, objetivo
2. **Tópicos** — tabela de temas com tempo e responsável
3. **Conteúdo** — subtítulo, objetivo e bullets de cada tópico
4. **Revisão** — preview completo + botão gerar

### Função `_run_generation(pauta)`

```python
# Chama o quality loop com callback de progresso
pptx_bytes, audit_report, iterations_log = run_quality_loop(
    pauta, out_path, ASSETS_DIR, progress_callback=_progress
)
```

O `_progress` atualiza um `st.empty()` a cada iteração mostrando o score em tempo real.

### Diretório de saída

```python
SAIDA_DIR = Path(tempfile.gettempdir()) / "reuniao_saida"
```

Usa o `/tmp` do sistema operacional para compatibilidade com ambientes cloud (Streamlit Community Cloud, Railway). Os arquivos são temporários — o usuário baixa via botão e não depende de persistência.

**Decisão:** originalmente `BASE_DIR / "saida"` (relativo ao projeto). Mudado para `tempfile` porque em ambientes cloud o diretório do repositório pode ser read-only.

---

## 13. Deploy

### Repositório GitHub
**URL:** https://github.com/rafaelxavier-cmyk/reuniao-diretores-
**Branch principal:** `main`

Para atualizar:
```bash
git add <arquivos>
git commit -m "descrição da mudança"
git push
```
O Streamlit Cloud redeploya automaticamente após o push.

### Streamlit Community Cloud (produção)
**Plataforma:** share.streamlit.io
**Configuração:**
- Repository: `rafaelxavier-cmyk/reuniao-diretores-`
- Branch: `main`
- Main file: `app/app.py`

**Por que Streamlit Cloud e não Vercel?**
Vercel é serverless (funções que iniciam e terminam em segundos). O Streamlit usa WebSockets e mantém estado por sessão — requer um processo Python rodando continuamente. O Streamlit Community Cloud foi construído especificamente para isso e é gratuito.

**Railway** também foi considerado mas é pago. **Render** tem plano gratuito mas dorme após 15 min de inatividade.

---

## 14. Decisões de Design Relevantes

### Por que python-pptx e não uma API de IA para gerar slides?
Controle total e determinístico sobre pixel/EMU. APIs como GPT ou Gemini para slides não garantem o posicionamento exato, paleta de cores e tipografia da identidade Matriz Educação.

### Por que criar slides do zero em vez de usar o Modelo.pptx como template?
O `Modelo.pptx` foi a referência visual inicial. Durante o desenvolvimento descobriu-se que clonar slides de um template via `python-pptx` introduz artefatos (estilos herdados, shapes invisíveis). A abordagem de construir cada slide com `Presentation()` puro é mais estável e previsível.

### Por que o loop de qualidade tem máximo de 4 iterações?
Trade-off entre qualidade e tempo de resposta. Na prática, 95% das pautas passam na primeira iteração. 4 iterações é um limite seguro que garante resposta em <60 segundos mesmo no pior caso.

### Por que a logo tem `top` negativo (`-116.672 EMU`)?
A logo da Matriz Educação tem espaço em branco na parte inferior. Para alinhar a parte visível com a barra lateral teal, a posição vertical foi ajustada negativamente. Isso é intencional e reproduz o layout do PPT original.

### Por que o footer fica fora do `SAFE_BOTTOM`?
`SAFE_BOTTOM = 6.450.000` é o limite para conteúdo dinâmico gerado. O footer está em `6.537.960` — ligeiramente abaixo. O footer é um elemento estático posicionado explicitamente, não conteúdo dinâmico, então não precisa respeitar esse limite.

---

## 15. Problemas Resolvidos (Histórico)

| Problema | Causa | Solução |
|----------|-------|---------|
| Parser capturava 20 itens na agenda em vez de 7 | Sem distinção entre fase agenda e fase conteúdo | Sistema de fases: `header → agenda → content` |
| Seções 4 e 5 não detectadas | Nome na agenda diferente do título no corpo | Match por número além do match por nome |
| Sub-item "3. Estaremos no Museal" vira seção 3 | Número 3 coincide com seção da agenda | Heurística de sequencialidade: `num == last+1 AND num ≤ 4` |
| Seção 5 não capturada depois dos sub-itens 1,2,3,4 | Sub-itens iam até 4, então 5 parecia sequencial (4+1=5) | Trava `num <= 4` na heurística de sequencialidade |
| Logo trocada por ícone genérico | Arquivo de logo ausente | Extraída manualmente do PPT correto e salva em `assets/logo.png` |
| Score do validador zerava (29 warnings) | Shapes decorativos/estruturais sendo auditados | Zonas de exclusão: HEADER_ZONE_BOTTOM, FOOTER_ZONE_TOP, `_is_label_shape()` |
| `slide_layouts[6]` não existe | Modelo.pptx tem apenas 1 layout | Mudado para `slide_layouts[0]` |
| Footer apontado como fora dos limites | Footer (6.537.960) > SAFE_BOTTOM (6.450.000) | `layout_audit.py`: zona de footer a partir de `6.430.000` excluída |
| `SAIDA_DIR.mkdir()` falha no cloud | Diretório do repo pode ser read-only | Mudado para `tempfile.gettempdir()` |

---

## 16. Como Fazer Backup Completo

Os arquivos essenciais para recriar o sistema do zero:

1. **Código-fonte** — todo o conteúdo do repositório GitHub (clone com `git clone`)
2. **Logo** — `app/assets/logo.png` (68 KB, md5=298b377c) — extraída manualmente do PPT corrigido pelo Rafael; se perder, precisa ser extraída novamente de um PPT que tenha a logo correta

O `Modelo.pptx` está no repositório por histórico, mas não é usado em runtime.

---

*Este documento deve ser atualizado a cada mudança significativa no código, infraestrutura ou decisões de design.*
