# AYSHA Toolkit — Analisador de programas Rockwell (L5X)

MVP de uma aplicação que recebe um programa de PLC Rockwell/Allen-Bradley exportado em **L5X**,
extrai a estrutura de forma **determinística** (`l5x_core`) e, nas fases seguintes, usa um LLM para
explicar, revisar, documentar e conversar sobre o código. O produto é de **leitura**: nunca gera,
altera ou envia código para o controlador.

Estado atual: **Fases 1 e 2 concluídas** (parser + análise determinística; camada LLM com prompts,
retrieval, cache e orçamento de contexto). Fase 3 (interface Streamlit) ainda não foi iniciada.
O plano completo está em [PROMPT_MVP_L5X.md](PROMPT_MVP_L5X.md).

## Como exportar o L5X no Studio 5000

1. Abra o projeto `.ACD` no Studio 5000 Logix Designer.
2. `File → Save As…`, em *Save as type* escolha **L5X** e salve.
3. Coloque o arquivo em `samples/` (pasta ignorada pelo git). Nunca versione um projeto real.

## Instalação e uso da CLI

```bash
pip install -r requirements.txt
python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --summary --findings
python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --findings --min-severity warning
python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --findings --rule routine_never_called --rule output_multiple_writes
python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --dump-routine DIAGNOSTICS/CPU_Status
python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --tags
python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --json saida.json
```

Testes:

```bash
python -m pytest -q
```

O teste `test_real_sample_parses_fast` roda contra `samples/P80_HULL_HCS01.L5X` se o arquivo existir
(e é pulado se não existir). Nos exemplos de 31 MB e 46 MB o parse completo leva cerca de 1 a 2 s.

## Camada LLM (Fase 2)

Configuração: copie `.env.example` para `.env` (ou `.streamlit/secrets.toml.example` para
`.streamlit/secrets.toml`) e preencha `ANTHROPIC_API_KEY`, `MODEL_FAST` (explicar, conversar) e
`MODEL_DEEP` (revisar, documentar). Os IDs de modelo **não** ficam no código: consulte
https://docs.claude.com/en/docs/about-claude/models. Variáveis opcionais: `LLM_EFFORT_FAST` /
`LLM_EFFORT_DEEP` (`low`…`max`; deixe vazio para modelos que não aceitam `output_config.effort`),
`MAX_CONTEXT_TOKENS` (padrão 60000), `MAX_OUTPUT_TOKENS` (16000), `APP_LANGUAGE` (`pt-BR`, `en`, `es`),
`LLM_LOG_PATH` (CSV de métricas; vazio desativa).

Teste manual com chave real:

```bash
python -m app.llm samples/P80_HULL_HCS01.L5X --mode explain --target DIAGNOSTICS/CPU_Status
python -m app.llm samples/P80_HULL_HCS01.L5X --mode review --target Control_Loops
python -m app.llm samples/P80_HULL_HCS01.L5X --mode document --target DIAGNOSTICS
python -m app.llm samples/P80_HULL_HCS01.L5X --mode chat --question "o que liga Blackout_Start_Activation?"
```

Como funciona:

- `app/llm.py`: `LLMConfig.from_env()`, `ProjectContext` (índice do projeto + tabelas determinísticas de
  timers e I/O), `LLMClient` com os quatro modos (`explain`, `review`, `document`, `chat`), streaming
  (`on_text`), map-reduce quando o escopo não cabe no orçamento, consolidação recursiva, tratamento de
  `refusal`/`max_tokens` e mapeamento de erros da API para mensagens em português (`LLMError.kind`).
- **Prompt caching**: o bloco de sistema é `[regras fixas (system.md), visão geral do projeto]` com
  `cache_control: ephemeral` no segundo bloco; o conteúdo variável (rotinas, findings, pergunta) vai na
  mensagem do usuário, então a visão geral é reaproveitada em todas as chamadas da sessão. Verifique com
  `cache_read_tokens` no CSV.
- **Orçamento**: `MAX_CONTEXT_TOKENS` menos o bloco de sistema e uma reserva. Rotinas são empacotadas em
  pedaços (greedy, na ordem do projeto); uma rotina maior que o orçamento é dividida em fronteiras de rung,
  linha ST ou tag. Revisão/documentação completas rodam rotina a rotina e consolidam (`consolidate.md`).
  Nos dois exemplos reais, nenhuma chamada passa de ~59 mil tokens estimados.
- `app/retrieval.py`: sem vetor. Pontua rotinas por nome de rotina/programa citado, tags citadas
  (resolvidas pelo xref para as rotinas que as usam), palavras da descrição e dos comentários; sem
  correspondência, cai nas rotinas principais dos programas agendados.
- `app/prompts/*.md`: `system.md` (aterramento, citação `Programa/Rotina rung N`, idioma, safety,
  protegido, FBD, "confirmado pelo parser" vs "hipótese do modelo", aviso final de revisão), `explain.md`,
  `review.md`, `document.md`, `chat.md`, `consolidate.md`.
- Log `llm_calls.csv`: timestamp, modo, modelo, tokens de entrada/saída, tokens lidos/gravados no cache,
  contexto estimado, latência, status. Nunca contém conteúdo do programa (há teste para isso).
- Testes (`tests/test_llm.py`, `tests/test_retrieval.py`) usam um cliente falso (`tests/fake_anthropic.py`);
  nenhum teste chama a API.

Decisões: o parâmetro `thinking` não é enviado (os modelos atuais já usam raciocínio adaptativo por
padrão); `fallbacks` de recusa no servidor não foi ativado porque depende do modelo configurado; uma
recusa (`stop_reason = refusal`) é devolvida como texto explicativo com `LLMResult.refusal = True`.

## Estrutura

```
l5x_core/
  model.py       dataclasses: Controller, Task, Program, Routine, Rung, Instruction, Tag, DataType, Aoi, Module, Finding, Metrics
  parser.py      L5X -> model (lxml, huge_tree); tolerante; EncodedData => protected=True
  rung_text.py   tokenizador do texto neutro dos rungs; tabela de leitura/escrita por instrução
  st_text.py     extração aproximada de tags e chamadas em Structured Text (regex)
  xref.py        tag -> usos (programa, rotina, rung, instrução, leitura|escrita); grafo de chamadas JSR
  analysis.py    regras determinísticas -> Finding; métricas
  serialize.py   project_overview(), routine_text(), tag_sheet(), to_json() com estimativa de tokens
  cli.py         linha de comando
app/
  llm.py         LLMConfig, ProjectContext, LLMClient (explain/review/document/chat), log CSV, CLI
  retrieval.py   seleção de rotinas relevantes para uma pergunta
  prompts/       system.md, explain.md, review.md, document.md, chat.md, consolidate.md
tests/
  data/minimal.l5x   L5X sintético que cobre todas as regras
  test_*.py
samples/           projetos reais (ignorado pelo git)
```

## Regras de análise (`Finding.rule`)

| Regra | Severidade | O que detecta |
|---|---|---|
| `jsr_missing_routine` | error | JSR/FOR para rotina que não existe no programa |
| `routine_never_called` | warning | rotina que não é main/fault e não aparece em nenhum JSR |
| `program_not_scheduled` / `program_disabled` | warning | programa fora de qualquer task / `Disabled="true"` |
| `output_multiple_writes` | warning | mesma saída com OTE em mais de um rung (a última vence) |
| `output_mixed_latch` | warning | OTE e OTL/OTU na mesma saída |
| `timer_no_preset` | warning | TON/TOF/RTO/CTU/CTD com preset literal 0 ou `PRE=0` no valor da tag |
| `tag_write_only` | warning | tag escrita e nunca lida (exclui I/O, produced/consumed, alias, parâmetros de programa) |
| `tag_read_only` | info | tag lida e nunca escrita (exclui as mesmas, e tags `Constant`) |
| `tag_unused` | info | tag declarada e nunca referenciada (exclui I/O e produced/consumed) |
| `rungs_without_comment` | info | por rotina: quantos rungs sem comentário e % comentados |
| `routine_no_description` | info | rotina sem descrição |
| `routine_protected` / `aoi_protected` | info | source protection (`EncodedData`): conteúdo não lido |
| `safety_content` | info | tasks/programas/tags `Class="Safety"`; todos os findings nesse conteúdo levam `safety=True` |
| `unknown_instruction` | info | mnemônico fora da tabela: operandos tratados como leitura |
| `unresolved_tags` | info | operandos que não resolvem para nenhuma tag (por rotina) |
| `rung_parse_error` | info | rung cujo texto não foi interpretado; o resto do arquivo segue normalmente |

Métricas (`Metrics`): programas, rotinas por tipo, rungs, linhas ST, instruções, tags por escopo, AOIs
(e protegidas), UDTs, módulos, tasks, maior rotina, profundidade máxima de ramos, % de rungs comentados.

## Diferenças observadas em relação à especificação (PROMPT)

Confirmadas nos dois L5X de exemplo (Studio 5000 v35):

- **Tags não trazem `Class`** nos exemplos (nem `Standard` nem `Safety`); `Program` e `Task` também não.
  O parser lê `Class` quando existe e ainda marca como safety os programas agendados em uma task `Class="Safety"`.
- **Não há tags `Alias`** nos exemplos; o suporte foi implementado e testado com o L5X sintético. Em L5X,
  tag alias não tem `DataType`, só `AliasFor`.
- **A maioria das rotinas é FBD** (774 de 849 no HULL, 2116 de 2346 no FGS). O PROMPT pede apenas registrar
  existência/tipo/tamanho; se fosse só isso, a regra "tag nunca referenciada" apontaria quase todas as tags.
  Por isso o parser extrai do FBD **referências aproximadas** (`IRef` = leitura, `ORef` = escrita,
  `Block`/`AddOnInstruction` `Operand` e `InOutParameter` = leitura+escrita), marcadas como `approximate`.
  A lógica interna do FBD continua não analisada.
- **`TON(T,?,?)` é a forma normal** do texto neutro em v35: preset e acumulado ficam no `<Data Format="Decorated">`
  da tag. A regra de preset lê `PRE` de lá (inclusive dentro de UDT e arrays). Valores só são guardados para tags
  cujo tipo contém TIMER/COUNTER, para economizar memória.
- **Chamada de AOI no rung** = tag de instância + os parâmetros `Required="true"` ou `Usage="InOut"`, na ordem
  da definição. Os demais parâmetros vivem dentro da tag de instância.
- `EncodedData` apareceu só em AOIs nos exemplos (8 no HULL, 23 no FGS), como elemento irmão de
  `AddOnInstructionDefinition` dentro de `AddOnInstructionDefinitions`. Rotinas protegidas
  (`Routine/EncodedData`) são suportadas e cobertas pelo L5X sintético.
- Módulos aparecem como operando de AOIs de diagnóstico (`L_ModuleSts(sts, NOME_DO_MODULO)`); esses nomes são
  resolvidos contra a lista de módulos e não geram `unresolved_tags`.
- Referências a I/O usam `Modulo:slot:I.Data` diretamente no rung; essas tags não estão em `<Tags>` e são
  registradas como uso de I/O do módulo.
- Instruções encontradas nos exemplos além da tabela mínima do PROMPT e já cobertas: NOP, AFI, TND, ONS/OSR/OSF,
  CMP, ABS, SQR, NEG, AND/OR/XOR/NOT, BTD, BSL/BSR, FAL/FSC, SIZE, SWPB, MVM, GSV/SSV, CONCAT/LOWER/UPPER/MID/
  DELETE/INSERT/FIND/DTOS/STOD/RTOS/STOR, PID, FOR/BRK, UID/UIE, MCR, EVENT, IOT. Nenhuma instrução
  desconhecida restou nos dois exemplos.

## Privacidade

- `samples/`, `.streamlit/secrets.toml`, `*.db` e `.env` estão no `.gitignore`.
- O parser nunca registra em log o conteúdo do programa; findings carregam apenas trechos curtos de evidência.
- Conteúdo `EncodedData` nunca é decodificado.

## Fase seguinte (não iniciada)

- **Fase 3**: `app/streamlit_app.py`, `app/credits.py` (códigos de acesso em SQLite), `streamlit run app/streamlit_app.py`.
