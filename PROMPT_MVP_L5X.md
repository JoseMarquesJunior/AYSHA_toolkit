# PROMPT — MVP "Analisador de programas Rockwell (L5X)"

> **Como usar (para você, não para o agente):** salve este arquivo como `PROMPT.md` na raiz de uma pasta vazia, abra a pasta no VS Code e diga ao agente de IA (Claude Code, Copilot Agent, Cursor etc.): *"Leia PROMPT.md por completo e execute a Fase 1. Não avance de fase sem me mostrar os testes passando."*
>
> **Antes de começar:** o arquivo `P80_HULL_HCS01.ACD` é binário e **não** é a entrada do produto. Exporte-o para L5X no Studio 5000 (`File → Save As → Save as type: L5X`) e coloque o resultado em `samples/P80_HULL_HCS01.L5X`. Essa pasta fica no `.gitignore` — nunca suba um projeto real para o repositório.

---

## 1. Papel e objetivo

Você é um engenheiro de software sênior com experiência em Python, parsing de XML e integração com a API da Anthropic. Vai construir um **MVP** de uma aplicação web que recebe um programa de PLC Rockwell/Allen-Bradley exportado em **L5X**, extrai a estrutura do programa de forma **determinística** e usa um LLM para **explicar, revisar, documentar e conversar** sobre o código. Interface em **Streamlit**. Idioma da interface e das respostas: **português do Brasil** por padrão, com opção de inglês e espanhol.

O produto é de **leitura**: ele nunca gera, altera ou envia código para o controlador.

## 2. Princípios que não se negociam

1. **Parser primeiro, LLM depois.** Nada do XML cru vai para o modelo. O modelo só recebe o índice estruturado produzido pelo parser (texto neutro dos rungs, tags, referências cruzadas). Toda afirmação do modelo deve citar `Programa / Rotina / Rung N`.
2. **Rotinas de segurança são somente-leitura.** Programas, tasks ou tags com `Class="Safety"` (GuardLogix) são marcados; o app explica e documenta, mas nunca sugere alteração neles, e exibe um aviso fixo.
3. **Conteúdo protegido é recusado.** Qualquer `<EncodedData>` (rotina ou AOI com source protection) é detectado, listado como "protegido" e **não** é decodificado nem enviado ao modelo.
4. **Privacidade.** O arquivo enviado vive só na sessão (diretório temporário, apagado ao fim). Nunca registre em log o conteúdo do programa. Chave da API vem de `.streamlit/secrets.toml` ou variável de ambiente, nunca do código.
5. **Simplicidade.** Dependências permitidas sem perguntar: `lxml`, `streamlit`, `anthropic`, `pytest`, `python-dotenv`. Qualquer outra, pergunte antes. Sem banco vetorial, sem fila, sem Docker nesta fase.

## 3. O que é um L5X (não invente a estrutura — abra o arquivo de exemplo e confirme)

L5X é o export XML do Studio 5000. Estrutura relevante:

```xml
<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="33.00" TargetName="..." TargetType="Controller" ...>
  <Controller Name="..." ProcessorType="1756-L8xE" MajorRev=".." MinorRev=".." ...>
    <DataTypes>            <!-- UDTs: <DataType Name=".." Class="User"><Members><Member Name=".." DataType=".." Dimension=".."/>... -->
    <Modules>              <!-- hardware: <Module Name=".." CatalogNumber=".." ParentModule=".." .../> -->
    <AddOnInstructionDefinitions>   <!-- AOIs: <AddOnInstructionDefinition Name=".."><Parameters/><LocalTags/><Routines/></...> -->
    <Tags>                 <!-- tags de escopo controlador: <Tag Name=".." TagType="Base|Alias|Produced|Consumed" DataType=".." AliasFor=".." Class="Standard|Safety"><Description><![CDATA[..]]></Description>... -->
    <Programs>
      <Program Name=".." MainRoutineName=".." FaultRoutineName=".." Disabled="false" Class="Standard|Safety">
        <Tags> ... </Tags>  <!-- tags de escopo programa -->
        <Routines>
          <Routine Name=".." Type="RLL">   <!-- ladder -->
            <RLLContent>
              <Rung Number="0" Type="N">
                <Comment><![CDATA[comentário do rung]]></Comment>
                <Text><![CDATA[XIC(Start)XIO(Stop)[OTE(Motor),TON(T1,?,?)];]]></Text>
              </Rung>
            </RLLContent>
          </Routine>
          <Routine Name=".." Type="ST"><STContent><Line Number="0"><![CDATA[texto estruturado]]></Line></STContent></Routine>
          <Routine Name=".." Type="FBD"> <FBDContent>...</FBDContent> </Routine>   <!-- tratar como "não analisado em detalhe" no MVP -->
          <Routine Name=".." Type="SFC"> ... </Routine>                              <!-- idem -->
          <Routine Name=".." Type="RLL"><EncodedData EncodedType="Routine" ...>base64</EncodedData></Routine>  <!-- protegida: recusar -->
        </Routines>
      </Program>
    </Programs>
    <Tasks>  <!-- <Task Name=".." Type="CONTINUOUS|PERIODIC|EVENT" Rate=".." Class=".."><ScheduledPrograms><ScheduledProgram Name=".."/></...> -->
  </Controller>
</RSLogix5000Content>
```

Gramática do texto neutro dos rungs (`<Text>`): sequência de instruções `NOME(operando1,operando2,...)`, ramos em paralelo entre `[` `]` separados por `,`, aninhamento permitido, rung termina em `;`. Operandos são tags (`Motor`, `Tanque.Nivel`, `Alarmes[3].Ativo`, `Timer.DN`), literais (`10`, `1.5`, `'texto'`) ou `?` (parâmetro não usado). Referência a tag de outro programa usa `\Programa.Tag`.

Tabela mínima de instruções e papel dos operandos (complete-a a partir do que encontrar no arquivo de exemplo):

| Instrução | Operandos | Escrita em |
|---|---|---|
| XIC, XIO, ONS | tag | — (ONS escreve no bit de storage) |
| OTE, OTL, OTU | tag | operando 1 |
| TON, TOF, RTO | timer, preset, accum | timer |
| CTU, CTD | counter, preset, accum | counter |
| RES | timer/counter | operando 1 |
| MOV | source, dest | dest (último) |
| COP, CPS | source, dest, length | dest |
| FLL | source, dest, length | dest |
| CLR | dest | dest |
| ADD, SUB, MUL, DIV, MOD | a, b, dest | dest (último) |
| CPT | dest, expressão | dest (primeiro) |
| EQU, NEQ, GRT, GEQ, LES, LEQ, LIM, MEQ | comparação | — |
| JSR | rotina, nº entradas, entradas..., retornos... | retornos |
| SBR, RET | parâmetros | — |
| JMP, LBL | label | — |
| MSG | control | control |
| AOI (nome não listado) | parâmetros conforme definição | parâmetros `Usage="Output"` ou `InOut` |
| desconhecida | — | tratar todos como leitura e registrar aviso |

## 4. Estrutura do repositório

```
plc-analyzer/
  l5x_core/                 # pacote Python independente de UI (reaproveitado depois em outra interface)
    __init__.py
    model.py                # dataclasses: Controller, Task, Program, Routine, Rung, Instruction, Tag, DataType, Aoi, Module, Finding
    parser.py               # L5X -> model (lxml); tolerante a elementos desconhecidos
    rung_text.py            # tokenizador/parsers do texto neutro: instruções, operandos, ramos
    xref.py                 # referências cruzadas: tag -> (programa, rotina, rung, instrução, leitura|escrita); JSR -> grafo de chamadas
    analysis.py             # regras determinísticas -> lista de Finding (ver Fase 1)
    serialize.py            # índice compacto para o LLM (JSON e texto), por projeto e por rotina
    cli.py                  # `python -m l5x_core.cli arquivo.l5x --summary --findings --dump-routine Programa/Rotina`
  app/
    streamlit_app.py        # UI
    llm.py                  # cliente Anthropic, caching, roteamento de modelo, orçamento de tokens
    retrieval.py            # seleção de rotinas relevantes para uma pergunta (sem vetor: nomes, tags, comentários)
    credits.py              # códigos de acesso e cotas (SQLite)
    prompts/
      system.md             # regras gerais (aterramento, citação, idioma, segurança)
      explain.md
      review.md
      document.md
      chat.md
  tests/
    data/minimal.l5x        # L5X sintético pequeno criado por você, cobrindo todos os casos abaixo
    test_parser.py
    test_rung_text.py
    test_xref.py
    test_analysis.py
  samples/                  # projetos reais — NO .gitignore
  .streamlit/
    config.toml             # server.maxUploadSize = 300
    secrets.toml.example
  .gitignore                # samples/, .streamlit/secrets.toml, *.db, .env
  requirements.txt
  README.md
```

## 5. Fase 1 — parser e análise determinística (`l5x_core`)

**Entregar primeiro, com testes passando, antes de qualquer código de UI ou LLM.**

### 5.1 Parser

- Ler L5X de até ~150 MB com `lxml` (`etree.parse` é aceitável; se memória for problema, `iterparse`). Nunca carregar o XML inteiro como string em memória duas vezes.
- Extrair: cabeçalho (`SoftwareRevision`, `TargetName`, `ExportDate`), controlador (`Name`, `ProcessorType`, `MajorRev`), tasks e programas agendados, programas (nome, `MainRoutineName`, `FaultRoutineName`, `Disabled`, `Class`), rotinas (nome, tipo, nº de rungs/linhas, protegida?), rungs (número, tipo, comentário, texto), tags por escopo (nome, tipo, data type, alias, descrição, `Class`), UDTs (nome, membros), AOIs (nome, revisão, parâmetros com `Usage`, nº de rotinas, protegida?), módulos (nome, catálogo, pai).
- Rotinas ST: guardar as linhas; extrair referências a tags por regex de identificadores (aceitar imprecisão, marcar como "aproximado").
- Rotinas FBD/SFC: registrar existência, tipo e tamanho; não analisar conteúdo no MVP.
- `EncodedData` em rotina ou AOI → `protected=True`, conteúdo descartado.
- Toda exceção de parsing de um rung vira um `Finding` de severidade "info" com o texto do rung; o parser nunca aborta o arquivo inteiro por um rung ruim.

### 5.2 Texto neutro dos rungs (`rung_text.py`)

- Tokenizar `NOME(...)`, respeitando parênteses aninhados, aspas e colchetes de ramo.
- Para cada instrução: nome, lista de operandos, posição no rung, se está em ramo.
- Para cada operando: nome base da tag (até o primeiro `.` ou `[`), caminho completo, se é literal, se é `?`.
- Classificar leitura/escrita pela tabela da seção 3; AOIs pela definição (`Usage` dos parâmetros); desconhecidas = leitura + aviso.

### 5.3 Referências cruzadas (`xref.py`)

- Resolver cada operando para uma tag: primeiro escopo do programa, depois controlador; `\Programa.Tag` resolve direto; aliases resolvem para o alvo (`AliasFor`) mantendo o nome original.
- Índice `tag -> lista de usos (programa, rotina, rung, instrução, leitura|escrita)`.
- Grafo de chamadas: `JSR` (ladder e ST) → `rotina chamadora -> rotina chamada`; raízes = `MainRoutineName`, `FaultRoutineName`, rotinas de AOI (`Logic`, `Prescan`, `EnableInFalse`).

### 5.4 Regras de análise (`analysis.py`) — cada uma gera `Finding(regra, severidade, programa, rotina, rung, tag, mensagem, evidência)`

Obrigatórias no MVP:

1. **Rotina nunca chamada**: não é main/fault, não aparece em nenhum JSR, não é rotina de AOI.
2. **Programa não agendado ou desabilitado**: não está em nenhuma task, ou `Disabled="true"`.
3. **Tag declarada e nunca referenciada** (por escopo; ignorar tags de módulo de I/O `Local:...` e tags consumidas/produzidas).
4. **Tag escrita e nunca lida** / **lida e nunca escrita** (excluindo I/O, produced/consumed e aliases).
5. **Saída escrita em mais de um lugar** com `OTE` (mesma tag em rungs diferentes — a última vence) e `OTE` + `OTL/OTU` misturados na mesma tag.
6. **Timer/contador sem preset ou com preset literal 0**.
7. **JSR para rotina inexistente**.
8. **Rung sem comentário** (métrica: % de rungs comentados por rotina) e **rotina sem descrição**.
9. **Rotina protegida (`EncodedData`)** — severidade "info", listada como não analisável.
10. **Conteúdo safety** — programas/tasks/tags `Class="Safety"`: marcar, não gerar sugestões.
11. **Métricas**: nº de programas, rotinas, rungs, instruções, tags por escopo, AOIs, UDTs, módulos; rungs por rotina (maior rotina), profundidade máxima de ramos.

### 5.5 Serialização para o LLM (`serialize.py`)

- `project_overview()` → texto compacto (≤ ~6 mil tokens): controlador, tasks, programas com rotinas e tamanhos, AOIs, UDTs, módulos, resumo dos findings.
- `routine_text(programa, rotina)` → rotina com rungs numerados, comentário e texto neutro, e as descrições das tags usadas nela (uma linha por tag).
- `tag_sheet()` → tabela de tags com tipo, descrição e contagem de leituras/escritas.
- Estimar tokens (≈ caracteres / 3,5) e expor no retorno.

### 5.6 CLI e testes

- CLI: `--summary`, `--findings` (tabela), `--dump-routine Programa/Rotina`, `--json saida.json`.
- Crie `tests/data/minimal.l5x` à mão com: 2 programas (um não agendado), 4 rotinas (main, uma chamada por JSR, uma órfã, uma `EncodedData`), tags de controlador e de programa, um alias, uma AOI com parâmetro Output, um UDT, uma saída com OTE duplicado, um timer sem preset, um rung sem comentário, um programa `Class="Safety"`.
- Testes cobrem cada regra da 5.4 e o tokenizador (ramos aninhados, literais com vírgula dentro de aspas, `?`).
- Rode também contra `samples/P80_HULL_HCS01.L5X` (se existir) e imprima o tempo de parsing e as métricas; não falhe o teste se o arquivo não existir.

**Critério de aceite da Fase 1:** `pytest` verde; `python -m l5x_core.cli samples/P80_HULL_HCS01.L5X --summary --findings` roda em menos de 30 s e lista rotinas órfãs, tags não usadas e saídas duplicadas com programa/rotina/rung.

## 6. Fase 2 — camada LLM (`app/llm.py`, `app/retrieval.py`, `app/prompts/`)

- SDK oficial `anthropic`. Dois modelos configuráveis por variável de ambiente: `MODEL_FAST` (explicar, conversar) e `MODEL_DEEP` (revisar, documentar). Não fixe IDs de modelo no código; leia de `.env`/secrets e documente no README que os IDs atuais estão em `docs.claude.com/en/docs/about-claude/models`.
- **Prompt caching**: o bloco de sistema é `[regras fixas, visão geral do projeto (cache_control ephemeral)]`; o contexto do projeto entra uma vez por sessão e é reaproveitado no chat.
- **Orçamento de contexto**: se `project_overview + rotinas selecionadas` ultrapassar o limite configurado (`MAX_CONTEXT_TOKENS`, padrão 60 mil), use `retrieval.py` para escolher as rotinas relevantes (pontuação por nome de rotina, tags citadas na pergunta, palavras do comentário) e, para revisão/documentação completas, processe rotina a rotina e consolide (map-reduce).
- **Quatro modos**, cada um com seu prompt em `prompts/`:
  - `explain`: explica uma rotina (ou o programa) para um técnico de manutenção; por rung ou por bloco funcional; sempre cita o rung.
  - `review`: recebe os `Finding` do parser **e** a rotina; prioriza, explica impacto e sugere correção genérica (nunca para conteúdo safety); separa "confirmado pelo parser" de "hipótese do modelo".
  - `document`: gera documentação em Markdown: descrição funcional por programa/rotina, lista de I/O (tags de módulo com descrição), lista de timers/contadores com presets, intertravamentos principais; exportável.
  - `chat`: pergunta livre, com histórico; usa retrieval para trazer rotinas.
- **Regras no `system.md`** (escreva-as explicitamente): responder no idioma configurado; só citar tags e rotinas presentes no índice — se não encontrar, dizer que não encontrou; citar `Programa/Rotina rung N` em toda afirmação sobre lógica; não inventar presets, endereços ou comportamento de hardware; para conteúdo safety, apenas explicar; terminar revisões com o aviso "análise automática, validar com engenheiro responsável".
- Registrar por chamada: modo, modelo, tokens de entrada/saída, tokens lidos do cache, latência — em um CSV local, **sem** conteúdo do programa.

## 7. Fase 3 — app Streamlit (`app/streamlit_app.py`, `app/credits.py`)

- Tela de entrada: campo de **código de acesso** (tabela SQLite `codes(code, owner, quota, used)`); sem código válido, nada funciona. Script `python -m app.credits add CODIGO --owner nome --quota 20`.
- Upload (`st.file_uploader`, apenas `.l5x`, limite em `config.toml`). Parsing com `st.cache_data` chaveado pelo hash do arquivo. Arquivo salvo em `tempfile.TemporaryDirectory` e apagado ao fim da sessão.
- Barra lateral: resumo (controlador, revisão, nº programas/rotinas/rungs/tags), contagem de findings por severidade, avisos de conteúdo protegido e safety, seletor de idioma (pt-BR/en/es), saldo de créditos.
- Área principal: seletor de programa/rotina; quatro botões — **Explicar**, **Revisar riscos e qualidade**, **Gerar documentação**, **Conversar** — e o chat (`st.chat_message`/`st.chat_input`) com histórico em `st.session_state`.
- Cada ação debita 1 crédito (documentação completa debita 3); mostrar o custo antes de executar.
- Botão para baixar o resultado em Markdown (`st.download_button`).
- Tratamento de erros visível ao usuário: arquivo inválido, rotina protegida, limite de tokens, falha de API.
- README com: como exportar L5X do Studio 5000, como rodar localmente (`streamlit run app/streamlit_app.py`), como criar códigos, como configurar modelos e chave, e como publicar no Streamlit Community Cloud.

**Critério de aceite da Fase 3:** com `samples/P80_HULL_HCS01.L5X`, um usuário com código válido consegue, em menos de 2 minutos, ver o resumo, pedir a explicação de uma rotina em português com citação de rungs, gerar o relatório de revisão com os findings do parser e baixar a documentação em Markdown.

## 8. Modo de trabalho

- Trabalhe **fase por fase**; ao terminar cada fase, mostre os testes passando e pare para revisão antes de seguir.
- Antes de escrever o parser, **abra o L5X de exemplo** e confirme nomes de elementos e atributos; se algo diferir desta especificação, siga o arquivo e anote a diferença no README.
- Nunca grave, versione ou envie para qualquer serviço o conteúdo de `samples/`.
- Código em inglês (identificadores, comentários curtos); textos de UI, prompts e README em português.
- Commits pequenos por fase: `feat(l5x_core): parser`, `feat(l5x_core): analysis rules`, `feat(app): llm layer`, `feat(app): streamlit ui`.
- Se precisar de uma decisão de produto (ex.: como tratar um tipo de rotina não previsto), pergunte em vez de assumir.
