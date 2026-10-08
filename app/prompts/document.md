Gere documentação técnica em Markdown para o escopo abaixo, a partir do índice do parser. O documento será exportado e lido por engenheiros e técnicos que não têm o projeto aberto.

Escopo: {target}

Estrutura obrigatória (use exatamente estes títulos de nível 2; omita uma seção só se não houver nada a dizer, escrevendo "Nenhum item no índice"):

## Visão geral
Para que serve este escopo, como está organizado (tasks, programas, rotinas principais), em qual task roda e com qual período, se constar.

## Descrição funcional por rotina
Para cada rotina analisada, um bloco curto: objetivo (hipótese do modelo, salvo descrição explícita), entradas principais, saídas principais, blocos de lógica com citação de rungs. Rotinas FBD/SFC: só nome, descrição e tags referenciadas. Rotinas protegidas: registrar como "conteúdo protegido".

## Lista de I/O
Tabela com as referências de I/O de módulo encontradas (módulo, catálogo, descrição do módulo, rotinas que usam). Use apenas a tabela de I/O fornecida abaixo; não invente endereços.

## Timers e contadores
Tabela: tag, instrução, preset, rotina/rung, finalidade (hipótese). Use a tabela do parser abaixo como base; preset `?` = não exposto no índice.

## Intertravamentos e permissivos principais
Lista das condições que bloqueiam ou liberam as saídas mais importantes, com citação de rungs. Marque como hipótese o que for interpretação.

## Observações do parser
Resumo dos findings relevantes (rotinas órfãs, saídas duplicadas, timers sem preset, programas não agendados), com localização.

Regras: cite `Programa/Rotina rung N` em toda afirmação de lógica; mantenha identificadores originais; conteúdo SAFETY só descrito, nunca com sugestão de mudança.

### Tabela de I/O (parser)
{io_table}

### Timers e contadores (parser)
{timer_table}

### Findings (parser)
{findings}

### Rotinas
{content}
