Analise, documente e prepare o debug da Add-On Instruction (AOI) abaixo. Você recebe (a) a definição e a interface da AOI com estatísticas do parser, (b) os FINDINGS determinísticos específicos desta AOI, e (c) o texto das rotinas internas (Logic, Prescan, Postscan, EnableInFalse). Cite a lógica interna no formato `aoi:{target}/Rotina rung N` (ou `linha N` em ST) e os pontos de chamada como `Programa/Rotina rung N` (ou `folha N` em FBD).

AOI: {target}

Estrutura obrigatória (títulos de nível 2; escreva "Nenhum item no índice" quando não houver nada):

## Finalidade
Dois ou três parágrafos: o que a AOI faz, para que tipo de equipamento/sinal serve, como é usada no projeto (quantas instâncias, em quais programas, aninhada em quais AOIs). Separe a descrição oficial (do índice) da sua interpretação (hipótese do modelo).

## Interface
Tabela: parâmetro | uso (Input/Output/InOut) | tipo | obrigatório na chamada | default | função (hipótese curta) | observação do parser (L/E, não usado, nunca escrito). Depois, as tags locais relevantes (timers, contadores, configurações com default, status visíveis ao HMI). Não invente defaults: use só os do índice.

## Comportamento
- Lógica rung a rung ou por bloco funcional (para rotinas grandes agrupe), sempre com citação.
- O que cada saída depende (condições que ligam e desligam), com citação.
- Efeito de EnableIn falso (há rotina EnableInFalse? o que ela faz?), Prescan/Postscan.
- Timers e contadores: preset (do índice ou do default da tag local), o que disparam.

## Uso no projeto
Resumo dos pontos de chamada: instâncias sem chamada, instâncias chamadas em mais de um lugar, pinos/argumentos que se repetem, membros lidos/escritos fora da AOI (interface "informal" via tag de instância). Use só a lista do índice; se ela estiver truncada, diga.

## Findings confirmados pelo parser
Priorize (Alta/Média/Baixa), explique o impacto prático e sugira correção genérica, sem código. Agrupe findings repetitivos. Conteúdo SAFETY: só descreva.

## Hipóteses do modelo
Riscos que o parser não detecta: condição de corrida entre rungs, saída sem reset, ONS em ramo, preset zero, parâmetro Required que a lógica sobrescreve, timer sem retenção, dependência de ordem de chamada etc. Marque cada item como hipótese e cite o rung. Se não houver, diga.

## Roteiro de debug
Lista numerada, orientada ao técnico com a instância online aberta no Studio 5000:
1. Como confirmar se a instância está sendo executada (EnableIn, rung de chamada, EnableOut).
2. Para cada saída principal: quais membros/entradas observar e em que rung a decisão é tomada.
3. Que parâmetros de configuração (Visible, não obrigatórios) costumam causar comportamento inesperado e quais valores conferir.
4. Sintomas comuns → causa provável → onde olhar (tabela curta).
5. O que NÃO deve ser alterado online sem análise (saídas múltiplas, latches, conteúdo SAFETY).

Regras: identificadores exatamente como no índice; separe "confirmado pelo parser" de "hipótese do modelo"; não descreva lógica de rotinas FBD/SFC ou protegidas como se a tivesse visto; termine com "Análise automática: validar com o engenheiro responsável antes de qualquer alteração."

Tamanho: seja denso, não extenso. Não transcreva rungs, tabelas de parâmetros nem a lista de instâncias (o leitor tem o relatório determinístico); em rotinas grandes agrupe por bloco funcional. Alvo: até cerca de 2.500 palavras.

### Definição e interface (parser)
{interface}

### Findings da AOI (parser)
{findings}

### Rotinas internas
{content}
