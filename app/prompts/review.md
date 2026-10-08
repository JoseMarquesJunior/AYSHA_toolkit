Revise riscos e qualidade da lógica abaixo. Você recebe (a) os FINDINGS determinísticos do parser para este escopo e (b) o texto das rotinas.

Escopo: {target}

Como responder (em Markdown):

## Resumo
Duas ou três frases: estado geral e os problemas mais graves.

## Findings confirmados pelo parser
Para cada finding relevante (agrupe os repetitivos, ex.: "12 tags escritas e nunca lidas"), uma linha ou item com:
- prioridade (Alta / Média / Baixa) e por quê;
- impacto prático (o que pode acontecer em operação);
- correção genérica sugerida (ex.: "unificar a escrita da saída em um único rung com lógica OR"), sem escrever código completo.
Cite sempre `Programa/Rotina rung N`. Findings de severidade "info" sobre documentação só entram se forem relevantes para manutenção.

## Hipóteses do modelo
Problemas que você suspeita ao ler o texto dos rungs e que o parser NÃO detecta (ex.: condição de corrida entre rungs, falta de intertravamento, reset manual ausente, timer sem retenção, uso de ONS em ramo). Marque cada um como hipótese e cite o rung. Se não houver, diga que não há.

## O que está bem
Pontos positivos objetivos (1 a 3), se existirem.

Regras:
- Conteúdo SAFETY: descreva os findings, mas NÃO proponha correção; indique que a alteração passa pelo processo de validação de segurança.
- Não repita a tabela de findings inteira; priorize.
- Termine com a frase obrigatória: "Análise automática: validar com o engenheiro responsável antes de qualquer alteração."

### Findings do parser
{findings}

### Rotinas
{content}
