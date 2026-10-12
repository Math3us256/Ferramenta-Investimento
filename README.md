# Ferramenta de Investimentos em FIIs

## Sobre o projeto

Esta pasta contém uma ferramenta de investimentos desenvolvida em Microsoft Excel. Ela permite fazer projeções de patrimônio a partir de aportes mensais e de uma taxa de rendimento mensal, estimar dividendos mensais e visualizar uma sugestão de distribuição do aporte entre categorias de Fundos de Investimento Imobiliário (FIIs), conforme um dos perfis cadastrados: **Conservador**, **Moderado** ou **Agressivo**.

Os resultados são simulações baseadas nas premissas inseridas na planilha. Não representam garantia de rentabilidade nem recomendação individual de investimento.

## Arquivos e abas

- `ferramenta-investimentos.xlsx`: arquivo principal do projeto.
- `README.md`: documentação das perguntas respondidas pela ferramenta, das fórmulas e dos parâmetros utilizados.

O arquivo Excel possui duas abas:

- **Ferramenta**: área principal para inserir ou consultar premissas, visualizar projeções por prazo e conferir a alocação sugerida por perfil.
- **Tabela de Apoio**: tabela de percentuais de alocação por perfil e categoria de FII, utilizada nas buscas de percentuais.

## Quais perguntas a ferramenta responde?

| Pergunta | Onde aparece a resposta | Como a resposta é obtida |
|---|---|---|
| Qual é a sugestão de investimento correspondente a 30% do salário? | Aba **Ferramenta**, célula `D15` | A fórmula calcula 30% do salário informado em `D13`. O resultado recebe sinal negativo na fórmula para representar uma saída/aporte. |
| Quanto será investido por mês na simulação? | Aba **Ferramenta**, célula `D18` | É o valor mensal informado para o aporte. A sugestão calculada em `D15` não substitui automaticamente esse valor. |
| Qual prazo será usado na projeção principal? | Aba **Ferramenta**, célula `D19` | Número de anos escolhido na planilha. |
| Qual taxa mensal é usada para projetar o patrimônio? | Aba **Ferramenta**, célula `D20` | Taxa informada como premissa da projeção. No arquivo fornecido, o valor é `0,01079`, ou aproximadamente **1,079% ao mês**. |
| Qual patrimônio acumulado é projetado para o prazo selecionado? | Aba **Ferramenta**, célula `D21` | A função de valor futuro (`VF`, chamada `FV` no arquivo) usa a taxa mensal, o prazo convertido em meses e o aporte mensal. |
| Qual é a estimativa de dividendos mensais do patrimônio projetado? | Aba **Ferramenta**, célula `D22` | Patrimônio projetado em `D21` multiplicado pelo rendimento mensal da carteira informado em `D14`. No arquivo, `D14` contém `0,0089`, ou **0,89% ao mês**. |
| Como patrimônio e dividendos podem variar em outros prazos? | Aba **Ferramenta**, linhas `25:29`, colunas `B:D` | A coluna `B` mostra os prazos (2, 5, 10, 20 e 30 anos); a coluna `C` calcula o patrimônio projetado; a coluna `D` estima os dividendos mensais. |
| Como distribuir o aporte entre as categorias de FIIs? | Aba **Ferramenta**, células `C31:D41` | O perfil é selecionado em `C31`; `C35:C40` mostra os percentuais por categoria; `D35:D40` mostra os valores correspondentes; `D41` soma os valores alocados. |

## Como a função VF entra nos cálculos

A função **VF** (Valor Futuro) projeta o valor futuro de uma série de aportes periódicos considerando uma taxa constante. No arquivo enviado, o nome da função aparece como `FV`, que é o nome em inglês equivalente a `VF`.

Na aba **Ferramenta**, a função aparece nestes locais:

- **`D21`**: `=FV(taxa_mensal,prazo_anos*12,Aporte)` — calcula o patrimônio projetado para o prazo escolhido em `D19`.
- **`C25:C29`**: fórmulas `FV` semelhantes, que usam o prazo de cada linha (`B25:B29`) multiplicado por 12 para obter o número de meses.

Em Excel configurado para português, a expressão equivalente é apresentada como `VF`, por exemplo:

```excel
=VF(taxa_mensal; prazo_anos*12; Aporte)
```

Os argumentos significam:

- `taxa_mensal`: taxa mensal de projeção, definida em `Ferramenta!D20`;
- `prazo_anos*12`: conversão do prazo de anos para meses;
- `Aporte`: valor de aporte mensal definido em `Ferramenta!D18`.

O aporte possui sinal negativo na fórmula porque representa dinheiro aplicado em cada período; assim, o resultado futuro é apresentado como positivo. Como o argumento opcional relativo ao momento do pagamento não é especificado, a função considera pagamentos no fim de cada período. A projeção pressupõe que a taxa mensal permaneça constante durante todo o prazo.

## Como a função PROCV entra nos cálculos

A função **PROCV** procura uma chave na primeira coluna de uma tabela e devolve um valor de outra coluna da mesma linha. No arquivo fornecido, o nome da função aparece em inglês como `VLOOKUP`, equivalente a `PROCV` no Excel em português.

Na aba **Ferramenta**, as fórmulas de `C35:C40` combinam o perfil selecionado (`perfil_selecionado`) e a categoria de FII da coluna `B` para criar chaves como `Conservador-PAPEL`. A busca é feita na tabela `Tabela de Apoio!A3:D21`, retorna a quarta coluna (percentual) e usa correspondência exata (`FALSE`, equivalente a `FALSO`).

Depois da busca:

- `C35:C40` apresenta o percentual de cada categoria para o perfil selecionado;
- `D35:D40` calcula o valor destinado a cada categoria com base no aporte mensal e no percentual retornado;
- `D41` soma os valores de `D35:D40`.

Na aba **Tabela de Apoio**, `A4` contém a fórmula `=B4&"-"&C4`, que concatena o perfil e a categoria para montar a chave. No arquivo atual, as demais chaves de `A5:A21` estão preenchidas como valores. Existe também uma busca auxiliar em `H4`, que usa a chave informada em `G4` e consulta a tabela de apoio.

## Intervalos nomeados

A pasta de trabalho contém os seguintes intervalos nomeados:

| Nome definido | Referência | Finalidade |
|---|---|---|
| `salario` | `Ferramenta!$D$13` | Salário usado na sugestão de investimento de 30%. |
| `Rendimento_Carteira` | `Ferramenta!$D$14` | Rendimento usado para estimar os dividendos mensais. |
| `sugestao_investimento` | `Ferramenta!$D$15` | Valor sugerido a partir de 30% do salário. |
| `Aporte` | `Ferramenta!$D$18` | Aporte mensal usado na projeção. |
| `prazo_anos` | `Ferramenta!$D$19` | Prazo da projeção, em anos. |
| `taxa_mensal` | `Ferramenta!$D$20` | Taxa mensal utilizada na função de valor futuro. |
| `patrimonio_acumulado` | `Ferramenta!$D$21` | Resultado da projeção de patrimônio. |
| `perfil_selecionado` | `Ferramenta!$C$31` | Perfil usado na busca dos percentuais de alocação. |
| `valor_investir_mensal` | `Ferramenta!$C$32` | Valor mensal referenciado nos cálculos de alocação. |

## Percentuais de cada perfil e origem

Os percentuais abaixo estão cadastrados na aba **Tabela de Apoio**, nas colunas `B:D`, e são consultados de acordo com o perfil selecionado na aba principal.

| Categoria de FII | Conservador | Moderado | Agressivo |
|---|---:|---:|---:|
| Papel | 30% | 32% | 50% |
| Tijolo | 50% | 35% | 10% |
| Híbridos | 10% | 8% | 5% |
| FOFs | 10% | 5% | 5% |
| Desenvolvimento | 0% | 10% | 20% |
| Hotelarias | 0% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

**Origem informada para este projeto:** os percentuais foram obtidos do **material ou das aulas do curso**. O arquivo Excel não registra uma referência específica, como nome de aula, página, módulo ou link. Caso a entrega exija uma referência bibliográfica detalhada, ela deve ser completada com os dados exatos do material consultado.

Esses percentuais são premissas utilizadas pela ferramenta para distribuir o aporte; a presença deles na planilha não significa que sejam uma recomendação financeira universal ou personalizada.

## O que mudou em relação à ferramenta do Expert?

Conforme informado para esta documentação, a alteração realizada em relação à ferramenta do Expert foi no **layout/apresentação visual**. Não foram fornecidos detalhes adicionais sobre elementos visuais específicos, portanto este README não atribui mudanças particulares de cores, gráficos ou componentes. Também não afirma que fórmulas ou percentuais foram modificados em relação à versão original, pois isso não foi informado.

## Como utilizar

1. Na aba **Ferramenta**, confira as premissas de salário, rendimento da carteira, aporte mensal, prazo e taxa mensal.
2. Consulte o patrimônio projetado e a estimativa de dividendos mensais.
3. Compare os cenários de 2, 5, 10, 20 e 30 anos.
4. Selecione o perfil e consulte os percentuais e valores sugeridos por categoria de FII.
5. Consulte a aba **Tabela de Apoio** para verificar os percentuais associados a cada perfil.

## Limitações da simulação

A projeção considera as taxas informadas como constantes e não representa necessariamente a volatilidade, os custos, os impostos, a inflação ou as mudanças de rendimento dos ativos ao longo do tempo. Os valores projetados devem ser interpretados como simulações matemáticas, não como resultados garantidos.
