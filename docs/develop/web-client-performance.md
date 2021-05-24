---
title: Melhorar o desempenho de seus Office Scripts
description: Crie scripts mais rápidos compreendendo a comunicação entre a Excel de trabalho e seu script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544988"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Melhorar o desempenho de seus Office Scripts

O objetivo Office scripts é automatizar séries de tarefas comumente executadas para economizar tempo. Um script lento pode parecer que ele não acelera seu fluxo de trabalho. Na maioria das vezes, seu script ficará perfeitamente bem e será executado conforme o esperado. No entanto, há alguns cenários evitáveis que podem afetar o desempenho.

O motivo mais comum para um script lento é a comunicação excessiva com a workbook. Seu script é executado em sua máquina local, enquanto a workbook existe na nuvem. Em determinados momentos, seu script sincroniza seus dados locais com os da workbook. Isso significa que todas as operações de gravação (como ) só serão aplicadas à agenda de trabalho quando essa sincronização de `workbook.addWorksheet()` bastidores acontecer. Da mesma forma, qualquer operação de leitura (como ) só obter dados da agenda de `myRange.getValues()` trabalho para o script naqueles momentos. Em ambos os casos, o script busca informações antes de agir nos dados. Por exemplo, o código a seguir registrará com precisão o número de linhas no intervalo usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office As APIs de scripts garantem que quaisquer dados na workbook ou script sejam precisos e atualizados quando necessário. Você não precisa se preocupar com essas sincronizações para que seu script seja executado corretamente. No entanto, uma conscientização dessa comunicação entre scripts e nuvem pode ajudá-lo a evitar chamadas de rede não precisas.

## <a name="performance-optimizations"></a>Otimizações de desempenho

Você pode aplicar técnicas simples para ajudar a reduzir a comunicação com a nuvem. Os padrões a seguir ajudam a acelerar seus scripts.

- Leia dados de uma vez em vez de repetidamente em um loop.
- Remova instruções `console.log` desnecessárias.
- Evite usar blocos try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Ler dados de uma agenda de trabalho fora de um loop

Qualquer método que obtém dados da agenda de trabalho pode disparar uma chamada de rede. Em vez de fazer a mesma chamada repetidamente, você deve salvar dados localmente sempre que possível. Isso é especialmente verdadeiro ao lidar com loops.

Considere um script para obter a contagem de números negativos no intervalo usado de uma planilha. O script precisa iterar em todas as células no intervalo usado. Para fazer isso, ele precisa do intervalo, do número de linhas e do número de colunas. Você deve armazená-los como variáveis locais antes de iniciar o loop. Caso contrário, cada iteração do loop força um retorno à workbook.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> Como um experimento, tente substituir `usedRangeValues` no loop por `usedRange.getValues()` . Você pode notar que o script leva consideravelmente mais tempo para ser executado ao lidar com intervalos grandes.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Evite usar `try...catch` blocos em loops ou ao redor

Não recomendamos o uso de [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instruções em loops ou loops ao redor. Isso ocorre pelo mesmo motivo pelo qual você deve evitar a leitura de dados em um loop: cada iteração força o script a sincronizar com a workbook para garantir que nenhum erro tenha sido lançado. A maioria dos erros pode ser evitada verificando objetos retornados da workbook. Por exemplo, o script a seguir verifica se a tabela retornada pela lista de trabalho existe antes de tentar adicionar uma linha.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

### <a name="remove-unnecessary-consolelog-statements"></a>Remover instruções `console.log` desnecessárias

O log de console é uma ferramenta vital [para depurar seus scripts.](../testing/troubleshooting.md) No entanto, ele força o script a sincronizar com a manual de trabalho para garantir que as informações registradas estejam atualizadas. Considere remover instruções de registro em log desnecessárias (como as usadas para testes) antes de compartilhar seu script. Isso normalmente não causará um problema de desempenho perceptível, a menos que `console.log()` a instrução esteja em um loop.

## <a name="case-by-case-help"></a>Ajuda caso a caso

À medida que Office plataforma scripts se expande para trabalhar com [Power Automate,](https://flow.microsoft.com/) [Cartões](/adaptive-cards)Adaptáveis e outros recursos entre produtos, os detalhes da comunicação script-workbook se tornam mais complexos. Se você precisar de ajuda para tornar seu script mais rápido, entre em contato com o [Microsoft Q&A](/answers/topics/office-scripts-dev.html). Certifique-se de marcar sua pergunta com "office-scripts-dev" para que os especialistas possam encontrá-la e ajudar.

## <a name="see-also"></a>Confira também

- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [Documentos da Web do MDN: Loops e iteração](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
