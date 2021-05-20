---
title: Melhore o desempenho de seus scripts de Office
description: Crie scripts mais rápidos entendendo a comunicação entre a Excel pasta de trabalho e seu script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544988"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Melhore o desempenho de seus scripts de Office

O objetivo do Office Scripts é automatizar séries comumente executadas de tarefas para economizar tempo. Um script lento pode parecer que não acelera seu fluxo de trabalho. Na maioria das vezes, seu roteiro estará perfeitamente bem e será executado como esperado. No entanto, existem alguns cenários evitáveis que podem afetar o desempenho.

A razão mais comum para um script lento é a comunicação excessiva com a pasta de trabalho. Seu script é executado em sua máquina local, enquanto a pasta de trabalho existe na nuvem. Em certos momentos, seu script sincroniza seus dados locais com os da pasta de trabalho. Isso significa que qualquer operação de gravação (como `workbook.addWorksheet()` ) só é aplicada à pasta de trabalho quando essa sincronização nos bastidores acontece. Da mesma forma, qualquer operação de leitura (como `myRange.getValues()` ) só recebe dados da pasta de trabalho para o script nesses momentos. Em ambos os casos, o script busca informações antes de agir sobre os dados. Por exemplo, o código a seguir registrará com precisão o número de linhas no intervalo usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office As APIs de scripts garantem que quaisquer dados na pasta de trabalho ou script são precisos e atualizados quando necessário. Você não precisa se preocupar com essas sincronizações para que seu script seja executado corretamente. No entanto, uma consciência dessa comunicação script-to-cloud pode ajudá-lo a evitar chamadas de rede não fornecidas.

## <a name="performance-optimizations"></a>Otimizações de desempenho

Você pode aplicar técnicas simples para ajudar a reduzir a comunicação à nuvem. Os seguintes padrões ajudam a acelerar seus scripts.

- Leia dados da pasta de trabalho uma vez em vez de repetidamente em um loop.
- Remova `console.log` declarações desnecessárias.
- Evite usar blocos de tentativa/captura.

### <a name="read-workbook-data-outside-of-a-loop"></a>Leia dados da pasta de trabalho fora de um loop

Qualquer método que obtenha dados da pasta de trabalho pode acionar uma chamada de rede. Em vez de fazer repetidamente a mesma chamada, você deve salvar os dados localmente sempre que possível. Isso é especialmente verdade quando se lida com loops.

Considere um script para obter a contagem de números negativos na faixa usada de uma planilha. O script precisa iterar sobre cada célula da gama usada. Para isso, precisa do intervalo, do número de linhas e do número de colunas. Você deve armazená-los como variáveis locais antes de iniciar o loop. Caso contrário, cada iteração do loop forçará um retorno à pasta de trabalho.

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
> Como um experimento, tente substituir `usedRangeValues` no loop com `usedRange.getValues()` . Você pode notar que o script leva consideravelmente mais tempo para ser executado ao lidar com grandes faixas.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Evite usar `try...catch` blocos em loops ou ao redor

Não recomendamos o uso [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) de declarações em loops ou loops circundantes. Isso é pela mesma razão que você deve evitar ler dados em um loop: cada iteração força o script a sincronizar com a pasta de trabalho para garantir que nenhum erro tenha sido jogado. A maioria dos erros pode ser evitada verificando objetos retornados da pasta de trabalho. Por exemplo, o script a seguir verifica se a tabela devolvida pela pasta de trabalho existe antes de tentar adicionar uma linha.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Remover `console.log` declarações desnecessárias

O registro de consoles é uma ferramenta vital para [depurar seus scripts](../testing/troubleshooting.md). No entanto, ele força o script a sincronizar com a pasta de trabalho para garantir que as informações registradas estão atualizadas. Considere remover declarações de registro desnecessárias (como as usadas para testes) antes de compartilhar seu script. Isso normalmente não causará um problema de desempenho perceptível, a menos que a `console.log()` instrução esteja em um loop.

## <a name="case-by-case-help"></a>Ajuda caso a caso

À medida que a plataforma Office Scripts se expande para trabalhar com [Power Automate,](https://flow.microsoft.com/) [Cartões Adaptativos](/adaptive-cards)e outros recursos entre produtos, os detalhes da comunicação script-workbook se tornam mais complexos. Se você precisar de ajuda para fazer seu script funcionar mais rápido, entre em contato com [o Microsoft Q&A](/answers/topics/office-scripts-dev.html). Certifique-se de marcar sua pergunta com "office-scripts-dev" para que os especialistas possam encontrá-la e ajudar.

## <a name="see-also"></a>Confira também

- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [MDN web docs: Loops e iteração](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
