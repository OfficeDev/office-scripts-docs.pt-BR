---
title: Melhorar o desempenho dos scripts do Office
description: Crie scripts mais rápidos compreendendo a comunicação entre a planilha do Excel e seu script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867867"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Melhorar o desempenho dos scripts do Office

O objetivo dos Scripts do Office é automatizar uma série de tarefas normalmente realizadas para economizar tempo. Um script lento pode parecer que ele não acelera seu fluxo de trabalho. Na maioria das vezes, seu script ficará perfeitamente bem e será executado conforme o esperado. No entanto, há alguns cenários que podem afetar o desempenho.

O motivo mais comum para um script lento é a comunicação excessiva com a agenda. O script é executado no computador local, enquanto a agenda existe na nuvem. Em determinados momentos, seu script sincroniza seus dados locais com os da agenda. Isso significa que todas as operações de gravação (como) serão aplicadas somente à plano de trabalho quando essa sincronização nos `workbook.addWorksheet()` bastidores ocorrer. Da mesma forma, qualquer operação de leitura (como) só obter dados da área de trabalho `myRange.getValues()` para o script nesses momentos. Em ambos os casos, o script busca informações antes de agir sobre os dados. Por exemplo, o código a seguir registrará com precisão o número de linhas no intervalo usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

As APIs de scripts do Office garantem que todos os dados da lista de trabalho ou script sejam precisos e atualizados quando necessário. Você não precisa se preocupar com essas sincronizações para que seu script seja executado corretamente. No entanto, um reconhecimento dessa comunicação entre scripts e nuvem pode ajudá-lo a evitar chamadas de rede não precisas.

## <a name="performance-optimizations"></a>Otimizações de desempenho

Você pode aplicar técnicas simples para ajudar a reduzir a comunicação com a nuvem. Os seguintes padrões ajudam a acelerar seus scripts.

- Ler dados de uma vez em vez de repetidamente em um loop.
- Remova instruções `console.log` desnecessárias.
- Evite usar blocos try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Ler dados da área de trabalho fora de um loop

Qualquer método que obtém dados da agenda pode disparar uma chamada de rede. Em vez de fazer repetidamente a mesma chamada, você deve salvar dados localmente sempre que possível. Isso é especialmente verdadeiro ao lidar com loops.

Considere um script para obter a contagem de números negativos no intervalo usado de uma planilha. O script precisa iterar em todas as células no intervalo usado. Para fazer isso, ele precisa do intervalo, do número de linhas e do número de colunas. Você deve armazená-los como variáveis locais antes de iniciar o loop. Caso contrário, cada iteração do loop força um retorno à agenda.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Remover instruções `console.log` desnecessárias

O log do console é uma ferramenta vital [para depurar seus scripts.](../testing/troubleshooting.md) No entanto, ele força o script a sincronizar com a agenda para garantir que as informações registradas estejam atualizadas. Considere remover instruções de registro em log desnecessárias (como aquelas usadas para teste) antes de compartilhar seu script. Isso normalmente não causará um problema de desempenho perceptível, a menos que `console.log()` a instrução esteja em um loop.

### <a name="avoid-using-trycatch-blocks"></a>Evite usar blocos try/catch

Não recomendamos o uso de [ `try` / `catch` blocos](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) como parte do fluxo de controle esperado de um script. A maioria dos erros pode ser evitada verificando objetos retornados da agenda. Por exemplo, o script a seguir verifica se a tabela retornada pela lista de trabalho existe antes de tentar adicionar uma linha.

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

## <a name="case-by-case-help"></a>Ajuda caso a caso

À medida que a plataforma de Scripts do Office se expande para trabalhar com o [Power Automate](https://flow.microsoft.com/), Cartões [Adaptáveis](/adaptive-cards)e outros recursos entre produtos, os detalhes da comunicação entre as guias de trabalho de script se tornam mais complexos. Se precisar de ajuda para fazer seu script ser executado mais rapidamente, entre em contato com o [Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts) Certifique-se de marcar sua pergunta com "office-scripts" para que os especialistas possam encontrá-la e ajudar.

## <a name="see-also"></a>Confira também

- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [Documentos da Web do MDN: Loops e iteração](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)