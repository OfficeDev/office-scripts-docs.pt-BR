---
title: Usar objetos internos do JavaScript nos scripts do Office
description: Como chamar APIs JavaScript integrados de um script Office em Excel na Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545044"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Use objetos JavaScript incorporados em scripts Office

JavaScript fornece vários objetos incorporados que você pode usar em seus scripts Office, independentemente de você estar fazendo scripts em JavaScript ou [TypeScript](../overview/code-editor-environment.md) (um superconjunto de JavaScript). Este artigo descreve como você pode usar alguns dos objetos JavaScript incorporados em scripts Office para Excel na Web.

> [!NOTE]
> Para obter uma lista completa de todos os objetos JavaScript incorporados, consulte o artigo [de objetos incorporados Padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) da Mozilla.

## <a name="array"></a>Matriz

O objeto [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) fornece uma maneira padronizada de trabalhar com arrays em seu script. Embora os arrays sejam construções JavaScript padrão, elas se relacionam com scripts Office de duas maneiras principais: intervalos e coleções.

### <a name="work-with-ranges"></a>Trabalhar com faixas

As faixas contêm várias matrizes bidimensionais que mapeiam diretamente para as células nesse intervalo. Essas matrizes contêm informações específicas sobre cada célula nesse intervalo. Por exemplo, `Range.getValues` retorna todos os valores nessas células (com as linhas e colunas do mapeamento bidimensional do array para as linhas e colunas dessa subseção da planilha). `Range.getFormulas` e `Range.getNumberFormats` são outros métodos frequentemente usados que retornam matrizes como `Range.getValues` .

O script a seguir pesquisa a faixa **A1:D4** para qualquer formato de número que contenha um "$". O script define a cor de preenchimento nessas células como "amarela".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="work-with-collections"></a>Trabalhar com coleções

Muitos objetos Excel estão contidos em uma coleção. A coleção é gerenciada pela API Office Scripts e exposta como uma matriz. Por exemplo, todas as formas em uma planilha estão [contidas](/javascript/api/office-scripts/excelscript/excelscript.shape) em um `Shape[]` que é devolvido pelo `Worksheet.getShapes` método. Você pode usar este array para ler valores da coleção ou acessar objetos específicos dos métodos do objeto `get*` pai.

> [!NOTE]
> Não adicione manualmente ou remova objetos dessas matrizes de coleta. Use os `add` métodos nos objetos-pai e os `delete` métodos nos objetos do tipo de coleta. Por exemplo, adicione uma [tabela](/javascript/api/office-scripts/excelscript/excelscript.table) a uma [planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) com o `Worksheet.addTable` método e remova o uso `Table` `Table.delete` .

O script a seguir registra o tipo de todas as formas na planilha atual.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

O script a seguir exclui a forma mais antiga da planilha atual.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Data

O objeto [Data](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fornece uma maneira padronizada de trabalhar com datas em seu script. `Date.now()` gera um objeto com a data e a hora atuais, o que é útil ao adicionar datamps de tempo à entrada de dados do seu script.

O script a seguir adiciona a data atual à planilha. Observe que, usando o `toLocaleDateString` método, Excel reconhece o valor como uma data e altera automaticamente o formato de número da célula.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

A seção [Trabalho com datas](../resources/samples/excel-samples.md#dates) das amostras tem mais scripts relacionados a datas.

## <a name="math"></a>Matemática

O objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fornece métodos e constantes para operações matemáticas comuns. Estes fornecem muitas funções também disponíveis em Excel, sem a necessidade de usar o mecanismo de cálculo da pasta de trabalho. Isso evita que seu script tenha que consultar a pasta de trabalho, o que melhora o desempenho.

O script a seguir usa `Math.min` para encontrar e registrar o menor número na faixa **A1:D4.** Observe que esta amostra pressupõe que toda a gama contém apenas números, não strings.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>O uso de bibliotecas JavaScript externas não é suportado

Office Scripts não suportam o uso de bibliotecas externas de terceiros. Seu script só pode usar os objetos JavaScript incorporados e as APIs Office Scripts.

## <a name="see-also"></a>Confira também

- [Objetos embutidos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Ambiente do Editor de Código de Scripts](../overview/code-editor-environment.md)
