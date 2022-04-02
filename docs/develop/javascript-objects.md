---
title: Usar objetos internos do JavaScript nos scripts do Office
description: Como chamar APIs JavaScript integrados de um script Office no Excel na Web.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 620b97660eb07fd1289ab3aafcae1acaed43ed2f
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585727"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Usar objetos JavaScript integrados em Office Scripts

JavaScript fornece vários objetos integrados que você pode usar em seus scripts de Office, independentemente de você estar fazendo scripts em JavaScript ou [TypeScript](../overview/code-editor-environment.md) (um superconjunto de JavaScript). Este artigo descreve como você pode usar alguns dos objetos JavaScript integrados Office Scripts para Excel na Web.

> [!NOTE]
> Para uma lista completa de todos os objetos JavaScript integrados, consulte o artigo Objetos integrados [Standard do](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla.

## <a name="array"></a>Array

O [objeto Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) fornece uma maneira padronizada de trabalhar com matrizes em seu script. Embora as matrizes sejam construções JavaScript padrão, elas se relacionam Office scripts de duas maneiras principais: intervalos e coleções.

### <a name="work-with-ranges"></a>Trabalhar com intervalos

Os intervalos contêm várias matrizes bidimensionais que mapeiam diretamente para as células nesse intervalo. Essas matrizes contêm informações específicas sobre cada célula nesse intervalo. Por exemplo, `Range.getValues` retorna todos os valores nessas células (com as linhas e colunas do mapeamento de matriz bidimensional para as linhas e colunas dessa subseção de planilha). `Range.getFormulas` e `Range.getNumberFormats` são outros métodos usados com frequência que retornam matrizes como `Range.getValues`.

O script a seguir pesquisa o intervalo **A1:D4** para qualquer formato de número que contenha um "$". O script define a cor de preenchimento nessas células como "amarela".

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

Muitos Excel objetos estão contidos em uma coleção. A coleção é gerenciada pela API Office Scripts e exposta como uma matriz. Por exemplo, todas [as Formas](/javascript/api/office-scripts/excelscript/excelscript.shape) em uma planilha estão contidas em um `Shape[]` que é retornado pelo `Worksheet.getShapes` método. Você pode usar essa matriz para ler valores da coleção ou acessar objetos específicos dos métodos do `get*` objeto pai.

> [!NOTE]
> Não adicione ou remova objetos manualmente dessas matrizes de coleção. Use os `add` métodos nos objetos pai e nos `delete` métodos nos objetos do tipo coleção. Por exemplo, adicione uma [Tabela](/javascript/api/office-scripts/excelscript/excelscript.table) a [uma Planilha com](/javascript/api/office-scripts/excelscript/excelscript.worksheet) o `Worksheet.addTable` método e remova o `Table` uso `Table.delete`.

O script a seguir registra o tipo de cada forma na planilha atual.

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

O [objeto Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fornece uma maneira padronizada de trabalhar com datas em seu script. `Date.now()` gera um objeto com a data e a hora atuais, o que é útil ao adicionar data/hora à entrada de dados do script.

O script a seguir adiciona a data atual à planilha. Observe que, usando o método`toLocaleDateString`, Excel reconhece o valor como uma data e altera automaticamente o formato de número da célula.

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

A [seção Trabalhar com datas](../resources/samples/excel-samples.md#dates) dos exemplos tem mais scripts relacionados à data.

## <a name="math"></a>Matemática

O [objeto Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fornece métodos e constantes para operações matemáticas comuns. Elas fornecem muitas funções também disponíveis no Excel, sem a necessidade de usar o mecanismo de cálculo da agenda de trabalho. Isso salva o script de ter que consultar a workbook, o que melhora o desempenho.

O script a seguir usa `Math.min` para encontrar e registrar o menor número no intervalo **A1:D4** . Observe que este exemplo supõe que todo o intervalo contém apenas números, não cadeias de caracteres.

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>Não há suporte para o uso de bibliotecas JavaScript externas

Office scripts não suportam o uso de bibliotecas externas de terceiros. Seu script só pode usar os objetos JavaScript integrados e as APIs Office Scripts.

## <a name="see-also"></a>Confira também

- [Objetos integrados padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office do Editor de Código de Scripts](../overview/code-editor-environment.md)
