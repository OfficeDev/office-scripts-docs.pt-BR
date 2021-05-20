---
title: Roteiros básicos para roteiros Office em Excel na Web
description: Uma coleção de amostras de código para usar com scripts Office em Excel na Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f252934a92126212b9520223826b3b2f5161ed57
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545756"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Roteiros básicos para roteiros Office em Excel na Web

As amostras a seguir são scripts simples para você experimentar em seus próprios livros de trabalho. Para usá-los em Excel na Web:

1. Abra a guia **Automação**.
2. **Editor de Código de Imprensa**.
3. Pressione **o novo script** no painel de tarefas do Editor de Código.
4. Substitua todo o script com a amostra de sua escolha.
5. Pressione **executar** no painel de tarefas do Editor de Código.

## <a name="script-basics"></a>Conceitos básicos

Essas amostras demonstram blocos fundamentais de construção para Office Scripts. Adicione isso aos seus scripts para estender sua solução e resolver problemas comuns.

### <a name="read-and-log-one-cell"></a>Leia e registre uma célula

Esta amostra lê o valor de **A1** e imprime-o no console.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a>Leia a célula ativa

Este script registra o valor da célula ativa atual. Se várias células forem selecionadas, a célula mais à esquerda será registrada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Alterar uma célula adjacente

Este script recebe células adjacentes usando referências relativas. Observe que se a célula ativa estiver na linha superior, parte do script falha, pois faz referência à célula acima da atualmente selecionada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a>Alterar todas as células adjacentes

Este script copia a formatação na célula ativa para as células vizinhas. Observe que este script só funciona quando a célula ativa não está na borda da planilha.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a>Alterar cada célula individual em um intervalo

Este script gira em loops ao longo do intervalo selecionado no momento. Ele limpa a formatação atual e define a cor de preenchimento em cada célula para uma cor aleatória.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Obter grupos de células com base em critérios especiais

Este script recebe todas as células em branco na faixa usada da planilha atual. Em seguida, destaca todas essas células com um fundo amarelo.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a>Coleções

Essas amostras trabalham com coleções de objetos na pasta de trabalho.

### <a name="iterate-over-collections"></a>Iterar sobre coleções

Este script recebe e registra os nomes de todas as planilhas na pasta de trabalho. Ele também define as cores de suas guias para uma cor aleatória.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="query-and-delete-from-a-collection"></a>Consultar e excluir de uma coleção

Este roteiro cria uma nova planilha. Ele verifica uma cópia existente da planilha e exclui-a antes de fazer uma nova folha.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a>Datas

As amostras nesta seção mostram como usar o [objeto](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) Data JavaScript.

A amostra a seguir obtém a data e a hora atuais e, em seguida, escreve esses valores para duas células na planilha ativa.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

A próxima amostra lê uma data armazenada em Excel e a traduz para um objeto JavaScript Date. Ele usa o [número de série numérico da data](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para a Data JavaScript.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a>Exibir dados

Essas amostras demonstram como trabalhar com dados de planilhas e fornecem aos usuários uma melhor visão ou organização.

### <a name="apply-conditional-formatting"></a>Aplicar formatação condicional

Esta amostra aplica formatação condicional à faixa utilizada atualmente na planilha. A formatação condicional é um preenchimento verde para os 10% mais altos dos valores.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a>Crie uma tabela classificada

Esta amostra cria uma tabela a partir da faixa usada da planilha atual e, em seguida, classifica-a com base na primeira coluna.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Registre os valores "Grande Total" de uma Tabela Dinâmica

Esta amostra encontra a primeira Tabela Dinâmica na pasta de trabalho e registra os valores nas células "Grande Total" (como destacado em verde na imagem abaixo).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Uma Tabela Dinâmica mostrando vendas de frutas com a linha Grand Total destacada verde":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a>Crie uma lista de drop-down usando validação de dados

Este script cria uma lista de seleção para uma célula. Ele usa os valores existentes da faixa selecionada como as opções para a lista.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Uma planilha mostrando uma gama de três células contendo escolhas de cores 'vermelho, azul, verde' e ao lado dela, as mesmas escolhas mostradas em uma lista de drop-down":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a>Fórmulas

Essas amostras usam fórmulas Excel e mostram como trabalhar com elas em scripts.

### <a name="single-formula"></a>Fórmula única

Este script define a fórmula de uma célula e, em seguida, exibe como Excel armazena a fórmula e o valor da célula separadamente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Manuseie um `#SPILL!` erro retornado de uma fórmula

Este script transpõe a faixa "A1:D2" para "A4:B7" usando a função TRANSPOSE. Se a transposição resultar em um `#SPILL` erro, ela limpa o intervalo de destino e aplica a fórmula novamente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a>Sugira novas amostras

Damos boas-vindas às sugestões de novas amostras. Se houver um cenário comum que ajude outros desenvolvedores de script, por favor, diga-nos na seção de feedback na parte inferior da página.

## <a name="see-also"></a>Confira também

* ["Range basics" de Sudhi Ramamurthy no YouTube](https://youtu.be/4emjkOFdLBA)
* [Office Scripts amostras e cenários](samples-overview.md)
* [Gravar, editar e criar scripts do Office no Excel na Web](../../tutorials/excel-tutorial.md)
