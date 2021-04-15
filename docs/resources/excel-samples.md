---
title: Scripts básicos para scripts do Office no Excel na Web
description: Uma coleção de exemplos de código a ser usado com scripts do Office no Excel na Web.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 32c380a2abdb0596bbfdb387768735fbfdcf7c08
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755088"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Scripts básicos para scripts do Office no Excel na Web

Os exemplos a seguir são scripts simples para você experimentar suas próprias workbooks. Para usá-los no Excel na Web:

1. Abra a guia **Automação**.
2. Pressione **Editor de Código**.
3. Pressione **Novo Script** no painel de tarefas do Editor de Código.
4. Substitua o script inteiro pelo exemplo de sua escolha.
5. Pressione **Executar** no painel de tarefas do Editor de Código.

## <a name="scripting-basics"></a>Noções básicas de scripts

Esses exemplos demonstram blocos de construção fundamentais para scripts do Office. Adicione-os aos scripts para estender sua solução e resolver problemas comuns.

### <a name="read-and-log-one-cell"></a>Ler e registrar uma célula

Este exemplo lê o valor **de A1** e o imprime no console.

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

### <a name="read-the-active-cell"></a>Ler a célula ativa

Esse script registra o valor da célula ativa atual. Se várias células forem selecionadas, a célula mais à esquerda será registrada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Alterar uma célula adjacente

Esse script obtém células adjacentes usando referências relativas. Observe que, se a célula ativa estiver na linha superior, parte do script falhará, pois faz referência à célula acima da selecionada no momento.

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

Esse script copia a formatação na célula ativa para as células vizinhas. Observe que esse script só funciona quando a célula ativa não está na borda da planilha.

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

Esse script loops sobre o intervalo de seleção no momento. Ele limpa a formatação atual e define a cor de preenchimento em cada célula como uma cor aleatória.

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

Esse script obtém todas as células em branco no intervalo usado da planilha atual. Em seguida, realça todas as células com um plano de fundo amarelo.

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

Esses exemplos funcionam com coleções de objetos na workbook.

### <a name="iterating-over-collections"></a>Iterating over collections

Esse script obtém e registra os nomes de todas as planilhas na pasta de trabalho. Ele também define as cores da guia como uma cor aleatória.

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

### <a name="querying-and-deleting-from-a-collection"></a>Consulta e exclusão de uma coleção

Este script cria uma nova planilha. Ele verifica uma cópia existente da planilha e a exclui antes de criar uma nova planilha.

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

Os exemplos nesta seção mostram como usar o objeto Data de [JavaScript.](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)

O exemplo a seguir obtém a data e a hora atuais e grava esses valores em duas células na planilha ativa.

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

O próximo exemplo lê uma data armazenada no Excel e a converte em um objeto Data JavaScript. Ele usa o [número de série](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) numérico da data como entrada para a Data JavaScript.

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

Esses exemplos demonstram como trabalhar com dados de planilha e fornecer aos usuários uma melhor exibição ou organização.

### <a name="apply-conditional-formatting"></a>Aplicar formatação condicional

Este exemplo aplica formatação condicional ao intervalo usado atualmente na planilha. A formatação condicional é um preenchimento verde para os 10% principais dos valores.

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

### <a name="create-a-sorted-table"></a>Criar uma tabela classificação

Este exemplo cria uma tabela do intervalo usado da planilha atual e classifica-a com base na primeira coluna.

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Registrar os valores "Grande Total" de uma tabela dinâmica

Este exemplo localiza a primeira Tabela Dinâmica na lista de trabalho e registra os valores nas células "Grand Total" (conforme realçado em verde na imagem abaixo).

:::image type="content" source="../images/sample-pivottable-grand-total-row.png" alt-text="Uma tabela dinâmica mostrando vendas de frutas com a linha Grand Total realçada verde.":::

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

### <a name="use-data-validation-to-create-a-drop-down-list"></a>Usar validação de dados para criar uma lista lista listada

Esse script cria uma lista de seleção listada para uma célula. Ele usa os valores existentes do intervalo selecionado como as opções da lista.

:::image type="content" source="../images/sample-data-validation.png" alt-text="Uma planilha mostrando um intervalo de três células que contêm opções de cores &quot;vermelho, azul, verde&quot; e ao lado dela, as mesmas opções mostradas em uma lista lista listada.":::

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

Esses exemplos usam fórmulas do Excel e mostram como trabalhar com elas em scripts.

### <a name="single-formula"></a>Fórmula única

Esse script define a fórmula de uma célula e exibe como o Excel armazena a fórmula e o valor da célula separadamente.

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

### <a name="spilling-results-from-a-formula"></a>Vazamento de resultados de uma fórmula

Esse script transpõe o intervalo "A1:D2" para "A4:B7" usando a função TRANSPOSE. Se a transposição resulta em um erro #SPILL, limpa o intervalo de destino e aplica a fórmula novamente.

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

## <a name="suggest-new-samples"></a>Sugerir novos exemplos

Recebemos sugestões de novos exemplos. Se houver um cenário comum que ajude outros desenvolvedores de scripts, conte-nos na seção comentários abaixo.
