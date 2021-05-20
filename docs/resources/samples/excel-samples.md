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
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="8ff83-103">Roteiros básicos para roteiros Office em Excel na Web</span><span class="sxs-lookup"><span data-stu-id="8ff83-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="8ff83-104">As amostras a seguir são scripts simples para você experimentar em seus próprios livros de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8ff83-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="8ff83-105">Para usá-los em Excel na Web:</span><span class="sxs-lookup"><span data-stu-id="8ff83-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="8ff83-106">Abra a guia **Automação**.</span><span class="sxs-lookup"><span data-stu-id="8ff83-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="8ff83-107">**Editor de Código de Imprensa**.</span><span class="sxs-lookup"><span data-stu-id="8ff83-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="8ff83-108">Pressione **o novo script** no painel de tarefas do Editor de Código.</span><span class="sxs-lookup"><span data-stu-id="8ff83-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="8ff83-109">Substitua todo o script com a amostra de sua escolha.</span><span class="sxs-lookup"><span data-stu-id="8ff83-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="8ff83-110">Pressione **executar** no painel de tarefas do Editor de Código.</span><span class="sxs-lookup"><span data-stu-id="8ff83-110">Press **Run** in the Code Editor's task pane.</span></span>

## <a name="script-basics"></a><span data-ttu-id="8ff83-111">Conceitos básicos</span><span class="sxs-lookup"><span data-stu-id="8ff83-111">Script basics</span></span>

<span data-ttu-id="8ff83-112">Essas amostras demonstram blocos fundamentais de construção para Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="8ff83-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="8ff83-113">Adicione isso aos seus scripts para estender sua solução e resolver problemas comuns.</span><span class="sxs-lookup"><span data-stu-id="8ff83-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="8ff83-114">Leia e registre uma célula</span><span class="sxs-lookup"><span data-stu-id="8ff83-114">Read and log one cell</span></span>

<span data-ttu-id="8ff83-115">Esta amostra lê o valor de **A1** e imprime-o no console.</span><span class="sxs-lookup"><span data-stu-id="8ff83-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="8ff83-116">Leia a célula ativa</span><span class="sxs-lookup"><span data-stu-id="8ff83-116">Read the active cell</span></span>

<span data-ttu-id="8ff83-117">Este script registra o valor da célula ativa atual.</span><span class="sxs-lookup"><span data-stu-id="8ff83-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="8ff83-118">Se várias células forem selecionadas, a célula mais à esquerda será registrada.</span><span class="sxs-lookup"><span data-stu-id="8ff83-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="8ff83-119">Alterar uma célula adjacente</span><span class="sxs-lookup"><span data-stu-id="8ff83-119">Change an adjacent cell</span></span>

<span data-ttu-id="8ff83-120">Este script recebe células adjacentes usando referências relativas.</span><span class="sxs-lookup"><span data-stu-id="8ff83-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="8ff83-121">Observe que se a célula ativa estiver na linha superior, parte do script falha, pois faz referência à célula acima da atualmente selecionada.</span><span class="sxs-lookup"><span data-stu-id="8ff83-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="8ff83-122">Alterar todas as células adjacentes</span><span class="sxs-lookup"><span data-stu-id="8ff83-122">Change all adjacent cells</span></span>

<span data-ttu-id="8ff83-123">Este script copia a formatação na célula ativa para as células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="8ff83-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="8ff83-124">Observe que este script só funciona quando a célula ativa não está na borda da planilha.</span><span class="sxs-lookup"><span data-stu-id="8ff83-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="8ff83-125">Alterar cada célula individual em um intervalo</span><span class="sxs-lookup"><span data-stu-id="8ff83-125">Change each individual cell in a range</span></span>

<span data-ttu-id="8ff83-126">Este script gira em loops ao longo do intervalo selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="8ff83-126">This script loops over the currently select range.</span></span> <span data-ttu-id="8ff83-127">Ele limpa a formatação atual e define a cor de preenchimento em cada célula para uma cor aleatória.</span><span class="sxs-lookup"><span data-stu-id="8ff83-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="8ff83-128">Obter grupos de células com base em critérios especiais</span><span class="sxs-lookup"><span data-stu-id="8ff83-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="8ff83-129">Este script recebe todas as células em branco na faixa usada da planilha atual.</span><span class="sxs-lookup"><span data-stu-id="8ff83-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="8ff83-130">Em seguida, destaca todas essas células com um fundo amarelo.</span><span class="sxs-lookup"><span data-stu-id="8ff83-130">It then highlights all those cells with a yellow background.</span></span>

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

## <a name="collections"></a><span data-ttu-id="8ff83-131">Coleções</span><span class="sxs-lookup"><span data-stu-id="8ff83-131">Collections</span></span>

<span data-ttu-id="8ff83-132">Essas amostras trabalham com coleções de objetos na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8ff83-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterate-over-collections"></a><span data-ttu-id="8ff83-133">Iterar sobre coleções</span><span class="sxs-lookup"><span data-stu-id="8ff83-133">Iterate over collections</span></span>

<span data-ttu-id="8ff83-134">Este script recebe e registra os nomes de todas as planilhas na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8ff83-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="8ff83-135">Ele também define as cores de suas guias para uma cor aleatória.</span><span class="sxs-lookup"><span data-stu-id="8ff83-135">It also sets the their tab colors to a random color.</span></span>

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

### <a name="query-and-delete-from-a-collection"></a><span data-ttu-id="8ff83-136">Consultar e excluir de uma coleção</span><span class="sxs-lookup"><span data-stu-id="8ff83-136">Query and delete from a collection</span></span>

<span data-ttu-id="8ff83-137">Este roteiro cria uma nova planilha.</span><span class="sxs-lookup"><span data-stu-id="8ff83-137">This script creates a new worksheet.</span></span> <span data-ttu-id="8ff83-138">Ele verifica uma cópia existente da planilha e exclui-a antes de fazer uma nova folha.</span><span class="sxs-lookup"><span data-stu-id="8ff83-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

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

## <a name="dates"></a><span data-ttu-id="8ff83-139">Datas</span><span class="sxs-lookup"><span data-stu-id="8ff83-139">Dates</span></span>

<span data-ttu-id="8ff83-140">As amostras nesta seção mostram como usar o [objeto](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) Data JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8ff83-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="8ff83-141">A amostra a seguir obtém a data e a hora atuais e, em seguida, escreve esses valores para duas células na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="8ff83-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="8ff83-142">A próxima amostra lê uma data armazenada em Excel e a traduz para um objeto JavaScript Date.</span><span class="sxs-lookup"><span data-stu-id="8ff83-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="8ff83-143">Ele usa o [número de série numérico da data](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para a Data JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8ff83-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="8ff83-144">Exibir dados</span><span class="sxs-lookup"><span data-stu-id="8ff83-144">Display data</span></span>

<span data-ttu-id="8ff83-145">Essas amostras demonstram como trabalhar com dados de planilhas e fornecem aos usuários uma melhor visão ou organização.</span><span class="sxs-lookup"><span data-stu-id="8ff83-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="8ff83-146">Aplicar formatação condicional</span><span class="sxs-lookup"><span data-stu-id="8ff83-146">Apply conditional formatting</span></span>

<span data-ttu-id="8ff83-147">Esta amostra aplica formatação condicional à faixa utilizada atualmente na planilha.</span><span class="sxs-lookup"><span data-stu-id="8ff83-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="8ff83-148">A formatação condicional é um preenchimento verde para os 10% mais altos dos valores.</span><span class="sxs-lookup"><span data-stu-id="8ff83-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="8ff83-149">Crie uma tabela classificada</span><span class="sxs-lookup"><span data-stu-id="8ff83-149">Create a sorted table</span></span>

<span data-ttu-id="8ff83-150">Esta amostra cria uma tabela a partir da faixa usada da planilha atual e, em seguida, classifica-a com base na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="8ff83-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="8ff83-151">Registre os valores "Grande Total" de uma Tabela Dinâmica</span><span class="sxs-lookup"><span data-stu-id="8ff83-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="8ff83-152">Esta amostra encontra a primeira Tabela Dinâmica na pasta de trabalho e registra os valores nas células "Grande Total" (como destacado em verde na imagem abaixo).</span><span class="sxs-lookup"><span data-stu-id="8ff83-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

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

### <a name="create-a-drop-down-list-using-data-validation"></a><span data-ttu-id="8ff83-154">Crie uma lista de drop-down usando validação de dados</span><span class="sxs-lookup"><span data-stu-id="8ff83-154">Create a drop-down list using data validation</span></span>

<span data-ttu-id="8ff83-155">Este script cria uma lista de seleção para uma célula.</span><span class="sxs-lookup"><span data-stu-id="8ff83-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="8ff83-156">Ele usa os valores existentes da faixa selecionada como as opções para a lista.</span><span class="sxs-lookup"><span data-stu-id="8ff83-156">It uses the existing values of the selected range as the choices for the list.</span></span>

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

## <a name="formulas"></a><span data-ttu-id="8ff83-158">Fórmulas</span><span class="sxs-lookup"><span data-stu-id="8ff83-158">Formulas</span></span>

<span data-ttu-id="8ff83-159">Essas amostras usam fórmulas Excel e mostram como trabalhar com elas em scripts.</span><span class="sxs-lookup"><span data-stu-id="8ff83-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="8ff83-160">Fórmula única</span><span class="sxs-lookup"><span data-stu-id="8ff83-160">Single formula</span></span>

<span data-ttu-id="8ff83-161">Este script define a fórmula de uma célula e, em seguida, exibe como Excel armazena a fórmula e o valor da célula separadamente.</span><span class="sxs-lookup"><span data-stu-id="8ff83-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a><span data-ttu-id="8ff83-162">Manuseie um `#SPILL!` erro retornado de uma fórmula</span><span class="sxs-lookup"><span data-stu-id="8ff83-162">Handle a `#SPILL!` error returned from a formula</span></span>

<span data-ttu-id="8ff83-163">Este script transpõe a faixa "A1:D2" para "A4:B7" usando a função TRANSPOSE.</span><span class="sxs-lookup"><span data-stu-id="8ff83-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="8ff83-164">Se a transposição resultar em um `#SPILL` erro, ela limpa o intervalo de destino e aplica a fórmula novamente.</span><span class="sxs-lookup"><span data-stu-id="8ff83-164">If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.</span></span>

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

## <a name="suggest-new-samples"></a><span data-ttu-id="8ff83-165">Sugira novas amostras</span><span class="sxs-lookup"><span data-stu-id="8ff83-165">Suggest new samples</span></span>

<span data-ttu-id="8ff83-166">Damos boas-vindas às sugestões de novas amostras.</span><span class="sxs-lookup"><span data-stu-id="8ff83-166">We welcome suggestions for new samples.</span></span> <span data-ttu-id="8ff83-167">Se houver um cenário comum que ajude outros desenvolvedores de script, por favor, diga-nos na seção de feedback na parte inferior da página.</span><span class="sxs-lookup"><span data-stu-id="8ff83-167">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="8ff83-168">Confira também</span><span class="sxs-lookup"><span data-stu-id="8ff83-168">See also</span></span>

* [<span data-ttu-id="8ff83-169">"Range basics" de Sudhi Ramamurthy no YouTube</span><span class="sxs-lookup"><span data-stu-id="8ff83-169">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="8ff83-170">Office Scripts amostras e cenários</span><span class="sxs-lookup"><span data-stu-id="8ff83-170">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="8ff83-171">Gravar, editar e criar scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="8ff83-171">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
