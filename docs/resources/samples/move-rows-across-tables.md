---
title: Mover linhas entre tabelas usando Office Scripts
description: Saiba como mover linhas entre tabelas salvando filtros e, em seguida, processamento e reaplicação dos filtros.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 860521de166108d5a8355ea246c1bfe77e0e064b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313754"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a><span data-ttu-id="dcfe4-103">Mover linhas entre tabelas salvando filtros e, em seguida, processamento e reaplicação dos filtros</span><span class="sxs-lookup"><span data-stu-id="dcfe4-103">Move rows across tables by saving filters, then processing and reapplying the filters</span></span>

<span data-ttu-id="dcfe4-104">Esse script faz o seguinte:</span><span class="sxs-lookup"><span data-stu-id="dcfe4-104">This script does the following:</span></span>

* <span data-ttu-id="dcfe4-105">Seleciona linhas da tabela de origem onde o valor em uma coluna é igual a _algum valor_.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-105">Selects rows from the source table where the value in a column is equal to _some value_.</span></span>
* <span data-ttu-id="dcfe4-106">Move todas as linhas selecionadas para outra tabela (destino) em outra planilha.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-106">Moves all selected rows into another (target) table on another worksheet.</span></span>
* <span data-ttu-id="dcfe4-107">Reaplica os filtros relevantes na tabela de origem.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-107">Reapplies the relevant filters on the source table.</span></span>

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="Capturas de tela da workbook antes e depois.":::

## <a name="sample-excel-file"></a><span data-ttu-id="dcfe4-109">Exemplo Excel arquivo</span><span class="sxs-lookup"><span data-stu-id="dcfe4-109">Sample Excel file</span></span>

<span data-ttu-id="dcfe4-110">Baixe o arquivo <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> para uma pasta de trabalho pronta para uso.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-110">Download the file <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="dcfe4-111">Adicione o seguinte script para experimentar o exemplo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="dcfe4-111">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-move-rows-using-range-values"></a><span data-ttu-id="dcfe4-112">Código de exemplo: Mover linhas usando valores de intervalo</span><span class="sxs-lookup"><span data-stu-id="dcfe4-112">Sample code: Move rows using range values</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a><span data-ttu-id="dcfe4-113">Vídeo de treinamento: Mover linhas entre tabelas</span><span class="sxs-lookup"><span data-stu-id="dcfe4-113">Training video: Move rows across tables</span></span>

<span data-ttu-id="dcfe4-114">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/_3t3Pk4i2L0).</span><span class="sxs-lookup"><span data-stu-id="dcfe4-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/_3t3Pk4i2L0).</span></span> <span data-ttu-id="dcfe4-115">Há dois scripts mostrados na solução do vídeo.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-115">There are two scripts shown in the video's solution.</span></span> <span data-ttu-id="dcfe4-116">A principal diferença é como as linhas são selecionadas.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-116">The main difference is how the rows are selected.</span></span>

* <span data-ttu-id="dcfe4-117">Na primeira variante, as linhas são selecionadas aplicando o filtro de tabela e lendo o intervalo visível.</span><span class="sxs-lookup"><span data-stu-id="dcfe4-117">In the first variant, the rows are selected by applying the table filter and reading the visible range.</span></span>
* <span data-ttu-id="dcfe4-118">No segundo, as linhas são selecionadas lendo os valores e extraindo os valores da linha (que é o que o exemplo nesta página usa).</span><span class="sxs-lookup"><span data-stu-id="dcfe4-118">In the second, the rows are selected by reading the values and extracting the row values (which is what the sample on this page uses).</span></span>
