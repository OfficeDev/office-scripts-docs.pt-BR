---
title: Limpar filtro de coluna de tabela com base no local ativo da célula
description: Saiba como limpar o filtro de coluna de tabela com base no local ativo da célula.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: f10e23b4ad948a28c5b749533ddedefe164d7142
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313887"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="d7266-103">Limpar filtro de coluna de tabela com base no local ativo da célula</span><span class="sxs-lookup"><span data-stu-id="d7266-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="d7266-104">Este exemplo limpa o filtro de coluna de tabela com base no local da célula ativa.</span><span class="sxs-lookup"><span data-stu-id="d7266-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="d7266-105">O script detecta se a célula faz parte de uma tabela, determina a coluna da tabela e limpa qualquer filtro aplicado a ela.</span><span class="sxs-lookup"><span data-stu-id="d7266-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="d7266-106">Se você quiser saber mais sobre como salvar o filtro antes de desmatá-lo (e reaplicação posterior), consulte Mover linhas entre tabelas salvando [filtros](move-rows-across-tables.md), um exemplo mais avançado.</span><span class="sxs-lookup"><span data-stu-id="d7266-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="d7266-107">_Antes de limpar o filtro de coluna (observe a célula ativa)_</span><span class="sxs-lookup"><span data-stu-id="d7266-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Uma célula ativa antes de limpar o filtro de coluna.":::

<span data-ttu-id="d7266-109">_Depois de limpar o filtro de coluna_</span><span class="sxs-lookup"><span data-stu-id="d7266-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Uma célula ativa após limpar o filtro de coluna.":::

## <a name="sample-excel-file"></a><span data-ttu-id="d7266-111">Exemplo Excel arquivo</span><span class="sxs-lookup"><span data-stu-id="d7266-111">Sample Excel file</span></span>

<span data-ttu-id="d7266-112">Baixe <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> para uma workbook pronta para uso.</span><span class="sxs-lookup"><span data-stu-id="d7266-112">Download <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="d7266-113">Adicione o seguinte script para experimentar o exemplo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="d7266-113">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="d7266-114">Código de exemplo: Limpar filtro de coluna de tabela com base na célula ativa</span><span class="sxs-lookup"><span data-stu-id="d7266-114">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="d7266-115">O script a seguir limpa o filtro de coluna de tabela com base no local da célula ativa e pode ser aplicado a qualquer arquivo Excel com uma tabela.</span><span class="sxs-lookup"><span data-stu-id="d7266-115">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
