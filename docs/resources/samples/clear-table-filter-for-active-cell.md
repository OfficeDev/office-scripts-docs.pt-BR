---
title: Limpar filtro de coluna de tabela com base no local ativo da célula
description: Saiba como limpar o filtro de coluna de tabela com base no local ativo da célula.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 06fba191a79f4641d4d1017bda332c7559b50e6d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285924"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="8b74e-103">Limpar filtro de coluna de tabela com base no local ativo da célula</span><span class="sxs-lookup"><span data-stu-id="8b74e-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="8b74e-104">Este exemplo limpa o filtro de coluna de tabela com base no local da célula ativa.</span><span class="sxs-lookup"><span data-stu-id="8b74e-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="8b74e-105">O script detecta se a célula faz parte de uma tabela, determina a coluna da tabela e limpa qualquer filtro aplicado a ela.</span><span class="sxs-lookup"><span data-stu-id="8b74e-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="8b74e-106">Se você quiser saber mais sobre como salvar o filtro antes de desmatá-lo (e reaplicação posterior), consulte Mover linhas entre tabelas salvando [filtros](move-rows-across-tables.md), um exemplo mais avançado.</span><span class="sxs-lookup"><span data-stu-id="8b74e-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="8b74e-107">_Antes de limpar o filtro de coluna (observe a célula ativa)_</span><span class="sxs-lookup"><span data-stu-id="8b74e-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Uma célula ativa antes de limpar o filtro de coluna":::

<span data-ttu-id="8b74e-109">_Depois de limpar o filtro de coluna_</span><span class="sxs-lookup"><span data-stu-id="8b74e-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Uma célula ativa após limpar o filtro de coluna":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="8b74e-111">Código de exemplo: Limpar filtro de coluna de tabela com base na célula ativa</span><span class="sxs-lookup"><span data-stu-id="8b74e-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="8b74e-112">O script a seguir limpa o filtro de coluna de tabela com base no local da célula ativa e pode ser aplicado a qualquer arquivo Excel com uma tabela.</span><span class="sxs-lookup"><span data-stu-id="8b74e-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="8b74e-113">Por conveniência, você pode baixar e usar <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span><span class="sxs-lookup"><span data-stu-id="8b74e-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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
