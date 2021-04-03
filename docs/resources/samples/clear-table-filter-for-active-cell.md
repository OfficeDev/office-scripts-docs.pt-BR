---
title: Limpar filtro de coluna de tabela com base no local ativo da célula
description: Saiba como limpar o filtro de coluna de tabela com base no local ativo da célula.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: 5da3cdfe753f254c7f9ec97bea384b77ccaff09e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571060"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="7f295-103">Limpar filtro de coluna de tabela com base no local ativo da célula</span><span class="sxs-lookup"><span data-stu-id="7f295-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="7f295-104">Este exemplo limpa o filtro de coluna de tabela com base no local da célula ativa.</span><span class="sxs-lookup"><span data-stu-id="7f295-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="7f295-105">O script detecta se a célula faz parte de uma tabela, determina a coluna da tabela e limpa qualquer filtro aplicado a ela.</span><span class="sxs-lookup"><span data-stu-id="7f295-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="7f295-106">Se você quiser saber mais sobre como salvar o filtro antes de desmatá-lo (e reaplicação posterior), consulte Mover linhas entre tabelas salvando [filtros](move-rows-across-tables.md), um exemplo mais avançado.</span><span class="sxs-lookup"><span data-stu-id="7f295-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="7f295-107">_Antes de limpar o filtro de coluna (observe a célula ativa)_</span><span class="sxs-lookup"><span data-stu-id="7f295-107">_Before clearing column filter (notice the active cell)_</span></span>

![Antes de limpar o filtro de coluna](../../images/before-filter-applied.png)

<span data-ttu-id="7f295-109">_Depois de limpar o filtro de coluna_</span><span class="sxs-lookup"><span data-stu-id="7f295-109">_After clearing column filter_</span></span>

![Depois de limpar o filtro de coluna](../../images/after-filter-cleared.png)

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="7f295-111">Código de exemplo: Limpar filtro de coluna de tabela com base na célula ativa</span><span class="sxs-lookup"><span data-stu-id="7f295-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="7f295-112">O script a seguir limpa o filtro de coluna de tabela com base no local da célula ativa e pode ser aplicado a qualquer arquivo do Excel com uma tabela.</span><span class="sxs-lookup"><span data-stu-id="7f295-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="7f295-113">Por conveniência, você pode baixar e usar <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span><span class="sxs-lookup"><span data-stu-id="7f295-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```

## <a name="training-video-clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="7f295-114">Vídeo de treinamento: Limpar filtro de coluna de tabela com base no local ativo da célula</span><span class="sxs-lookup"><span data-stu-id="7f295-114">Training video: Clear table column filter based on active cell location</span></span>

<span data-ttu-id="7f295-115">Para ver um exemplo de como trabalhar com intervalos, consulte [Range basics training videos](range-basics.md#training-videos-range-basics).</span><span class="sxs-lookup"><span data-stu-id="7f295-115">For an example of how to work with ranges, see [Range basics training videos](range-basics.md#training-videos-range-basics).</span></span>
