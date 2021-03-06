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
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Limpar filtro de coluna de tabela com base no local ativo da célula

Este exemplo limpa o filtro de coluna de tabela com base no local da célula ativa. O script detecta se a célula faz parte de uma tabela, determina a coluna da tabela e limpa qualquer filtro aplicado a ela.

Se você quiser saber mais sobre como salvar o filtro antes de desmatá-lo (e reaplicação posterior), consulte Mover linhas entre tabelas salvando [filtros](move-rows-across-tables.md), um exemplo mais avançado.

_Antes de limpar o filtro de coluna (observe a célula ativa)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Uma célula ativa antes de limpar o filtro de coluna.":::

_Depois de limpar o filtro de coluna_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Uma célula ativa após limpar o filtro de coluna.":::

## <a name="sample-excel-file"></a>Exemplo Excel arquivo

Baixe <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> para uma workbook pronta para uso. Adicione o seguinte script para experimentar o exemplo você mesmo!

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Código de exemplo: Limpar filtro de coluna de tabela com base na célula ativa

O script a seguir limpa o filtro de coluna de tabela com base no local da célula ativa e pode ser aplicado a qualquer arquivo Excel com uma tabela.

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
