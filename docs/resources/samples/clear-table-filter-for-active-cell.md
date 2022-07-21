---
title: Remover filtros de coluna de tabela
description: Saiba como limpar o filtro de coluna de tabela com base no local da célula ativa.
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21a79abfdd4aeac79af4a0f9ea4a581d45b9706b
ms.sourcegitcommit: dd632402cb46ec8407a1c98456f1bc9ab96ffa46
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/21/2022
ms.locfileid: "66918808"
---
# <a name="remove-table-column-filters"></a>Remover filtros de coluna de tabela

Este exemplo remove os filtros de uma coluna de tabela, com base no local da célula ativa. O script detecta se a célula faz parte de uma tabela, determina a coluna da tabela e limpa qualquer filtro aplicado nela.

Se você quiser saber mais sobre como salvar o filtro antes de desmarcar (e aplicar novamente mais tarde), consulte Mover linhas entre [tabelas salvando filtros](move-rows-across-tables.md), um exemplo mais avançado.

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> para uma pasta de trabalho pronta para uso. Adicione o script a seguir para experimentar o exemplo por conta própria!

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Código de exemplo: limpar o filtro de coluna de tabela com base na célula ativa

O script a seguir limpa o filtro de coluna da tabela com base no local da célula ativa e pode ser aplicado a qualquer arquivo do Excel com uma tabela.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>Antes de limpar o filtro de coluna (observe a célula ativa)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Uma célula ativa antes de limpar o filtro de coluna.":::

## <a name="after-clearing-column-filter"></a>Depois de limpar o filtro de coluna

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Uma célula ativa após limpar o filtro de coluna.":::
