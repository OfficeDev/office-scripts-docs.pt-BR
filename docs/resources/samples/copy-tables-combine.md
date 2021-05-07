---
title: Combinar dados de várias Excel tabelas em uma única tabela
description: Saiba como usar Office Scripts para combinar dados de várias Excel tabelas em uma única tabela.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ac8c7d0a3f0f4f3d7d3217ffac31aff1a5595d17
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232442"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Combinar dados de várias Excel tabelas em uma única tabela

Este exemplo combina dados de várias Excel tabelas em uma única tabela que inclui todas as linhas. Supõe que todas as tabelas que estão sendo usadas tenham a mesma estrutura.

Há duas variações deste script:

1. O [primeiro script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combina todas as tabelas no arquivo Excel.
1. O [segundo script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) obtém tabelas seletivamente em um conjunto de planilhas.

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Código de exemplo: combinar dados de várias Excel tabelas em uma única tabela

Baixe o arquivo de <a href="tables-copy.xlsx"> exemplotables-copy.xlsx</a> e use-o com o script a seguir para experimentar você mesmo!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Código de exemplo: combinar dados de várias Excel tabelas em selecionar planilhas em uma única tabela

Baixe o arquivo de <a href="tables-select-copy.xlsx"> exemplotables-select-copy.xlsx</a> e use-o com o script a seguir para experimentar você mesmo!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Vídeo de treinamento: Combinar dados de várias Excel tabelas em uma única tabela

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/di-8JukK3Lc).
