---
title: Combinar dados de várias tabelas do Excel em uma única tabela
description: Saiba como usar scripts do Office para combinar dados de várias tabelas do Excel em uma única tabela.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 2f3f7232216f686946861d8c2cdec44013333ec7
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571080"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Combinar dados de várias tabelas do Excel em uma única tabela

Este exemplo combina dados de várias tabelas do Excel em uma única tabela que inclui todas as linhas. Supõe que todas as tabelas que estão sendo usadas tenham a mesma estrutura.

Há duas variações deste script:

1. O [primeiro script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combina todas as tabelas no arquivo do Excel.
1. O [segundo script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) obtém tabelas seletivamente em um conjunto de planilhas.

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Código de exemplo: combinar dados de várias tabelas do Excel em uma única tabela

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Código de exemplo: Combinar dados de várias tabelas do Excel em planilhas selecionadas em uma única tabela

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Vídeo de treinamento: combinar dados de várias tabelas do Excel em uma única tabela

[![Assista ao vídeo passo a passo sobre como combinar dados de várias tabelas do Excel em uma única tabela](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "Vídeo passo a passo sobre como combinar dados de várias tabelas do Excel em uma única tabela")
