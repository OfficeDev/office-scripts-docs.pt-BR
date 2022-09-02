---
title: Combinar dados de várias tabelas do Excel em uma única tabela
description: Saiba como usar Scripts do Office para combinar dados de várias tabelas do Excel em uma única tabela.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3db510514c676b9012fd47abc2a7e92492a9cf87
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572448"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Combinar dados de várias tabelas do Excel em uma única tabela

Este exemplo combina dados de várias tabelas do Excel em uma única tabela que inclui todas as linhas. Ele pressupõe que todas as tabelas que estão sendo usadas tenham a mesma estrutura.

Há duas variações desse script:

1. O [primeiro script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combina todas as tabelas no arquivo do Excel.
1. O [segundo script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) obtém seletivamente tabelas dentro de um conjunto de planilhas.

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe [tables-copy.xlsx](tables-copy.xlsx) para uma pasta de trabalho pronta para uso. Adicione os scripts a seguir para experimentar o exemplo por conta própria!

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Código de exemplo: combinar dados de várias tabelas do Excel em uma única tabela

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Código de exemplo: combinar dados de várias tabelas do Excel em planilhas selecionadas em uma única tabela

Baixe o arquivo de [tables-select-copy.xlsx](tables-select-copy.xlsx) e use-o com o script a seguir para experimentá-lo por conta própria!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Vídeo de treinamento: Combinar dados de várias tabelas do Excel em uma única tabela

[Veja Sudhi Ramamurthy percorrer este exemplo no YouTube](https://youtu.be/di-8JukK3Lc).
