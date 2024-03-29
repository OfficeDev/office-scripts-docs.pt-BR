---
title: Contar linhas em branco em planilhas
description: Saiba como usar Office Scripts para detectar se há linhas em branco em vez de dados em planilhas e, em seguida, relatar a contagem de linhas em branco a ser usada em um fluxo Power Automate.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1ae513928b885994dc7f6d1b8ad66d694b61e7b7
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585559"
---
# <a name="count-blank-rows-on-sheets"></a>Contar linhas em branco em planilhas

Este projeto inclui dois scripts:

* [Contar linhas em branco em uma determinada planilha](#sample-code-count-blank-rows-on-a-given-sheet): percorre o intervalo usado em uma determinada planilha e retorna uma contagem de linhas em branco.
* [Contar linhas em branco em todas as planilhas](#sample-code-count-blank-rows-on-all-sheets): percorre o intervalo usado em todas as _planilhas_ e retorna uma contagem de linhas em branco.

> [!NOTE]
> Para nosso script, uma linha em branco é qualquer linha onde não há dados. A linha pode ter formatação.

_Esta planilha retorna a contagem de 4 linhas em branco_

:::image type="content" source="../../images/blank-rows.png" alt-text="Uma planilha mostrando dados com linhas em branco.":::

_Esta planilha retorna a contagem de 0 linhas em branco (todas as linhas têm alguns dados)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Uma planilha mostrando dados sem linhas em branco.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Código de exemplo: Contar linhas em branco em uma determinada planilha

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Código de exemplo: Contar linhas em branco em todas as planilhas

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```
