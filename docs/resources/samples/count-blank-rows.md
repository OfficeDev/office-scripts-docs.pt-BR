---
title: Contar linhas em branco em planilhas
description: Saiba como usar Office Scripts para detectar se há linhas em branco em vez de dados em planilhas e, em seguida, relatar a contagem de linhas em branco a ser usada em um fluxo Power Automate.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: e5b60779d2ca2de5f4cf4e03ddd6ff7372515ad6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313803"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="93f51-103">Contar linhas em branco em planilhas</span><span class="sxs-lookup"><span data-stu-id="93f51-103">Count blank rows on sheets</span></span>

<span data-ttu-id="93f51-104">Este projeto inclui dois scripts:</span><span class="sxs-lookup"><span data-stu-id="93f51-104">This project includes two scripts:</span></span>

* <span data-ttu-id="93f51-105">[Contar linhas em branco em uma determinada planilha:](#sample-code-count-blank-rows-on-a-given-sheet)percorre o intervalo usado em uma determinada planilha e retorna uma contagem de linhas em branco.</span><span class="sxs-lookup"><span data-stu-id="93f51-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="93f51-106">[Contar linhas em branco em todas as planilhas](#sample-code-count-blank-rows-on-all-sheets): percorre o intervalo usado em todas as _planilhas_ e retorna uma contagem de linhas em branco.</span><span class="sxs-lookup"><span data-stu-id="93f51-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="93f51-107">Para nosso script, uma linha em branco é qualquer linha onde não há dados.</span><span class="sxs-lookup"><span data-stu-id="93f51-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="93f51-108">A linha pode ter formatação.</span><span class="sxs-lookup"><span data-stu-id="93f51-108">The row can have formatting.</span></span>

<span data-ttu-id="93f51-109">_Esta planilha retorna a contagem de 4 linhas em branco_</span><span class="sxs-lookup"><span data-stu-id="93f51-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Uma planilha mostrando dados com linhas em branco.":::

<span data-ttu-id="93f51-111">_Esta planilha retorna a contagem de 0 linhas em branco (todas as linhas têm alguns dados)_</span><span class="sxs-lookup"><span data-stu-id="93f51-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Uma planilha mostrando dados sem linhas em branco.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="93f51-113">Código de exemplo: Contar linhas em branco em uma determinada planilha</span><span class="sxs-lookup"><span data-stu-id="93f51-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="93f51-114">Código de exemplo: Contar linhas em branco em todas as planilhas</span><span class="sxs-lookup"><span data-stu-id="93f51-114">Sample code: Count blank rows on all sheets</span></span>

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
