---
title: Contar linhas em branco em planilhas
description: Saiba como usar scripts do Office para detectar se há linhas em branco em vez de dados em planilhas e, em seguida, relatar a contagem de linhas em branco a ser usada em um fluxo do Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 088ab97c686484ca5c13c875b80431ac28d20736
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754828"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="8eb25-103">Contar linhas em branco em planilhas</span><span class="sxs-lookup"><span data-stu-id="8eb25-103">Count blank rows on sheets</span></span>

<span data-ttu-id="8eb25-104">Este projeto inclui dois scripts:</span><span class="sxs-lookup"><span data-stu-id="8eb25-104">This project includes two scripts:</span></span>

* <span data-ttu-id="8eb25-105">[Contar linhas em branco em uma determinada planilha:](#sample-code-count-blank-rows-on-a-given-sheet)percorre o intervalo usado em uma determinada planilha e retorna uma contagem de linhas em branco.</span><span class="sxs-lookup"><span data-stu-id="8eb25-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="8eb25-106">[Contar linhas em branco em todas as planilhas](#sample-code-count-blank-rows-on-all-sheets): percorre o intervalo usado em todas as _planilhas_ e retorna uma contagem de linhas em branco.</span><span class="sxs-lookup"><span data-stu-id="8eb25-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="8eb25-107">Para nosso script, uma linha em branco é qualquer linha onde não há dados.</span><span class="sxs-lookup"><span data-stu-id="8eb25-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="8eb25-108">A linha pode ter formatação.</span><span class="sxs-lookup"><span data-stu-id="8eb25-108">The row can have formatting.</span></span>

<span data-ttu-id="8eb25-109">_Esta planilha retorna a contagem de 4 linhas em branco_</span><span class="sxs-lookup"><span data-stu-id="8eb25-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Uma planilha mostrando dados com linhas em branco.":::

<span data-ttu-id="8eb25-111">_Esta planilha retorna a contagem de 0 linhas em branco (todas as linhas têm alguns dados)_</span><span class="sxs-lookup"><span data-stu-id="8eb25-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Uma planilha mostrando dados sem linhas em branco.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="8eb25-113">Código de exemplo: Contar linhas em branco em uma determinada planilha</span><span class="sxs-lookup"><span data-stu-id="8eb25-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="8eb25-114">Código de exemplo: Contar linhas em branco em todas as planilhas</span><span class="sxs-lookup"><span data-stu-id="8eb25-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a><span data-ttu-id="8eb25-115">Usar com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="8eb25-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Um fluxo do Power Automate mostrando como configurar para executar um Script do Office.":::
