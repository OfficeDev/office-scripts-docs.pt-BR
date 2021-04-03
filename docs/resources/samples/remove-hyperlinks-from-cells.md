---
title: Remover hiperlinks de cada célula em uma planilha do Excel
description: Saiba como usar scripts do Office para remover hiperlinks de cada célula em uma planilha do Excel.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 07b670aac3368e38b9b93283404befee608391a7
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571039"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="812b3-103">Remover hiperlinks de cada célula em uma planilha do Excel</span><span class="sxs-lookup"><span data-stu-id="812b3-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="812b3-104">Este exemplo limpa todos os hiperlinks da planilha atual.</span><span class="sxs-lookup"><span data-stu-id="812b3-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="812b3-105">Ele percorre a planilha e, se houver algum hiperlink associado à célula, ele limpará o hiperlink e ainda manterá o valor da célula como está.</span><span class="sxs-lookup"><span data-stu-id="812b3-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="812b3-106">Também registra o tempo necessário para concluir a transição.</span><span class="sxs-lookup"><span data-stu-id="812b3-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="812b3-107">Isso só funcionará se a contagem de células for < 10k.</span><span class="sxs-lookup"><span data-stu-id="812b3-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="812b3-108">Código de exemplo: Remover hiperlinks</span><span class="sxs-lookup"><span data-stu-id="812b3-108">Sample code: Remove hyperlinks</span></span>

<span data-ttu-id="812b3-109">Baixe o arquivo <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> usado neste exemplo e experimente você mesmo!</span><span class="sxs-lookup"><span data-stu-id="812b3-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> used in this sample and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {

  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);
  const targetRange = sheet.getUsedRange(true);
  if (!targetRange) {
    console.log(`There is no data in the worksheet. `)
    return;
  }
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  const totalCells = rowCount * colCount;
  if (totalCells > 10000) {
    console.log("Too many cells to operate with. Consider editing script to use selected range and then remove hyperlinks in batches. " + targetRange.getAddress());
    return;
  }
  // Call the helper function to remove the hyperlinks. 
  removeHyperLink(targetRange);
  return;
}

/**
 * Removes hyperlink for each cell in the target range. Logs the time it takes to complete traversal.
 * @param targetRange Target range to clear the hyperlinks from.
 */
function removeHyperLink(targetRange: ExcelScript.Range): void {
  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);
  let clearedCount = 0;
  let cellsVisited = 0;

  let groupStart = new Date().getTime();
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      cellsVisited++;
      if (cellsVisited % 50 === 0) {
        let groupEnd = new Date().getTime();
        console.log(`Completed ${cellsVisited} cells out of ${rowCount * colCount}. This group took: ${(groupEnd - groupStart) / 1000} seconds to complete.`);
        groupStart = new Date().getTime();
      }
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }
  console.log(`Done. Inspected ${cellsVisited} cells. Cleared hyperlinks in: ${clearedCount} cells`);
  return;
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="812b3-110">Vídeo de treinamento: Remover hiperlinks de cada célula em uma planilha do Excel</span><span class="sxs-lookup"><span data-stu-id="812b3-110">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="812b3-111">[![Assista a um vídeo passo a passo sobre como remover hiperlinks de cada célula em uma planilha do Excel](../../images/hyperlinks-vid.jpg)](https://youtu.be/v20fdinxpHU "Vídeo passo a passo sobre como remover hiperlinks de cada célula em uma planilha do Excel")</span><span class="sxs-lookup"><span data-stu-id="812b3-111">[![Watch step-by-step video on how to remove hyperlinks from each cell in an Excel worksheet](../../images/hyperlinks-vid.jpg)](https://youtu.be/v20fdinxpHU "Step-by-step video on how to remove hyperlinks from each cell in an Excel worksheet")</span></span>
