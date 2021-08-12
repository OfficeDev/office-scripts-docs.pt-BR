---
title: Remover hiperlinks de cada célula em uma Excel de trabalho
description: Saiba como usar Office Scripts para remover hiperlinks de cada célula em uma Excel de trabalho.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 498d55ea1ee7926ab124d00795825660005c5e38e73ed5d90fe8f9208a583908
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847426"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Remover hiperlinks de cada célula em uma Excel de trabalho

 Este exemplo limpa todos os hiperlinks da planilha atual. Ele percorre a planilha e, se houver algum hiperlink associado à célula, ele limpará o hiperlink e ainda manterá o valor da célula como está. Também registra o tempo necessário para concluir a transição.

> [!NOTE]
> Isso só funcionará se a contagem de células for < 10k.

## <a name="sample-excel-file"></a>Exemplo Excel arquivo

Baixe o arquivo <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> para uma pasta de trabalho pronta para uso. Adicione o seguinte script para experimentar o exemplo você mesmo!

## <a name="sample-code-remove-hyperlinks"></a>Código de exemplo: Remover hiperlinks

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
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

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Vídeo de treinamento: Remover hiperlinks de cada célula em uma Excel de trabalho

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/v20fdinxpHU).
