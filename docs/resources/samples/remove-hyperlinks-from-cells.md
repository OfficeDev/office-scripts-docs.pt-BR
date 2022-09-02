---
title: Remover hiperlinks de cada célula em uma planilha do Excel
description: Saiba como usar Scripts do Office para remover hiperlinks de cada célula em uma planilha do Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1445988b1e6a85fcab8914ffeaaef80a07a52f5e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572623"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Remover hiperlinks de cada célula em uma planilha do Excel

 Este exemplo limpa todos os hiperlinks da planilha atual. Ele atravessa a planilha e, se houver algum hiperlink associado à célula, ele limpará o hiperlink e ainda manterá o valor da célula como está. Também registra o tempo necessário para concluir a passagem.

> [!NOTE]
> Isso só funcionará se a contagem de células for < 10 mil.

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe o arquivo [remove-hyperlinks.xlsx](remove-hyperlinks.xlsx) para uma pasta de trabalho pronta para uso. Adicione o script a seguir para experimentar o exemplo por conta própria!

## <a name="sample-code-remove-hyperlinks"></a>Código de exemplo: remover hiperlinks

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Vídeo de treinamento: Remover hiperlinks de cada célula em uma planilha do Excel

[Veja Sudhi Ramamurthy percorrer este exemplo no YouTube](https://youtu.be/v20fdinxpHU).
