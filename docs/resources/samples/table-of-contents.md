---
title: Criar um índice de conteúdo de uma workbook
description: Saiba como criar um índice de conteúdo com links para cada planilha.
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 658143e9e1e6a43cff19eac36abeec88310cda25
ms.sourcegitcommit: 161229492c85f3519c899573cf5022140026e7b8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62220411"
---
# <a name="create-a-workbook-table-of-contents"></a>Criar um índice de conteúdo de uma workbook

Este exemplo mostra como criar um índice de conteúdo para a workbook. Cada entrada no índice de conteúdo é um hiperlink para uma das planilhas da pasta de trabalho.

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="A planilha de conteúdo mostrando links para as outras planilhas.":::

## <a name="sample-excel-file"></a>Exemplo Excel arquivo

Baixe <a href="table-of-contents.xlsx">table-of-contents.xlsx</a> para uma workbook pronta para uso. Adicione o script a seguir e experimente o exemplo você mesmo!

## <a name="sample-code-create-a-workbook-table-of-contents"></a>Código de exemplo: Criar um índice de conteúdo de uma área de trabalho

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Insert a new worksheet at the beginning of the workbook.
  let tocSheet = workbook.addWorksheet();
  tocSheet.setPosition(0);
  tocSheet.setName("Table of Contents");

  // Give the worksheet a title in the sheet.
  tocSheet.getRange("A1").setValue("Table of Contents");
  tocSheet.getRange("A1").getFormat().getFont().setBold(true);

  // Create the table of contents headers.
  let tocRange = tocSheet.getRange("A2:B2")
  tocRange.setValues([["#", "Name"]]);

  // Get the range for the table of contents entries.
  let worksheets = workbook.getWorksheets();
  tocRange = tocRange.getResizedRange(worksheets.length, 0);

  // Loop through all worksheets in the workbook, except the first one.
  for (let i = 1; i < worksheets.length; i++) {
    // Create a row for each worksheet with its index and linked name.
    tocRange.getCell(i, 0).setValue(i);
    tocRange.getCell(i, 1).setHyperlink({
      textToDisplay: worksheets[i].getName(),
      documentReference: `'${worksheets[i].getName()}'!A1`
    });
  };

  // Activate the table of contents worksheet.
  tocSheet.activate();
}
```
