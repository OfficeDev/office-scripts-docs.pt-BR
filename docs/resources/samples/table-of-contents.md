---
title: Criar um sumário da pasta de trabalho
description: Saiba como criar um sumário com links para cada planilha.
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5b158160ecb9ac29df547c6da6552e21c9875be3
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572511"
---
# <a name="create-a-workbook-table-of-contents"></a>Criar um sumário da pasta de trabalho

Este exemplo mostra como criar um sumário para a pasta de trabalho. Cada entrada no sumário é um hiperlink para uma das planilhas na pasta de trabalho.

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="A planilha de sumário mostrando links para as outras planilhas.":::

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe [table-of-contents.xlsx](table-of-contents.xlsx) para uma pasta de trabalho pronta para uso. Adicione o script a seguir e experimente o exemplo por conta própria!

## <a name="sample-code-create-a-workbook-table-of-contents"></a>Código de exemplo: Criar um sumário da pasta de trabalho

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
