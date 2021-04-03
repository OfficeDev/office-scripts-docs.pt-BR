---
title: Executar um script em todos os arquivos do Excel em uma pasta
description: Saiba como executar um script em todos os arquivos do Excel em uma pasta no OneDrive for Business.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: a11876e8241a069a7c640bbcf2c36b4842d3bd90
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571144"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Executar um script em todos os arquivos do Excel em uma pasta

Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta no OneDrive for Business. Ele também pode ser usado em uma pasta do SharePoint.
Ele executa cálculos nos arquivos do Excel, adiciona formatação e insere um comentário que @mentions [um](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) colega.

## <a name="sample-code-add-formatting-and-insert-comment"></a>Código de exemplo: Adicionar formatação e inserir comentário

Baixe o arquivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraia os arquivos para uma pasta intitulada **Vendas** usada neste exemplo e experimente você mesmo!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vídeo de treinamento: execute um script em todos os arquivos do Excel em uma pasta

[Assista a um vídeo passo](https://youtu.be/xMg711o7k6w) a passo sobre como executar um script em todos os arquivos do Excel em uma pasta do OneDrive for Business ou do SharePoint.
