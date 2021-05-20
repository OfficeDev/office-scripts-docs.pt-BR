---
title: Executar um script em todos os arquivos do Excel em uma pasta
description: Aprenda a executar um script em todos os arquivos Excel em uma pasta em OneDrive for Business.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545785"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Executar um script em todos os arquivos do Excel em uma pasta

Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta em OneDrive for Business. Também pode ser usado em uma pasta SharePoint.
Ele realiza cálculos sobre os arquivos Excel, adiciona formatação e insere um comentário que [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) um colega.

Baixe o arquivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraia os arquivos para uma pasta intitulada **Sales** usado nesta amostra e experimente você mesmo!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Código de amostra: Adicionar formatação e inserir comentário

Este é o script que é executado em cada livro de trabalho individual.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>fluxo Power Automate: Execute o script em cada pasta de trabalho da pasta

Esse fluxo executa o script em cada pasta de trabalho da pasta "Vendas".

1. Crie um novo **fluxo de nuvens instantâneas.**
1. Selecione **Acionar manualmente um fluxo** e pressionar **Criar**.
1. Adicione uma **nova etapa** que usa o **conector OneDrive for Business** e os arquivos Lista em ação de **pasta.**

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="O conector OneDrive for Business concluído em Power Automate":::
1. Selecione a pasta "Vendas" com as pastas de trabalho extraídas.
1. Para garantir que apenas as pastas de trabalho sejam selecionadas, escolha **Nova etapa,** selecione **Condição** e defina os seguintes valores:
    1. **Nome** (o valor do nome do arquivo OneDrive)
    1. "termina com"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="O bloco de condição Power Automate que aplica ações subsequentes a cada arquivo":::
1. Na filial **If yes,** adicione o **conector Excel Online (Business)** com a ação **do script Run.** Use os seguintes valores para a ação:
    1. **Localização**: OneDrive for Business
    1. **Biblioteca de Documentos**: OneDrive
    1. **Arquivo**: **Id** (o valor de ID do arquivo OneDrive)
    1. **Script**: Seu nome de roteiro

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="O conector Excel Online (Business) completo em Power Automate":::
1. Guarde o fluxo e experimente.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vídeo de treinamento: Execute um script em todos os arquivos Excel em uma pasta

[Assista Sudhi Ramamurthy andar através desta amostra no YouTube](https://youtu.be/xMg711o7k6w).
