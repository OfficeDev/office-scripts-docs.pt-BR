---
title: Executar um script em todos os arquivos do Excel em uma pasta
description: Saiba como executar um script em todos os arquivos Excel em uma pasta em OneDrive for Business.
ms.date: 04/02/2021
localization_priority: Normal
ms.openlocfilehash: 6376dcac0eb36c04c2b60b2717d18cd730a0a8ee
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026835"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Executar um script em todos os arquivos do Excel em uma pasta

Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta OneDrive for Business. Ele também pode ser usado em uma pasta SharePoint de dados.
Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions [um](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) colega.

Baixe o arquivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraia os arquivos para uma pasta intitulada **Vendas** usada neste exemplo e experimente você mesmo!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Código de exemplo: Adicionar formatação e inserir comentário

Este é o script que é executado em cada manual de trabalho individual.

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate fluxo: execute o script em cada pasta de trabalho na pasta

Esse fluxo executa o script em cada pasta de trabalho na pasta "Vendas".

1. Criar um novo **fluxo de nuvem instantânea.**
1. Selecione **Disparar manualmente um fluxo e** pressione **Criar**.
1. Adicione uma **nova etapa que** usa o conector **OneDrive for Business** e os arquivos list na **ação de** pasta.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="O conector OneDrive for Business no Power Automate.":::
1. Selecione a pasta "Vendas" com as pastas de trabalho extraídas.
1. Para garantir que apenas as guias de trabalho sejam selecionadas, escolha **Nova etapa** e, em seguida, selecione **Condição** e de definir os seguintes valores:
    1. **Nome** (o valor OneDrive nome do arquivo)
    1. "termina com"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="O Power Automate de condição que aplica ações subsequentes a cada arquivo.":::
1. Na **ramificação Se sim,** adicione o **conector Excel Online (Business)** com a ação **Executar script (visualização).** Use os seguintes valores para a ação:
    1. **Localização**: OneDrive for Business
    1. **Biblioteca de Documentos**: OneDrive
    1. **Arquivo**: **Id** (o valor OneDrive ID do arquivo)
    1. **Script**: Seu nome de script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. Salve o fluxo e experimente-o.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vídeo de treinamento: execute um script em todos os Excel arquivos em uma pasta

[Assista a um vídeo passo](https://youtu.be/xMg711o7k6w) a passo sobre como executar um script em todos os arquivos Excel em uma pasta OneDrive for Business ou SharePoint.
