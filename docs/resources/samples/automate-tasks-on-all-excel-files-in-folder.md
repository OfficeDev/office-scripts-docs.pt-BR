---
title: Executar um script em todos os arquivos do Excel em uma pasta
description: Saiba como executar um script em todos os arquivos Excel em uma pasta em OneDrive for Business.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: a595c31c9e0fa7066d6e18aff4d3778f727714b6
ms.sourcegitcommit: 9d00ee1c11cdf897410e5232692ee985f01ee098
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53772320"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Executar um script em todos os arquivos do Excel em uma pasta

Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta OneDrive for Business. Ele também pode ser usado em uma pasta SharePoint de dados.
Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions [um](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) colega.

## <a name="sample-excel-files"></a>Exemplo Excel arquivos

Baixe <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> todas as guias de trabalho que você precisará para este exemplo. Extraia esses arquivos para uma pasta intitulada **Vendas**. Adicione o seguinte script à sua coleção de scripts para experimentar o exemplo você mesmo!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Código de exemplo: Adicionar formatação e inserir comentário

Este é o script que é executado em cada manual de trabalho individual.

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate fluxo: execute o script em cada pasta de trabalho na pasta

Esse fluxo executa o script em cada pasta de trabalho na pasta "Vendas".

1. Criar um novo **fluxo de nuvem instantânea.**
1. Escolha **Disparar manualmente um fluxo e** selecione **Criar**.
1. Adicione uma **nova etapa que** usa o conector **OneDrive for Business** e os arquivos list na **ação de** pasta.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="O conector OneDrive for Business no Power Automate.":::
1. Selecione a pasta "Vendas" com as pastas de trabalho extraídas.
1. Para garantir que apenas as guias de trabalho sejam selecionadas, escolha **Nova etapa** e, em seguida, **selecione Condição**. Use os seguintes valores para a condição.
    1. **Nome** (o valor OneDrive nome do arquivo)
    1. "termina com"
    1. "xlsx"

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="O Power Automate de condição que aplica ações subsequentes a cada arquivo.":::
1. Na **ramificação Se sim,** adicione o **conector Excel Online (Business)** com a **ação Executar script.** Use os seguintes valores para a ação.
    1. **Localização**: OneDrive for Business
    1. **Biblioteca de Documentos**: OneDrive
    1. **Arquivo**: **Id** (o valor OneDrive ID do arquivo)
    1. **Script**: Seu nome de script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vídeo de treinamento: execute um script em todos os Excel arquivos em uma pasta

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/xMg711o7k6w).
