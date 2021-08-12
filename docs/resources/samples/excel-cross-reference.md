---
title: Referência cruzada Excel arquivos com Power Automate
description: Saiba como usar Office scripts e Power Automate para fazer referência cruzada e formatar um arquivo Excel.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: ddbcdd25791e0c1a80fedfc36ebbfbd5dd940ec6f55ef2fe2bce0cf23b6bcb61
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847237"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Referência cruzada Excel arquivos com Power Automate

Esta solução mostra como comparar dados entre dois arquivos Excel para encontrar discrepâncias. Ele usa Office scripts para analisar dados e Power Automate para se comunicar entre as guias de trabalho.

## <a name="example-scenario"></a>Cenário de exemplo

Você é um coordenador de eventos que está agendando palestrantes para próximas conferências. Você mantém os dados do evento em uma planilha e os registros do alto-falante em outra. Para garantir que as duas guias de trabalho sejam mantidas em sincronia, use um fluxo com Office Scripts para realçar quaisquer possíveis problemas.

## <a name="sample-excel-files"></a>Exemplo Excel arquivos

Baixe os arquivos a seguir para obter pastas de trabalho prontas para uso para o exemplo.

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

Adicione os scripts a seguir para experimentar o exemplo você mesmo!

## <a name="sample-code-get-event-data"></a>Código de exemplo: Obter dados de evento

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a>Código de exemplo: Validar registros de alto-falantes

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate fluxo: Verifique se há inconsistências nas guias de trabalho

Esse fluxo extrai as informações de evento da primeira workbook e usa esses dados para validar a segunda workbook.

1. Entre [Power Automate](https://flow.microsoft.com) e crie um novo fluxo **de nuvem instantâneo.**
1. Escolha **Disparar manualmente um fluxo e** selecione **Criar**.
1. Adicione uma **nova etapa** que usa o conector Excel **Online (Business)** com a **ação Executar script.** Use os seguintes valores para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: event-data.xlsx ([selecionado com o seledor de arquivo](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Obter dados de evento

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="O conector Excel online (Business) concluído para o primeiro script no Power Automate.":::

1. Adicione uma segunda **nova etapa que** usa o conector Excel Online **(Business)** com a **ação Executar script.** Use os seguintes valores para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: speaker-registration.xlsx ([selecionado com o seledor de arquivo](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Validar o registro de alto-falantes

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="O conector Excel online (Business) concluído para o segundo script no Power Automate.":::
1. Este exemplo usa Outlook como cliente de email. Você pode usar qualquer conector de email Power Automate suporte. Adicione uma **nova etapa que** usa o conector **Office 365 Outlook** e a ação Enviar e **email (V2).** Use os seguintes valores para a ação.
    * **Para**: sua conta de email de teste (ou email pessoal)
    * **Assunto**: Resultados da validação de eventos
    * **Body**: result (_conteúdo dinâmico de Executar script **2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="O conector Office 365 Outlook no Power Automate.":::
1. Salve o fluxo. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.
1. Você deve receber um email dizendo "Incompatibilidade encontrada. Os dados exigem sua revisão." Isso indica que há diferenças entre linhas em **speaker-registrations.xlsx** e linhas em **event-data.xlsx**. Abra **speaker-registrations.xlsx** para ver várias células realçadas onde há possíveis problemas com as listagem de registro do alto-falante.
