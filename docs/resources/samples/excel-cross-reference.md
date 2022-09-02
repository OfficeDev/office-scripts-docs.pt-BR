---
title: Arquivos do Excel de referência cruzada com o Power Automate
description: Saiba como usar scripts do Office e o Power Automate para fazer referência cruzada e formatar um arquivo do Excel.
ms.date: 06/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b32249dc7cb1e8c1b841a4db6caaff3b4d2998ec
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572672"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Arquivos do Excel de referência cruzada com o Power Automate

Esta solução mostra como comparar dados entre dois arquivos do Excel para encontrar discrepâncias. Ele usa Scripts do Office para analisar dados e o Power Automate para se comunicar entre as pastas de trabalho.

Este exemplo passa dados entre pastas de trabalho usando [objetos JSON](https://www.w3schools.com/whatis/whatis_json.asp) . Para obter mais informações sobre como trabalhar com JSON, [leia Usar JSON para passar dados de e para scripts do Office](../../develop/use-json.md).

## <a name="example-scenario"></a>Cenário de exemplo

Você é um coordenador de eventos que está agendando palestrantes para conferências futuras. Você mantém os dados do evento em uma planilha e os registros do locutor em outra. Para garantir que as duas pastas de trabalho sejam mantidas em sincronia, use um fluxo com scripts do Office para realçar possíveis problemas.

## <a name="sample-excel-files"></a>Arquivos de exemplo do Excel

Baixe os arquivos a seguir para obter pastas de trabalho prontas para uso para o exemplo.

1. [event-data.xlsx](event-data.xlsx)
1. [speaker-registrations.xlsx](speaker-registrations.xlsx)

Adicione os scripts a seguir para experimentar o exemplo por conta própria!

## <a name="sample-code-get-event-data"></a>Código de exemplo: obter dados de evento

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

## <a name="sample-code-validate-speaker-registrations"></a>Código de exemplo: Validar registros do locutor

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Fluxo do Power Automate: verifique se há inconsistências nas pastas de trabalho

Esse fluxo extrai as informações de evento da primeira pasta de trabalho e usa esses dados para validar a segunda pasta de trabalho.

1. Entre no [Power Automate e](https://flow.microsoft.com) crie um novo fluxo **de nuvem instantâneo**.
1. Escolha **Disparar um fluxo manualmente e** selecione **Criar**.
1. Adicione uma **nova etapa que** usa o conector **do Excel Online (Business)** com a **ação Executar script** . Use os valores a seguir para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: event-data.xlsx ([selecionado com o seletor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Obter dados de evento

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="O conector completo do Excel Online (Business) para o primeiro script no Power Automate.":::

1. Adicione uma segunda **nova etapa que** usa o conector **do Excel Online (Business)** com a **ação Executar script** . Isso usa os valores retornados do script **obter dados de** evento como entrada para o script **validar dados de** evento. Use os valores a seguir para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: speaker-registration.xlsx ([selecionado com o seletor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Validar o registro do locutor
    * **keys**: result (_dynamic content from **Run script**_)

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="O conector completo do Excel Online (Business) para o segundo script no Power Automate.":::
1. Este exemplo usa o Outlook como o cliente de email. Você pode usar qualquer conector de email compatível com o Power Automate. Adicione uma **nova etapa que** usa o **Office 365 outlook** e a ação **Enviar e email (V2**). Isso usa os valores retornados do script **de registro validar locutor** como o conteúdo do corpo do email. Use os valores a seguir para a ação.
    * **Para**: sua conta de email de teste (ou email pessoal)
    * **Assunto**: resultados da validação de evento
    * **Corpo**: resultado (_conteúdo dinâmico do **script de execução 2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="O conector Office 365 Outlook completo no Power Automate.":::
1. Salve o fluxo. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo por meio da **guia Meus fluxos** . Certifique-se de permitir o acesso quando solicitado.
1. Você deve receber um email dizendo "Incompatibilidade encontrada. Os dados exigem sua revisão." Isso indica que há diferenças entre linhas **emspeaker-registrations.xlsxe** linhas em **event-data.xlsx**. Abra **speaker-registrations.xlsx** para ver várias células realçadas em que há possíveis problemas com as listagem de registro do locutor.
