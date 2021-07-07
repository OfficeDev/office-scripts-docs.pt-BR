---
title: Referência cruzada Excel arquivos com Power Automate
description: Saiba como usar Office scripts e Power Automate para fazer referência cruzada e formatar um arquivo Excel.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 0776ce49cacecfa15339cc7c0cd4866daad789ff
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313957"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="29eee-103">Referência cruzada Excel arquivos com Power Automate</span><span class="sxs-lookup"><span data-stu-id="29eee-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="29eee-104">Esta solução mostra como comparar dados entre dois arquivos Excel para encontrar discrepâncias.</span><span class="sxs-lookup"><span data-stu-id="29eee-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="29eee-105">Ele usa Office scripts para analisar dados e Power Automate para se comunicar entre as guias de trabalho.</span><span class="sxs-lookup"><span data-stu-id="29eee-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="29eee-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="29eee-106">Example scenario</span></span>

<span data-ttu-id="29eee-107">Você é um coordenador de eventos que está agendando palestrantes para próximas conferências.</span><span class="sxs-lookup"><span data-stu-id="29eee-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="29eee-108">Você mantém os dados do evento em uma planilha e os registros do alto-falante em outra.</span><span class="sxs-lookup"><span data-stu-id="29eee-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="29eee-109">Para garantir que as duas guias de trabalho sejam mantidas em sincronia, use um fluxo com Office Scripts para realçar quaisquer possíveis problemas.</span><span class="sxs-lookup"><span data-stu-id="29eee-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="29eee-110">Exemplo Excel arquivos</span><span class="sxs-lookup"><span data-stu-id="29eee-110">Sample Excel files</span></span>

<span data-ttu-id="29eee-111">Baixe os arquivos a seguir para obter pastas de trabalho prontas para uso para o exemplo.</span><span class="sxs-lookup"><span data-stu-id="29eee-111">Download the following files to get ready-to-use workbooks for the sample.</span></span>

1. <span data-ttu-id="29eee-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="29eee-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="29eee-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="29eee-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

<span data-ttu-id="29eee-114">Adicione os scripts a seguir para experimentar o exemplo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="29eee-114">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="29eee-115">Código de exemplo: Obter dados de evento</span><span class="sxs-lookup"><span data-stu-id="29eee-115">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="29eee-116">Código de exemplo: Validar registros de alto-falantes</span><span class="sxs-lookup"><span data-stu-id="29eee-116">Sample code: Validate speaker registrations</span></span>

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="29eee-117">Power Automate fluxo: Verifique se há inconsistências nas guias de trabalho</span><span class="sxs-lookup"><span data-stu-id="29eee-117">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="29eee-118">Esse fluxo extrai as informações de evento da primeira workbook e usa esses dados para validar a segunda workbook.</span><span class="sxs-lookup"><span data-stu-id="29eee-118">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="29eee-119">Entre [Power Automate](https://flow.microsoft.com) e crie um novo fluxo **de nuvem instantâneo.**</span><span class="sxs-lookup"><span data-stu-id="29eee-119">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="29eee-120">Escolha **Disparar manualmente um fluxo e** selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="29eee-120">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="29eee-121">Adicione uma **nova etapa** que usa o conector Excel **Online (Business)** com a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="29eee-121">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="29eee-122">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="29eee-122">Use the following values for the action:</span></span>
    * <span data-ttu-id="29eee-123">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="29eee-123">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="29eee-124">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="29eee-124">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="29eee-125">**Arquivo**: event-data.xlsx ([selecionado com o seledor de arquivo](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="29eee-125">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="29eee-126">**Script**: Obter dados de evento</span><span class="sxs-lookup"><span data-stu-id="29eee-126">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="O conector Excel online (Business) concluído para o primeiro script no Power Automate.":::

1. <span data-ttu-id="29eee-128">Adicione uma segunda **nova etapa que** usa o conector Excel Online **(Business)** com a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="29eee-128">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="29eee-129">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="29eee-129">Use the following values for the action:</span></span>
    * <span data-ttu-id="29eee-130">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="29eee-130">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="29eee-131">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="29eee-131">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="29eee-132">**Arquivo**: speaker-registration.xlsx ([selecionado com o seledor de arquivo](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="29eee-132">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="29eee-133">**Script**: Validar o registro de alto-falantes</span><span class="sxs-lookup"><span data-stu-id="29eee-133">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="O conector Excel online (Business) concluído para o segundo script no Power Automate.":::
1. <span data-ttu-id="29eee-135">Este exemplo usa Outlook como cliente de email.</span><span class="sxs-lookup"><span data-stu-id="29eee-135">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="29eee-136">Você pode usar qualquer conector de email Power Automate suporte.</span><span class="sxs-lookup"><span data-stu-id="29eee-136">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="29eee-137">Adicione uma **nova etapa que** usa o conector **Office 365 Outlook** e a ação Enviar e **email (V2).**</span><span class="sxs-lookup"><span data-stu-id="29eee-137">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="29eee-138">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="29eee-138">Use the following values for the action:</span></span>
    * <span data-ttu-id="29eee-139">**Para**: sua conta de email de teste (ou email pessoal)</span><span class="sxs-lookup"><span data-stu-id="29eee-139">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="29eee-140">**Assunto**: Resultados da validação de eventos</span><span class="sxs-lookup"><span data-stu-id="29eee-140">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="29eee-141">**Body**: result (_conteúdo dinâmico de Executar script **2**_)</span><span class="sxs-lookup"><span data-stu-id="29eee-141">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="O conector Office 365 Outlook no Power Automate.":::
1. <span data-ttu-id="29eee-143">Salve o fluxo.</span><span class="sxs-lookup"><span data-stu-id="29eee-143">Save the flow.</span></span> <span data-ttu-id="29eee-144">Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.</span><span class="sxs-lookup"><span data-stu-id="29eee-144">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>
1. <span data-ttu-id="29eee-145">Você deve receber um email dizendo "Incompatibilidade encontrada.</span><span class="sxs-lookup"><span data-stu-id="29eee-145">You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="29eee-146">Os dados exigem sua revisão."</span><span class="sxs-lookup"><span data-stu-id="29eee-146">Data requires your review."</span></span> <span data-ttu-id="29eee-147">Isso indica que há diferenças entre linhas em **speaker-registrations.xlsx** e linhas em **event-data.xlsx**.</span><span class="sxs-lookup"><span data-stu-id="29eee-147">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="29eee-148">Abra **speaker-registrations.xlsx** para ver várias células realçadas onde há possíveis problemas com as listagem de registro do alto-falante.</span><span class="sxs-lookup"><span data-stu-id="29eee-148">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
