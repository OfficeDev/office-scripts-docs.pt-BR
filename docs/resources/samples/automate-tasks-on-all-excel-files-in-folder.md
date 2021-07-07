---
title: Executar um script em todos os arquivos do Excel em uma pasta
description: Saiba como executar um script em todos os arquivos Excel em uma pasta em OneDrive for Business.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: bf9c0c486dacced5c3017b267ea65dfd215a5197
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313894"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="79239-103">Executar um script em todos os arquivos do Excel em uma pasta</span><span class="sxs-lookup"><span data-stu-id="79239-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="79239-104">Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta OneDrive for Business.</span><span class="sxs-lookup"><span data-stu-id="79239-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="79239-105">Ele também pode ser usado em uma pasta SharePoint de dados.</span><span class="sxs-lookup"><span data-stu-id="79239-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="79239-106">Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions [um](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) colega.</span><span class="sxs-lookup"><span data-stu-id="79239-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="79239-107">Exemplo Excel arquivos</span><span class="sxs-lookup"><span data-stu-id="79239-107">Sample Excel files</span></span>

<span data-ttu-id="79239-108">Baixe <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> todas as guias de trabalho que você precisará para este exemplo.</span><span class="sxs-lookup"><span data-stu-id="79239-108">Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample.</span></span> <span data-ttu-id="79239-109">Extraia esses arquivos para uma pasta intitulada **Vendas**.</span><span class="sxs-lookup"><span data-stu-id="79239-109">Extract those files to a folder titled **Sales**.</span></span> <span data-ttu-id="79239-110">Adicione o seguinte script à sua coleção de scripts para experimentar o exemplo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="79239-110">Add the following script to your script collection to try the sample yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="79239-111">Código de exemplo: Adicionar formatação e inserir comentário</span><span class="sxs-lookup"><span data-stu-id="79239-111">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="79239-112">Este é o script que é executado em cada manual de trabalho individual.</span><span class="sxs-lookup"><span data-stu-id="79239-112">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="79239-113">Power Automate fluxo: execute o script em cada pasta de trabalho na pasta</span><span class="sxs-lookup"><span data-stu-id="79239-113">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="79239-114">Esse fluxo executa o script em cada pasta de trabalho na pasta "Vendas".</span><span class="sxs-lookup"><span data-stu-id="79239-114">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="79239-115">Criar um novo **fluxo de nuvem instantânea.**</span><span class="sxs-lookup"><span data-stu-id="79239-115">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="79239-116">Escolha **Disparar manualmente um fluxo e** selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="79239-116">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="79239-117">Adicione uma **nova etapa que** usa o conector **OneDrive for Business** e os arquivos list na **ação de** pasta.</span><span class="sxs-lookup"><span data-stu-id="79239-117">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="O conector OneDrive for Business no Power Automate.":::
1. <span data-ttu-id="79239-119">Selecione a pasta "Vendas" com as pastas de trabalho extraídas.</span><span class="sxs-lookup"><span data-stu-id="79239-119">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="79239-120">Para garantir que apenas as guias de trabalho sejam selecionadas, escolha **Nova etapa** e, em seguida, selecione **Condição** e de definir os seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="79239-120">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="79239-121">**Nome** (o valor OneDrive nome do arquivo)</span><span class="sxs-lookup"><span data-stu-id="79239-121">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="79239-122">"termina com"</span><span class="sxs-lookup"><span data-stu-id="79239-122">"ends with"</span></span>
    1. <span data-ttu-id="79239-123">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="79239-123">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="O Power Automate de condição que aplica ações subsequentes a cada arquivo.":::
1. <span data-ttu-id="79239-125">Na **ramificação Se sim,** adicione o **conector Excel Online (Business)** com a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="79239-125">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="79239-126">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="79239-126">Use the following values for the action:</span></span>
    1. <span data-ttu-id="79239-127">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="79239-127">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="79239-128">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="79239-128">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="79239-129">**Arquivo**: **Id** (o valor OneDrive ID do arquivo)</span><span class="sxs-lookup"><span data-stu-id="79239-129">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="79239-130">**Script**: Seu nome de script</span><span class="sxs-lookup"><span data-stu-id="79239-130">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. <span data-ttu-id="79239-132">Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.</span><span class="sxs-lookup"><span data-stu-id="79239-132">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="79239-133">Vídeo de treinamento: execute um script em todos os Excel arquivos em uma pasta</span><span class="sxs-lookup"><span data-stu-id="79239-133">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="79239-134">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="79239-134">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
