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
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="2ddbd-103">Executar um script em todos os arquivos do Excel em uma pasta</span><span class="sxs-lookup"><span data-stu-id="2ddbd-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="2ddbd-104">Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta OneDrive for Business.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="2ddbd-105">Ele também pode ser usado em uma pasta SharePoint de dados.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="2ddbd-106">Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions [um](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) colega.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="2ddbd-107">Baixe o arquivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraia os arquivos para uma pasta intitulada **Vendas** usada neste exemplo e experimente você mesmo!</span><span class="sxs-lookup"><span data-stu-id="2ddbd-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="2ddbd-108">Código de exemplo: Adicionar formatação e inserir comentário</span><span class="sxs-lookup"><span data-stu-id="2ddbd-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="2ddbd-109">Este é o script que é executado em cada manual de trabalho individual.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-109">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="2ddbd-110">Power Automate fluxo: execute o script em cada pasta de trabalho na pasta</span><span class="sxs-lookup"><span data-stu-id="2ddbd-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="2ddbd-111">Esse fluxo executa o script em cada pasta de trabalho na pasta "Vendas".</span><span class="sxs-lookup"><span data-stu-id="2ddbd-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="2ddbd-112">Criar um novo **fluxo de nuvem instantânea.**</span><span class="sxs-lookup"><span data-stu-id="2ddbd-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="2ddbd-113">Selecione **Disparar manualmente um fluxo e** pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="2ddbd-114">Adicione uma **nova etapa que** usa o conector **OneDrive for Business** e os arquivos list na **ação de** pasta.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="O conector OneDrive for Business no Power Automate.":::
1. <span data-ttu-id="2ddbd-116">Selecione a pasta "Vendas" com as pastas de trabalho extraídas.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="2ddbd-117">Para garantir que apenas as guias de trabalho sejam selecionadas, escolha **Nova etapa** e, em seguida, selecione **Condição** e de definir os seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="2ddbd-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="2ddbd-118">**Nome** (o valor OneDrive nome do arquivo)</span><span class="sxs-lookup"><span data-stu-id="2ddbd-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="2ddbd-119">"termina com"</span><span class="sxs-lookup"><span data-stu-id="2ddbd-119">"ends with"</span></span>
    1. <span data-ttu-id="2ddbd-120">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="2ddbd-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="O Power Automate de condição que aplica ações subsequentes a cada arquivo.":::
1. <span data-ttu-id="2ddbd-122">Na **ramificação Se sim,** adicione o **conector Excel Online (Business)** com a ação **Executar script (visualização).**</span><span class="sxs-lookup"><span data-stu-id="2ddbd-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="2ddbd-123">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="2ddbd-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="2ddbd-124">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="2ddbd-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="2ddbd-125">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2ddbd-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="2ddbd-126">**Arquivo**: **Id** (o valor OneDrive ID do arquivo)</span><span class="sxs-lookup"><span data-stu-id="2ddbd-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="2ddbd-127">**Script**: Seu nome de script</span><span class="sxs-lookup"><span data-stu-id="2ddbd-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. <span data-ttu-id="2ddbd-129">Salve o fluxo e experimente-o.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="2ddbd-130">Vídeo de treinamento: execute um script em todos os Excel arquivos em uma pasta</span><span class="sxs-lookup"><span data-stu-id="2ddbd-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="2ddbd-131">[Assista a um vídeo passo](https://youtu.be/xMg711o7k6w) a passo sobre como executar um script em todos os arquivos Excel em uma pasta OneDrive for Business ou SharePoint.</span><span class="sxs-lookup"><span data-stu-id="2ddbd-131">[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.</span></span>
