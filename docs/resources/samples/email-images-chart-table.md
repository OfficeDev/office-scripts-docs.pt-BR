---
title: Envie por email as imagens de um Excel gráfico e tabela
description: Saiba como usar Office scripts e Power Automate extrair e enviar por email as imagens de um gráfico Excel e tabela.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 50bc65c82df7f5fc68dbebf942c4f607bb6af60a
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313838"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="827e3-103">Use Office scripts e Power Automate para enviar imagens de email de um gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="827e3-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="827e3-104">Este exemplo usa Office scripts e Power Automate para criar um gráfico.</span><span class="sxs-lookup"><span data-stu-id="827e3-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="827e3-105">Em seguida, envia em email imagens do gráfico e de sua tabela base.</span><span class="sxs-lookup"><span data-stu-id="827e3-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="827e3-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="827e3-106">Example scenario</span></span>

* <span data-ttu-id="827e3-107">Calcule para obter os resultados mais recentes.</span><span class="sxs-lookup"><span data-stu-id="827e3-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="827e3-108">Criar gráfico.</span><span class="sxs-lookup"><span data-stu-id="827e3-108">Create chart.</span></span>
* <span data-ttu-id="827e3-109">Obter imagens de gráfico e tabela.</span><span class="sxs-lookup"><span data-stu-id="827e3-109">Get chart and table images.</span></span>
* <span data-ttu-id="827e3-110">Envie um email para as imagens Power Automate.</span><span class="sxs-lookup"><span data-stu-id="827e3-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="827e3-111">_Dados de entrada_</span><span class="sxs-lookup"><span data-stu-id="827e3-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada.":::

<span data-ttu-id="827e3-113">_Gráfico de saída_</span><span class="sxs-lookup"><span data-stu-id="827e3-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de coluna criado mostrando o valor devido pelo cliente.":::

<span data-ttu-id="827e3-115">_Email recebido por meio de Power Automate fluxo_</span><span class="sxs-lookup"><span data-stu-id="827e3-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="O email enviado pelo fluxo mostrando o gráfico Excel incorporado no corpo.":::

## <a name="solution"></a><span data-ttu-id="827e3-117">Solução</span><span class="sxs-lookup"><span data-stu-id="827e3-117">Solution</span></span>

<span data-ttu-id="827e3-118">Esta solução tem duas partes:</span><span class="sxs-lookup"><span data-stu-id="827e3-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="827e3-119">Um Office script para calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="827e3-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="827e3-120">Um Power Automate fluxo para invocar o script e enviar por email os resultados.</span><span class="sxs-lookup"><span data-stu-id="827e3-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="827e3-121">Para ver um exemplo sobre como fazer isso, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="827e3-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="827e3-122">Exemplo Excel arquivo</span><span class="sxs-lookup"><span data-stu-id="827e3-122">Sample Excel file</span></span>

<span data-ttu-id="827e3-123">Baixe <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> para uma workbook pronta para uso.</span><span class="sxs-lookup"><span data-stu-id="827e3-123">Download <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="827e3-124">Adicione o seguinte script para experimentar o exemplo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="827e3-124">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="827e3-125">Código de exemplo: Calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="827e3-125">Sample code: Calculate and extract Excel chart and table</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="827e3-126">Power Automate fluxo: envie por email as imagens do gráfico e da tabela</span><span class="sxs-lookup"><span data-stu-id="827e3-126">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="827e3-127">Esse fluxo executa o script e envia emails para as imagens retornadas.</span><span class="sxs-lookup"><span data-stu-id="827e3-127">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="827e3-128">Criar um novo **fluxo de nuvem instantânea.**</span><span class="sxs-lookup"><span data-stu-id="827e3-128">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="827e3-129">Escolha **Disparar manualmente um fluxo e** selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="827e3-129">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="827e3-130">Adicione uma **nova etapa** que usa o conector Excel **Online (Business)** com a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="827e3-130">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="827e3-131">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="827e3-131">Use the following values for the action:</span></span>
    * <span data-ttu-id="827e3-132">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="827e3-132">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="827e3-133">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="827e3-133">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="827e3-134">**Arquivo**: sua pasta de trabalho ([selecionada com o seledor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="827e3-134">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="827e3-135">**Script**: Seu nome de script</span><span class="sxs-lookup"><span data-stu-id="827e3-135">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. <span data-ttu-id="827e3-137">Este exemplo usa Outlook como cliente de email.</span><span class="sxs-lookup"><span data-stu-id="827e3-137">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="827e3-138">Você pode usar qualquer conector de email Power Automate suporte, mas o restante das etapas supõe que você Outlook.</span><span class="sxs-lookup"><span data-stu-id="827e3-138">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="827e3-139">Adicione uma **nova etapa que** usa o conector **Office 365 Outlook** e a ação Enviar e **email (V2).**</span><span class="sxs-lookup"><span data-stu-id="827e3-139">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="827e3-140">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="827e3-140">Use the following values for the action:</span></span>
    * <span data-ttu-id="827e3-141">**Para**: sua conta de email de teste (ou email pessoal)</span><span class="sxs-lookup"><span data-stu-id="827e3-141">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="827e3-142">**Assunto**: Revise dados do relatório</span><span class="sxs-lookup"><span data-stu-id="827e3-142">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="827e3-143">Para o **campo Corpo,** selecione "Exibição de Código" ( `</>` ) e insira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="827e3-143">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="O conector Office 365 Outlook no Power Automate.":::
1. <span data-ttu-id="827e3-145">Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.</span><span class="sxs-lookup"><span data-stu-id="827e3-145">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="827e3-146">Vídeo de treinamento: Extrair e enviar imagens de email de gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="827e3-146">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="827e3-147">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="827e3-147">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
