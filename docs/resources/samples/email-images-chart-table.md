---
title: Envie por email as imagens de um Excel gráfico e tabela
description: Saiba como usar Office scripts e Power Automate extrair e enviar por email as imagens de um gráfico Excel e tabela.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 5eb20025462614d62774ae6c088bdf397dcfb39d
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074589"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="4afce-103">Use Office scripts e Power Automate para enviar imagens de email de um gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="4afce-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="4afce-104">Este exemplo usa Office scripts e Power Automate para criar um gráfico.</span><span class="sxs-lookup"><span data-stu-id="4afce-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="4afce-105">Em seguida, envia em email imagens do gráfico e de sua tabela base.</span><span class="sxs-lookup"><span data-stu-id="4afce-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="4afce-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="4afce-106">Example scenario</span></span>

* <span data-ttu-id="4afce-107">Calcule para obter os resultados mais recentes.</span><span class="sxs-lookup"><span data-stu-id="4afce-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="4afce-108">Criar gráfico.</span><span class="sxs-lookup"><span data-stu-id="4afce-108">Create chart.</span></span>
* <span data-ttu-id="4afce-109">Obter imagens de gráfico e tabela.</span><span class="sxs-lookup"><span data-stu-id="4afce-109">Get chart and table images.</span></span>
* <span data-ttu-id="4afce-110">Envie um email para as imagens Power Automate.</span><span class="sxs-lookup"><span data-stu-id="4afce-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="4afce-111">_Dados de entrada_</span><span class="sxs-lookup"><span data-stu-id="4afce-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada.":::

<span data-ttu-id="4afce-113">_Gráfico de saída_</span><span class="sxs-lookup"><span data-stu-id="4afce-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de coluna criado mostrando o valor devido pelo cliente.":::

<span data-ttu-id="4afce-115">_Email recebido por meio de Power Automate fluxo_</span><span class="sxs-lookup"><span data-stu-id="4afce-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="O email enviado pelo fluxo mostrando o gráfico Excel incorporado no corpo.":::

## <a name="solution"></a><span data-ttu-id="4afce-117">Solução</span><span class="sxs-lookup"><span data-stu-id="4afce-117">Solution</span></span>

<span data-ttu-id="4afce-118">Esta solução tem duas partes:</span><span class="sxs-lookup"><span data-stu-id="4afce-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="4afce-119">Um Office script para calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="4afce-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="4afce-120">Um Power Automate fluxo para invocar o script e enviar por email os resultados.</span><span class="sxs-lookup"><span data-stu-id="4afce-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="4afce-121">Para ver um exemplo sobre como fazer isso, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="4afce-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="4afce-122">Código de exemplo: Calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="4afce-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="4afce-123">O script a seguir calcula e extrai um Excel gráfico e tabela.</span><span class="sxs-lookup"><span data-stu-id="4afce-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="4afce-124">Baixe o arquivo de <a href="email-chart-table.xlsx"> exemploemail-chart-table.xlsx</a> e use-o com este script para experimentar você mesmo!</span><span class="sxs-lookup"><span data-stu-id="4afce-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="4afce-125">Power Automate fluxo: envie por email as imagens do gráfico e da tabela</span><span class="sxs-lookup"><span data-stu-id="4afce-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="4afce-126">Esse fluxo executa o script e envia emails para as imagens retornadas.</span><span class="sxs-lookup"><span data-stu-id="4afce-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="4afce-127">Criar um novo **fluxo de nuvem instantânea.**</span><span class="sxs-lookup"><span data-stu-id="4afce-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="4afce-128">Selecione **Disparar manualmente um fluxo e** pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="4afce-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="4afce-129">Adicione uma **nova etapa** que usa o conector Excel **Online (Business)** com a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="4afce-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="4afce-130">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="4afce-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="4afce-131">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="4afce-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="4afce-132">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="4afce-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="4afce-133">**Arquivo**: sua pasta de trabalho ([selecionada com o seledor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="4afce-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="4afce-134">**Script**: Seu nome de script</span><span class="sxs-lookup"><span data-stu-id="4afce-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. <span data-ttu-id="4afce-136">Este exemplo usa Outlook como cliente de email.</span><span class="sxs-lookup"><span data-stu-id="4afce-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="4afce-137">Você pode usar qualquer conector de email Power Automate suporte, mas o restante das etapas supõe que você Outlook.</span><span class="sxs-lookup"><span data-stu-id="4afce-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="4afce-138">Adicione uma **nova etapa que** usa o conector **Office 365 Outlook** e a ação Enviar e **email (V2).**</span><span class="sxs-lookup"><span data-stu-id="4afce-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="4afce-139">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="4afce-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="4afce-140">**Para**: sua conta de email de teste (ou email pessoal)</span><span class="sxs-lookup"><span data-stu-id="4afce-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="4afce-141">**Assunto**: Revise dados do relatório</span><span class="sxs-lookup"><span data-stu-id="4afce-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="4afce-142">Para o **campo Corpo,** selecione "Exibição de Código" ( `</>` ) e insira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="4afce-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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
1. <span data-ttu-id="4afce-144">Salve o fluxo e experimente-o.</span><span class="sxs-lookup"><span data-stu-id="4afce-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="4afce-145">Vídeo de treinamento: Extrair e enviar imagens de email de gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="4afce-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="4afce-146">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="4afce-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
