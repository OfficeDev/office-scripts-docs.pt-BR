---
title: Envie por email as imagens de um Excel gráfico e tabela
description: Saiba como usar Office scripts e Power Automate extrair e enviar por email as imagens de um gráfico Excel e tabela.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: 0265250f7fd885cb4899d0b9493b4285496965ff
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026860"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="1e696-103">Use Office scripts e Power Automate para enviar imagens de email de um gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="1e696-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="1e696-104">Este exemplo usa Office scripts e Power Automate para criar um gráfico.</span><span class="sxs-lookup"><span data-stu-id="1e696-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="1e696-105">Em seguida, envia em email imagens do gráfico e de sua tabela base.</span><span class="sxs-lookup"><span data-stu-id="1e696-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="1e696-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="1e696-106">Example scenario</span></span>

* <span data-ttu-id="1e696-107">Calcule para obter os resultados mais recentes.</span><span class="sxs-lookup"><span data-stu-id="1e696-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="1e696-108">Criar gráfico.</span><span class="sxs-lookup"><span data-stu-id="1e696-108">Create chart.</span></span>
* <span data-ttu-id="1e696-109">Obter imagens de gráfico e tabela.</span><span class="sxs-lookup"><span data-stu-id="1e696-109">Get chart and table images.</span></span>
* <span data-ttu-id="1e696-110">Envie um email para as imagens Power Automate.</span><span class="sxs-lookup"><span data-stu-id="1e696-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="1e696-111">_Dados de entrada_</span><span class="sxs-lookup"><span data-stu-id="1e696-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada.":::

<span data-ttu-id="1e696-113">_Gráfico de saída_</span><span class="sxs-lookup"><span data-stu-id="1e696-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de coluna criado mostrando o valor devido pelo cliente.":::

<span data-ttu-id="1e696-115">_Email recebido por meio de Power Automate fluxo_</span><span class="sxs-lookup"><span data-stu-id="1e696-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="O email enviado pelo fluxo mostrando o gráfico Excel incorporado no corpo.":::

## <a name="solution"></a><span data-ttu-id="1e696-117">Solução</span><span class="sxs-lookup"><span data-stu-id="1e696-117">Solution</span></span>

<span data-ttu-id="1e696-118">Esta solução tem duas partes:</span><span class="sxs-lookup"><span data-stu-id="1e696-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="1e696-119">Um Office script para calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="1e696-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="1e696-120">Um Power Automate fluxo para invocar o script e enviar por email os resultados.</span><span class="sxs-lookup"><span data-stu-id="1e696-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="1e696-121">Para ver um exemplo sobre como fazer isso, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="1e696-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="1e696-122">Código de exemplo: Calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="1e696-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="1e696-123">O script a seguir calcula e extrai um Excel gráfico e tabela.</span><span class="sxs-lookup"><span data-stu-id="1e696-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="1e696-124">Baixe o arquivo de <a href="email-chart-table.xlsx"> exemploemail-chart-table.xlsx</a> e use-o com este script para experimentar você mesmo!</span><span class="sxs-lookup"><span data-stu-id="1e696-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="1e696-125">Power Automate fluxo: envie por email as imagens do gráfico e da tabela</span><span class="sxs-lookup"><span data-stu-id="1e696-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="1e696-126">Esse fluxo executa o script e envia emails para as imagens retornadas.</span><span class="sxs-lookup"><span data-stu-id="1e696-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="1e696-127">Criar um novo **fluxo de nuvem instantânea.**</span><span class="sxs-lookup"><span data-stu-id="1e696-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="1e696-128">Selecione **Disparar manualmente um fluxo e** pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="1e696-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="1e696-129">Adicione uma **nova etapa** que usa o conector Excel **Online (Business)** com a **ação Executar script (visualização).**</span><span class="sxs-lookup"><span data-stu-id="1e696-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="1e696-130">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="1e696-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="1e696-131">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="1e696-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="1e696-132">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="1e696-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="1e696-133">**Arquivo**: sua pasta de trabalho ([selecionada com o seledor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="1e696-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="1e696-134">**Script**: Seu nome de script</span><span class="sxs-lookup"><span data-stu-id="1e696-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. <span data-ttu-id="1e696-136">Este exemplo usa Outlook como cliente de email.</span><span class="sxs-lookup"><span data-stu-id="1e696-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="1e696-137">Você pode usar qualquer conector de email Power Automate suporte, mas o restante das etapas supõe que você Outlook.</span><span class="sxs-lookup"><span data-stu-id="1e696-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="1e696-138">Adicione uma **nova etapa que** usa o conector **Office 365 Outlook** e a ação Enviar e **email (V2).**</span><span class="sxs-lookup"><span data-stu-id="1e696-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="1e696-139">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="1e696-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="1e696-140">**Para**: sua conta de email de teste (ou email pessoal)</span><span class="sxs-lookup"><span data-stu-id="1e696-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="1e696-141">**Assunto**: Revise dados do relatório</span><span class="sxs-lookup"><span data-stu-id="1e696-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="1e696-142">Para o **campo Corpo,** selecione "Exibição de Código" ( `</>` ) e insira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="1e696-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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
1. <span data-ttu-id="1e696-144">Salve o fluxo e experimente-o.</span><span class="sxs-lookup"><span data-stu-id="1e696-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="1e696-145">Vídeo de treinamento: Extrair e enviar imagens de email de gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="1e696-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="1e696-146">[![Assista ao vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela")</span><span class="sxs-lookup"><span data-stu-id="1e696-146">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
