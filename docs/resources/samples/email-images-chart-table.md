---
title: Envie por email as imagens de um gráfico e tabela do Excel
description: Saiba como usar os Scripts do Office e o Power Automate para extrair e enviar por email as imagens de um gráfico e tabela do Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 7eb12526f97d72de31acdc3c9a4228c670875e2b
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571066"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="37bed-103">Usar Scripts do Office e Power Automate para enviar imagens de email de um gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="37bed-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="37bed-104">Este exemplo usa Scripts do Office e Power Automate para criar um gráfico.</span><span class="sxs-lookup"><span data-stu-id="37bed-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="37bed-105">Em seguida, envia em email imagens do gráfico e de sua tabela base.</span><span class="sxs-lookup"><span data-stu-id="37bed-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="37bed-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="37bed-106">Example scenario</span></span>

* <span data-ttu-id="37bed-107">Calcule para obter os resultados mais recentes.</span><span class="sxs-lookup"><span data-stu-id="37bed-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="37bed-108">Criar gráfico.</span><span class="sxs-lookup"><span data-stu-id="37bed-108">Create chart.</span></span>
* <span data-ttu-id="37bed-109">Obter imagens de gráfico e tabela.</span><span class="sxs-lookup"><span data-stu-id="37bed-109">Get chart and table images.</span></span>
* <span data-ttu-id="37bed-110">Envie um email para as imagens com o Power Automate.</span><span class="sxs-lookup"><span data-stu-id="37bed-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="37bed-111">_Dados de entrada_</span><span class="sxs-lookup"><span data-stu-id="37bed-111">_Input data_</span></span>

![Dados de entrada](../../images/input-data.png)

<span data-ttu-id="37bed-113">_Gráfico de saída_</span><span class="sxs-lookup"><span data-stu-id="37bed-113">_Output chart_</span></span>

![Gráfico criado](../../images/chart-created.png)

<span data-ttu-id="37bed-115">_Email recebido por meio do fluxo do Power Automate_</span><span class="sxs-lookup"><span data-stu-id="37bed-115">_Email that was received through Power Automate flow_</span></span>

![Email recebido](../../images/email-received.png)

## <a name="solution"></a><span data-ttu-id="37bed-117">Solução</span><span class="sxs-lookup"><span data-stu-id="37bed-117">Solution</span></span>

<span data-ttu-id="37bed-118">Esta solução tem duas partes:</span><span class="sxs-lookup"><span data-stu-id="37bed-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="37bed-119">Um Script do Office para calcular e extrair gráfico e tabela do Excel</span><span class="sxs-lookup"><span data-stu-id="37bed-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="37bed-120">Um fluxo do Power Automate para invocar o script e enviar por email os resultados.</span><span class="sxs-lookup"><span data-stu-id="37bed-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="37bed-121">Para ver um exemplo sobre como fazer isso, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="37bed-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="37bed-122">Código de exemplo: Calcular e extrair gráfico e tabela do Excel</span><span class="sxs-lookup"><span data-stu-id="37bed-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="37bed-123">O script a seguir calcula e extrai um gráfico e tabela do Excel.</span><span class="sxs-lookup"><span data-stu-id="37bed-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="37bed-124">Baixe o arquivo de <a href="email-chart-table.xlsx"> exemploemail-chart-table.xlsx</a> e use-o com este script para experimentar você mesmo!</span><span class="sxs-lookup"><span data-stu-id="37bed-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="37bed-125">Vídeo de treinamento: Extrair e enviar imagens de email de gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="37bed-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="37bed-126">[![Assista ao vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela")</span><span class="sxs-lookup"><span data-stu-id="37bed-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
