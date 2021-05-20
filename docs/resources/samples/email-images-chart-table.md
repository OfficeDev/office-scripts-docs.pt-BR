---
title: Envie por e-mail as imagens de um gráfico e tabela de Excel
description: Saiba como usar Office Scripts e Power Automate para extrair e enviar e-mails as imagens de um gráfico e tabela Excel.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545770"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="920cb-103">Use Office Scripts e Power Automate para enviar imagens de e-mail de um gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="920cb-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="920cb-104">Esta amostra usa Office Scripts e Power Automate para criar um gráfico.</span><span class="sxs-lookup"><span data-stu-id="920cb-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="920cb-105">Em seguida, envia imagens do gráfico e sua tabela base.</span><span class="sxs-lookup"><span data-stu-id="920cb-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="920cb-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="920cb-106">Example scenario</span></span>

* <span data-ttu-id="920cb-107">Calcule para obter os resultados mais recentes.</span><span class="sxs-lookup"><span data-stu-id="920cb-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="920cb-108">Criar gráfico.</span><span class="sxs-lookup"><span data-stu-id="920cb-108">Create chart.</span></span>
* <span data-ttu-id="920cb-109">Obter gráfico e imagens de tabela.</span><span class="sxs-lookup"><span data-stu-id="920cb-109">Get chart and table images.</span></span>
* <span data-ttu-id="920cb-110">Envie as imagens por e-mail com Power Automate.</span><span class="sxs-lookup"><span data-stu-id="920cb-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="920cb-111">_Dados de entrada_</span><span class="sxs-lookup"><span data-stu-id="920cb-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada":::

<span data-ttu-id="920cb-113">_Gráfico de saída_</span><span class="sxs-lookup"><span data-stu-id="920cb-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de colunas criado mostrando o valor devido pelo cliente":::

<span data-ttu-id="920cb-115">_E-mail recebido através de fluxo Power Automate_</span><span class="sxs-lookup"><span data-stu-id="920cb-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="O e-mail enviado pelo fluxo mostrando o gráfico de Excel incorporado no corpo":::

## <a name="solution"></a><span data-ttu-id="920cb-117">Solução</span><span class="sxs-lookup"><span data-stu-id="920cb-117">Solution</span></span>

<span data-ttu-id="920cb-118">Esta solução tem duas partes:</span><span class="sxs-lookup"><span data-stu-id="920cb-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="920cb-119">Um script Office para calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="920cb-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="920cb-120">Um fluxo Power Automate para invocar o script e enviar os resultados por e-mail.</span><span class="sxs-lookup"><span data-stu-id="920cb-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="920cb-121">Para obter um exemplo sobre como fazer isso, consulte [Criar um fluxo de trabalho automatizado com Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="920cb-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="920cb-122">Código amostral: Calcular e extrair Excel gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="920cb-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="920cb-123">O script a seguir calcula e extrai um gráfico e tabela Excel.</span><span class="sxs-lookup"><span data-stu-id="920cb-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="920cb-124">Baixe o arquivo de amostra <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> e use-o com este script para experimentá-lo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="920cb-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="920cb-125">fluxo de Power Automate: Envie um e-mail para o gráfico e as imagens da tabela</span><span class="sxs-lookup"><span data-stu-id="920cb-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="920cb-126">Esse fluxo executa o script e envia e-mails para as imagens devolvidas.</span><span class="sxs-lookup"><span data-stu-id="920cb-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="920cb-127">Crie um novo **fluxo de nuvens instantâneas.**</span><span class="sxs-lookup"><span data-stu-id="920cb-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="920cb-128">Selecione **Acionar manualmente um fluxo** e pressionar **Criar**.</span><span class="sxs-lookup"><span data-stu-id="920cb-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="920cb-129">Adicione uma **nova etapa** que usa o **conector Excel Online (Business)** com a ação **do script Run.**</span><span class="sxs-lookup"><span data-stu-id="920cb-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="920cb-130">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="920cb-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="920cb-131">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="920cb-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="920cb-132">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="920cb-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="920cb-133">**Arquivo**: Sua pasta de trabalho ([selecionada com o seletor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="920cb-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="920cb-134">**Script**: Seu nome de roteiro</span><span class="sxs-lookup"><span data-stu-id="920cb-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector Excel Online (Business) completo em Power Automate":::
1. <span data-ttu-id="920cb-136">Esta amostra usa Outlook como cliente de e-mail.</span><span class="sxs-lookup"><span data-stu-id="920cb-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="920cb-137">Você pode usar qualquer conector de e-mail Power Automate suportes, mas o resto das etapas presumem que você escolheu Outlook.</span><span class="sxs-lookup"><span data-stu-id="920cb-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="920cb-138">Adicione uma **nova etapa** que usa o **conector Office 365 Outlook** e a ação Enviar e **e-mail (V2).**</span><span class="sxs-lookup"><span data-stu-id="920cb-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="920cb-139">Use os seguintes valores para a ação:</span><span class="sxs-lookup"><span data-stu-id="920cb-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="920cb-140">**Para**: Sua conta de e-mail de teste (ou e-mail pessoal)</span><span class="sxs-lookup"><span data-stu-id="920cb-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="920cb-141">**Assunto**: Por favor, revise os dados do relatório</span><span class="sxs-lookup"><span data-stu-id="920cb-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="920cb-142">Para o campo **Corpo,** selecione "Code View" `</>` () e digite o seguinte:</span><span class="sxs-lookup"><span data-stu-id="920cb-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="O conector Office 365 Outlook concluído em Power Automate":::
1. <span data-ttu-id="920cb-144">Guarde o fluxo e experimente.</span><span class="sxs-lookup"><span data-stu-id="920cb-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="920cb-145">Vídeo de treinamento: Extrato e e-mail imagens de gráfico e tabela</span><span class="sxs-lookup"><span data-stu-id="920cb-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="920cb-146">[Assista Sudhi Ramamurthy andar através desta amostra no YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="920cb-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
