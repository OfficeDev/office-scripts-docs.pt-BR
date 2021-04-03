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
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Usar Scripts do Office e Power Automate para enviar imagens de email de um gráfico e tabela

Este exemplo usa Scripts do Office e Power Automate para criar um gráfico. Em seguida, envia em email imagens do gráfico e de sua tabela base.

## <a name="example-scenario"></a>Cenário de exemplo

* Calcule para obter os resultados mais recentes.
* Criar gráfico.
* Obter imagens de gráfico e tabela.
* Envie um email para as imagens com o Power Automate.

_Dados de entrada_

![Dados de entrada](../../images/input-data.png)

_Gráfico de saída_

![Gráfico criado](../../images/chart-created.png)

_Email recebido por meio do fluxo do Power Automate_

![Email recebido](../../images/email-received.png)

## <a name="solution"></a>Solução

Esta solução tem duas partes:

1. [Um Script do Office para calcular e extrair gráfico e tabela do Excel](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Um fluxo do Power Automate para invocar o script e enviar por email os resultados. Para ver um exemplo sobre como fazer isso, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de exemplo: Calcular e extrair gráfico e tabela do Excel

O script a seguir calcula e extrai um gráfico e tabela do Excel.

Baixe o arquivo de <a href="email-chart-table.xlsx"> exemploemail-chart-table.xlsx</a> e use-o com este script para experimentar você mesmo!

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

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de treinamento: Extrair e enviar imagens de email de gráfico e tabela

[![Assista ao vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela")
