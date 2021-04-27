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
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Use Office scripts e Power Automate para enviar imagens de email de um gráfico e tabela

Este exemplo usa Office scripts e Power Automate para criar um gráfico. Em seguida, envia em email imagens do gráfico e de sua tabela base.

## <a name="example-scenario"></a>Cenário de exemplo

* Calcule para obter os resultados mais recentes.
* Criar gráfico.
* Obter imagens de gráfico e tabela.
* Envie um email para as imagens Power Automate.

_Dados de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada.":::

_Gráfico de saída_

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de coluna criado mostrando o valor devido pelo cliente.":::

_Email recebido por meio de Power Automate fluxo_

:::image type="content" source="../../images/email-received.png" alt-text="O email enviado pelo fluxo mostrando o gráfico Excel incorporado no corpo.":::

## <a name="solution"></a>Solução

Esta solução tem duas partes:

1. [Um Office script para calcular e extrair Excel gráfico e tabela](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Um Power Automate fluxo para invocar o script e enviar por email os resultados. Para ver um exemplo sobre como fazer isso, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de exemplo: Calcular e extrair Excel gráfico e tabela

O script a seguir calcula e extrai um Excel gráfico e tabela.

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate fluxo: envie por email as imagens do gráfico e da tabela

Esse fluxo executa o script e envia emails para as imagens retornadas.

1. Criar um novo **fluxo de nuvem instantânea.**
1. Selecione **Disparar manualmente um fluxo e** pressione **Criar**.
1. Adicione uma **nova etapa** que usa o conector Excel **Online (Business)** com a **ação Executar script (visualização).** Use os seguintes valores para a ação:
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: sua pasta de trabalho ([selecionada com o seledor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Seu nome de script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. Este exemplo usa Outlook como cliente de email. Você pode usar qualquer conector de email Power Automate suporte, mas o restante das etapas supõe que você Outlook. Adicione uma **nova etapa que** usa o conector **Office 365 Outlook** e a ação Enviar e **email (V2).** Use os seguintes valores para a ação:
    * **Para**: sua conta de email de teste (ou email pessoal)
    * **Assunto**: Revise dados do relatório
    * Para o **campo Corpo,** selecione "Exibição de Código" ( `</>` ) e insira o seguinte:

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
1. Salve o fluxo e experimente-o.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de treinamento: Extrair e enviar imagens de email de gráfico e tabela

[![Assista ao vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vídeo passo a passo sobre como extrair e enviar imagens de email de gráfico e tabela")
