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
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Use Office Scripts e Power Automate para enviar imagens de e-mail de um gráfico e tabela

Esta amostra usa Office Scripts e Power Automate para criar um gráfico. Em seguida, envia imagens do gráfico e sua tabela base.

## <a name="example-scenario"></a>Cenário de exemplo

* Calcule para obter os resultados mais recentes.
* Criar gráfico.
* Obter gráfico e imagens de tabela.
* Envie as imagens por e-mail com Power Automate.

_Dados de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada":::

_Gráfico de saída_

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de colunas criado mostrando o valor devido pelo cliente":::

_E-mail recebido através de fluxo Power Automate_

:::image type="content" source="../../images/email-received.png" alt-text="O e-mail enviado pelo fluxo mostrando o gráfico de Excel incorporado no corpo":::

## <a name="solution"></a>Solução

Esta solução tem duas partes:

1. [Um script Office para calcular e extrair Excel gráfico e tabela](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Um fluxo Power Automate para invocar o script e enviar os resultados por e-mail. Para obter um exemplo sobre como fazer isso, consulte [Criar um fluxo de trabalho automatizado com Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código amostral: Calcular e extrair Excel gráfico e tabela

O script a seguir calcula e extrai um gráfico e tabela Excel.

Baixe o arquivo de amostra <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> e use-o com este script para experimentá-lo você mesmo!

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>fluxo de Power Automate: Envie um e-mail para o gráfico e as imagens da tabela

Esse fluxo executa o script e envia e-mails para as imagens devolvidas.

1. Crie um novo **fluxo de nuvens instantâneas.**
1. Selecione **Acionar manualmente um fluxo** e pressionar **Criar**.
1. Adicione uma **nova etapa** que usa o **conector Excel Online (Business)** com a ação **do script Run.** Use os seguintes valores para a ação:
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: Sua pasta de trabalho ([selecionada com o seletor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Seu nome de roteiro

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector Excel Online (Business) completo em Power Automate":::
1. Esta amostra usa Outlook como cliente de e-mail. Você pode usar qualquer conector de e-mail Power Automate suportes, mas o resto das etapas presumem que você escolheu Outlook. Adicione uma **nova etapa** que usa o **conector Office 365 Outlook** e a ação Enviar e **e-mail (V2).** Use os seguintes valores para a ação:
    * **Para**: Sua conta de e-mail de teste (ou e-mail pessoal)
    * **Assunto**: Por favor, revise os dados do relatório
    * Para o campo **Corpo,** selecione "Code View" `</>` () e digite o seguinte:

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
1. Guarde o fluxo e experimente.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de treinamento: Extrato e e-mail imagens de gráfico e tabela

[Assista Sudhi Ramamurthy andar através desta amostra no YouTube](https://youtu.be/152GJyqc-Kw).
