---
title: Email imagens de um gráfico e tabela do Excel
description: Saiba como usar os Scripts do Office e o Power Automate para extrair e enviar por email as imagens de um gráfico e tabela do Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572462"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Usar scripts do Office e o Power Automate para enviar imagens por email de um gráfico e tabela

Este exemplo usa Scripts do Office e o Power Automate para criar um gráfico. Em seguida, ele envia por email imagens do gráfico e sua tabela base.

## <a name="example-scenario"></a>Cenário de exemplo

* Calcule para obter os resultados mais recentes.
* Criar gráfico.
* Obter imagens de gráfico e tabela.
* Email as imagens com o Power Automate.

_Dados de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Uma planilha mostrando uma tabela de dados de entrada.":::

_Gráfico de saída_

:::image type="content" source="../../images/chart-created.png" alt-text="O gráfico de colunas criado mostrando o valor devido pelo cliente.":::

_Email que foi recebido por meio do fluxo do Power Automate_

:::image type="content" source="../../images/email-received.png" alt-text="O email enviado pelo fluxo mostrando o gráfico do Excel inserido no corpo.":::

## <a name="solution"></a>Solução

Essa solução tem duas partes:

1. [Um Script do Office para calcular e extrair o gráfico e a tabela do Excel](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Um fluxo do Power Automate para invocar o script e enviar por email os resultados. Para obter um exemplo de como fazer isso, consulte [Criar um fluxo de trabalho automatizado com o Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe [email-chart-table.xlsx](email-chart-table.xlsx) para uma pasta de trabalho pronta para uso. Adicione o script a seguir para experimentar o exemplo por conta própria!

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de exemplo: calcular e extrair gráfico e tabela do Excel

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Fluxo do Power Automate: Email imagens do gráfico e da tabela

Esse fluxo executa o script e envia por email as imagens retornadas.

1. Crie um fluxo **de nuvem instantâneo**.
1. Escolha **Disparar um fluxo manualmente e** selecione **Criar**.
1. Adicione uma **nova etapa que** usa o conector **do Excel Online (Business)** com a **ação Executar script** . Use os valores a seguir para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: sua pasta de trabalho ([selecionada com o seletor de arquivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: seu nome de script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="O conector completo do Excel Online (Business) no Power Automate.":::
1. Este exemplo usa o Outlook como o cliente de email. Você pode usar qualquer conector de email compatível com o Power Automate, mas o restante das etapas pressupõe que você escolheu o Outlook. Adicione uma **nova etapa que** usa o **Office 365 outlook** e a ação **Enviar e email (V2**). Use os valores a seguir para a ação.
    * **Para**: sua conta de email de teste (ou email pessoal)
    * **Assunto**: Examine os dados do relatório
    * Para o **campo Corpo** , selecione "Modo de Exibição de Código" (`</>`) e insira o seguinte:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="O conector Office 365 Outlook completo no Power Automate.":::
1. Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo por meio da **guia Meus fluxos** . Certifique-se de permitir o acesso quando solicitado.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de treinamento: Extrair e enviar por email imagens de gráfico e tabela

[Veja Sudhi Ramamurthy percorrer este exemplo no YouTube](https://youtu.be/152GJyqc-Kw).
