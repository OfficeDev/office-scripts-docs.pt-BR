---
title: 'Cenário de exemplo de Scripts do Office: Analisar downloads da Web'
description: Um exemplo que usa dados brutos de tráfego da Internet em uma pasta de trabalho do Excel e determina o local de origem, antes de organizar essas informações em uma tabela.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0ef368c5193fe65c0a01676aa2a8b3a2c5cf3bdc
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572420"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Cenário de exemplo de Scripts do Office: Analisar downloads da Web

Nesse cenário, você tem a tarefa de analisar relatórios de download do site da sua empresa. O objetivo dessa análise é determinar se o tráfego da Web está vindo da Estados Unidos ou de outro lugar do mundo.

Seus colegas carregam os dados brutos em sua pasta de trabalho. O conjunto de dados de cada semana tem sua própria planilha. Há também a planilha **Resumo** com uma tabela e um gráfico que mostra tendências semana a semana.

Você desenvolverá um script que analisa os dados de downloads semanais na planilha ativa. Ele analisará o endereço IP associado a cada download e determinará se ele veio ou não dos EUA. A resposta será inserida na planilha como um valor booliano ("TRUE" ou "FALSE") e a formatação condicional será aplicada a essas células. Os resultados do local do endereço IP serão totalados na planilha e copiados para a tabela de resumo.

## <a name="scripting-skills-covered"></a>Habilidades de script cobertas

- Análise de texto
- Subfunções em scripts
- Formatação condicional
- Tabelas

## <a name="setup-instructions"></a>Instruções de instalação

1. Baixe [analyze-web-downloads.xlsx](analyze-web-downloads.xlsx) para o OneDrive.

1. Abra a pasta de trabalho com Excel para a Web.

1. Na guia **Automatizar** , selecione **Novo Script** e cole o script a seguir no editor.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
        * End the script early if either object is not in the workbook.
        */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }

      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }

      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();

      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
          "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
        .getLastColumn()
        .getOffsetRange(0, 1);

      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues() as number[][];
      let logRangeValues = logRange.getValues() as string[][];
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);

      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol: (boolean | string)[][] = [];

      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }

      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];

      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);

      // Call the local function to add summary data to the worksheet.
      addSummaryData();

      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);

      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();

      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();

      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;

      /**
       * A function to add summary data on the worksheet.
        */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
          "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
          "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet.getRange("J1:M1");
        let summaryContentRow = currentWorksheet.getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
          .getOffsetRange(0, 2)
          .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
        ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
        ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
        formula1: "=TRUE",
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
        formula1: "=FALSE",
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
      // Split the IP address into octets.
      let octets = ipAddress.split(".");

      // Create a number for each octet and do the math to create the integer value of the IP address.
      let fullNum =
        // Define an arbitrary number for the last octet.
        111 +
        parseInt(octets[2]) * 256 +
        parseInt(octets[1]) * 65536 +
        parseInt(octets[0]) * 16777216;
      return fullNum;
    }
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

1. Renomeie o script para **Analisar Downloads da Web** e salve-o.

## <a name="running-the-script"></a>Executando o script

Navegue até qualquer uma das **planilhas\*\*** semanais e execute o script Analisar **Downloads da** Web. O script aplicará a formatação condicional e a rotulagem de local na planilha atual. Ele também atualizará **a planilha** Resumo.

### <a name="before-running-the-script"></a>Antes de executar o script

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Uma planilha que mostra dados brutos de tráfego da Web.":::

### <a name="after-running-the-script"></a>Depois de executar o script

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Uma planilha que mostra informações de localização de IP formatada com as linhas de tráfego da Web anteriores.":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="A tabela de resumo e o gráfico que resume as planilhas nas quais o script foi executado.":::
