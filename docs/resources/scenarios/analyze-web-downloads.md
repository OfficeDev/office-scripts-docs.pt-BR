---
title: 'Office Cenário de exemplo de scripts: Analisar downloads da Web'
description: Um exemplo que coleta dados brutos de tráfego da Internet em uma Excel de trabalho e determina o local de origem, antes de organizar essas informações em uma tabela.
ms.date: 04/27/2021
localization_priority: Normal
ms.openlocfilehash: bdd6b43290e5432d87c4a85a35fbaf32967fbf03
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074456"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="7db1e-103">Office Cenário de exemplo de scripts: Analisar downloads da Web</span><span class="sxs-lookup"><span data-stu-id="7db1e-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="7db1e-104">Nesse cenário, você tem a tarefa de analisar relatórios de download do site da sua empresa.</span><span class="sxs-lookup"><span data-stu-id="7db1e-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="7db1e-105">O objetivo dessa análise é determinar se o tráfego da Web está vindo dos Estados Unidos ou de qualquer outro lugar do mundo.</span><span class="sxs-lookup"><span data-stu-id="7db1e-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="7db1e-106">Seus colegas carregam os dados brutos na sua workbook.</span><span class="sxs-lookup"><span data-stu-id="7db1e-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="7db1e-107">O conjunto de dados de cada semana tem sua própria planilha.</span><span class="sxs-lookup"><span data-stu-id="7db1e-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="7db1e-108">Há também a planilha **Resumo** com uma tabela e um gráfico que mostra tendências semana após semana.</span><span class="sxs-lookup"><span data-stu-id="7db1e-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="7db1e-109">Você desenvolverá um script que analisa os dados de downloads semanais na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="7db1e-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="7db1e-110">Ele analisará o endereço IP associado a cada download e determinará se ele veio ou não dos EUA.</span><span class="sxs-lookup"><span data-stu-id="7db1e-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="7db1e-111">A resposta será inserida na planilha como um valor booleano ("TRUE" ou "FALSE") e a formatação condicional será aplicada a essas células.</span><span class="sxs-lookup"><span data-stu-id="7db1e-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="7db1e-112">Os resultados do local do endereço IP serão totalados na planilha e copiados para a tabela de resumo.</span><span class="sxs-lookup"><span data-stu-id="7db1e-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="7db1e-113">Habilidades de script abordadas</span><span class="sxs-lookup"><span data-stu-id="7db1e-113">Scripting skills covered</span></span>

- <span data-ttu-id="7db1e-114">Análise de texto</span><span class="sxs-lookup"><span data-stu-id="7db1e-114">Text parsing</span></span>
- <span data-ttu-id="7db1e-115">Subfunções em scripts</span><span class="sxs-lookup"><span data-stu-id="7db1e-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="7db1e-116">Formatação condicional</span><span class="sxs-lookup"><span data-stu-id="7db1e-116">Conditional formatting</span></span>
- <span data-ttu-id="7db1e-117">Tabelas</span><span class="sxs-lookup"><span data-stu-id="7db1e-117">Tables</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="7db1e-118">Instruções de instalação</span><span class="sxs-lookup"><span data-stu-id="7db1e-118">Setup instructions</span></span>

1. <span data-ttu-id="7db1e-119">Baixe <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> para seu OneDrive.</span><span class="sxs-lookup"><span data-stu-id="7db1e-119">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="7db1e-120">Abra a workbook com Excel para a Web.</span><span class="sxs-lookup"><span data-stu-id="7db1e-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="7db1e-121">Na guia **Automatizar,** abra **Todos os Scripts.**</span><span class="sxs-lookup"><span data-stu-id="7db1e-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="7db1e-122">No painel de tarefas Editor de **Código,** pressione **Novo Script** e colar o seguinte script no editor.</span><span class="sxs-lookup"><span data-stu-id="7db1e-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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
      let newCol = [];
  
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

5. <span data-ttu-id="7db1e-123">Renomeie o script para **Analisar Downloads da Web** e salve-o.</span><span class="sxs-lookup"><span data-stu-id="7db1e-123">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="7db1e-124">Executando o script</span><span class="sxs-lookup"><span data-stu-id="7db1e-124">Running the script</span></span>

<span data-ttu-id="7db1e-125">Navegue até qualquer uma **das planilhas \* \* semanais** e execute o script Analisar **Downloads da Web.**</span><span class="sxs-lookup"><span data-stu-id="7db1e-125">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="7db1e-126">O script aplicará a formatação condicional e a rotulagem de local na planilha atual.</span><span class="sxs-lookup"><span data-stu-id="7db1e-126">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="7db1e-127">Ele também atualizará **a planilha Resumo.**</span><span class="sxs-lookup"><span data-stu-id="7db1e-127">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="7db1e-128">Antes de executar o script</span><span class="sxs-lookup"><span data-stu-id="7db1e-128">Before running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Uma planilha que mostra dados brutos de tráfego da Web.":::

### <a name="after-running-the-script"></a><span data-ttu-id="7db1e-130">Depois de executar o script</span><span class="sxs-lookup"><span data-stu-id="7db1e-130">After running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Uma planilha que mostra informações de localização IP formatada com as linhas de tráfego da Web anteriores.":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="A tabela de resumo e o gráfico que resume as planilhas nas quais o script foi executado.":::
