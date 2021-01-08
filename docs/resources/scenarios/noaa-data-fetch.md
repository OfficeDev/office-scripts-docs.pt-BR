---
title: 'Cenário de exemplo de scripts do Office: dados em nível de água do Graph do NOAA'
description: Um exemplo que busca dados JSON de um banco de dados NOAA e os usa para criar um gráfico.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: d2afcd05125ea66c028d8e21bcc878371c20fcc3
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784176"
---
# <a name="office-scripts-sample-scenario-graph-water-level-data-from-noaa"></a><span data-ttu-id="4c7a4-103">Cenário de exemplo de scripts do Office: dados em nível de água do Graph do NOAA</span><span class="sxs-lookup"><span data-stu-id="4c7a4-103">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>

<span data-ttu-id="4c7a4-104">Neste cenário, você precisa plotar o nível de água na estação de Seattle da Administração Nacional do Oceano e da Administração [Administrativa DeSuce.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)</span><span class="sxs-lookup"><span data-stu-id="4c7a4-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="4c7a4-105">Você usará dados externos para preencher uma planilha e criar um gráfico.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="4c7a4-106">Você desenvolverá um script que usa o comando para consultar o banco de dados `fetch` [noAA - NoAA -](https://tidesandcurrents.noaa.gov/)Banco de dados currents.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="4c7a4-107">Isso obterá o nível de água registrado em um determinado intervalo de tempo.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="4c7a4-108">As informações serão retornadas como JSON, portanto, parte do script as converterá em valores de intervalo.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="4c7a4-109">Quando os dados estão na planilha, eles serão usados para fazer um gráfico.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="4c7a4-110">Habilidades de script cobertas</span><span class="sxs-lookup"><span data-stu-id="4c7a4-110">Scripting skills covered</span></span>

- <span data-ttu-id="4c7a4-111">Chamadas de API externa ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="4c7a4-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="4c7a4-112">Análise JSON</span><span class="sxs-lookup"><span data-stu-id="4c7a4-112">JSON parsing</span></span>
- <span data-ttu-id="4c7a4-113">Gráficos</span><span class="sxs-lookup"><span data-stu-id="4c7a4-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="4c7a4-114">Instruções de instalação</span><span class="sxs-lookup"><span data-stu-id="4c7a4-114">Setup instructions</span></span>

1. <span data-ttu-id="4c7a4-115">Abra a planilha com o Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="4c7a4-116">Na guia **Automatizar,** selecione **Todos os Scripts.**</span><span class="sxs-lookup"><span data-stu-id="4c7a4-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="4c7a4-117">No painel **de tarefas editor** de código, selecione Novo **script** e colar o seguinte script no editor.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
      const noaaData = JSON.parse(stringifiedJson);
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.data.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData.data[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
      
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth,dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);

      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1", 
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    }
    ```

1. <span data-ttu-id="4c7a4-118">Renomeie o script para Gráfico **de Nível de Água NOAA** e salve-o.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="4c7a4-119">Executando o script</span><span class="sxs-lookup"><span data-stu-id="4c7a4-119">Running the script</span></span>

<span data-ttu-id="4c7a4-120">Em qualquer planilha, execute o script Gráfico de Nível de **Água NOAA.**</span><span class="sxs-lookup"><span data-stu-id="4c7a4-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="4c7a4-121">O script busca os dados de nível de água de 25 de dezembro de 2020 a 27 de dezembro de 2020.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="4c7a4-122">As `const` variáveis no início do script podem ser alteradas para usar datas diferentes ou obter informações de estação diferentes.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="4c7a4-123">A [API de CO-OPS para recuperação de](https://api.tidesandcurrents.noaa.gov/api/prod/) dados descreve como obter todos esses dados.</span><span class="sxs-lookup"><span data-stu-id="4c7a4-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="4c7a4-124">Depois de executar o script</span><span class="sxs-lookup"><span data-stu-id="4c7a4-124">After running the script</span></span>

![A planilha após a execução do script mostra alguns dados de nível de água e um gráfico.](../../images/scenario-noaa-water-level-after.png)
