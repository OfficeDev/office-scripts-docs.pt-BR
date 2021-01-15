---
title: 'Cenário de exemplo de scripts do Office: dados em nível de água do Graph do NOAA'
description: Um exemplo que busca dados JSON de um banco de dados NOAA e os usa para criar um gráfico.
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 5b0b4e3675cbe053368f63123d819f0dab626e60
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867874"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Cenário de exemplo de scripts do Office: buscar e gráfico dados no nível d'água do NOAA

Neste cenário, você precisa plotar o nível de água na estação de Seattle da Administração Nacional do Oceano e da Administração [Administrativa DeSuce.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130) Você usará dados externos para preencher uma planilha e criar um gráfico.

Você desenvolverá um script que usa o comando para consultar o banco de dados `fetch` [noAA - NoAA -](https://tidesandcurrents.noaa.gov/)Banco de dados currents. Isso obterá o nível de água registrado em um determinado intervalo de tempo. As informações serão retornadas como JSON, portanto, parte do script as converterá em valores de intervalo. Quando os dados estão na planilha, eles serão usados para fazer um gráfico.

## <a name="scripting-skills-covered"></a>Habilidades de script cobertas

- Chamadas de API externa ( `fetch` )
- Análise JSON
- Gráficos

## <a name="setup-instructions"></a>Instruções de instalação

1. Abra a planilha com o Excel na Web.

1. Na guia **Automatizar,** selecione **Todos os Scripts.**

1. No painel **de tarefas editor** de código, selecione Novo **script** e colar o seguinte script no editor.

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

1. Renomeie o script para Gráfico **de Nível de Água NOAA** e salve-o.

## <a name="running-the-script"></a>Executando o script

Em qualquer planilha, execute o script Gráfico de Nível de **Água NOAA.** O script busca os dados de nível de água de 25 de dezembro de 2020 a 27 de dezembro de 2020. As `const` variáveis no início do script podem ser alteradas para usar datas diferentes ou obter informações de estação diferentes. A [API de CO-OPS para recuperação de](https://api.tidesandcurrents.noaa.gov/api/prod/) dados descreve como obter todos esses dados.

### <a name="after-running-the-script"></a>Depois de executar o script

![A planilha após a execução do script mostra alguns dados de nível de água e um gráfico.](../../images/scenario-noaa-water-level-after.png)
