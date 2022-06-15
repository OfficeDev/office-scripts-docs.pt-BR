---
title: 'Office de exemplo de scripts: Graph dados de nível de água do NOAA'
description: Um exemplo que busca dados JSON de um banco de dados NOAA e os usa para criar um gráfico.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4181edae7d8a46ae381ddfb1a2893b03faffd9b
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088096"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office de exemplo de scripts: buscar e grafar dados em nível de água do NOAA

Nesse cenário, você precisa plotar o nível da água na estação de Seattle da Administração Oceânico Nacional [e Atmosférica](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130). Você usará dados externos para popular uma planilha e criar um gráfico.

Você desenvolverá um script que usa o `fetch` comando para consultar o banco de [dados NoAA Tides e Currents](https://tidesandcurrents.noaa.gov/). Isso obterá o nível de água registrado em um determinado período de tempo. As informações serão retornadas como [JSON](https://www.w3schools.com/whatis/whatis_json.asp), portanto, parte do script converterá isso em valores de intervalo. Depois que os dados estão na planilha, eles serão usados para criar um gráfico.

Para obter mais informações sobre como trabalhar com JSON, [leia Usar JSON](../../develop/use-json.md) para passar dados de e para Office Scripts.

## <a name="scripting-skills-covered"></a>Habilidades de script cobertas

- Chamadas à API externa (`fetch`)
- Análise JSON
- Gráficos

## <a name="setup-instructions"></a>Instruções de instalação

1. Abra a pasta de trabalho com Excel na Web.

1. Na guia **Automatizar** , selecione **Novo Script** e cole o script a seguir no editor.

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook) {
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
      const rawJson: string = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
    
      // Note that we're only taking the data part of the JSON and excluding the metadata.
      const noaaData: NOAAData[] = JSON.parse(stringifiedJson).data;
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
    
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth, dataRange);
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
    
      /**
       * An interface to wrap the parts of the JSON we need.
       * These properties must match the names used in the JSON.
       */ 
      interface NOAAData {
        t: string; // Time
        v: number; // Level
      }
    }
    ```

1. Renomeie o script como **Gráfico de Nível de Água NOAA** e salve-o.

## <a name="running-the-script"></a>Executando o script

Em qualquer planilha, execute o script gráfico **de nível de água NOAA** . O script busca os dados de nível de água de 25 de dezembro de 2020 a 27 de dezembro de 2020. As `const` variáveis no início do script podem ser alteradas para usar datas diferentes ou obter informações de estação diferentes. A [API de CO-OPS para recuperação de](https://api.tidesandcurrents.noaa.gov/api/prod/) dados descreve como obter todos esses dados.

### <a name="after-running-the-script"></a>Depois de executar o script

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="A planilha depois de executar o script mostra alguns dados de nível de água e um gráfico.":::
