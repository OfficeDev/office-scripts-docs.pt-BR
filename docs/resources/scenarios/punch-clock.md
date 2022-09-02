---
title: 'Cenário de exemplo de Scripts do Office: botão Depurar relógio'
description: Este exemplo adiciona um botão de relógio de ponche e permite que um usuário entre e saia usando a hora atual.
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: ac128a33b653506b6168bd4acfe1713bf6d26759
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572679"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>Cenário de exemplo de Scripts do Office: botão Depurar relógio

A ideia de cenário e o script usados neste exemplo foram contribuição do membro da comunidade scripts do Office [Brian Gonzalez](https://github.com/b-gonzalez).

Nesse cenário, você criará uma folha de horários para um funcionário que permite que ele registre seus horários de início e término com a pressionamento de um [botão](../../develop/script-buttons.md). Com base no que foi gravado anteriormente, pressionar o botão iniciará o dia (entrada do relógio) ou terminará o dia (saída do relógio). O exemplo funciona para o Excel na Web e no Windows.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Uma tabela com três colunas ('Clock In', 'Clock Out' e 'Duration') e um botão rotulado 'Punch clock' na pasta de trabalho.":::

## <a name="setup-instructions"></a>Instruções de instalação

1. Baixe [punch-clock-sample.xlsx](punch-clock-sample.xlsx) para o OneDrive.

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="Uma tabela com três colunas: 'Clock In', 'Clock Out' e 'Duration'.":::

1. Abra a pasta de trabalho Excel na Web.

1. Na guia **Automatizar** , selecione **Novo Script** e cole o script a seguir no editor.

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. Renomeie o script como "Relógio de ponche".

1. Salve o script.

1. Na pasta de trabalho, selecione a célula **E2**.

1. Adicionar um botão de script. Vá para o menu **Mais opções (...) na** página **Detalhes do script** e selecione **o botão Adicionar**.

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="O menu &quot;Mais opções&quot; e o botão &quot;Adicionar&quot;.":::

1. Salve a pasta de trabalho.

## <a name="run-the-script"></a>Executar o script

Pressione o **botão Relógio de** ponche para executar o script. Ele registra a hora atual em "Clock In" ou "Clock Out", dependendo do que foi inserido anteriormente.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="A tabela e o botão &quot;Relógio de ponche&quot; na pasta de trabalho.":::

> [!NOTE]
> A duração só será registrada se for maior que um minuto. Edite manualmente a hora "Clock In" para testar durações maiores.
