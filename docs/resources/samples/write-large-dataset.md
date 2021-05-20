---
title: Escrever um grande conjuntos de dados
description: Aprenda a dividir um grande conjunto de dados em operações de gravação menores em scripts Office.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 06abb58c61c18620d638ab3eb61ea68398bf20aa
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545620"
---
# <a name="write-a-large-dataset"></a>Escrever um grande conjuntos de dados

A `Range.setValues()` API coloca os dados em um intervalo. Essa API tem limitações dependendo de vários fatores, como tamanho de dados e configurações de rede. Isso significa que se você tentar escrever uma quantidade maciça de informações para uma pasta de trabalho como uma única operação, você precisará escrever os dados em lotes menores, a fim de atualizar de forma confiável uma [grande gama](../../testing/platform-limits.md).

Para o básico de desempenho em Office Scripts, leia [Melhorar o desempenho de seus scripts Office](../../develop/web-client-performance.md).

## <a name="sample-code-write-a-large-dataset"></a>Código de amostra: Escreva um grande conjunto de dados

Este script escreve linhas de uma gama em partes menores. Ele seleciona 1000 células para escrever de cada vez. Execute o script em uma planilha em branco para ver os lotes de atualização em ação. A saída do console dá mais informações sobre o que está acontecendo.

> [!NOTE]
> Você pode alterar o número de linhas totais que estão sendo escritas alterando o valor de `SAMPLE_ROWS` . Você pode alterar o número de células para escrever como uma única ação alterando o valor de `CELLS_IN_BATCH` .

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const SAMPLE_ROWS = 100000;
  const CELLS_IN_BATCH = 10000;

  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  console.log(`Generating data...`)
  let data: (string | number | boolean)[][] = [];
  // Generate six columns of random data per row. 
  for (let i = 0; i < SAMPLE_ROWS; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInBatches(sheet.getRange("B2"), data, CELLS_IN_BATCH);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInBatches(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInBatch: number
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per batch setting: ${cellsInBatch}`);

  // Determine the total number of cells to write.
  const totalCells = values.length * values[0].length;
  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInBatch) {
    console.log(`No need to batch -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  // Determine how many rows to write at once.
  const rowsPerBatch = Math.floor(cellsInBatch / values[0].length);
  console.log("Rows per batch: " + rowsPerBatch);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let batchCount = 0;

  // Write each batch of rows.
  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerBatch) {
      batchCount++;
      console.log(`Calling update next batch function. Batch#: ${batchCount}`);
      updateNextBatch(startCell, values, rowsPerBatch, totalRowsUpdated);

      // Write a completion percentage to help the user understand the progress.
      rowCount = 0;
      totalRowsUpdated += rowsPerBatch;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);
    }
  }
  
  console.log(`Updating remaining rows -- last batch: ${rowCount}`)
  if (rowCount > 0) {
    updateNextBatch(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInBatch).toFixed(8)} seconds per ${cellsInBatch} cells-batch.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */
function updateNextBatch(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerBatch: number,
  totalRowsUpdated: number
) {
  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerBatch - 1, data[0].length - 1);
  console.log(`Updating batch at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerBatch);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the batch range: ${JSON.stringify(e)}`;
  }
  return;
}

/**
 * A helper function that computes the target range given the target range's starting cell
 * and selected range and updates the values.
 */
function updateTargetRange(
  targetCell: ExcelScript.Range,
  values: (string | boolean | number)[][]
) {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range: ${targetRange.getAddress()}`);
  try {
    targetRange.setValues(values);
  } catch (e) {
    throw `Error while updating the whole range: ${JSON.stringify(e)}`;
  }
  return;
}

// Credit: https://www.codegrepper.com/code-examples/javascript/random+text+generator+javascript
function getRandomString(length: number): string {
  var randomChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += randomChars.charAt(Math.floor(Math.random() * randomChars.length));
  }
  return result;
}
```

## <a name="training-video-write-a-large-dataset"></a>Vídeo de treinamento: Escreva um grande conjunto de dados

[Assista Sudhi Ramamurthy andar através desta amostra no YouTube](https://youtu.be/BP9Kp0Ltj7U).
