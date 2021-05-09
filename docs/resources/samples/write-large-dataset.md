---
title: Otimização de desempenho ao escrever um grande conjuntos de dados
description: Saiba como otimizar o desempenho ao escrever um grande conjuntos de dados em Office Scripts.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 9622494378a24db16ea43b5500d6efa156726ff8
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285945"
---
# <a name="performance-optimization-when-writing-a-large-dataset"></a><span data-ttu-id="3426b-103">Otimização de desempenho ao escrever um grande conjuntos de dados</span><span class="sxs-lookup"><span data-stu-id="3426b-103">Performance optimization when writing a large dataset</span></span>

## <a name="basic-performance-optimization"></a><span data-ttu-id="3426b-104">Otimização básica de desempenho</span><span class="sxs-lookup"><span data-stu-id="3426b-104">Basic performance optimization</span></span>

<span data-ttu-id="3426b-105">Para obter noções básicas Office scripts, consulte a seção [de](getting-started.md#basic-performance-considerations) desempenho do artigo Getting Started.</span><span class="sxs-lookup"><span data-stu-id="3426b-105">For performance basics in Office Scripts, see the [performance section](getting-started.md#basic-performance-considerations) of the Getting Started article.</span></span>

## <a name="sample-code-optimize-performance-of-a-large-dataset"></a><span data-ttu-id="3426b-106">Código de exemplo: otimizar o desempenho de um grande conjuntos de dados</span><span class="sxs-lookup"><span data-stu-id="3426b-106">Sample code: Optimize performance of a large dataset</span></span>

<span data-ttu-id="3426b-107">A `setValues()` API range permite definir os valores de um intervalo.</span><span class="sxs-lookup"><span data-stu-id="3426b-107">The `setValues()` Range API allows setting the values of a range.</span></span> <span data-ttu-id="3426b-108">Essa API tem limitações de dados dependendo de vários fatores, como tamanho de dados, configurações de rede, etc. Para atualizar com confiança uma grande variedade de dados, você precisará pensar em fazer atualizações de dados em partes menores.</span><span class="sxs-lookup"><span data-stu-id="3426b-108">This API has data limitations depending on various factors such as data size, network settings, etc. In order to reliably update a large range of data, you'll need to think about doing data updates in smaller chunks.</span></span> <span data-ttu-id="3426b-109">Esse script tenta fazer isso e grava linhas de um intervalo em partes para que, se um intervalo grande precisar ser atualizado, ele possa ser feito em partes menores.</span><span class="sxs-lookup"><span data-stu-id="3426b-109">This script attempts to do this and writes rows of a range in chunks so that if a large range needs to be updated, it can be done in smaller parts.</span></span> <span data-ttu-id="3426b-110">**Aviso:** ele não foi testado em vários tamanhos, portanto, esteja ciente disso se você quiser usá-lo em seu script.</span><span class="sxs-lookup"><span data-stu-id="3426b-110">**Warning**: It has not been tested across various sizes so be aware of that if you want to use this in your script.</span></span> <span data-ttu-id="3426b-111">À medida que temos a oportunidade de testar, atualizaremos com as descobertas sobre como ele é feito para vários tamanhos de dados.</span><span class="sxs-lookup"><span data-stu-id="3426b-111">As we have opportunity to test, we'll update with findings around how it performs for various data sizes.</span></span>

<span data-ttu-id="3426b-112">Este script seleciona células 1K por parte, mas você pode substituir para testar como ele funciona para você.</span><span class="sxs-lookup"><span data-stu-id="3426b-112">This script selects 1K cells per chunk but you can override to test out how it works for you.</span></span> <span data-ttu-id="3426b-113">Ele atualiza linhas de 100 mil com 6 colunas de dados.</span><span class="sxs-lookup"><span data-stu-id="3426b-113">It updates 100k rows with 6 columns of data.</span></span> <span data-ttu-id="3426b-114">Execute isso em uma planilha em branco para examinar.</span><span class="sxs-lookup"><span data-stu-id="3426b-114">Run this on a blank sheet to examine.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  let data: (string | number | boolean)[][] = [];
  // Number of rows in the random data (x 6 columns).
  const sampleRows = 100000;

  console.log(`Generating data...`)
  // Dynamically generate some random data for testing purpose. 
  for (let i = 0; i < sampleRows; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInChunks(sheet.getRange("B2"), data);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInChunks(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInChunk: number = 10000
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per chunk setting: ${cellsInChunk}`);
  if (!values) {
    console.log(`Invalid input values to update.`);
    return false;
  }
  if (values.length === 0 || values[0].length === 0) {
    console.log(`Empty data -- nothing to update.`);
    return true;
  }
  const totalCells = values.length * values[0].length;

  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInChunk) {
    console.log(`No need to chunk -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  const rowsPerChunk = Math.floor(cellsInChunk / values[0].length);
  console.log("Rows per chunk: " + rowsPerChunk);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let chunkCount = 0;

  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerChunk) {
      chunkCount++;
      console.log(`Calling update next chunk function. Chunk#: ${chunkCount}`);
      updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
      rowCount = 0;
      totalRowsUpdated += rowsPerChunk;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);

    }
  }
  console.log(`Updating remaining rows -- last chunk: ${rowCount}`)
  if (rowCount > 0) {
    updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInChunk).toFixed(8)} seconds per ${cellsInChunk} cells-chunk.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */

function updateNextChunk(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerChunk: number,
  totalRowsUpdated: number
) {

  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
  console.log(`Updating chunk at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the chunk range: ${JSON.stringify(e)}`;
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

## <a name="training-video-optimize-performance-when-writing-a-large-dataset"></a><span data-ttu-id="3426b-115">Vídeo de treinamento: otimizar o desempenho ao escrever um grande conjuntos de dados</span><span class="sxs-lookup"><span data-stu-id="3426b-115">Training video: Optimize performance when writing a large dataset</span></span>

<span data-ttu-id="3426b-116">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/BP9Kp0Ltj7U).</span><span class="sxs-lookup"><span data-stu-id="3426b-116">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/BP9Kp0Ltj7U).</span></span>
