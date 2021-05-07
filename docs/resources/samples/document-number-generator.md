---
title: Gerar um identificador exclusivo em uma workbook
description: Saiba como usar Office Scripts para gerar um identificador exclusivo e adicionar uma linha a uma tabela e intervalo.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 62c930bfc638dc46b36daf81b6d1ec976c90a8d0
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232477"
---
# <a name="generate-a-unique-identifier-in-a-workbook"></a><span data-ttu-id="8ab23-103">Gerar um identificador exclusivo em uma workbook</span><span class="sxs-lookup"><span data-stu-id="8ab23-103">Generate a unique identifier in a workbook</span></span>

<span data-ttu-id="8ab23-104">Esse cenário ajuda um usuário a gerar um número de documento exclusivo com um formato específico e a adicioná-lo como uma entrada a um intervalo ou tabela.</span><span class="sxs-lookup"><span data-stu-id="8ab23-104">This scenario helps a user generate a unique document number with a specific format and add it as an entry to a range or table.</span></span> <span data-ttu-id="8ab23-105">A nova entrada ou linha adicionada conterá o número de documento exclusivo recém-gerado e alguns outros atributos passados para o script.</span><span class="sxs-lookup"><span data-stu-id="8ab23-105">The new entry or row added will contain the newly generated unique document number and a few other attributes passed to the script.</span></span>

<span data-ttu-id="8ab23-106">Há duas versões do exemplo para este cenário.</span><span class="sxs-lookup"><span data-stu-id="8ab23-106">There are two versions of the sample for this scenario.</span></span>

* [<span data-ttu-id="8ab23-107">Versão 1: ler e adicionar uma linha a uma planilha que contém intervalo simples</span><span class="sxs-lookup"><span data-stu-id="8ab23-107">Version 1: Read and add a row to a worksheet containing plain range</span></span>](#sample-code-generate-key-and-add-row-to-range)

    <span data-ttu-id="8ab23-108">_Antes que a nova linha seja adicionada_</span><span class="sxs-lookup"><span data-stu-id="8ab23-108">_Before the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-range-before.png" alt-text="Uma planilha mostrando um intervalo de dados antes que a linha seja adicionada":::

    <span data-ttu-id="8ab23-110">_Depois que a nova linha é adicionada_</span><span class="sxs-lookup"><span data-stu-id="8ab23-110">_After the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-range-after.png" alt-text="Uma planilha mostrando um intervalo de dados depois que a linha é adicionada":::

* [<span data-ttu-id="8ab23-112">Versão 2: ler e adicionar uma linha a uma tabela</span><span class="sxs-lookup"><span data-stu-id="8ab23-112">Version 2: Read and add a row to a table</span></span>](#sample-code-generate-key-and-add-row-to-table)

    <span data-ttu-id="8ab23-113">_Antes que a nova linha seja adicionada_</span><span class="sxs-lookup"><span data-stu-id="8ab23-113">_Before the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-table-before.png" alt-text="Uma planilha mostrando uma tabela antes de uma linha ser adicionada":::

    <span data-ttu-id="8ab23-115">_Depois que a nova linha é adicionada_</span><span class="sxs-lookup"><span data-stu-id="8ab23-115">_After the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-table-after.png" alt-text="Uma planilha mostrando uma tabela depois que uma linha é adicionada":::

## <a name="sample-excel-file"></a><span data-ttu-id="8ab23-117">Exemplo Excel arquivo</span><span class="sxs-lookup"><span data-stu-id="8ab23-117">Sample Excel file</span></span>

<span data-ttu-id="8ab23-118">Baixe o arquivo <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> usado nesta solução para experimentar você mesmo!</span><span class="sxs-lookup"><span data-stu-id="8ab23-118">Download the file <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> used in this solution to try it out yourself!</span></span>

## <a name="sample-code-generate-key-and-add-row-to-range"></a><span data-ttu-id="8ab23-119">Código de exemplo: Gerar chave e adicionar linha ao intervalo</span><span class="sxs-lookup"><span data-stu-id="8ab23-119">Sample code: Generate key and add row to range</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX  = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input:RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' && 
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('PlainSheet'); /* plain range sheet */
    const range = sheet.getUsedRange();

    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it. 
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length-1];

    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;
    
    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);
    
    // Get last row and compute next row address.
    const last = range.getLastRow();
    const target = last.getOffsetRange(1, 0);

    // Add a row with incoming data plus the computed key value.
    target.setValues([
      [
        nextKey, 
        /* Capitalize the document type. */
        input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
        input.documentName
      ]
    ])
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```

## <a name="sample-code-generate-key-and-add-row-to-table"></a><span data-ttu-id="8ab23-120">Código de exemplo: Gerar chave e adicionar linha à tabela</span><span class="sxs-lookup"><span data-stu-id="8ab23-120">Sample code: Generate key and add row to table</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input: RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' &&
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('TableSheet'); /* table sheet */
    const table = sheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it.
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length - 1];


    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;

    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);

    // Add a row with incoming data plus the computed key value.
    table.addRow(-1, [
            nextKey,
            /* Capitalize the document type. */
            input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
            input.documentName
        ]);
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```
