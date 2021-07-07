---
title: Saída Excel dados como JSON
description: Saiba como Excel dados de tabela como JSON a ser usado Power Automate.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 63379d1323f5e2084f4aa39af3f4b6e5e6d7e7bb
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313943"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="3844b-103">Dados Excel de tabela de saída como JSON para uso em Power Automate</span><span class="sxs-lookup"><span data-stu-id="3844b-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="3844b-104">Excel dados de tabela podem ser representados como uma matriz de objetos na forma de JSON.</span><span class="sxs-lookup"><span data-stu-id="3844b-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="3844b-105">Cada objeto representa uma linha na tabela.</span><span class="sxs-lookup"><span data-stu-id="3844b-105">Each object represents a row in the table.</span></span> <span data-ttu-id="3844b-106">Isso ajuda a extrair os dados Excel em um formato consistente que é visível para o usuário.</span><span class="sxs-lookup"><span data-stu-id="3844b-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="3844b-107">Em seguida, os dados podem ser dados a outros sistemas por meio Power Automate fluxos.</span><span class="sxs-lookup"><span data-stu-id="3844b-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="3844b-108">_Dados da tabela de entrada_</span><span class="sxs-lookup"><span data-stu-id="3844b-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="Uma planilha mostrando dados da tabela de entrada.":::

<span data-ttu-id="3844b-110">Uma variação desse exemplo também inclui os hiperlinks em uma das colunas da tabela.</span><span class="sxs-lookup"><span data-stu-id="3844b-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="3844b-111">Isso permite que níveis adicionais de dados de células sejam a superfície no JSON.</span><span class="sxs-lookup"><span data-stu-id="3844b-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="3844b-112">_Dados da tabela de entrada que incluem hiperlinks_</span><span class="sxs-lookup"><span data-stu-id="3844b-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Uma planilha mostrando uma coluna de dados de tabela formatados como hiperlinks.":::

<span data-ttu-id="3844b-114">_Caixa de diálogo para editar hiperlink_</span><span class="sxs-lookup"><span data-stu-id="3844b-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="A caixa de diálogo Editar Hiperlink exibe opções para alterar hiperlinks.":::

## <a name="sample-excel-file"></a><span data-ttu-id="3844b-116">Exemplo Excel arquivo</span><span class="sxs-lookup"><span data-stu-id="3844b-116">Sample Excel file</span></span>

<span data-ttu-id="3844b-117">Baixe o arquivo <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> para uma pasta de trabalho pronta para uso.</span><span class="sxs-lookup"><span data-stu-id="3844b-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="3844b-118">Adicione o seguinte script para experimentar o exemplo você mesmo!</span><span class="sxs-lookup"><span data-stu-id="3844b-118">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="3844b-119">Código de exemplo: Retornar dados da tabela como JSON</span><span class="sxs-lookup"><span data-stu-id="3844b-119">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="3844b-120">Você pode alterar a `interface TableData` estrutura para corresponder às colunas da tabela.</span><span class="sxs-lookup"><span data-stu-id="3844b-120">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="3844b-121">Observe que, para nomes de coluna com espaços, certifique-se de colocar sua chave entre aspas, como com `"Event ID"` no exemplo.</span><span class="sxs-lookup"><span data-stu-id="3844b-121">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

// This function converts a 2D-array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a><span data-ttu-id="3844b-122">Saída de exemplo da planilha "PlainTable"</span><span class="sxs-lookup"><span data-stu-id="3844b-122">Sample output from the "PlainTable" worksheet</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="3844b-123">Código de exemplo: Retornar dados da tabela como JSON com texto de hiperlink</span><span class="sxs-lookup"><span data-stu-id="3844b-123">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="3844b-124">O script sempre extrai hiperlinks da 4ª coluna (índice 0) da tabela.</span><span class="sxs-lookup"><span data-stu-id="3844b-124">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="3844b-125">Você pode alterar essa ordem ou incluir várias colunas como dados de hiperlink modificando o código no comentário `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="3844b-125">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        object[objectKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        object[objectKeys[j]] = values[i][j];
      }
    }

    objectArray.push(object);
  }
  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output-from-the-withhyperlink-worksheet"></a><span data-ttu-id="3844b-126">Exemplo de saída da planilha "WithHyperLink"</span><span class="sxs-lookup"><span data-stu-id="3844b-126">Sample output from the "WithHyperLink" worksheet</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="3844b-127">Usar no Power Automate</span><span class="sxs-lookup"><span data-stu-id="3844b-127">Use in Power Automate</span></span>

<span data-ttu-id="3844b-128">Para saber como usar esse script em Power Automate, consulte [Create a automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="3844b-128">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
