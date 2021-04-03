---
title: Filtrar tabela do Excel e obter intervalo visível
description: Saiba como usar scripts do Office para filtrar uma tabela do Excel e obter o intervalo visível como uma matriz de objetos.
ms.date: 03/16/2021
localization_priority: Normal
ms.openlocfilehash: c0a5842af4a62162225e3fc10203c261b91e010a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571097"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="3d789-103">Filtrar a tabela do Excel e obter o intervalo visível como um objeto JSON</span><span class="sxs-lookup"><span data-stu-id="3d789-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="3d789-104">Este exemplo filtra uma tabela do Excel e retorna o intervalo visível como um objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="3d789-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="3d789-105">Esse JSON pode ser fornecido a um fluxo do Power Automate como parte de uma solução maior.</span><span class="sxs-lookup"><span data-stu-id="3d789-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="3d789-106">Cenário de exemplo</span><span class="sxs-lookup"><span data-stu-id="3d789-106">Example scenario</span></span>

* <span data-ttu-id="3d789-107">Aplique um filtro a uma coluna de tabela.</span><span class="sxs-lookup"><span data-stu-id="3d789-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="3d789-108">Extraia o intervalo visível após a filtragem.</span><span class="sxs-lookup"><span data-stu-id="3d789-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="3d789-109">Montar e retornar um objeto com uma [estrutura JSON específica.](#sample-json)</span><span class="sxs-lookup"><span data-stu-id="3d789-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="3d789-110">Código de exemplo: filtrar uma tabela e obter intervalo visível</span><span class="sxs-lookup"><span data-stu-id="3d789-110">Sample code: Filter a table and get visible range</span></span>

<span data-ttu-id="3d789-111">O script a seguir filtra uma tabela e obtém o intervalo visível.</span><span class="sxs-lookup"><span data-stu-id="3d789-111">The following script filters a table and gets the visible range.</span></span>

<span data-ttu-id="3d789-112">Baixe o arquivo de <a href="table-filter.xlsx"> exemplotable-filter.xlsx</a> e use-o com este script para experimentar você mesmo!</span><span class="sxs-lookup"><span data-stu-id="3d789-112">Download the sample file <a href="table-filter.xlsx">table-filter.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
}
```

### <a name="sample-json"></a><span data-ttu-id="3d789-113">Exemplo JSON</span><span class="sxs-lookup"><span data-stu-id="3d789-113">Sample JSON</span></span>

<span data-ttu-id="3d789-114">Cada chave representa um valor exclusivo de uma tabela.</span><span class="sxs-lookup"><span data-stu-id="3d789-114">Each key represents a unique value of a table.</span></span> <span data-ttu-id="3d789-115">Cada instância de matriz representa a linha que fica visível quando o filtro correspondente é aplicado.</span><span class="sxs-lookup"><span data-stu-id="3d789-115">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="3d789-116">Vídeo de treinamento: filtrar uma tabela do Excel e obter o intervalo visível</span><span class="sxs-lookup"><span data-stu-id="3d789-116">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="3d789-117">[![Assista ao vídeo passo a passo sobre como filtrar uma tabela do Excel e obter o intervalo visível](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Vídeo passo a passo sobre como filtrar uma tabela do Excel e obter o intervalo visível")</span><span class="sxs-lookup"><span data-stu-id="3d789-117">[![Watch step-by-step video on how to filter an Excel table and get the visible range](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Step-by-step video on how to filter an Excel table and get the visible range")</span></span>
