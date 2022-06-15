---
title: Usar JSON para passar dados de e para Office Scripts
description: Saiba como estruturar dados em objetos JSON para uso com chamadas externas e Power Automate
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088145"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>Usar JSON para passar dados de e para Office Scripts

[JSON (JavaScript Object Notation)](https://www.w3schools.com/whatis/whatis_json.asp) é um formato para armazenar e transferir dados. Cada objeto JSON é uma coleção de pares nome/valor que podem ser definidos quando criados. O JSON é útil com Office scripts porque pode lidar com a complexidade arbitrária de intervalos, tabelas e outros padrões de dados Excel. O JSON permite analisar dados de entrada de serviços [Web](external-calls.md) e passar objetos complexos [por meio Power Automate fluxos](power-automate-integration.md).

Este artigo se concentra no uso de JSON com Office Scripts. É recomendável que você primeiro saiba mais sobre o formato de artigos como [Introdução JSON](https://www.w3schools.com/js/js_json_intro.asp) do W3 Schools.

## <a name="parse-json-data-into-a-range-or-table"></a>Analisar dados JSON em um intervalo ou tabela

Matrizes de objetos JSON fornecem uma maneira consistente de passar linhas de dados de tabela entre aplicativos e serviços Web. Nesses casos, cada objeto JSON representa uma linha, enquanto as propriedades representam as colunas. Um Office script pode executar um loop em uma matriz JSON e reassemmbá-la como uma matriz 2D. Essa matriz é definida como os valores de um intervalo e armazenada em uma pasta de trabalho. Os nomes de propriedade também podem ser adicionados como cabeçalhos para criar uma tabela.

O script a seguir mostra os dados JSON sendo convertidos em uma tabela. Observe que os dados não são obtidos de uma fonte externa. Isso será abordado mais adiante neste artigo.

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> Se você souber a estrutura do JSON, poderá criar sua própria interface para facilitar a criação de propriedades específicas. Você pode substituir as etapas de conversão JSON para matriz por referências de tipo seguro. O snippet de código a seguir mostra essas etapas (agora comentadas) substituídas por chamadas que usam uma nova `ActionRow` interface. Observe que isso torna a `convertJsonToRow` função não mais necessária.
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>Obter dados JSON de fontes externas

Há duas maneiras de importar dados JSON para sua pasta de trabalho por meio de Office Script.

- Como um [parâmetro](power-automate-integration.md#main-parameters-pass-data-to-a-script) com um Power Automate fluxo.
- Com uma `fetch` chamada para um [serviço Web externo](external-calls.md).

#### <a name="modify-the-sample-to-work-with-power-automate"></a>Modifique o exemplo para trabalhar com Power Automate

Os dados JSON Power Automate podem ser passados como uma matriz de objetos genéricos. Adicione uma `object[]` propriedade ao script para aceitar esses dados.

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

Em seguida, você verá uma opção no conector Power Automate para adicionar `jsonData` à ação **Executar script**.

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="Um Excel online (Business) mostrando uma ação Executar script com o parâmetro jsonData.":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>Modificar o exemplo para usar uma `fetch` chamada

Os serviços Web podem responder a `fetch` chamadas com dados JSON. Isso fornece ao script os dados necessários enquanto mantém você Excel. Saiba mais sobre e `fetch` chamadas externas lendo [o suporte a chamadas à API externa Office Scripts](external-calls.md).

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>Criar JSON de um intervalo

As linhas e colunas de uma planilha geralmente implicam relações entre seus valores de dados. Uma linha de uma tabela mapeia conceitualmente para um objeto de programação, com cada coluna sendo uma propriedade desse objeto. Considere a tabela de dados a seguir. Cada linha representa uma transação registrada na planilha.

|ID |Data     |Valor |Fornecedor                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |US$ 43,54 |Melhor para você, Empresa orgânica |
|2  |6/3/2022 |US$ 67,23 |Liberty Bakery and Cafe       |
|3  |6/3/2022 |US$ 37,12 |Melhor para você, Empresa orgânica |
|4  |6/6/2022 |US$ 86,95 |Vinícola Coho                 |
|5  |6/7/2022 |US$ 13,64 |Liberty Bakery and Cafe       |

Cada transação (cada linha) tem um conjunto de propriedades associadas a ela: "ID", "Date", "Amount" e "Vendor". Isso pode ser modelado em um script Office como um objeto.

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

As linhas na tabela de exemplo correspondem às propriedades na interface, para que um script possa converter facilmente cada linha em um `Transaction` objeto. Isso é útil ao gerar os dados para Power Automate. O script a seguir itera em cada linha na tabela e a adiciona a um `Transaction[]`.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="A saída do console do script anterior que mostra os valores de propriedade do objeto.":::

### <a name="use-a-generic-object"></a>Usar um objeto genérico

O exemplo anterior pressupõe que os valores de cabeçalho da tabela são consistentes. Se a tabela tiver colunas variáveis, você precisará criar um objeto JSON genérico. O script a seguir mostra um script que registra qualquer tabela como JSON.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>Confira também

- [Chamada de API externa nos scripts do Office](external-calls.md)
- [Exemplo: usar chamadas de busca externas Office Scripts](../resources/samples/external-fetch-calls.md)
- [Executar Office scripts com Power Automate](power-automate-integration.md)