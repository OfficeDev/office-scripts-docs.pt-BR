---
title: Fundamentos de script para scripts do Office no Excel na Web
description: Informações sobre o modelo de objeto e outros fundamentos para saber mais antes de escrever scripts do Office.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: b5038dde38550e63bae872b39b9222d3defe9943ccefad85a469a5c0717fb2ef
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846674"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Fundamentos de script para Scripts do Office no Excel na Web

Este artigo apresentará os aspectos técnicos dos scripts do Office. Você saberá como os objetos do Excel funcionam em conjunto e como o editor de código se sincroniza com uma pasta de trabalho.

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript: A linguagem dos Scripts do Office

Os Scripts do Office são escritos em [TypeScript](https://www.typescriptlang.org/docs/home.html), que é um superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Se você está familiarizado com o JavaScript, seus conhecimentos serão aproveitados porque muito do código é o mesmo em ambas as linguagens. Recomendamos que você tenha algum conhecimento de programação de nível iniciante antes de iniciar sua jornada de codificação nos Scripts do Office. Os recursos a seguir podem ajudá-lo a entender o lado da codificação dos Scripts do Office.

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>Função `main`: O ponto de partida do script

Cada script deve conter uma função `main` com o tipo `ExcelScript.Workbook` como seu primeiro parâmetro. Quando a função é executada, o aplicativo Excel invoca a função `main` fornecendo a pasta de trabalho como seu primeiro parâmetro. Um `ExcelScript.Workbook` deve sempre ser o primeiro parâmetro.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

O código dentro da função `main` é executado quando o script é executado. `main` pode chamar outras funções em seu script, mas o código que não estiver contido em uma função não será executado. Os scripts não podem invocar ou chamar outros Scripts do Office.

O [Power Automate](https://flow.microsoft.com) permite que você conecte scripts em fluxos. Os dados são passados entre os scripts e o fluxo por meio dos parâmetros e retornos do método `main`. Como integrar os Scripts do Office com o Power Automate é abordado em detalhes em [Executar Scripts do Office com o Power Automate ](power-automate-integration.md).

## <a name="object-model-overview"></a>Visão geral do modelo de objeto

Para escrever um script, você precisa entender como as APIs do Scripts do Office se encaixam. Os componentes de uma pasta de trabalho têm relações específicas entre si. De várias maneiras, essas relações correspondem às da interface do usuário do Excel.

- Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.
- Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.
- Um **Intervalo** representa um grupo de células contíguas.
- Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.
- Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.
- As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.

## <a name="workbook"></a>Pasta de Trabalho

Todo script é fornecido com um `workbook` objeto do tipo `Workbook` pela função `main`. Isso representa o objeto de nível superior por meio do qual seu script interage com a pasta de trabalho do Excel.

O script a seguir obtém a planilha ativa da pasta de trabalho e registra seu nome.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a>Intervalos

Um intervalo é um grupo de células contíguas na pasta de trabalho. Os scripts costumam usar uma notação estilo A1 (por ex.: **B3** para a única célula na coluna **B** e linha **3** ou **C2:F4** para as células das colunas **C** a **F** e linhas **2** a **4**) para definir intervalos.

Os intervalos têm três propriedades principais: valores, fórmulas e formato. Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células. Eles são acessados através de `getValues`, `getFormulas` e `getFormat`. Os valores e fórmulas podem ser alterados com `setValues` e `setFormulas`, enquanto o formato é um objeto `RangeFormat` composto de vários objetos menores que são configurados individualmente.

Os intervalo usam matrizes bidimensionais para gerenciar informações. Para obter mais informações sobre como lidar com matrizes na estrutura de Scripts do Office, consulte [Trabalhar com intervalos](javascript-objects.md#work-with-ranges).

### <a name="range-sample"></a>Exemplo de intervalo

O exemplo a seguir mostra como criar registros de vendas. Este script usa `Range` objetos para definir os valores, fórmulas e partes do formato.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.54],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

Executar este script cria os seguintes dados na planilha atual:

:::image type="content" source="../images/range-sample.png" alt-text="Uma planilha contendo um registro de vendas composto por linhas de valor, uma coluna de fórmulas e os cabeçalhos formatados.":::

### <a name="the-types-of-range-values"></a>Os tipos de valores do Intervalo

Cada célula tem um valor. Este valor é o valor subjacente inserido na célula, que pode ser diferente do texto exibido no Excel. Por exemplo, você pode ver "2/5/2021" exibido na célula como uma data, mas o valor real é 44318. Esta exibição pode ser alterada com o formato de número, mas o valor real e o tipo na célula mudam apenas quando um novo valor é definido.

Quando você estiver usando o valor da célula, é importante informar ao TypeScript qual valor você espera obter de uma célula ou intervalo. Uma célula contém um dos seguintes tipos: `string`, `number`, ou `boolean`. Para que seu script trate os valores retornados como um desses tipos, você deve declarar o tipo.

O script a seguir obtém o preço médio da tabela do exemplo anterior. Observe o código `priceRange.getValues() as number[][]`. Isso [declara](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) o tipo dos valores do intervalo como um `number[][]`. Todos os valores nessa matriz podem então ser tratados como números no script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>Gráficos, tabelas e outros objetos de dados

Os scripts podem criar e manipular estruturas de dados e visualizações no Excel. As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais. Eles são armazenados em coleções, que serão discutidas mais adiante neste artigo.

### <a name="create-a-table"></a>Criar uma tabela

Crie tabelas usando intervalos preenchidos com dados. A formatação e os controles de tabela (como filtros) são automaticamente aplicados ao intervalo.

O script a seguir cria uma tabela usando os intervalos do exemplo anterior.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

Executar esse script na planilha com os dados anteriores cria a tabela a seguir:

:::image type="content" source="../images/table-sample.png" alt-text="Uma planilha contendo uma tabela feita do registro de vendas anterior.":::

### <a name="create-a-chart"></a>Criar um gráfico

Crie gráficos para visualizar os dados em um intervalo. Os scripts permitem inúmeras variedades de gráficos que podem ser personalizadas de acordo com suas necessidades.

O script a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

Executar este script na planilha com a tabela anterior cria o seguinte gráfico:

:::image type="content" source="../images/chart-sample.png" alt-text="Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.":::

## <a name="collections"></a>Coleções

Quando um objeto do Excel tem uma coleção de um ou mais objetos do mesmo tipo, ele os armazena em uma matriz. Por exemplo, um objeto `Workbook` contém um `Worksheet[]`. Esta matriz é acessada pelo método `Workbook.getWorksheets()`. Os métodos `get` que são plurais, como `Worksheet.getCharts()`, retornam toda a coleção de objetos como uma matriz. Você verá este padrão em todas as APIs de Scripts do Office: o objeto `Worksheet` tem um método `getTables()` que retorna um `Table[]`, o objeto `Table` tem um método `getColumns()` que retorna um `TableColumn[]`, como assim em diante.

A matriz retornada é uma matriz normal, portanto todas as operações regulares de matriz estão disponíveis para seu script. Você também pode acessar objetos individuais na coleção usando o valor do índice da matriz. Por exemplo, `workbook.getTables()[0]` retorna a primeira tabela da coleção. Para saber mais sobre o uso da funcionalidade de matriz interna com a estrutura de Scripts do Office, consulte [Trabalhar com coleções](javascript-objects.md#work-with-collections). 

Objetos individuais também são acessados a partir da coleção por meio de um método `get`. Os métodos `get` que são singulares, como `Worksheet.getTable(name)`, retornam um único objeto e requerem uma ID ou nome para o objeto específico. Esse ID ou nome geralmente é definido pelo script ou por meio da IU do Excel.

O seguinte roteiro recebe todas as tabelas na pasta de trabalho. Assim, ele garante que os cabeçalhos sejam exibidos, os botões de filtro sejam visíveis e o estilo da tabela seja definido como "TableStyleLight1".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a>Adicionar objetos do Excel com um script

Você pode adicionar programaticamente objetos de documento, como tabelas ou gráficos, chamando o método `add` correspondente disponível no objeto pai.

> [!IMPORTANT]
> Não adicione manualmente objetos as matrizes de coleção. Use os métodos `add` nos objetos pai, por exemplo, adicione `Table` a `Worksheet` com o método `Worksheet.addTable`.

O script a seguir cria, no Excel, uma tabela na primeira planilha da pasta de trabalho. Observe que a tabela criada é enviada de volta pelo método `addTable`.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> A maioria dos objetos do Excel possui um método `setName`. Isso fornece uma maneira fácil de acessar objetos do Excel posteriormente no script ou em outros scripts para a mesma pasta de trabalho.

### <a name="verify-an-object-exists-in-the-collection"></a>Verifique se existe um objeto na coleção

Os scripts geralmente precisam verificar se uma tabela ou objeto semelhante existe antes de continuar. Use os nomes dados por scripts ou por meio da IU do Excel para identificar os objetos necessários e agir de acordo. O métodos `get` retornam `undefined` quando o objeto solicitado não está na coleção.

O script a seguir solicita uma tabela chamada "MinhaTabela" e utiliza uma instrução `if...else` para verificar se a tabela foi encontrada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

Um padrão comum em Scripts do Office é recriar uma tabela, gráfico ou outro objeto sempre que o script for executado. Se você não precisa de dados antigos, é melhor excluir o objeto antigo antes de criar o novo. Isso evita conflitos de nome ou outras diferenças que possam ter sido introduzidas por outros usuários.

O script a seguir remove a tabela chamada "MinhaTabela", se houver, e adiciona uma nova tabela com o mesmo nome.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a>Remova objetos do Excel com um script

Para excluir um objeto, chame o método `delete` do objeto.

> [!NOTE]
> Como na adição de objetos, não remova manualmente objetos de matrizes de coleção. Use os métodos `delete` nos objetos do tipo coleção. Por exemplo, remova um `Table` de um `Worksheet` usando `Table.delete`.

O script a seguir remove a primeira planilha da pasta de trabalho.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a>Leituras adicionais sobre o modelo de objeto

A [documentação de referência de API dos scripts do Office](/javascript/api/office-scripts/overview) é uma lista completa dos objetos usados nos scripts do Office. Lá, você pode usar o sumário para navegar para qualquer classe da qual quiser saber mais. Estas são várias páginas exibidas com frequência.

- [Gráfico](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [Comentário](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range)
- [RangeFormat](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [Formato](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [Table](/javascript/api/office-scripts/excelscript/excelscript.table)
- [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a>Confira também

- [Gravar, editar e criar scripts do Office no Excel na Web](../tutorials/excel-tutorial.md)
- [Ler os dados da pasta de trabalho com scripts do Office no Excel na Web](../tutorials/excel-read-tutorial.md)
- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)
- [Práticas recomendadas nos Scripts do Office ](best-practices.md)
- [Centro de Desenvolvimento de Scripts do Office](https://developer.microsoft.com/office-scripts)
