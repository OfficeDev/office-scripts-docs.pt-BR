---
title: Fundamentos de script para scripts do Office no Excel na Web
description: Informações sobre o modelo de objeto e outros fundamentos para saber mais antes de escrever scripts do Office.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: acbeec69a5d9ae9e3ebfa95c9070033d1cca2265
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933270"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Fundamentos de script para scripts do Office no Excel na Web (visualização)

Este artigo apresentará os aspectos técnicos dos scripts do Office. Você saberá como os objetos do Excel funcionam em conjunto e como o editor de código se sincroniza com uma pasta de trabalho.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>função `main`

Cada Script do Office precisa incluir uma função `main` tendo o tipo de `ExcelScript.Workbook` como seu primeiro parâmetro. Quando a função é executada, o aplicativo Excel invoca esta função `main` fornecendo a pasta de trabalho como seu primeiro parâmetro. Portanto, é importante não modificar a assinatura básica da função `main` depois de gravar o script ou criar um script a partir do editor de código.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

O código dentro da função `main` é executado quando o script é executado. `main` pode chamar outras funções em seu script, mas o código que não estiver contido em uma função não será executado.

> [!CAUTION]
> Se sua função `main` se parece com `async function main(context: Excel.RequestContext)`, seu script está usando um modelo antigo de API assíncrona. Para obter mais informações (inclusive sobre como converter seu script para o modelo de API atual), confira o artigo [Suporte a Scripts do Office mais antigos que utilizam APIs Assíncronas](excel-async-model.md).

## <a name="object-model"></a>Modelo de objetos

Para escrever um script, você precisa entender como as APIs dos Scripts do Office se encaixam. Os componentes de uma pasta de trabalho têm relações específicas entre si. De várias maneiras, essas relações correspondem às da interface do usuário do Excel.

- Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.
- Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.
- Um **Intervalo** representa um grupo de células contíguas.
- Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.
- Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.
- As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.

### <a name="workbook"></a>Pasta de Trabalho

Todo script é fornecido com um `workbook` objeto do tipo `Workbook` pela função `main`. Isso representa o objeto de nível superior por meio do qual seu script interage com a pasta de trabalho do Excel.

O script a seguir obtém a planilha ativa da pasta de trabalho e registra seu nome.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a>Intervalos

Um intervalo é um grupo de células contíguas na pasta de trabalho. Os scripts costumam usar uma notação estilo A1 (por ex.: **B3** para a única célula na coluna **B** e linha **3** ou **C2:F4** para as células das colunas **C** a **F** e linhas **2** a **4**) para definir intervalos.

Os intervalos têm três propriedades principais: valores, fórmulas e formato. Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células. Eles são acessados através de `getValues`, `getFormulas` e `getFormat`. Os valores e fórmulas podem ser alterados com `setValues` e `setFormulas`, enquanto o formato é um objeto `RangeFormat` composto de vários objetos menores que são configurados individualmente.

Os intervalo usam matrizes bidimensionais para gerenciar informações. Leia a [Trabalhando com intervalos da seção Usando objetos JavaScript incorporados nos Scripts do Office](javascript-objects.md#working-with-ranges) para obter mais informações sobre como lidar com essas matrizes na estrutura de Scripts do Office.

#### <a name="range-sample"></a>Exemplo de intervalo

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
        ["Chocolate", 10, 9.56],
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

![Um registro de vendas mostrando as linhas de valores, uma coluna de fórmulas e cabeçalhos formatados.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Gráficos, tabelas e outros objetos de dados

Os scripts podem criar e manipular estruturas de dados e visualizações no Excel. As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais. Eles são armazenados em coleções, que serão discutidas mais adiante neste artigo.

#### <a name="creating-a-table"></a>Criar uma tabela

Criar tabelas usando intervalos de dados preenchidos. Controles de formatação e tabela (por exemplo, filtros) são aplicados automaticamente ao intervalo.

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

![Uma tabela criada a partir do registro de vendas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Criar um gráfico

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

![Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a>Coleções e outras relações de objeto

Qualquer objeto filho pode ser acessado através do objeto pai. Por exemplo, você pode ler `Worksheets` do objeto `Workbook`. Haverá um método `get` relacionado na classe mãe (por exemplo, `Workbook.getWorksheets()` ou `Workbook.getWorksheet(name)`). Os métodos `get` singulares retornam um único objeto e requerem um ID ou nome para o objeto específico (como o nome de uma planilha). Os métodos `get` que são plurais retornam toda a coleção de objetos como uma matriz. Se a coleção estiver vazia, você obterá uma matriz vazia (`[]`).

Depois que a coleção é recuperada, você pode usar operações regulares de matriz, como obter seus `length` ou usar `for`, `for..of`, `while` loops para iteração ou métodos de matriz TypeScript como `map`, `forEach`. Você também pode acessar objetos individuais na coleção usando o valor do índice da matriz. Por exemplo, `workbook.getTables()[0]` retorna a primeira tabela da coleção. Leia a seção [Trabalhando com coleções de Usando objetos JavaScript nos Scripts do Office](javascript-objects.md#working-with-collections) para aprender mais sobre o uso da funcionalidade de matriz incorporada com a estrutura de Scripts do Office.

O script a seguir obtém todas as tabelas na pasta de trabalho. Em seguida, garante que os cabeçalhos sejam exibidos, os botões de filtro estejam visíveis e o estilo da tabela seja definido como "TableStyleLight1".

```typescript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a>Adicionando objetos do Excel com um script

Você pode adicionar programaticamente objetos de documento, como tabelas ou gráficos, chamando o método `add` correspondente disponível no objeto pai.

> [!NOTE]
> Não adicione manualmente objetos as matrizes de coleção. Use os métodos `add` nos objetos pai, por exemplo, adicione `Table` a `Worksheet` com o método `Worksheet.addTable`.

O script a seguir cria, no Excel, uma tabela na primeira planilha da pasta de trabalho. Observe que a tabela criada é enviada de volta pelo método `addTable`.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a>Removendo objetos do Excel com um script

Para excluir um objeto, chame o método `delete` do objeto.

> [!NOTE]
> Como na adição de objetos, não remova manualmente objetos de matrizes de coleção. Use os métodos `delete` nos objetos do tipo coleção. Por exemplo, remova um `Table` de um `Worksheet` usando `Table.delete`.

O script a seguir remove a primeira planilha da pasta de trabalho.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a>Leituras adicionais sobre o modelo de objeto

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
