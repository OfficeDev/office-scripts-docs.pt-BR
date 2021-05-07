---
title: Fundamentos de script para scripts do Office no Excel na Web
description: Informações sobre o modelo de objeto e outros fundamentos para saber mais antes de escrever scripts do Office.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: 685f83952fa6aecc660524a95dec57e149522820
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232386"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="2b51e-103">Fundamentos de script para scripts do Office no Excel na Web (visualização)</span><span class="sxs-lookup"><span data-stu-id="2b51e-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="2b51e-104">Este artigo apresentará os aspectos técnicos dos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="2b51e-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="2b51e-105">Você saberá como os objetos do Excel funcionam em conjunto e como o editor de código se sincroniza com uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2b51e-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a><span data-ttu-id="2b51e-106">função `main`</span><span class="sxs-lookup"><span data-stu-id="2b51e-106">`main` function</span></span>

<span data-ttu-id="2b51e-107">Cada Script do Office precisa incluir uma função `main` tendo o tipo de `ExcelScript.Workbook` como seu primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2b51e-107">Each Office Script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="2b51e-108">Quando a função é executada, o aplicativo Excel invoca esta função `main` fornecendo a pasta de trabalho como seu primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2b51e-108">When the function is executed, the Excel application invokes this `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="2b51e-109">Portanto, é importante não modificar a assinatura básica da função `main` depois de gravar o script ou criar um script a partir do editor de código.</span><span class="sxs-lookup"><span data-stu-id="2b51e-109">Hence, it is important to not modify the basic signature of the `main` function once you have either recorded the script or created a new script from the code editor.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="2b51e-110">O código dentro da função `main` é executado quando o script é executado.</span><span class="sxs-lookup"><span data-stu-id="2b51e-110">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="2b51e-111">`main` pode chamar outras funções em seu script, mas o código que não estiver contido em uma função não será executado.</span><span class="sxs-lookup"><span data-stu-id="2b51e-111">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

> [!CAUTION]
> <span data-ttu-id="2b51e-112">Se sua função `main` se parece com `async function main(context: Excel.RequestContext)`, seu script está usando um modelo antigo de API assíncrona.</span><span class="sxs-lookup"><span data-stu-id="2b51e-112">If your `main` function looks like `async function main(context: Excel.RequestContext)`, your script is using the older async API model.</span></span> <span data-ttu-id="2b51e-113">Para obter mais informações (inclusive sobre como converter seu script para o modelo de API atual), confira o artigo [Suporte a Scripts do Office mais antigos que utilizam APIs Assíncronas](excel-async-model.md).</span><span class="sxs-lookup"><span data-stu-id="2b51e-113">For more information (including how to convert your script to the current API model), refer to [Support older Office Scripts that use the Async APIs](excel-async-model.md).</span></span>

## <a name="object-model"></a><span data-ttu-id="2b51e-114">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="2b51e-114">Object model</span></span>

<span data-ttu-id="2b51e-115">Para escrever um script, você precisa entender como as APIs dos Scripts do Office se encaixam.</span><span class="sxs-lookup"><span data-stu-id="2b51e-115">To write a script, you need to understand how the Office Script APIs fit together.</span></span> <span data-ttu-id="2b51e-116">Os componentes de uma pasta de trabalho têm relações específicas entre si.</span><span class="sxs-lookup"><span data-stu-id="2b51e-116">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="2b51e-117">De várias maneiras, essas relações correspondem às da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="2b51e-117">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="2b51e-118">Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.</span><span class="sxs-lookup"><span data-stu-id="2b51e-118">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="2b51e-119">Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.</span><span class="sxs-lookup"><span data-stu-id="2b51e-119">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="2b51e-120">Um **Intervalo** representa um grupo de células contíguas.</span><span class="sxs-lookup"><span data-stu-id="2b51e-120">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="2b51e-121">Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.</span><span class="sxs-lookup"><span data-stu-id="2b51e-121">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="2b51e-122">Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.</span><span class="sxs-lookup"><span data-stu-id="2b51e-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="2b51e-123">As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.</span><span class="sxs-lookup"><span data-stu-id="2b51e-123">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="workbook"></a><span data-ttu-id="2b51e-124">Pasta de Trabalho</span><span class="sxs-lookup"><span data-stu-id="2b51e-124">Workbook</span></span>

<span data-ttu-id="2b51e-125">Todo script é fornecido com um `workbook` objeto do tipo `Workbook` pela função `main`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-125">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="2b51e-126">Isso representa o objeto de nível superior por meio do qual seu script interage com a pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="2b51e-126">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="2b51e-127">O script a seguir obtém a planilha ativa da pasta de trabalho e registra seu nome.</span><span class="sxs-lookup"><span data-stu-id="2b51e-127">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a><span data-ttu-id="2b51e-128">Intervalos</span><span class="sxs-lookup"><span data-stu-id="2b51e-128">Ranges</span></span>

<span data-ttu-id="2b51e-129">Um intervalo é um grupo de células contíguas na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2b51e-129">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="2b51e-130">Os scripts costumam usar uma notação estilo A1 (por ex.: **B3** para a única célula na coluna **B** e linha **3** ou **C2:F4** para as células das colunas **C** a **F** e linhas **2** a **4**) para definir intervalos.</span><span class="sxs-lookup"><span data-stu-id="2b51e-130">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="2b51e-131">Os intervalos têm três propriedades principais: valores, fórmulas e formato.</span><span class="sxs-lookup"><span data-stu-id="2b51e-131">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="2b51e-132">Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células.</span><span class="sxs-lookup"><span data-stu-id="2b51e-132">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="2b51e-133">Eles são acessados através de `getValues`, `getFormulas` e `getFormat`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-133">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="2b51e-134">Os valores e fórmulas podem ser alterados com `setValues` e `setFormulas`, enquanto o formato é um objeto `RangeFormat` composto de vários objetos menores que são configurados individualmente.</span><span class="sxs-lookup"><span data-stu-id="2b51e-134">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="2b51e-135">Os intervalo usam matrizes bidimensionais para gerenciar informações.</span><span class="sxs-lookup"><span data-stu-id="2b51e-135">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="2b51e-136">Leia a [Trabalhando com intervalos da seção Usando objetos JavaScript incorporados nos Scripts do Office](javascript-objects.md#working-with-ranges) para obter mais informações sobre como lidar com essas matrizes na estrutura de Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="2b51e-136">Read the [Working with ranges section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-ranges) for more information on handling those arrays in the Office Scripts framework.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="2b51e-137">Exemplo de intervalo</span><span class="sxs-lookup"><span data-stu-id="2b51e-137">Range sample</span></span>

<span data-ttu-id="2b51e-138">O exemplo a seguir mostra como criar registros de vendas.</span><span class="sxs-lookup"><span data-stu-id="2b51e-138">The following sample shows how to create sales records.</span></span> <span data-ttu-id="2b51e-139">Este script usa `Range` objetos para definir os valores, fórmulas e partes do formato.</span><span class="sxs-lookup"><span data-stu-id="2b51e-139">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="2b51e-140">Executar este script cria os seguintes dados na planilha atual:</span><span class="sxs-lookup"><span data-stu-id="2b51e-140">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="Uma planilha contendo um registro de vendas composto de linhas de valor, uma coluna de fórmula e cabeçalhos formatados":::

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="2b51e-142">Gráficos, tabelas e outros objetos de dados</span><span class="sxs-lookup"><span data-stu-id="2b51e-142">Charts, tables, and other data objects</span></span>

<span data-ttu-id="2b51e-143">Os scripts podem criar e manipular estruturas de dados e visualizações no Excel.</span><span class="sxs-lookup"><span data-stu-id="2b51e-143">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="2b51e-144">As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais.</span><span class="sxs-lookup"><span data-stu-id="2b51e-144">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="2b51e-145">Eles são armazenados em coleções, que serão discutidas mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="2b51e-145">These are stored in collections, which will be discussed later in this article.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="2b51e-146">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="2b51e-146">Creating a table</span></span>

<span data-ttu-id="2b51e-p112">Crie tabelas usando intervalos preenchidos com dados. A formatação e os controles de tabela (como filtros) são automaticamente aplicados ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="2b51e-p112">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="2b51e-149">O script a seguir cria uma tabela usando os intervalos do exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="2b51e-149">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="2b51e-150">Executar esse script na planilha com os dados anteriores cria a tabela a seguir:</span><span class="sxs-lookup"><span data-stu-id="2b51e-150">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="Uma planilha contendo uma tabela feita com base no registro de vendas anterior":::

#### <a name="creating-a-chart"></a><span data-ttu-id="2b51e-152">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="2b51e-152">Creating a chart</span></span>

<span data-ttu-id="2b51e-153">Crie gráficos para visualizar os dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="2b51e-153">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="2b51e-154">Os scripts permitem inúmeras variedades de gráficos que podem ser personalizadas de acordo com suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="2b51e-154">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="2b51e-155">O script a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="2b51e-155">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="2b51e-156">Executar este script na planilha com a tabela anterior cria o seguinte gráfico:</span><span class="sxs-lookup"><span data-stu-id="2b51e-156">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="Um gráfico de colunas exibindo a quantidade de três itens do registro de vendas anterior":::

### <a name="collections-and-other-object-relations"></a><span data-ttu-id="2b51e-158">Coleções e outras relações de objeto</span><span class="sxs-lookup"><span data-stu-id="2b51e-158">Collections and other object relations</span></span>

<span data-ttu-id="2b51e-159">Qualquer objeto filho pode ser acessado através do objeto pai.</span><span class="sxs-lookup"><span data-stu-id="2b51e-159">Any child object can be accessed through its parent object.</span></span> <span data-ttu-id="2b51e-160">Por exemplo, você pode ler `Worksheets` do objeto `Workbook`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-160">For example, you can read `Worksheets` from the `Workbook` object.</span></span> <span data-ttu-id="2b51e-161">Haverá um método `get` relacionado na classe mãe (por exemplo, `Workbook.getWorksheets()` ou `Workbook.getWorksheet(name)`).</span><span class="sxs-lookup"><span data-stu-id="2b51e-161">There will be a related `get` method on the parent class that (e.g., `Workbook.getWorksheets()` or `Workbook.getWorksheet(name)`).</span></span> <span data-ttu-id="2b51e-162">Os métodos `get` singulares retornam um único objeto e requerem um ID ou nome para o objeto específico (como o nome de uma planilha).</span><span class="sxs-lookup"><span data-stu-id="2b51e-162">`get` methods that are singular return a single object and require an ID or name for the specific object (such as the name of a worksheet).</span></span> <span data-ttu-id="2b51e-163">Os métodos `get` que são plurais retornam toda a coleção de objetos como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="2b51e-163">`get` methods that are plural return the entire object collection as an array.</span></span> <span data-ttu-id="2b51e-164">Se a coleção estiver vazia, você obterá uma matriz vazia (`[]`).</span><span class="sxs-lookup"><span data-stu-id="2b51e-164">If the collection is empty, you'll get an empty array (`[]`).</span></span>

<span data-ttu-id="2b51e-165">Depois que a coleção é recuperada, você pode usar operações regulares de matriz, como obter seus `length` ou usar `for`, `for..of`, `while` loops para iteração ou métodos de matriz TypeScript como `map`, `forEach`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-165">Once the collection is retrieved, you can use regular array operations such as getting its `length` or use `for`, `for..of`, `while` loops for iteration or use TypeScript array methods such as `map`, `forEach` on them.</span></span> <span data-ttu-id="2b51e-166">Você também pode acessar objetos individuais na coleção usando o valor do índice da matriz.</span><span class="sxs-lookup"><span data-stu-id="2b51e-166">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="2b51e-167">Por exemplo, `workbook.getTables()[0]` retorna a primeira tabela da coleção.</span><span class="sxs-lookup"><span data-stu-id="2b51e-167">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="2b51e-168">Leia a seção [Trabalhando com coleções de Usando objetos JavaScript nos Scripts do Office](javascript-objects.md#working-with-collections) para aprender mais sobre o uso da funcionalidade de matriz incorporada com a estrutura de Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="2b51e-168">Read the [Working with collections section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-collections) to learn more about using built-in array functionality with the Office Scripts framework.</span></span>

<span data-ttu-id="2b51e-p116">O seguinte roteiro recebe todas as tabelas na pasta de trabalho. Assim, ele garante que os cabeçalhos sejam exibidos, os botões de filtro sejam visíveis e o estilo da tabela seja definido como "TableStyleLight1".</span><span class="sxs-lookup"><span data-stu-id="2b51e-p116">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

```TypeScript
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

#### <a name="adding-excel-objects-with-a-script"></a><span data-ttu-id="2b51e-171">Adicionando objetos do Excel com um script</span><span class="sxs-lookup"><span data-stu-id="2b51e-171">Adding Excel objects with a script</span></span>

<span data-ttu-id="2b51e-172">Você pode adicionar programaticamente objetos de documento, como tabelas ou gráficos, chamando o método `add` correspondente disponível no objeto pai.</span><span class="sxs-lookup"><span data-stu-id="2b51e-172">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!NOTE]
> <span data-ttu-id="2b51e-173">Não adicione manualmente objetos as matrizes de coleção.</span><span class="sxs-lookup"><span data-stu-id="2b51e-173">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="2b51e-174">Use os métodos `add` nos objetos pai, por exemplo, adicione `Table` a `Worksheet` com o método `Worksheet.addTable`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-174">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="2b51e-175">O script a seguir cria, no Excel, uma tabela na primeira planilha da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2b51e-175">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="2b51e-176">Observe que a tabela criada é enviada de volta pelo método `addTable`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-176">Note that the created table is returned by the `addTable` method.</span></span>

```TypeScript
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

## <a name="removing-excel-objects-with-a-script"></a><span data-ttu-id="2b51e-177">Removendo objetos do Excel com um script</span><span class="sxs-lookup"><span data-stu-id="2b51e-177">Removing Excel objects with a script</span></span>

<span data-ttu-id="2b51e-178">Para excluir um objeto, chame o método `delete` do objeto.</span><span class="sxs-lookup"><span data-stu-id="2b51e-178">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="2b51e-179">Como na adição de objetos, não remova manualmente objetos de matrizes de coleção.</span><span class="sxs-lookup"><span data-stu-id="2b51e-179">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="2b51e-180">Use os métodos `delete` nos objetos do tipo coleção.</span><span class="sxs-lookup"><span data-stu-id="2b51e-180">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="2b51e-181">Por exemplo, remova um `Table` de um `Worksheet` usando `Table.delete`.</span><span class="sxs-lookup"><span data-stu-id="2b51e-181">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="2b51e-182">O script a seguir remove a primeira planilha da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2b51e-182">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="2b51e-183">Leituras adicionais sobre o modelo de objeto</span><span class="sxs-lookup"><span data-stu-id="2b51e-183">Further reading on the object model</span></span>

<span data-ttu-id="2b51e-184">A [documentação de referência de API dos scripts do Office](/javascript/api/office-scripts/overview) é uma lista completa dos objetos usados nos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="2b51e-184">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="2b51e-185">Lá, você pode usar o sumário para navegar para qualquer classe da qual quiser saber mais.</span><span class="sxs-lookup"><span data-stu-id="2b51e-185">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="2b51e-186">Estas são várias páginas exibidas com frequência.</span><span class="sxs-lookup"><span data-stu-id="2b51e-186">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="2b51e-187">Gráfico</span><span class="sxs-lookup"><span data-stu-id="2b51e-187">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="2b51e-188">Comentário</span><span class="sxs-lookup"><span data-stu-id="2b51e-188">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="2b51e-189">PivotTable</span><span class="sxs-lookup"><span data-stu-id="2b51e-189">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="2b51e-190">Range</span><span class="sxs-lookup"><span data-stu-id="2b51e-190">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="2b51e-191">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="2b51e-191">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="2b51e-192">Formato</span><span class="sxs-lookup"><span data-stu-id="2b51e-192">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="2b51e-193">Table</span><span class="sxs-lookup"><span data-stu-id="2b51e-193">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="2b51e-194">Pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="2b51e-194">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="2b51e-195">Planilha</span><span class="sxs-lookup"><span data-stu-id="2b51e-195">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="2b51e-196">Confira também</span><span class="sxs-lookup"><span data-stu-id="2b51e-196">See also</span></span>

- [<span data-ttu-id="2b51e-197">Gravar, editar e criar scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="2b51e-197">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="2b51e-198">Ler os dados da pasta de trabalho com scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="2b51e-198">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="2b51e-199">Referência da API de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="2b51e-199">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="2b51e-200">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="2b51e-200">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
