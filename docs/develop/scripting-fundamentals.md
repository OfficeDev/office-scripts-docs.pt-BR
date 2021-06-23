---
title: Fundamentos de script para scripts do Office no Excel na Web
description: Informações sobre o modelo de objeto e outros fundamentos para saber mais antes de escrever scripts do Office.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 9c3c10e283e40f1e719e73106bcdacfcff44dbc9
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074505"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="74256-103">Fundamentos de script para Scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="74256-103">Scripting fundamentals for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="74256-104">Este artigo apresentará os aspectos técnicos dos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="74256-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="74256-105">Você saberá como os objetos do Excel funcionam em conjunto e como o editor de código se sincroniza com uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="74256-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="74256-106">TypeScript: A linguagem dos Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="74256-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="74256-107">Os Scripts do Office são escritos em [TypeScript](https://www.typescriptlang.org/docs/home.html), que é um superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span><span class="sxs-lookup"><span data-stu-id="74256-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="74256-108">Se você está familiarizado com o JavaScript, seus conhecimentos serão aproveitados porque muito do código é o mesmo em ambas as linguagens.</span><span class="sxs-lookup"><span data-stu-id="74256-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="74256-109">Recomendamos que você tenha algum conhecimento de programação de nível iniciante antes de iniciar sua jornada de codificação nos Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="74256-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="74256-110">Os recursos a seguir podem ajudá-lo a entender o lado da codificação dos Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="74256-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="74256-111">Função `main`: O ponto de partida do script</span><span class="sxs-lookup"><span data-stu-id="74256-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="74256-112">Cada script deve conter uma função `main` com o tipo `ExcelScript.Workbook` como seu primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="74256-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="74256-113">Quando a função é executada, o aplicativo Excel invoca a função `main` fornecendo a pasta de trabalho como seu primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="74256-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="74256-114">Um `ExcelScript.Workbook` deve sempre ser o primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="74256-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="74256-115">O código dentro da função `main` é executado quando o script é executado.</span><span class="sxs-lookup"><span data-stu-id="74256-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="74256-116">`main` pode chamar outras funções em seu script, mas o código que não estiver contido em uma função não será executado.</span><span class="sxs-lookup"><span data-stu-id="74256-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="74256-117">Os scripts não podem invocar ou chamar outros Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="74256-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="74256-118">O [Power Automate](https://flow.microsoft.com) permite que você conecte scripts em fluxos.</span><span class="sxs-lookup"><span data-stu-id="74256-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="74256-119">Os dados são passados entre os scripts e o fluxo por meio dos parâmetros e retornos do método `main`.</span><span class="sxs-lookup"><span data-stu-id="74256-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="74256-120">Como integrar os Scripts do Office com o Power Automate é abordado em detalhes em [Executar Scripts do Office com o Power Automate ](power-automate-integration.md).</span><span class="sxs-lookup"><span data-stu-id="74256-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="74256-121">Visão geral do modelo de objeto</span><span class="sxs-lookup"><span data-stu-id="74256-121">Object model overview</span></span>

<span data-ttu-id="74256-122">Para escrever um script, você precisa entender como as APIs do Scripts do Office se encaixam.</span><span class="sxs-lookup"><span data-stu-id="74256-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="74256-123">Os componentes de uma pasta de trabalho têm relações específicas entre si.</span><span class="sxs-lookup"><span data-stu-id="74256-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="74256-124">De várias maneiras, essas relações correspondem às da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="74256-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="74256-125">Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.</span><span class="sxs-lookup"><span data-stu-id="74256-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="74256-126">Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.</span><span class="sxs-lookup"><span data-stu-id="74256-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="74256-127">Um **Intervalo** representa um grupo de células contíguas.</span><span class="sxs-lookup"><span data-stu-id="74256-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="74256-128">Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.</span><span class="sxs-lookup"><span data-stu-id="74256-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="74256-129">Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.</span><span class="sxs-lookup"><span data-stu-id="74256-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="74256-130">As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.</span><span class="sxs-lookup"><span data-stu-id="74256-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="74256-131">Pasta de Trabalho</span><span class="sxs-lookup"><span data-stu-id="74256-131">Workbook</span></span>

<span data-ttu-id="74256-132">Todo script é fornecido com um `workbook` objeto do tipo `Workbook` pela função `main`.</span><span class="sxs-lookup"><span data-stu-id="74256-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="74256-133">Isso representa o objeto de nível superior por meio do qual seu script interage com a pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="74256-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="74256-134">O script a seguir obtém a planilha ativa da pasta de trabalho e registra seu nome.</span><span class="sxs-lookup"><span data-stu-id="74256-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="74256-135">Intervalos</span><span class="sxs-lookup"><span data-stu-id="74256-135">Ranges</span></span>

<span data-ttu-id="74256-136">Um intervalo é um grupo de células contíguas na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="74256-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="74256-137">Os scripts costumam usar uma notação estilo A1 (por ex.: **B3** para a única célula na coluna **B** e linha **3** ou **C2:F4** para as células das colunas **C** a **F** e linhas **2** a **4**) para definir intervalos.</span><span class="sxs-lookup"><span data-stu-id="74256-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="74256-138">Os intervalos têm três propriedades principais: valores, fórmulas e formato.</span><span class="sxs-lookup"><span data-stu-id="74256-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="74256-139">Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células.</span><span class="sxs-lookup"><span data-stu-id="74256-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="74256-140">Eles são acessados através de `getValues`, `getFormulas` e `getFormat`.</span><span class="sxs-lookup"><span data-stu-id="74256-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="74256-141">Os valores e fórmulas podem ser alterados com `setValues` e `setFormulas`, enquanto o formato é um objeto `RangeFormat` composto de vários objetos menores que são configurados individualmente.</span><span class="sxs-lookup"><span data-stu-id="74256-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="74256-142">Os intervalo usam matrizes bidimensionais para gerenciar informações.</span><span class="sxs-lookup"><span data-stu-id="74256-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="74256-143">Para obter mais informações sobre como lidar com matrizes na estrutura de Scripts do Office, consulte [Trabalhar com intervalos](javascript-objects.md#work-with-ranges).</span><span class="sxs-lookup"><span data-stu-id="74256-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="74256-144">Exemplo de intervalo</span><span class="sxs-lookup"><span data-stu-id="74256-144">Range sample</span></span>

<span data-ttu-id="74256-145">O exemplo a seguir mostra como criar registros de vendas.</span><span class="sxs-lookup"><span data-stu-id="74256-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="74256-146">Este script usa `Range` objetos para definir os valores, fórmulas e partes do formato.</span><span class="sxs-lookup"><span data-stu-id="74256-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="74256-147">Executar este script cria os seguintes dados na planilha atual:</span><span class="sxs-lookup"><span data-stu-id="74256-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="Uma planilha contendo um registro de vendas composto por linhas de valor, uma coluna de fórmulas e os cabeçalhos formatados.":::

### <a name="the-types-of-range-values"></a><span data-ttu-id="74256-149">Os tipos de valores do Intervalo</span><span class="sxs-lookup"><span data-stu-id="74256-149">The types of Range values</span></span>

<span data-ttu-id="74256-150">Cada célula tem um valor.</span><span class="sxs-lookup"><span data-stu-id="74256-150">Each cell has value.</span></span> <span data-ttu-id="74256-151">Este valor é o valor subjacente inserido na célula, que pode ser diferente do texto exibido no Excel.</span><span class="sxs-lookup"><span data-stu-id="74256-151">This value is the underlying value entered into the cell, which may be different from the text displayed in Excel.</span></span> <span data-ttu-id="74256-152">Por exemplo, você pode ver "2/5/2021" exibido na célula como uma data, mas o valor real é 44318.</span><span class="sxs-lookup"><span data-stu-id="74256-152">For example, you might see "5/2/2021" displayed in the cell as a date, but the actual value is 44318.</span></span> <span data-ttu-id="74256-153">Esta exibição pode ser alterada com o formato de número, mas o valor real e o tipo na célula mudam apenas quando um novo valor é definido.</span><span class="sxs-lookup"><span data-stu-id="74256-153">This display can be changed with the number format, but the actual value and type in the cell only changes when a new value is set.</span></span>

<span data-ttu-id="74256-154">Quando você estiver usando o valor da célula, é importante informar ao TypeScript qual valor você espera obter de uma célula ou intervalo.</span><span class="sxs-lookup"><span data-stu-id="74256-154">When you are using the cell value, it's important to tell TypeScript what value you are expecting to get from a cell or range.</span></span> <span data-ttu-id="74256-155">Uma célula contém um dos seguintes tipos: `string`, `number`, ou `boolean`.</span><span class="sxs-lookup"><span data-stu-id="74256-155">A cell contains one of the following types: `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="74256-156">Para que seu script trate os valores retornados como um desses tipos, você deve declarar o tipo.</span><span class="sxs-lookup"><span data-stu-id="74256-156">In order for your script to treat the returned values as one of those types, you must declare the type.</span></span>

<span data-ttu-id="74256-157">O script a seguir obtém o preço médio da tabela do exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="74256-157">The following script gets the average price from the table in the previous sample.</span></span> <span data-ttu-id="74256-158">Observe o código `priceRange.getValues() as number[][]`.</span><span class="sxs-lookup"><span data-stu-id="74256-158">Note the code `priceRange.getValues() as number[][]`.</span></span> <span data-ttu-id="74256-159">Isso [declara](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) o tipo dos valores do intervalo como um `number[][]`.</span><span class="sxs-lookup"><span data-stu-id="74256-159">This [asserts](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) the type of the range values to be a `number[][]`.</span></span> <span data-ttu-id="74256-160">Todos os valores nessa matriz podem então ser tratados como números no script.</span><span class="sxs-lookup"><span data-stu-id="74256-160">All the values in that array can then be treated as numbers in the script.</span></span>

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

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="74256-161">Gráficos, tabelas e outros objetos de dados</span><span class="sxs-lookup"><span data-stu-id="74256-161">Charts, tables, and other data objects</span></span>

<span data-ttu-id="74256-162">Os scripts podem criar e manipular estruturas de dados e visualizações no Excel.</span><span class="sxs-lookup"><span data-stu-id="74256-162">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="74256-163">As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais.</span><span class="sxs-lookup"><span data-stu-id="74256-163">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="74256-164">Eles são armazenados em coleções, que serão discutidas mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="74256-164">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="74256-165">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="74256-165">Create a table</span></span>

<span data-ttu-id="74256-p116">Crie tabelas usando intervalos preenchidos com dados. A formatação e os controles de tabela (como filtros) são automaticamente aplicados ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="74256-p116">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="74256-168">O script a seguir cria uma tabela usando os intervalos do exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="74256-168">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="74256-169">Executar esse script na planilha com os dados anteriores cria a tabela a seguir:</span><span class="sxs-lookup"><span data-stu-id="74256-169">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="Uma planilha contendo uma tabela feita do registro de vendas anterior.":::

### <a name="create-a-chart"></a><span data-ttu-id="74256-171">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="74256-171">Create a chart</span></span>

<span data-ttu-id="74256-172">Crie gráficos para visualizar os dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="74256-172">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="74256-173">Os scripts permitem inúmeras variedades de gráficos que podem ser personalizadas de acordo com suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="74256-173">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="74256-174">O script a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="74256-174">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="74256-175">Executar este script na planilha com a tabela anterior cria o seguinte gráfico:</span><span class="sxs-lookup"><span data-stu-id="74256-175">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.":::

## <a name="collections"></a><span data-ttu-id="74256-177">Coleções</span><span class="sxs-lookup"><span data-stu-id="74256-177">Collections</span></span>

<span data-ttu-id="74256-178">Quando um objeto do Excel tem uma coleção de um ou mais objetos do mesmo tipo, ele os armazena em uma matriz.</span><span class="sxs-lookup"><span data-stu-id="74256-178">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="74256-179">Por exemplo, um objeto `Workbook` contém um `Worksheet[]`.</span><span class="sxs-lookup"><span data-stu-id="74256-179">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="74256-180">Esta matriz é acessada pelo método `Workbook.getWorksheets()`.</span><span class="sxs-lookup"><span data-stu-id="74256-180">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="74256-181">Os métodos `get` que são plurais, como `Worksheet.getCharts()`, retornam toda a coleção de objetos como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="74256-181">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="74256-182">Você verá este padrão em todas as APIs de Scripts do Office: o objeto `Worksheet` tem um método `getTables()` que retorna um `Table[]`, o objeto `Table` tem um método `getColumns()` que retorna um `TableColumn[]`, como assim em diante.</span><span class="sxs-lookup"><span data-stu-id="74256-182">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="74256-183">A matriz retornada é uma matriz normal, portanto todas as operações regulares de matriz estão disponíveis para seu script.</span><span class="sxs-lookup"><span data-stu-id="74256-183">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="74256-184">Você também pode acessar objetos individuais na coleção usando o valor do índice da matriz.</span><span class="sxs-lookup"><span data-stu-id="74256-184">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="74256-185">Por exemplo, `workbook.getTables()[0]` retorna a primeira tabela da coleção.</span><span class="sxs-lookup"><span data-stu-id="74256-185">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="74256-186">Para saber mais sobre o uso da funcionalidade de matriz interna com a estrutura de Scripts do Office, consulte [Trabalhar com coleções](javascript-objects.md#work-with-collections).</span><span class="sxs-lookup"><span data-stu-id="74256-186">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="74256-187">Objetos individuais também são acessados a partir da coleção por meio de um método `get`.</span><span class="sxs-lookup"><span data-stu-id="74256-187">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="74256-188">Os métodos `get` que são singulares, como `Worksheet.getTable(name)`, retornam um único objeto e requerem uma ID ou nome para o objeto específico.</span><span class="sxs-lookup"><span data-stu-id="74256-188">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="74256-189">Esse ID ou nome geralmente é definido pelo script ou por meio da IU do Excel.</span><span class="sxs-lookup"><span data-stu-id="74256-189">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="74256-p121">O seguinte roteiro recebe todas as tabelas na pasta de trabalho. Assim, ele garante que os cabeçalhos sejam exibidos, os botões de filtro sejam visíveis e o estilo da tabela seja definido como "TableStyleLight1".</span><span class="sxs-lookup"><span data-stu-id="74256-p121">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="74256-192">Adicionar objetos do Excel com um script</span><span class="sxs-lookup"><span data-stu-id="74256-192">Add Excel objects with a script</span></span>

<span data-ttu-id="74256-193">Você pode adicionar programaticamente objetos de documento, como tabelas ou gráficos, chamando o método `add` correspondente disponível no objeto pai.</span><span class="sxs-lookup"><span data-stu-id="74256-193">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="74256-194">Não adicione manualmente objetos as matrizes de coleção.</span><span class="sxs-lookup"><span data-stu-id="74256-194">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="74256-195">Use os métodos `add` nos objetos pai, por exemplo, adicione `Table` a `Worksheet` com o método `Worksheet.addTable`.</span><span class="sxs-lookup"><span data-stu-id="74256-195">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="74256-196">O script a seguir cria, no Excel, uma tabela na primeira planilha da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="74256-196">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="74256-197">Observe que a tabela criada é enviada de volta pelo método `addTable`.</span><span class="sxs-lookup"><span data-stu-id="74256-197">Note that the created table is returned by the `addTable` method.</span></span>

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
> <span data-ttu-id="74256-198">A maioria dos objetos do Excel possui um método `setName`.</span><span class="sxs-lookup"><span data-stu-id="74256-198">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="74256-199">Isso fornece uma maneira fácil de acessar objetos do Excel posteriormente no script ou em outros scripts para a mesma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="74256-199">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="74256-200">Verifique se existe um objeto na coleção</span><span class="sxs-lookup"><span data-stu-id="74256-200">Verify an object exists in the collection</span></span>

<span data-ttu-id="74256-201">Os scripts geralmente precisam verificar se uma tabela ou objeto semelhante existe antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="74256-201">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="74256-202">Use os nomes dados por scripts ou por meio da IU do Excel para identificar os objetos necessários e agir de acordo.</span><span class="sxs-lookup"><span data-stu-id="74256-202">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="74256-203">O métodos `get` retornam `undefined` quando o objeto solicitado não está na coleção.</span><span class="sxs-lookup"><span data-stu-id="74256-203">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="74256-204">O script a seguir solicita uma tabela chamada "MinhaTabela" e utiliza uma instrução `if...else` para verificar se a tabela foi encontrada.</span><span class="sxs-lookup"><span data-stu-id="74256-204">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

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

<span data-ttu-id="74256-205">Um padrão comum em Scripts do Office é recriar uma tabela, gráfico ou outro objeto sempre que o script for executado.</span><span class="sxs-lookup"><span data-stu-id="74256-205">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="74256-206">Se você não precisa de dados antigos, é melhor excluir o objeto antigo antes de criar o novo.</span><span class="sxs-lookup"><span data-stu-id="74256-206">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="74256-207">Isso evita conflitos de nome ou outras diferenças que possam ter sido introduzidas por outros usuários.</span><span class="sxs-lookup"><span data-stu-id="74256-207">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="74256-208">O script a seguir remove a tabela chamada "MinhaTabela", se houver, e adiciona uma nova tabela com o mesmo nome.</span><span class="sxs-lookup"><span data-stu-id="74256-208">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

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

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="74256-209">Remova objetos do Excel com um script</span><span class="sxs-lookup"><span data-stu-id="74256-209">Remove Excel objects with a script</span></span>

<span data-ttu-id="74256-210">Para excluir um objeto, chame o método `delete` do objeto.</span><span class="sxs-lookup"><span data-stu-id="74256-210">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="74256-211">Como na adição de objetos, não remova manualmente objetos de matrizes de coleção.</span><span class="sxs-lookup"><span data-stu-id="74256-211">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="74256-212">Use os métodos `delete` nos objetos do tipo coleção.</span><span class="sxs-lookup"><span data-stu-id="74256-212">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="74256-213">Por exemplo, remova um `Table` de um `Worksheet` usando `Table.delete`.</span><span class="sxs-lookup"><span data-stu-id="74256-213">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="74256-214">O script a seguir remove a primeira planilha da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="74256-214">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="74256-215">Leituras adicionais sobre o modelo de objeto</span><span class="sxs-lookup"><span data-stu-id="74256-215">Further reading on the object model</span></span>

<span data-ttu-id="74256-216">A [documentação de referência de API dos scripts do Office](/javascript/api/office-scripts/overview) é uma lista completa dos objetos usados nos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="74256-216">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="74256-217">Lá, você pode usar o sumário para navegar para qualquer classe da qual quiser saber mais.</span><span class="sxs-lookup"><span data-stu-id="74256-217">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="74256-218">Estas são várias páginas exibidas com frequência.</span><span class="sxs-lookup"><span data-stu-id="74256-218">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="74256-219">Gráfico</span><span class="sxs-lookup"><span data-stu-id="74256-219">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="74256-220">Comentário</span><span class="sxs-lookup"><span data-stu-id="74256-220">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="74256-221">PivotTable</span><span class="sxs-lookup"><span data-stu-id="74256-221">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="74256-222">Range</span><span class="sxs-lookup"><span data-stu-id="74256-222">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="74256-223">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="74256-223">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="74256-224">Formato</span><span class="sxs-lookup"><span data-stu-id="74256-224">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="74256-225">Table</span><span class="sxs-lookup"><span data-stu-id="74256-225">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="74256-226">Pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="74256-226">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="74256-227">Planilha</span><span class="sxs-lookup"><span data-stu-id="74256-227">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="74256-228">Confira também</span><span class="sxs-lookup"><span data-stu-id="74256-228">See also</span></span>

- [<span data-ttu-id="74256-229">Gravar, editar e criar scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="74256-229">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="74256-230">Ler os dados da pasta de trabalho com scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="74256-230">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="74256-231">Referência da API de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="74256-231">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="74256-232">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="74256-232">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="74256-233">Práticas recomendadas nos Scripts do Office </span><span class="sxs-lookup"><span data-stu-id="74256-233">Best practices in Office Scripts</span></span>](best-practices.md)
