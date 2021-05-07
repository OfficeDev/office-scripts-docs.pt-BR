---
title: Dar suporte Office scripts mais antigos que usam as APIs assíncronas
description: Uma cartilha nas APIs Office Scripts Async e como usar o padrão de carga/sincronização para scripts mais antigos.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 437fb2e389d6d3963f93cdb44c5529749c4d3569
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232407"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="75def-103">Dar suporte Office scripts mais antigos que usam as APIs assíncronas</span><span class="sxs-lookup"><span data-stu-id="75def-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="75def-104">Este artigo ensinará como manter e atualizar scripts que usam ASIs assíncronas do modelo mais antigo.</span><span class="sxs-lookup"><span data-stu-id="75def-104">This article will teach you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="75def-105">Essas APIs têm a mesma funcionalidade principal que as APIs de Scripts Office padrão agora padrão, mas exigem que seu script controle a sincronização de dados entre o script e a workbook.</span><span class="sxs-lookup"><span data-stu-id="75def-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="75def-106">O modelo assíncrono só pode ser usado com scripts criados antes da implementação do modelo [de API atual.](scripting-fundamentals.md)</span><span class="sxs-lookup"><span data-stu-id="75def-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md).</span></span> <span data-ttu-id="75def-107">Os scripts são permanentemente bloqueados para o modelo de API que eles têm após a criação.</span><span class="sxs-lookup"><span data-stu-id="75def-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="75def-108">Isso também significa que, se você quiser converter um script antigo para o novo modelo, deverá criar um novo script.</span><span class="sxs-lookup"><span data-stu-id="75def-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="75def-109">Recomendamos que você atualize seus scripts antigos para o novo modelo ao fazer alterações, já que o modelo atual é mais fácil de usar.</span><span class="sxs-lookup"><span data-stu-id="75def-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="75def-110">A [seção Converter scripts assíncronos para o modelo](#converting-async-scripts-to-the-current-model) atual tem conselhos sobre como fazer essa transição.</span><span class="sxs-lookup"><span data-stu-id="75def-110">The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="75def-111">função `main`</span><span class="sxs-lookup"><span data-stu-id="75def-111">`main` function</span></span>

<span data-ttu-id="75def-112">Scripts que usam as APIs assíncronas têm uma função `main` diferente.</span><span class="sxs-lookup"><span data-stu-id="75def-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="75def-113">É uma função `async` que tem um como o primeiro `Excel.RequestContext` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="75def-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="75def-114">Contexto</span><span class="sxs-lookup"><span data-stu-id="75def-114">Context</span></span>

<span data-ttu-id="75def-115">A função `main` aceita um parâmetro `Excel.RequestContext`, chamado `context`.</span><span class="sxs-lookup"><span data-stu-id="75def-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="75def-116">Imagine `context` como a ponte entre o seu script e a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="75def-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="75def-117">Seu script acessa a pasta de trabalho com o objeto `context` e usa esse `context` para troca de dados.</span><span class="sxs-lookup"><span data-stu-id="75def-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="75def-118">O objeto `context` é necessário porque o script e o Excel estão sendo executados em processos e locais diferentes.</span><span class="sxs-lookup"><span data-stu-id="75def-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="75def-119">O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem.</span><span class="sxs-lookup"><span data-stu-id="75def-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="75def-120">O objeto `context` gerencia essas transações.</span><span class="sxs-lookup"><span data-stu-id="75def-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="75def-121">Sincronizar e carregar</span><span class="sxs-lookup"><span data-stu-id="75def-121">Sync and Load</span></span>

<span data-ttu-id="75def-122">Como o seu script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois levará algum tempo.</span><span class="sxs-lookup"><span data-stu-id="75def-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="75def-123">Na API assíncrona, os comandos são enraizado até que o script chama explicitamente a operação para `sync` sincronizar o script e a workbook.</span><span class="sxs-lookup"><span data-stu-id="75def-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="75def-124">Seu script pode trabalhar de forma independente até que precise executar uma das seguintes ações:</span><span class="sxs-lookup"><span data-stu-id="75def-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="75def-125">Leia os dados da pasta de trabalho (seguindo uma `load` operação ou método que retorne um [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span><span class="sxs-lookup"><span data-stu-id="75def-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span></span>
- <span data-ttu-id="75def-126">Gravar dados na pasta de trabalho (geralmente porque o script terminou).</span><span class="sxs-lookup"><span data-stu-id="75def-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="75def-127">A imagem a seguir mostra um exemplo de fluxo de controle entre o script e a pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="75def-127">The following image shows an example control flow between the script and workbook:</span></span>

:::image type="content" source="../images/load-sync.png" alt-text="Um diagrama mostrando operações de leitura e gravação indo para a área de trabalho do script":::

### <a name="sync"></a><span data-ttu-id="75def-129">Sincronizar</span><span class="sxs-lookup"><span data-stu-id="75def-129">Sync</span></span>

<span data-ttu-id="75def-130">Sempre que o script assíncrono precisar ler dados ou gravar dados na workbook, chame o método `RequestContext.sync` conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="75def-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="75def-131">`context.sync()` é chamado implicitamente quando um script termina.</span><span class="sxs-lookup"><span data-stu-id="75def-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="75def-132">Após a conclusão da operação `sync`, a pasta de trabalho será atualizada para refletir as operações de gravação especificados por esse script.</span><span class="sxs-lookup"><span data-stu-id="75def-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="75def-133">Uma operação de gravação está definindo qualquer propriedade em um objeto Excel (por exemplo, ) ou chamando um método que altera uma `range.format.fill.color = "red"` propriedade (por exemplo, `range.format.autoFitColumns()` ).</span><span class="sxs-lookup"><span data-stu-id="75def-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="75def-134">A `sync` operação também lê todos os valores da pasta de trabalho que o script solicitou usando uma `load` operação ou um método que retorna a `ClientResult` (conforme discutido nas próximas seções).</span><span class="sxs-lookup"><span data-stu-id="75def-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="75def-135">A sincronização do seu script com a pasta de trabalho pode demorar, dependendo da sua rede.</span><span class="sxs-lookup"><span data-stu-id="75def-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="75def-136">Minimize o número de `sync` chamadas para ajudar seu script a executar rapidamente.</span><span class="sxs-lookup"><span data-stu-id="75def-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="75def-137">Caso contrário, as APIs assíncronas não são mais rápidas que as APIs padrão e síncrona.</span><span class="sxs-lookup"><span data-stu-id="75def-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="75def-138">Carregar</span><span class="sxs-lookup"><span data-stu-id="75def-138">Load</span></span>

<span data-ttu-id="75def-139">Um script assíncrono deve carregar dados da workbook antes de lê-lo.</span><span class="sxs-lookup"><span data-stu-id="75def-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="75def-140">No entanto, carregar dados de toda a workbook reduziria consideravelmente a velocidade do script.</span><span class="sxs-lookup"><span data-stu-id="75def-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="75def-141">O `load` método permite que seu script estado especificamente quais dados devem ser recuperados da workbook.</span><span class="sxs-lookup"><span data-stu-id="75def-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="75def-142">O método `load` está disponível em todos os objetos do Excel.</span><span class="sxs-lookup"><span data-stu-id="75def-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="75def-143">Seu script deve carregar as propriedades de um objeto para poder lê-lo.</span><span class="sxs-lookup"><span data-stu-id="75def-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="75def-144">Não fazer isso resulta em um erro.</span><span class="sxs-lookup"><span data-stu-id="75def-144">Not doing so results in an error.</span></span>

<span data-ttu-id="75def-145">Os exemplos a seguir usam um objeto `Range` para mostrar as três maneiras de usar o método `load` para carregar dados.</span><span class="sxs-lookup"><span data-stu-id="75def-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="75def-146">Finalidade</span><span class="sxs-lookup"><span data-stu-id="75def-146">Intent</span></span> |<span data-ttu-id="75def-147">Comando de exemplo</span><span class="sxs-lookup"><span data-stu-id="75def-147">Example Command</span></span> | <span data-ttu-id="75def-148">Efeito</span><span class="sxs-lookup"><span data-stu-id="75def-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="75def-149">Carregar uma propriedade</span><span class="sxs-lookup"><span data-stu-id="75def-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="75def-150">Carrega uma única propriedade, neste caso, a matriz bidimensional de valores nesse intervalo.</span><span class="sxs-lookup"><span data-stu-id="75def-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="75def-151">Carregar várias propriedades</span><span class="sxs-lookup"><span data-stu-id="75def-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="75def-152">Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e de colunas.</span><span class="sxs-lookup"><span data-stu-id="75def-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="75def-153">Carregar tudo</span><span class="sxs-lookup"><span data-stu-id="75def-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="75def-154">Carrega todas as propriedades no intervalo.</span><span class="sxs-lookup"><span data-stu-id="75def-154">Loads all the properties on the range.</span></span> <span data-ttu-id="75def-155">Essa não é uma solução recomendada, pois reduzirá a velocidade do script ao obter dados desnecessários.</span><span class="sxs-lookup"><span data-stu-id="75def-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="75def-156">Use isso somente durante o teste do script ou se você precisar de todas as propriedades do objeto.</span><span class="sxs-lookup"><span data-stu-id="75def-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="75def-157">Seu script deve chamar `context.sync()` antes de ler os valores carregados.</span><span class="sxs-lookup"><span data-stu-id="75def-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

<span data-ttu-id="75def-158">Você também pode carregar as propriedades em uma coleção.</span><span class="sxs-lookup"><span data-stu-id="75def-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="75def-159">Cada objeto de coleção na API assíncrona tem uma propriedade que `items` é uma matriz que contém os objetos nessa coleção.</span><span class="sxs-lookup"><span data-stu-id="75def-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="75def-160">Usar `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carrega as propriedades especificadas em cada um desses itens.</span><span class="sxs-lookup"><span data-stu-id="75def-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="75def-161">O exemplo a seguir carrega a propriedade `resolved` em cada objeto `Comment` no objeto `CommentCollection` de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="75def-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a><span data-ttu-id="75def-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="75def-162">ClientResult</span></span>

<span data-ttu-id="75def-163">Os métodos na API assíncrona que retornam informações da agenda de trabalho têm um padrão semelhante ao `load` / `sync` paradigma.</span><span class="sxs-lookup"><span data-stu-id="75def-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="75def-164">Por exemplo, `TableCollection.getCount` obtém o número de tabelas da coleção.</span><span class="sxs-lookup"><span data-stu-id="75def-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="75def-165">`getCount` retorna um `ClientResult<number>`, o que significa que a propriedade `value` em [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) retornado é um número.</span><span class="sxs-lookup"><span data-stu-id="75def-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) is a number.</span></span> <span data-ttu-id="75def-166">Seu script não pode acessar esse valor até que `context.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="75def-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="75def-167">Assim como carregar uma propriedade, o `value` é um valor local "vazio" até a `sync` chamada.</span><span class="sxs-lookup"><span data-stu-id="75def-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="75def-168">O script a seguir obtém o número total de tabelas na pasta de trabalho e registra esse número no console.</span><span class="sxs-lookup"><span data-stu-id="75def-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-async-scripts-to-the-current-model"></a><span data-ttu-id="75def-169">Converter scripts assíncronos no modelo atual</span><span class="sxs-lookup"><span data-stu-id="75def-169">Converting async scripts to the current model</span></span>

<span data-ttu-id="75def-170">O modelo de API atual não usa `load` , `sync` ou um `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="75def-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="75def-171">Isso torna os scripts muito mais fáceis de gravar e manter.</span><span class="sxs-lookup"><span data-stu-id="75def-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="75def-172">Seu melhor recurso para converter scripts antigos é [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="75def-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="75def-173">Lá, você pode pedir ajuda à comunidade com cenários específicos.</span><span class="sxs-lookup"><span data-stu-id="75def-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="75def-174">As diretrizes a seguir devem ajudar a delinear as etapas gerais que você precisará seguir.</span><span class="sxs-lookup"><span data-stu-id="75def-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="75def-175">Crie um novo script e copie o código assíncrono antigo para ele.</span><span class="sxs-lookup"><span data-stu-id="75def-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="75def-176">Certifique-se de não incluir a assinatura `main` do método antigo, usando a `function main(workbook: ExcelScript.Workbook)` atual.</span><span class="sxs-lookup"><span data-stu-id="75def-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="75def-177">Remova todas as `load` chamadas `sync` e.</span><span class="sxs-lookup"><span data-stu-id="75def-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="75def-178">Eles não são mais necessários.</span><span class="sxs-lookup"><span data-stu-id="75def-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="75def-179">Todas as propriedades foram removidas.</span><span class="sxs-lookup"><span data-stu-id="75def-179">All properties have been removed.</span></span> <span data-ttu-id="75def-180">Agora você acessa esses objetos por meio de métodos e, portanto, precisará alternar essas referências `get` de propriedade para chamadas de `set` método.</span><span class="sxs-lookup"><span data-stu-id="75def-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="75def-181">Por exemplo, em vez de definir a cor de preenchimento de uma célula por meio do acesso a propriedades como este: , agora você usará `mySheet.getRange("A2:C2").format.fill.color = "blue";` métodos como este: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="75def-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="75def-182">Classes de coleção foram substituídas por matrizes.</span><span class="sxs-lookup"><span data-stu-id="75def-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="75def-183">Os métodos e dessas classes de coleção foram movidos para o objeto que possuía a coleção, portanto, suas referências devem `add` `get` ser atualizadas de acordo.</span><span class="sxs-lookup"><span data-stu-id="75def-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="75def-184">Por exemplo, para obter um gráfico chamado "MyChart" da primeira planilha da pasta de trabalho, use o seguinte código: `workbook.getWorksheets()[0].getChart("MyChart");` .</span><span class="sxs-lookup"><span data-stu-id="75def-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="75def-185">Observe o `[0]` para acessar o primeiro valor do retornado por `Worksheet[]` `getWorksheets()` .</span><span class="sxs-lookup"><span data-stu-id="75def-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="75def-186">Alguns métodos foram renomeados para clareza e adicionados por conveniência.</span><span class="sxs-lookup"><span data-stu-id="75def-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="75def-187">Consulte a referência [da API Office Scripts](/javascript/api/office-scripts/overview) para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="75def-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="75def-188">Office Documentação de referência da API assíncrona de scripts</span><span class="sxs-lookup"><span data-stu-id="75def-188">Office Scripts async API reference documentation</span></span>

<span data-ttu-id="75def-189">As APIs assíncronas são equivalentes às usadas em Office de complementos. A documentação de referência é encontrada na seção Excel da referência da API JavaScript de Office de [complementos.](/javascript/api/excel?view=excel-js-online&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="75def-189">The async APIs are equivalent to those used in Office Add-ins. The reference documentation is found in [the Excel section of the Office Add-ins JavaScript API reference](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>
