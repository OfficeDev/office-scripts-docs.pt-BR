---
title: Suporte a scripts do Office mais antigos que usam APIs assíncronas
description: Uma cartilha sobre as APIs Assíncronas de Scripts do Office e como usar o padrão de carregamento/sincronização para scripts mais antigos.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: be7847efe59dc6026875b8a8e3b3c93e0eb82e4d
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242022"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="c687d-103">Suporte a scripts do Office mais antigos que usam APIs assíncronas</span><span class="sxs-lookup"><span data-stu-id="c687d-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="c687d-104">Este artigo ensina como manter e atualizar scripts que usam as APIs assíncronas do modelo mais antigo.</span><span class="sxs-lookup"><span data-stu-id="c687d-104">This article will teach you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="c687d-105">Essas APIs têm a mesma funcionalidade principal das APIs de Scripts do Office agora padrão e síncronas, mas exigem que seu script controle a sincronização de dados entre o script e a agenda.</span><span class="sxs-lookup"><span data-stu-id="c687d-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c687d-106">O modelo assíncrono só pode ser usado com scripts criados antes da implementação do modelo de [API atual.](scripting-fundamentals.md?view=office-scripts&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="c687d-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts&preserve-view=true).</span></span> <span data-ttu-id="c687d-107">Os scripts são permanentemente bloqueados para o modelo de API que eles têm na criação.</span><span class="sxs-lookup"><span data-stu-id="c687d-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="c687d-108">Isso também significa que, se você quiser converter um script antigo para o novo modelo, deverá criar um novo script.</span><span class="sxs-lookup"><span data-stu-id="c687d-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="c687d-109">Recomendamos que você atualize seus scripts antigos para o novo modelo ao fazer alterações, já que o modelo atual é mais fácil de usar.</span><span class="sxs-lookup"><span data-stu-id="c687d-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="c687d-110">A [seção conversão de scripts assíncronos](#converting-async-scripts-to-the-current-model) para o modelo atual tem conselhos sobre como fazer essa transição.</span><span class="sxs-lookup"><span data-stu-id="c687d-110">The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="c687d-111">função `main`</span><span class="sxs-lookup"><span data-stu-id="c687d-111">`main` function</span></span>

<span data-ttu-id="c687d-112">Scripts que usam as APIs assíncronas têm uma função `main` diferente.</span><span class="sxs-lookup"><span data-stu-id="c687d-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="c687d-113">É uma função `async` que tem um como primeiro `Excel.RequestContext` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="c687d-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="c687d-114">Contexto</span><span class="sxs-lookup"><span data-stu-id="c687d-114">Context</span></span>

<span data-ttu-id="c687d-115">A função `main` aceita um parâmetro `Excel.RequestContext`, chamado `context`.</span><span class="sxs-lookup"><span data-stu-id="c687d-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="c687d-116">Imagine `context` como a ponte entre o seu script e a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c687d-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="c687d-117">Seu script acessa a pasta de trabalho com o objeto `context` e usa esse `context` para troca de dados.</span><span class="sxs-lookup"><span data-stu-id="c687d-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="c687d-118">O objeto `context` é necessário porque o script e o Excel estão sendo executados em processos e locais diferentes.</span><span class="sxs-lookup"><span data-stu-id="c687d-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="c687d-119">O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem.</span><span class="sxs-lookup"><span data-stu-id="c687d-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="c687d-120">O objeto `context` gerencia essas transações.</span><span class="sxs-lookup"><span data-stu-id="c687d-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="c687d-121">Sincronizar e carregar</span><span class="sxs-lookup"><span data-stu-id="c687d-121">Sync and Load</span></span>

<span data-ttu-id="c687d-122">Como o seu script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois levará algum tempo.</span><span class="sxs-lookup"><span data-stu-id="c687d-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="c687d-123">Na API assíncrona, os comandos ficam na fila até que o script chama explicitamente a operação para sincronizar o script e a `sync` agenda.</span><span class="sxs-lookup"><span data-stu-id="c687d-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="c687d-124">Seu script pode trabalhar de forma independente até que precise executar uma das seguintes ações:</span><span class="sxs-lookup"><span data-stu-id="c687d-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="c687d-125">Leia os dados da pasta de trabalho (seguindo uma `load` operação ou método que retorne um [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span><span class="sxs-lookup"><span data-stu-id="c687d-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span></span>
- <span data-ttu-id="c687d-126">Gravar dados na pasta de trabalho (geralmente porque o script terminou).</span><span class="sxs-lookup"><span data-stu-id="c687d-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="c687d-127">A imagem a seguir mostra um exemplo de fluxo de controle entre o script e a pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="c687d-127">The following image shows an example control flow between the script and workbook:</span></span>

![Um diagrama mostrando operações de leitura e gravação saindo do script e indo para a pasta de trabalho.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="c687d-129">Sincronizar</span><span class="sxs-lookup"><span data-stu-id="c687d-129">Sync</span></span>

<span data-ttu-id="c687d-130">Sempre que seu script assíncrono precisar ler ou gravar dados na agenda, chame o método `RequestContext.sync` conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="c687d-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="c687d-131">`context.sync()` é chamado implicitamente quando um script termina.</span><span class="sxs-lookup"><span data-stu-id="c687d-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="c687d-132">Após a conclusão da operação `sync`, a pasta de trabalho será atualizada para refletir as operações de gravação especificados por esse script.</span><span class="sxs-lookup"><span data-stu-id="c687d-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="c687d-133">Uma operação de gravação está definindo qualquer propriedade em um objeto do Excel (por exemplo, ) ou chamando um método que altera uma propriedade `range.format.fill.color = "red"` (por exemplo, `range.format.autoFitColumns()` ).</span><span class="sxs-lookup"><span data-stu-id="c687d-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="c687d-134">A `sync` operação também lê todos os valores da pasta de trabalho que o script solicitou usando uma `load` operação ou um método que retorna a `ClientResult` (conforme discutido nas próximas seções).</span><span class="sxs-lookup"><span data-stu-id="c687d-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="c687d-135">A sincronização do seu script com a pasta de trabalho pode demorar, dependendo da sua rede.</span><span class="sxs-lookup"><span data-stu-id="c687d-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="c687d-136">Minimize o número de chamadas `sync` para ajudar seu script a ser executado rapidamente.</span><span class="sxs-lookup"><span data-stu-id="c687d-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="c687d-137">Caso contrário, as APIs assíncronas não são mais rápidas do que as APIs síncronas padrão.</span><span class="sxs-lookup"><span data-stu-id="c687d-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="c687d-138">Carregar</span><span class="sxs-lookup"><span data-stu-id="c687d-138">Load</span></span>

<span data-ttu-id="c687d-139">Um script assíncrono deve carregar dados da área de trabalho antes de lê-lo.</span><span class="sxs-lookup"><span data-stu-id="c687d-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="c687d-140">No entanto, o carregamento de dados de toda a área de trabalho reduziria significativamente a velocidade do script.</span><span class="sxs-lookup"><span data-stu-id="c687d-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="c687d-141">O `load` método permite que seu script a estado específico quais dados devem ser recuperados da agenda de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c687d-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="c687d-142">O método `load` está disponível em todos os objetos do Excel.</span><span class="sxs-lookup"><span data-stu-id="c687d-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="c687d-143">Seu script deve carregar as propriedades de um objeto para poder lê-lo.</span><span class="sxs-lookup"><span data-stu-id="c687d-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="c687d-144">Não fazer isso resulta em um erro.</span><span class="sxs-lookup"><span data-stu-id="c687d-144">Not doing so results in an error.</span></span>

<span data-ttu-id="c687d-145">Os exemplos a seguir usam um objeto `Range` para mostrar as três maneiras de usar o método `load` para carregar dados.</span><span class="sxs-lookup"><span data-stu-id="c687d-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="c687d-146">Finalidade</span><span class="sxs-lookup"><span data-stu-id="c687d-146">Intent</span></span> |<span data-ttu-id="c687d-147">Comando de exemplo</span><span class="sxs-lookup"><span data-stu-id="c687d-147">Example Command</span></span> | <span data-ttu-id="c687d-148">Efeito</span><span class="sxs-lookup"><span data-stu-id="c687d-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="c687d-149">Carregar uma propriedade</span><span class="sxs-lookup"><span data-stu-id="c687d-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="c687d-150">Carrega uma única propriedade, neste caso, a matriz bidimensional de valores nesse intervalo.</span><span class="sxs-lookup"><span data-stu-id="c687d-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="c687d-151">Carregar várias propriedades</span><span class="sxs-lookup"><span data-stu-id="c687d-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="c687d-152">Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e de colunas.</span><span class="sxs-lookup"><span data-stu-id="c687d-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="c687d-153">Carregar tudo</span><span class="sxs-lookup"><span data-stu-id="c687d-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="c687d-154">Carrega todas as propriedades no intervalo.</span><span class="sxs-lookup"><span data-stu-id="c687d-154">Loads all the properties on the range.</span></span> <span data-ttu-id="c687d-155">Essa não é uma solução recomendada, pois reduzirá a velocidade do script ao obter dados desnecessários.</span><span class="sxs-lookup"><span data-stu-id="c687d-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="c687d-156">Use isso somente durante o teste do script ou se você precisar de todas as propriedades do objeto.</span><span class="sxs-lookup"><span data-stu-id="c687d-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="c687d-157">Seu script deve chamar `context.sync()` antes de ler os valores carregados.</span><span class="sxs-lookup"><span data-stu-id="c687d-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

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

<span data-ttu-id="c687d-158">Você também pode carregar as propriedades em uma coleção.</span><span class="sxs-lookup"><span data-stu-id="c687d-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="c687d-159">Cada objeto de coleção na API assíncrona tem uma propriedade que é `items` uma matriz que contém os objetos dessa coleção.</span><span class="sxs-lookup"><span data-stu-id="c687d-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="c687d-160">Usar `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carrega as propriedades especificadas em cada um desses itens.</span><span class="sxs-lookup"><span data-stu-id="c687d-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="c687d-161">O exemplo a seguir carrega a propriedade `resolved` em cada objeto `Comment` no objeto `CommentCollection` de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="c687d-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

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

### <a name="clientresult"></a><span data-ttu-id="c687d-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="c687d-162">ClientResult</span></span>

<span data-ttu-id="c687d-163">Os métodos na API assíncrona que retornam informações da agenda têm um padrão semelhante ao `load` / `sync` paradigma.</span><span class="sxs-lookup"><span data-stu-id="c687d-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="c687d-164">Por exemplo, `TableCollection.getCount` obtém o número de tabelas da coleção.</span><span class="sxs-lookup"><span data-stu-id="c687d-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="c687d-165">`getCount` retorna um `ClientResult<number>` , o que significa que a propriedade no retornado é um `value` [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) número.</span><span class="sxs-lookup"><span data-stu-id="c687d-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) is a number.</span></span> <span data-ttu-id="c687d-166">Seu script não pode acessar esse valor até que `context.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="c687d-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="c687d-167">Assim como carregar uma propriedade, o `value` é um valor local "vazio" até a `sync` chamada.</span><span class="sxs-lookup"><span data-stu-id="c687d-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="c687d-168">O script a seguir obtém o número total de tabelas na pasta de trabalho e registra esse número no console.</span><span class="sxs-lookup"><span data-stu-id="c687d-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

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

## <a name="converting-async-scripts-to-the-current-model"></a><span data-ttu-id="c687d-169">Converter scripts assíncronos para o modelo atual</span><span class="sxs-lookup"><span data-stu-id="c687d-169">Converting async scripts to the current model</span></span>

<span data-ttu-id="c687d-170">O modelo de API atual não usa `load` , `sync` ou um `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="c687d-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="c687d-171">Isso facilita muito a gravação e a manutenção dos scripts.</span><span class="sxs-lookup"><span data-stu-id="c687d-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="c687d-172">Seu melhor recurso para converter scripts antigos é [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="c687d-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="c687d-173">Lá, você pode pedir ajuda à comunidade em cenários específicos.</span><span class="sxs-lookup"><span data-stu-id="c687d-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="c687d-174">As diretrizes a seguir devem ajudar a delinear as etapas gerais que você precisará seguir.</span><span class="sxs-lookup"><span data-stu-id="c687d-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="c687d-175">Crie um novo script e copie o código assíncrono antigo nele.</span><span class="sxs-lookup"><span data-stu-id="c687d-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="c687d-176">Certifique-se de não incluir a assinatura `main` do método antigo, usando a `function main(workbook: ExcelScript.Workbook)` atual.</span><span class="sxs-lookup"><span data-stu-id="c687d-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="c687d-177">Remova todas as `load` chamadas `sync` e todas.</span><span class="sxs-lookup"><span data-stu-id="c687d-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="c687d-178">Eles não são mais necessários.</span><span class="sxs-lookup"><span data-stu-id="c687d-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="c687d-179">Todas as propriedades foram removidas.</span><span class="sxs-lookup"><span data-stu-id="c687d-179">All properties have been removed.</span></span> <span data-ttu-id="c687d-180">Agora você acessa esses objetos `get` por meio de métodos e, portanto, precisará alternar essas referências `set` de propriedade para chamadas de método.</span><span class="sxs-lookup"><span data-stu-id="c687d-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="c687d-181">Por exemplo, em vez de definir a cor de preenchimento de uma célula por meio do acesso a propriedades como este: , agora você usará `mySheet.getRange("A2:C2").format.fill.color = "blue";` métodos como este: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="c687d-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="c687d-182">As classes de coleção foram substituídas por matrizes.</span><span class="sxs-lookup"><span data-stu-id="c687d-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="c687d-183">Os métodos e as classes dessas coleções foram movidos para o objeto que era proprietário da coleção, portanto, suas referências `add` `get` devem ser atualizadas de acordo.</span><span class="sxs-lookup"><span data-stu-id="c687d-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="c687d-184">Por exemplo, para obter um gráfico chamado "MyChart" da primeira planilha na pasta de trabalho, use o seguinte código: `workbook.getWorksheets()[0].getChart("MyChart");` .</span><span class="sxs-lookup"><span data-stu-id="c687d-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="c687d-185">Observe que `[0]` para acessar o primeiro valor do retornado por `Worksheet[]` `getWorksheets()` .</span><span class="sxs-lookup"><span data-stu-id="c687d-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="c687d-186">Alguns métodos foram renomeados para maior clareza e adicionados por conveniência.</span><span class="sxs-lookup"><span data-stu-id="c687d-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="c687d-187">Consulte a referência [da API de Scripts do Office](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="c687d-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="c687d-188">Documentação de referência da API assíncrona de Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="c687d-188">Office Scripts async API reference documentation</span></span>

<span data-ttu-id="c687d-189">As APIs assíncronas são equivalentes às usadas nos Complementos do Office. A documentação de referência é encontrada na seção do Excel da referência da API JavaScript de [Complementos do Office.](/javascript/api/excel?view=excel-js-online&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="c687d-189">The async APIs are equivalent to those used in Office Add-ins. The reference documentation is found in [the Excel section of the Office Add-ins JavaScript API reference](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>
