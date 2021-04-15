---
title: Suporte a scripts do Office mais antigos que usam as APIs assíncronas
description: Uma cartilha nas APIs assíncronas de Scripts do Office e como usar o padrão de carga/sincronização para scripts mais antigos.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 143f52a7ffefb4f19ee36ba4343fd7c2f1cbdffe
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755074"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Suporte a scripts do Office mais antigos que usam as APIs assíncronas

Este artigo ensinará como manter e atualizar scripts que usam ASIs assíncronas do modelo mais antigo. Essas APIs têm a mesma funcionalidade principal que as APIs de Scripts do Office agora padrão, mas exigem que seu script controle a sincronização de dados entre o script e a workbook.

> [!IMPORTANT]
> O modelo assíncrono só pode ser usado com scripts criados antes da implementação do modelo [de API atual.](scripting-fundamentals.md) Os scripts são permanentemente bloqueados para o modelo de API que eles têm após a criação. Isso também significa que, se você quiser converter um script antigo para o novo modelo, deverá criar um novo script. Recomendamos que você atualize seus scripts antigos para o novo modelo ao fazer alterações, já que o modelo atual é mais fácil de usar. A [seção Converter scripts assíncronos para o modelo](#converting-async-scripts-to-the-current-model) atual tem conselhos sobre como fazer essa transição.

## <a name="main-function"></a>função `main`

Scripts que usam as APIs assíncronas têm uma função `main` diferente. É uma função `async` que tem um como o primeiro `Excel.RequestContext` parâmetro.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Contexto

A função `main` aceita um parâmetro `Excel.RequestContext`, chamado `context`. Imagine `context` como a ponte entre o seu script e a pasta de trabalho. Seu script acessa a pasta de trabalho com o objeto `context` e usa esse `context` para troca de dados.

O objeto `context` é necessário porque o script e o Excel estão sendo executados em processos e locais diferentes. O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem. O objeto `context` gerencia essas transações.

## <a name="sync-and-load"></a>Sincronizar e carregar

Como o seu script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois levará algum tempo. Na API assíncrona, os comandos são enraizado até que o script chama explicitamente a operação para `sync` sincronizar o script e a workbook. Seu script pode trabalhar de forma independente até que precise executar uma das seguintes ações:

- Leia os dados da pasta de trabalho (seguindo uma `load` operação ou método que retorne um [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Gravar dados na pasta de trabalho (geralmente porque o script terminou).

A imagem a seguir mostra um exemplo de fluxo de controle entre o script e a pasta de trabalho:

:::image type="content" source="../images/load-sync.png" alt-text="Um diagrama mostrando operações de leitura e gravação saindo do script e indo para a pasta de trabalho.":::

### <a name="sync"></a>Sincronizar

Sempre que o script assíncrono precisar ler dados ou gravar dados na workbook, chame o método `RequestContext.sync` conforme mostrado aqui:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` é chamado implicitamente quando um script termina.

Após a conclusão da operação `sync`, a pasta de trabalho será atualizada para refletir as operações de gravação especificados por esse script. Uma operação de gravação está definindo qualquer propriedade em um objeto excel (por exemplo, ) ou chamando um método que altera `range.format.fill.color = "red"` uma propriedade (por exemplo, `range.format.autoFitColumns()` ). A `sync` operação também lê todos os valores da pasta de trabalho que o script solicitou usando uma `load` operação ou um método que retorna a `ClientResult` (conforme discutido nas próximas seções).

A sincronização do seu script com a pasta de trabalho pode demorar, dependendo da sua rede. Minimize o número de `sync` chamadas para ajudar seu script a executar rapidamente. Caso contrário, as APIs assíncronas não são mais rápidas que as APIs padrão e síncrona.

### <a name="load"></a>Carregar

Um script assíncrono deve carregar dados da workbook antes de lê-lo. No entanto, carregar dados de toda a workbook reduziria consideravelmente a velocidade do script. O `load` método permite que seu script estado especificamente quais dados devem ser recuperados da workbook.

O método `load` está disponível em todos os objetos do Excel. Seu script deve carregar as propriedades de um objeto para poder lê-lo. Não fazer isso resulta em um erro.

Os exemplos a seguir usam um objeto `Range` para mostrar as três maneiras de usar o método `load` para carregar dados.

|Finalidade |Comando de exemplo | Efeito |
|:--|:--|:--|
|Carregar uma propriedade |`myRange.load("values");` | Carrega uma única propriedade, neste caso, a matriz bidimensional de valores nesse intervalo. |
|Carregar várias propriedades |`myRange.load("values, rowCount, columnCount");`| Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e de colunas. |
|Carregar tudo | `myRange.load();`|Carrega todas as propriedades no intervalo. Essa não é uma solução recomendada, pois reduzirá a velocidade do script ao obter dados desnecessários. Use isso somente durante o teste do script ou se você precisar de todas as propriedades do objeto. |

Seu script deve chamar `context.sync()` antes de ler os valores carregados.

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

Você também pode carregar as propriedades em uma coleção. Cada objeto de coleção na API assíncrona tem uma propriedade que `items` é uma matriz que contém os objetos nessa coleção. Usar `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carrega as propriedades especificadas em cada um desses itens. O exemplo a seguir carrega a propriedade `resolved` em cada objeto `Comment` no objeto `CommentCollection` de uma planilha.

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

### <a name="clientresult"></a>ClientResult

Os métodos na API assíncrona que retornam informações da agenda de trabalho têm um padrão semelhante ao `load` / `sync` paradigma. Por exemplo, `TableCollection.getCount` obtém o número de tabelas da coleção. `getCount` retorna um `ClientResult<number>`, o que significa que a propriedade `value` em [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) retornado é um número. Seu script não pode acessar esse valor até que `context.sync()` seja chamado. Assim como carregar uma propriedade, o `value` é um valor local "vazio" até a `sync` chamada.

O script a seguir obtém o número total de tabelas na pasta de trabalho e registra esse número no console.

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

## <a name="converting-async-scripts-to-the-current-model"></a>Converter scripts assíncronos no modelo atual

O modelo de API atual não usa `load` , `sync` ou um `RequestContext` . Isso torna os scripts muito mais fáceis de gravar e manter. Seu melhor recurso para converter scripts antigos é [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). Lá, você pode pedir ajuda à comunidade com cenários específicos. As diretrizes a seguir devem ajudar a delinear as etapas gerais que você precisará seguir.

1. Crie um novo script e copie o código assíncrono antigo para ele. Certifique-se de não incluir a assinatura `main` do método antigo, usando a `function main(workbook: ExcelScript.Workbook)` atual.

2. Remova todas as `load` chamadas `sync` e. Eles não são mais necessários.

3. Todas as propriedades foram removidas. Agora você acessa esses objetos por meio de métodos e, portanto, precisará alternar essas referências `get` de propriedade para chamadas de `set` método. Por exemplo, em vez de definir a cor de preenchimento de uma célula por meio do acesso a propriedades como este: , agora você usará `mySheet.getRange("A2:C2").format.fill.color = "blue";` métodos como este: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Classes de coleção foram substituídas por matrizes. Os métodos e dessas classes de coleção foram movidos para o objeto que possuía a coleção, portanto, suas referências devem `add` `get` ser atualizadas de acordo. Por exemplo, para obter um gráfico chamado "MyChart" da primeira planilha da pasta de trabalho, use o seguinte código: `workbook.getWorksheets()[0].getChart("MyChart");` . Observe o `[0]` para acessar o primeiro valor do retornado por `Worksheet[]` `getWorksheets()` .

5. Alguns métodos foram renomeados para clareza e adicionados por conveniência. Consulte a referência [da API de Scripts do Office](/javascript/api/office-scripts/overview) para obter mais detalhes.

## <a name="office-scripts-async-api-reference-documentation"></a>Documentação de referência da API assíncrona de Scripts do Office

As APIs assíncronas são equivalentes às usadas em Complementos do Office. A documentação de referência é encontrada na seção Excel da referência da API JavaScript de [Complementos do Office.](/javascript/api/excel?view=excel-js-online&preserve-view=true)
