---
title: Suporte scripts de Office mais antigos que usam as APIs assíncia
description: Um primer no Office Scripts Async APIs e como usar o padrão de carga/sincronização para scripts mais antigos.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 80a1c0dec5393d8882ddb37eea5f81ef23b1ebb1
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545072"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Suporte scripts de Office mais antigos que usam as APIs assíncia

Este artigo ensina como manter e atualizar scripts que usam as APIs assídas do modelo mais antigo. Essas APIs têm a mesma funcionalidade principal que as APIs de Office scripts agora padrão e síncronos, mas exigem que seu script controle a sincronização de dados entre o script e a pasta de trabalho.

> [!IMPORTANT]
> O modelo async só pode ser usado com scripts criados antes da implementação do modelo atual de [API](scripting-fundamentals.md). Os scripts estão permanentemente bloqueados ao modelo de API que eles têm após a criação. Isso também significa que se você quiser converter um script antigo para o novo modelo, você deve criar um novo script. Recomendamos que você atualize seus scripts antigos para o novo modelo ao fazer alterações, já que o modelo atual é mais fácil de usar. Os [scripts de conversão de assínc para a](#convert-async-scripts-to-the-current-model) seção modelo atual tem conselhos sobre como fazer essa transição.

## <a name="older-main-function-signature"></a>Assinatura de função mais `main` antiga

Scripts que usam as APIs assíncia têm uma `main` função diferente. É uma `async` função que tem como `Excel.RequestContext` primeiro parâmetro.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Contexto

A função `main` aceita um parâmetro `Excel.RequestContext`, chamado `context`. Imagine `context` como a ponte entre o seu script e a pasta de trabalho. Seu script acessa a pasta de trabalho com o objeto `context` e usa esse `context` para troca de dados.

O objeto `context` é necessário porque o script e o Excel estão sendo executados em processos e locais diferentes. O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem. O objeto `context` gerencia essas transações.

## <a name="sync-and-load"></a>Sincronizar e carregar

Como o seu script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois levará algum tempo. Na API assíndia, os comandos são enfileiados até que o script chame explicitamente a `sync` operação para sincronizar o script e a pasta de trabalho. Seu script pode trabalhar de forma independente até que precise executar uma das seguintes ações:

- Leia os dados da pasta de trabalho (seguindo uma `load` operação ou método que retorne um [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Gravar dados na pasta de trabalho (geralmente porque o script terminou).

A imagem a seguir mostra um exemplo de fluxo de controle entre o script e a pasta de trabalho:

:::image type="content" source="../images/load-sync.png" alt-text="Um diagrama mostrando operações de leitura e gravação indo para a pasta de trabalho do script":::

### <a name="sync"></a>Sincronizar

Sempre que o seu script async precisar ler dados ou gravar dados para a pasta de trabalho, ligue para o `RequestContext.sync` método como mostrado no trecho de código a seguir:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` é chamado implicitamente quando um script termina.

Após a conclusão da operação `sync`, a pasta de trabalho será atualizada para refletir as operações de gravação especificados por esse script. Uma operação de gravação está definindo qualquer propriedade em um objeto Excel (por exemplo) `range.format.fill.color = "red"` ou chamando um método que altera uma propriedade (por exemplo, `range.format.autoFitColumns()` ). A `sync` operação também lê todos os valores da pasta de trabalho que o script solicitou usando uma `load` operação ou um método que retorna a `ClientResult` (conforme discutido nas próximas seções).

A sincronização do seu script com a pasta de trabalho pode demorar, dependendo da sua rede. Minimize o número de `sync` chamadas para ajudar seu script a funcionar rapidamente. Caso contrário, as APIs assínias não são mais rápidas das APIs padrão e síncronia.

### <a name="load"></a>Carregar

Um script async deve carregar dados da pasta de trabalho antes de lê-los. No entanto, carregar dados de toda a pasta de trabalho reduziria muito a velocidade do script. O `load` método permite que seu script diga especificamente quais dados devem ser recuperados da pasta de trabalho.

O método `load` está disponível em todos os objetos do Excel. Seu script deve carregar as propriedades de um objeto para poder lê-lo. Não fazer isso resulta em um erro.

Os exemplos a seguir usam um objeto `Range` para mostrar as três maneiras de usar o método `load` para carregar dados.

|Finalidade |Comando de exemplo | Efeito |
|:--|:--|:--|
|Carregar uma propriedade |`myRange.load("values");` | Carrega uma única propriedade, neste caso, a matriz bidimensional de valores nesse intervalo. |
|Carregar várias propriedades |`myRange.load("values, rowCount, columnCount");`| Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e de colunas. |
|Carregar tudo | `myRange.load();`|Carrega todas as propriedades no intervalo. Esta não é uma solução recomendada, pois irá retardar seu script obtendo dados desnecessários. Use-o somente enquanto testa seu script ou se você precisar de todas as propriedades do objeto. |

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

Você também pode carregar as propriedades em uma coleção. Cada objeto de coleta na API assíncia tem uma `items` propriedade que é uma matriz contendo os objetos nessa coleção. Usar `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carrega as propriedades especificadas em cada um desses itens. O exemplo a seguir carrega a propriedade `resolved` em cada objeto `Comment` no objeto `CommentCollection` de uma planilha.

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

Os métodos na API assincron que retornam informações da pasta de trabalho têm um padrão semelhante ao `load` / `sync` paradigma. Por exemplo, `TableCollection.getCount` obtém o número de tabelas da coleção. `getCount` retorna um `ClientResult<number>`, o que significa que a propriedade `value` em [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) retornado é um número. Seu script não pode acessar esse valor até que `context.sync()` seja chamado. Assim como carregar uma propriedade, o `value` é um valor local "vazio" até a `sync` chamada.

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

## <a name="convert-async-scripts-to-the-current-model"></a>Converta scripts async para o modelo atual

O modelo de API atual não `load` `sync` usa, ou um `RequestContext` . Isso torna os scripts muito mais fáceis de escrever e manter. Seu melhor recurso para converter scripts antigos é [o Microsoft Q&A](/answers/topics/office-scripts-dev.html). Lá, você pode pedir ajuda à comunidade com cenários específicos. A orientação a seguir deve ajudar a delinear os passos gerais que você precisará tomar.

1. Crie um novo script e copie o antigo código assínco nele. Certifique-se de não incluir a assinatura do `main` método antigo, usando a corrente `function main(workbook: ExcelScript.Workbook)` em vez disso.

2. Remova todas as `load` `sync` chamadas. Eles não são mais necessários.

3. Todas as propriedades foram removidas. Agora você acessa esses objetos através `get` e `set` métodos, então você precisará mudar essas referências de propriedade para chamadas de método. Por exemplo, em vez de definir a cor de preenchimento de uma célula através do acesso à propriedade como este: `mySheet.getRange("A2:C2").format.fill.color = "blue";` , agora você usará métodos como este: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. As aulas de coleta foram substituídas por matrizes. Os `add` `get` métodos dessas aulas de coleta foram movidos para o objeto que possuía a coleção, de modo que suas referências devem ser atualizadas em conformidade. Por exemplo, para obter um gráfico chamado "MyChart" da primeira planilha na pasta de trabalho, use o seguinte código: `workbook.getWorksheets()[0].getChart("MyChart");` . Observe `[0]` o acesso ao primeiro valor do `Worksheet[]` retornado por `getWorksheets()` .

5. Alguns métodos foram renomeados para clareza e adicionados por conveniência. Consulte a [referência de API Office Scripts](/javascript/api/office-scripts/overview) para obter mais detalhes.

## <a name="office-scripts-async-api-reference-documentation"></a>Office Scripts async API documentação de referência

As APIs assíncas são equivalentes às usadas em Office Add-ins. A documentação de referência é encontrada [na seção Excel da referência de API JavaScript Office Add-ins](/javascript/api/excel?view=excel-js-online&preserve-view=true).
