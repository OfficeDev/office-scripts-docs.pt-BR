---
title: Chamada de API externa nos scripts do Office
description: Suporte e diretrizes para fazer chamadas de API externas em um Office Script.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b847400893184533c250ab99b640563ff0cbdb3e
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088040"
---
# <a name="external-api-call-support-in-office-scripts"></a>Chamada de API externa nos scripts do Office

Os scripts dão suporte a chamadas para serviços externos. Use esses serviços para fornecer dados e outras informações à sua pasta de trabalho.

> [!CAUTION]
> Chamadas externas podem fazer com que dados confidenciais sejam expostos a pontos de extremidade indesejáveis. O administrador pode estabelecer proteção de firewall contra essas chamadas.

> [!IMPORTANT]
> As chamadas para APIs externas só podem ser feitas por meio do aplicativo Excel, não por meio Power Automate [em circunstâncias normais](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Configurar o script para chamadas externas

Chamadas externas [são assíncronas e](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) exigem que o script seja marcado como `async`. Adicione o `async` prefixo à função `main` e retorne um `Promise`, conforme mostrado aqui:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Scripts que retornam outras informações podem retornar `Promise` um desse tipo. Por exemplo, se o script precisar retornar um `Employee` objeto, a assinatura de retorno será `: Promise <Employee>`

Você precisará aprender as interfaces do serviço externo para fazer chamadas a esse serviço. Se você estiver usando `fetch` ou [APIs REST](https://wikipedia.org/wiki/Representational_state_transfer), precisará determinar a estrutura JSON dos dados retornados. Para entrada e saída do script, considere fazer uma correspondência `interface` com as estruturas JSON necessárias. Isso dá ao script mais segurança de tipo. Você pode ver um exemplo disso em Usar [fetch de Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitações com chamadas externas de Office Scripts

* Não há como entrar ou usar o tipo OAuth2 de fluxos de autenticação. Todas as chaves e credenciais precisam ser codificadas (ou lidas de outra fonte).
* Não há nenhuma infraestrutura para armazenar credenciais e chaves de API. Isso precisará ser gerenciado pelo usuário.
* Não há suporte `localStorage`para `sessionStorage` cookies de documento e objetos.
* Chamadas externas podem fazer com que dados confidenciais sejam expostos a pontos de extremidade indesejáveis ou que dados externos sejam trazidos para pastas de trabalho internas. O administrador pode estabelecer proteção de firewall contra essas chamadas. Verifique com as políticas locais antes de depender de chamadas externas.
* Verifique a quantidade de taxa de transferência de dados antes de assumir uma dependência. Por exemplo, extrair todo o conjunto de dados externo pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em partes.

## <a name="retrieve-information-with-fetch"></a>Recuperar informações com `fetch`

A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos. É uma `async` API, portanto, você precisa ajustar a assinatura `main` do script. Faça a `main` função `async`. Você também deve se certificar da `await` `fetch` chamada `json` e da recuperação. Isso garante que essas operações são concluídas antes que o script termine.

Todos os dados JSON recuperados devem `fetch` corresponder a uma interface definida no script. O valor retornado deve ser atribuído a um tipo específico porque [Office scripts não dão suporte ao `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts). Você deve consultar a documentação do serviço para ver quais são os nomes e tipos das propriedades retornadas. Em seguida, adicione a interface ou as interfaces correspondentes ao script.

O script a seguir usa `fetch` para recuperar dados JSON do servidor de teste na URL fornecida. Observe a `JSONData` interface para armazenar os dados como um tipo correspondente.

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a>Outros `fetch` exemplos

* O [exemplo Usar chamadas de busca externas Office scripts](../resources/samples/external-fetch-calls.md) mostra como obter informações básicas sobre os repositórios de GitHub de um usuário.
* O cenário de exemplo Office Scripts: Graph dados em nível de água do [NOAA](../resources/scenarios/noaa-data-fetch.md) demonstra o comando fetch que está sendo usado para recuperar registros do banco de dados De Marés e Correntes da Administração Oceânico e Atmosférica Nacional.

## <a name="external-calls-from-power-automate"></a>Chamadas externas de Power Automate

Qualquer chamada de API externa falha quando um script é executado com Power Automate. Essa é uma diferença comportamental entre executar um script por meio Excel aplicativo e por meio Power Automate. Certifique-se de verificar seus scripts para essas referências antes de compilá-las em um fluxo.

Você precisará usar [HTTP](/connectors/webcontents/) com Azure AD ou outras ações equivalentes para efetuar pull de dados ou efetuar push para um serviço externo.

> [!WARNING]
> As chamadas externas feitas por meio Power Automate [Excel online](/connectors/excelonlinebusiness) falham para ajudar a manter as políticas de prevenção contra perda de dados existentes. No entanto, os scripts executados Power Automate são feitos fora da sua organização e fora dos firewalls da sua organização. Para obter proteção adicional contra usuários mal-intencionados nesse ambiente externo, o administrador pode controlar o uso de scripts Office usuário. O administrador pode desabilitar o conector Excel Online no Power Automate ou desativar scripts do Office para Excel na Web por meio dos controles de administrador do [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Confira também

* [Usar JSON para passar dados de e para Office Scripts](use-json.md)
* [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)
* [Usar chamadas de busca externa em Scripts do Office](../resources/samples/external-fetch-calls.md)
* [Office de exemplo de scripts: Graph dados de nível de água do NOAA](../resources/scenarios/noaa-data-fetch.md)
