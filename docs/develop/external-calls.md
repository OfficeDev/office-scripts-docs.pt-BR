---
title: Chamada de API externa nos scripts do Office
description: Suporte e orientação para fazer chamadas de API externas em um Script Office.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545079"
---
# <a name="external-api-call-support-in-office-scripts"></a>Chamada de API externa nos scripts do Office

Os autores de script não devem esperar um comportamento consistente ao usar [APIs externas](https://developer.mozilla.org/docs/Web/API) durante a fase de visualização da plataforma. Como tal, não conte com APIs externas para cenários críticos de script.

As chamadas para APIs externas só podem ser feitas através do aplicativo Excel, não através de Power Automate [em circunstâncias normais](#external-calls-from-power-automate).

> [!CAUTION]
> Chamadas externas podem resultar em dados confidenciais expostos a pontos finais indesejáveis. Seu administrador pode estabelecer proteção de firewall contra tais chamadas.

## <a name="configure-your-script-for-external-calls"></a>Configure seu script para chamadas externas

Chamadas [externas são assíncrodas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) e exigem que seu script seja marcado como `async` . Adicione o `async` prefixo à sua `main` função e peça para devolvê-lo, como mostrado `Promise` aqui:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Scripts que retornam outras informações podem retornar um `Promise` desse tipo. Por exemplo, se o seu script precisar retornar um `Employee` objeto, a assinatura de retorno seria `: Promise <Employee>`

Você precisará aprender as interfaces do serviço externo para fazer chamadas para esse serviço. Se você estiver usando `fetch` ou [REST APIs,](https://wikipedia.org/wiki/Representational_state_transfer)você precisa determinar a estrutura JSON dos dados retornados. Para entrada e saída do seu script, considere fazer um `interface` para corresponder às estruturas JSON necessárias. Isso dá ao script mais segurança do tipo. Você pode ver um exemplo disso em [Usar buscar de Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitações com chamadas externas de scripts Office

* Não há como fazer login ou usar fluxos de autenticação do tipo OAuth2. Todas as chaves e credenciais devem ser codificadas (ou lidas de outra fonte).
* Não há infraestrutura para armazenar credenciais e chaves de API. Isso terá que ser gerenciado pelo usuário.
* Cookies de documentos `localStorage` e `sessionStorage` objetos não são suportados. 
* Chamadas externas podem resultar em dados confidenciais expostos a pontos finais indesejáveis ou dados externos a serem trazidos para pastas de trabalho internas. Seu administrador pode estabelecer proteção de firewall contra tais chamadas. Certifique-se de verificar com as políticas locais antes de depender de chamadas externas.
* Certifique-se de verificar a quantidade de throughput de dados antes de tomar uma dependência. Por exemplo, puxar todo o conjunto de dados externo pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em pedaços.

## <a name="retrieve-information-with-fetch"></a>Recuperar informações com `fetch`

A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos. É uma `async` API, então você precisa ajustar a `main` assinatura do seu script. Faça a `main` função e faça com que `async` devolva um `Promise<void>` . Você também deve ter certeza `await` da `fetch` chamada e `json` recuperação. Isso garante que essas operações são concluídas antes do fim do script.

Qualquer dado JSON recuperado `fetch` deve corresponder a uma interface definida no script. O valor retornado deve ser atribuído a um tipo específico porque [Office Scripts não suportam o `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts). Você deve consultar a documentação do seu serviço para ver quais são os nomes e tipos das propriedades devolvidas. Em seguida, adicione a interface ou interfaces correspondentes ao seu script.

O script a seguir usa `fetch` para recuperar dados JSON do servidor de teste na URL dada. Observe a `JSONData` interface para armazenar os dados como um tipo de correspondência.

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
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

### <a name="other-fetch-samples"></a>Outras `fetch` amostras

* A amostra [de busca externa use chamadas de Office Scripts](../resources/samples/external-fetch-calls.md) mostra como obter informações básicas sobre os repositórios de GitHub do usuário.
* O [cenário amostral de Office Scripts: Graph dados de nível de água da NOAA](../resources/scenarios/noaa-data-fetch.md) demonstram o comando de busca que está sendo usado para recuperar registros do banco de dados de Marés e Correntes da Administração Oceânica Nacional Oceânica e Atmosférica.

## <a name="external-calls-from-power-automate"></a>Chamadas externas de Power Automate

Qualquer chamada de API externa falha quando um script é executado com Power Automate. Esta é uma diferença comportamental entre executar um script através do aplicativo Excel e através de Power Automate. Certifique-se de verificar seus scripts para obter tais referências antes de construí-las em um fluxo.

Você terá que usar [HTTP com o Azure AD](/connectors/webcontents/) ou outras ações equivalentes para extrair dados ou empurrá-los para um serviço externo.

> [!WARNING]
> Chamadas externas feitas através do Power Automate [Excel conector Online](/connectors/excelonlinebusiness) falham para ajudar a manter as políticas de prevenção de perda de dados existentes. No entanto, os scripts que são executados através Power Automate são feitos fora da sua organização, e fora dos firewalls da sua organização. Para proteção adicional contra usuários mal-intencionados neste ambiente externo, o administrador pode controlar o uso de scripts Office. O administrador pode desativar o conector online Excel em Power Automate ou desativar Office Scripts para Excel na Web através dos [controles de administrador de scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Confira também

* [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)
* [Usar chamadas de busca externa em Scripts do Office](../resources/samples/external-fetch-calls.md)
* [Office Cenário da amostra de scripts: Graph dados de nível de água da NOAA](../resources/scenarios/noaa-data-fetch.md)
