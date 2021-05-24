---
title: Chamada de API externa nos scripts do Office
description: Suporte e orientação para fazer chamadas de API externas em Office Script.
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

Os autores de script não devem esperar comportamento consistente ao usar [APIs externas](https://developer.mozilla.org/docs/Web/API) durante a fase de visualização da plataforma. Dessa forma, não confie em APIs externas para cenários críticos de script.

As chamadas para APIs externas só podem ser feitas por meio do aplicativo Excel, não por meio Power Automate [em circunstâncias normais.](#external-calls-from-power-automate)

> [!CAUTION]
> Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis. O administrador pode estabelecer proteção de firewall contra essas chamadas.

## <a name="configure-your-script-for-external-calls"></a>Configurar seu script para chamadas externas

Chamadas externas [são assíncronas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) e exigem que seu script seja marcado como `async` . Adicione o `async` prefixo à `main` sua função e retorne um , conforme mostrado `Promise` aqui:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Scripts que retornam outras informações podem retornar `Promise` um desse tipo. Por exemplo, se o script precisar retornar um `Employee` objeto, a assinatura de retorno será `: Promise <Employee>`

Você precisará aprender as interfaces do serviço externo para fazer chamadas para esse serviço. Se você estiver usando ou APIs REST , será necessário determinar `fetch` a estrutura JSON dos dados retornados. [](https://wikipedia.org/wiki/Representational_state_transfer) Para entrada e saída do script, considere fazer uma para corresponder às `interface` estruturas JSON necessárias. Isso oferece ao script mais segurança de tipo. Você pode ver um exemplo disso em [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitações com chamadas externas de Office Scripts

* Não há nenhuma maneira de entrar ou usar o tipo OAuth2 de fluxos de autenticação. Todas as chaves e credenciais devem ser codificadas (ou leitura de outra fonte).
* Não há infraestrutura para armazenar credenciais e chaves da API. Isso terá que ser gerenciado pelo usuário.
* Cookies de documento `localStorage` e objetos não são `sessionStorage` suportados. 
* Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis ou dados externos a serem trazidos para as guias de trabalho internas. O administrador pode estabelecer proteção de firewall contra essas chamadas. Verifique as políticas locais antes de confiar em chamadas externas.
* Verifique a quantidade de transferência de dados antes de assumir uma dependência. Por exemplo, retirar todo o conjuntos de dados externos pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em partes.

## <a name="retrieve-information-with-fetch"></a>Recuperar informações com `fetch`

A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos. É uma `async` API, portanto, você precisa ajustar a `main` assinatura do seu script. Faça a `main` função e faça com que ela retorne um `async` `Promise<void>` . Você também deve se certificar `await` da `fetch` chamada e `json` recuperação. Isso garante que essas operações terminem antes do script terminar.

Todos os dados JSON recuperados por `fetch` devem corresponder a uma interface definida no script. O valor retornado deve ser atribuído a um tipo específico porque [Office scripts não suportam o `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts). Você deve consultar a documentação do seu serviço para ver quais são os nomes e tipos das propriedades retornadas. Em seguida, adicione a interface ou interfaces correspondentes ao seu script.

O script a seguir `fetch` usa para recuperar dados JSON do servidor de teste na URL determinada. Observe a `JSONData` interface para armazenar os dados como um tipo correspondente.

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

### <a name="other-fetch-samples"></a>Outros `fetch` exemplos

* O exemplo Usar chamadas de busca externa [Office scripts](../resources/samples/external-fetch-calls.md) mostra como obter informações básicas sobre os repositórios de GitHub do usuário.
* O Office de exemplo scripts: Graph dados de nível de água do [NOAA](../resources/scenarios/noaa-data-fetch.md) demonstra o comando fetch que está sendo usado para recuperar registros do banco de dados De onda e currents da Administração Nacional Oceânica e Atmosférico.

## <a name="external-calls-from-power-automate"></a>Chamadas externas de Power Automate

Qualquer chamada de API externa falha quando um script é executado com Power Automate. Essa é uma diferença comportamental entre executar um script por meio do aplicativo Excel e por meio Power Automate. Certifique-se de verificar seus scripts para essas referências antes de ad construi-las em um fluxo.

Você terá que usar HTTP com o [Azure AD](/connectors/webcontents/) ou outras ações equivalentes para puxar dados ou pressioná-los para um serviço externo.

> [!WARNING]
> As chamadas externas feitas por meio do conector Power Automate [Excel Online](/connectors/excelonlinebusiness) falham para ajudar a manter políticas de prevenção contra perda de dados existentes. No entanto, os scripts que são executados por Power Automate são feitos fora da sua organização e fora dos firewalls da sua organização. Para obter proteção adicional contra usuários mal-intencionados nesse ambiente externo, o administrador pode controlar o uso de Office Scripts. O administrador pode desabilitar o conector Excel Online no Power Automate ou desativar scripts do Office para Excel na Web por meio dos controles de administrador Office [Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Confira também

* [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)
* [Usar chamadas de busca externa em Scripts do Office](../resources/samples/external-fetch-calls.md)
* [Office Cenário de exemplo de scripts: Graph dados de nível de água do NOAA](../resources/scenarios/noaa-data-fetch.md)
