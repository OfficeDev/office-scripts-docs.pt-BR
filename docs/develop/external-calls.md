---
title: Chamada de API externa nos scripts do Office
description: Suporte e orientação para fazer chamadas de API externas em um Script do Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570308"
---
# <a name="external-api-call-support-in-office-scripts"></a>Chamada de API externa nos scripts do Office

Os autores de script não devem esperar comportamento consistente ao usar [APIs externas](https://developer.mozilla.org/docs/Web/API) durante a fase de visualização da plataforma. Dessa forma, não confie em APIs externas para cenários críticos de script.

As chamadas para APIs externas só podem ser feitas por meio do aplicativo Excel, não por meio do Power Automate [em circunstâncias normais.](#external-calls-from-power-automate)

> [!CAUTION]
> Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis. O administrador pode estabelecer proteção de firewall contra essas chamadas.

## <a name="working-with-fetch"></a>Trabalhando com `fetch`

A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos. É uma `async` API, portanto, você precisará ajustar a `main` assinatura do seu script. Faça a `main` função e faça com que ela retorne um `async` `Promise<void>` . Você também deve se certificar `await` da `fetch` chamada e `json` recuperação. Isso garante que essas operações terminem antes do script terminar.

O script a seguir `fetch` usa para recuperar dados JSON do servidor de teste na URL determinada.

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

O cenário de exemplo de Scripts do Office: gráfico de dados de nível de água do [NOAA](../resources/scenarios/noaa-data-fetch.md) demonstra o comando fetch que está sendo usado para recuperar registros do banco de dados Nacional oceanico e atmosférico da administração.

## <a name="external-calls-from-power-automate"></a>Chamadas externas do Power Automate

Qualquer chamada de API externa falha quando um script é executado com o Power Automate. Essa é uma diferença comportamental entre executar um script por meio do cliente do Excel e por meio do Power Automate. Certifique-se de verificar seus scripts para essas referências antes de ad construi-las em um fluxo.

> [!WARNING]
> As chamadas externas feitas por meio do conector do Power Automate [Excel Online](/connectors/excelonlinebusiness) falham para ajudar a manter políticas de prevenção contra perda de dados existentes. No entanto, os scripts executados por meio do Power Automate são feitos fora da sua organização e fora dos firewalls da sua organização. Para obter proteção adicional contra usuários mal-intencionados nesse ambiente externo, o administrador pode controlar o uso de Scripts do Office. O administrador pode desabilitar o conector do Excel Online no Power Automate ou desativar os Scripts do Office para Excel na Web por meio dos controles de administrador [do Office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Confira também

- [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)
- [Cenário de exemplo de Scripts do Office: gráfico de dados de nível de água do NOAA](../resources/scenarios/noaa-data-fetch.md)
