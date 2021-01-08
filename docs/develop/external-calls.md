---
title: Chamada de API externa nos scripts do Office
description: Suporte e diretrizes para fazer chamadas de API externa em um Script do Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784141"
---
# <a name="external-api-call-support-in-office-scripts"></a>Chamada de API externa nos scripts do Office

Os autores de scripts não devem esperar um comportamento consistente ao usar [APIs externas](https://developer.mozilla.org/docs/Web/API) durante a fase de visualização da plataforma. Dessa forma, não confie em APIs externas para cenários de script críticos.

As chamadas para APIs externas só podem ser feitas por meio do aplicativo Excel, não por meio do Power Automate [em circunstâncias normais.](#external-calls-from-power-automate)

> [!CAUTION]
> Chamadas externas podem resultar na exposição de dados confidenciais a pontos de extremidade indesejáveis. Seu administrador pode estabelecer proteção de firewall contra essas chamadas.

## <a name="working-with-fetch"></a>Trabalhando com `fetch`

A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos. É uma `async` API, portanto, você precisará ajustar a `main` assinatura do seu script. Make the `main` function and have it return a `async` `Promise<void>` . Você também deve ter certeza `await` da `fetch` chamada e `json` recuperação. Isso garante que essas operações terminem antes do script terminar.

O script a seguir `fetch` usa para recuperar dados JSON do servidor de teste na URL determinada.

```typescript
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

O cenário de exemplo de Scripts do Office: os dados gráficos no nível d'água do [NOAA](../resources/scenarios/noaa-data-fetch.md) demonstram o comando de busca que está sendo usado para recuperar registros do banco de dados National Oceanic and Administrations Currents.

## <a name="external-calls-from-power-automate"></a>Chamadas externas do Power Automate

Todas as chamadas de API externa falham quando um script é executado com o Power Automate. Essa é uma diferença comportamental entre executar um script por meio do cliente do Excel e por meio do Power Automate. Certifique-se de verificar seus scripts em busca dessas referências antes de building-los em um fluxo.

> [!WARNING]
> As chamadas externas feitas por meio do conector do Power Automate [Excel Online](/connectors/excelonlinebusiness) falham para ajudar a preservar as políticas de prevenção contra perda de dados existentes. No entanto, os scripts executados por meio do Power Automate são feitos fora da sua organização e fora dos firewalls da sua organização. Para obter proteção adicional contra usuários mal-intencionados nesse ambiente externo, o administrador pode controlar o uso de scripts do Office. O administrador pode desabilitar o conector do Excel Online no Power Automate ou desativar os Scripts do Office para Excel na Web por meio dos controles de administrador [de Scripts do Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Confira também

- [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)
- [Cenário de exemplo de scripts do Office: dados em nível de água do Graph do NOAA](../resources/scenarios/noaa-data-fetch.md)
