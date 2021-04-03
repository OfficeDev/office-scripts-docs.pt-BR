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
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="40adb-103">Chamada de API externa nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="40adb-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="40adb-104">Os autores de script não devem esperar comportamento consistente ao usar [APIs externas](https://developer.mozilla.org/docs/Web/API) durante a fase de visualização da plataforma.</span><span class="sxs-lookup"><span data-stu-id="40adb-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="40adb-105">Dessa forma, não confie em APIs externas para cenários críticos de script.</span><span class="sxs-lookup"><span data-stu-id="40adb-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="40adb-106">As chamadas para APIs externas só podem ser feitas por meio do aplicativo Excel, não por meio do Power Automate [em circunstâncias normais.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="40adb-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="40adb-107">Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis.</span><span class="sxs-lookup"><span data-stu-id="40adb-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="40adb-108">O administrador pode estabelecer proteção de firewall contra essas chamadas.</span><span class="sxs-lookup"><span data-stu-id="40adb-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="40adb-109">Trabalhando com `fetch`</span><span class="sxs-lookup"><span data-stu-id="40adb-109">Working with `fetch`</span></span>

<span data-ttu-id="40adb-110">A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos.</span><span class="sxs-lookup"><span data-stu-id="40adb-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="40adb-111">É uma `async` API, portanto, você precisará ajustar a `main` assinatura do seu script.</span><span class="sxs-lookup"><span data-stu-id="40adb-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="40adb-112">Faça a `main` função e faça com que ela retorne um `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="40adb-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="40adb-113">Você também deve se certificar `await` da `fetch` chamada e `json` recuperação.</span><span class="sxs-lookup"><span data-stu-id="40adb-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="40adb-114">Isso garante que essas operações terminem antes do script terminar.</span><span class="sxs-lookup"><span data-stu-id="40adb-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="40adb-115">O script a seguir `fetch` usa para recuperar dados JSON do servidor de teste na URL determinada.</span><span class="sxs-lookup"><span data-stu-id="40adb-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

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

<span data-ttu-id="40adb-116">O cenário de exemplo de Scripts do Office: gráfico de dados de nível de água do [NOAA](../resources/scenarios/noaa-data-fetch.md) demonstra o comando fetch que está sendo usado para recuperar registros do banco de dados Nacional oceanico e atmosférico da administração.</span><span class="sxs-lookup"><span data-stu-id="40adb-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="40adb-117">Chamadas externas do Power Automate</span><span class="sxs-lookup"><span data-stu-id="40adb-117">External calls from Power Automate</span></span>

<span data-ttu-id="40adb-118">Qualquer chamada de API externa falha quando um script é executado com o Power Automate.</span><span class="sxs-lookup"><span data-stu-id="40adb-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="40adb-119">Essa é uma diferença comportamental entre executar um script por meio do cliente do Excel e por meio do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="40adb-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="40adb-120">Certifique-se de verificar seus scripts para essas referências antes de ad construi-las em um fluxo.</span><span class="sxs-lookup"><span data-stu-id="40adb-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="40adb-121">As chamadas externas feitas por meio do conector do Power Automate [Excel Online](/connectors/excelonlinebusiness) falham para ajudar a manter políticas de prevenção contra perda de dados existentes.</span><span class="sxs-lookup"><span data-stu-id="40adb-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="40adb-122">No entanto, os scripts executados por meio do Power Automate são feitos fora da sua organização e fora dos firewalls da sua organização.</span><span class="sxs-lookup"><span data-stu-id="40adb-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="40adb-123">Para obter proteção adicional contra usuários mal-intencionados nesse ambiente externo, o administrador pode controlar o uso de Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="40adb-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="40adb-124">O administrador pode desabilitar o conector do Excel Online no Power Automate ou desativar os Scripts do Office para Excel na Web por meio dos controles de administrador [do Office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="40adb-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="40adb-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="40adb-125">See also</span></span>

- [<span data-ttu-id="40adb-126">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="40adb-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="40adb-127">Cenário de exemplo de Scripts do Office: gráfico de dados de nível de água do NOAA</span><span class="sxs-lookup"><span data-stu-id="40adb-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
