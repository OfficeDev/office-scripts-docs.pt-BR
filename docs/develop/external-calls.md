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
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="a8d41-103">Chamada de API externa nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="a8d41-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="a8d41-104">Os autores de script não devem esperar um comportamento consistente ao usar [APIs externas](https://developer.mozilla.org/docs/Web/API) durante a fase de visualização da plataforma.</span><span class="sxs-lookup"><span data-stu-id="a8d41-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="a8d41-105">Como tal, não conte com APIs externas para cenários críticos de script.</span><span class="sxs-lookup"><span data-stu-id="a8d41-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="a8d41-106">As chamadas para APIs externas só podem ser feitas através do aplicativo Excel, não através de Power Automate [em circunstâncias normais](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="a8d41-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="a8d41-107">Chamadas externas podem resultar em dados confidenciais expostos a pontos finais indesejáveis.</span><span class="sxs-lookup"><span data-stu-id="a8d41-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="a8d41-108">Seu administrador pode estabelecer proteção de firewall contra tais chamadas.</span><span class="sxs-lookup"><span data-stu-id="a8d41-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="a8d41-109">Configure seu script para chamadas externas</span><span class="sxs-lookup"><span data-stu-id="a8d41-109">Configure your script for external calls</span></span>

<span data-ttu-id="a8d41-110">Chamadas [externas são assíncrodas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) e exigem que seu script seja marcado como `async` .</span><span class="sxs-lookup"><span data-stu-id="a8d41-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="a8d41-111">Adicione o `async` prefixo à sua `main` função e peça para devolvê-lo, como mostrado `Promise` aqui:</span><span class="sxs-lookup"><span data-stu-id="a8d41-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="a8d41-112">Scripts que retornam outras informações podem retornar um `Promise` desse tipo.</span><span class="sxs-lookup"><span data-stu-id="a8d41-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="a8d41-113">Por exemplo, se o seu script precisar retornar um `Employee` objeto, a assinatura de retorno seria `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="a8d41-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="a8d41-114">Você precisará aprender as interfaces do serviço externo para fazer chamadas para esse serviço.</span><span class="sxs-lookup"><span data-stu-id="a8d41-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="a8d41-115">Se você estiver usando `fetch` ou [REST APIs,](https://wikipedia.org/wiki/Representational_state_transfer)você precisa determinar a estrutura JSON dos dados retornados.</span><span class="sxs-lookup"><span data-stu-id="a8d41-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="a8d41-116">Para entrada e saída do seu script, considere fazer um `interface` para corresponder às estruturas JSON necessárias.</span><span class="sxs-lookup"><span data-stu-id="a8d41-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="a8d41-117">Isso dá ao script mais segurança do tipo.</span><span class="sxs-lookup"><span data-stu-id="a8d41-117">This gives the script more type safety.</span></span> <span data-ttu-id="a8d41-118">Você pode ver um exemplo disso em [Usar buscar de Office Scripts](../resources/samples/external-fetch-calls.md).</span><span class="sxs-lookup"><span data-stu-id="a8d41-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="a8d41-119">Limitações com chamadas externas de scripts Office</span><span class="sxs-lookup"><span data-stu-id="a8d41-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="a8d41-120">Não há como fazer login ou usar fluxos de autenticação do tipo OAuth2.</span><span class="sxs-lookup"><span data-stu-id="a8d41-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="a8d41-121">Todas as chaves e credenciais devem ser codificadas (ou lidas de outra fonte).</span><span class="sxs-lookup"><span data-stu-id="a8d41-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="a8d41-122">Não há infraestrutura para armazenar credenciais e chaves de API.</span><span class="sxs-lookup"><span data-stu-id="a8d41-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="a8d41-123">Isso terá que ser gerenciado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="a8d41-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="a8d41-124">Cookies de documentos `localStorage` e `sessionStorage` objetos não são suportados.</span><span class="sxs-lookup"><span data-stu-id="a8d41-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span> 
* <span data-ttu-id="a8d41-125">Chamadas externas podem resultar em dados confidenciais expostos a pontos finais indesejáveis ou dados externos a serem trazidos para pastas de trabalho internas.</span><span class="sxs-lookup"><span data-stu-id="a8d41-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="a8d41-126">Seu administrador pode estabelecer proteção de firewall contra tais chamadas.</span><span class="sxs-lookup"><span data-stu-id="a8d41-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="a8d41-127">Certifique-se de verificar com as políticas locais antes de depender de chamadas externas.</span><span class="sxs-lookup"><span data-stu-id="a8d41-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="a8d41-128">Certifique-se de verificar a quantidade de throughput de dados antes de tomar uma dependência.</span><span class="sxs-lookup"><span data-stu-id="a8d41-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="a8d41-129">Por exemplo, puxar todo o conjunto de dados externo pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em pedaços.</span><span class="sxs-lookup"><span data-stu-id="a8d41-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="a8d41-130">Recuperar informações com `fetch`</span><span class="sxs-lookup"><span data-stu-id="a8d41-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="a8d41-131">A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos.</span><span class="sxs-lookup"><span data-stu-id="a8d41-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="a8d41-132">É uma `async` API, então você precisa ajustar a `main` assinatura do seu script.</span><span class="sxs-lookup"><span data-stu-id="a8d41-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="a8d41-133">Faça a `main` função e faça com que `async` devolva um `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="a8d41-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="a8d41-134">Você também deve ter certeza `await` da `fetch` chamada e `json` recuperação.</span><span class="sxs-lookup"><span data-stu-id="a8d41-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="a8d41-135">Isso garante que essas operações são concluídas antes do fim do script.</span><span class="sxs-lookup"><span data-stu-id="a8d41-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="a8d41-136">Qualquer dado JSON recuperado `fetch` deve corresponder a uma interface definida no script.</span><span class="sxs-lookup"><span data-stu-id="a8d41-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="a8d41-137">O valor retornado deve ser atribuído a um tipo específico porque [Office Scripts não suportam o `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts).</span><span class="sxs-lookup"><span data-stu-id="a8d41-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="a8d41-138">Você deve consultar a documentação do seu serviço para ver quais são os nomes e tipos das propriedades devolvidas.</span><span class="sxs-lookup"><span data-stu-id="a8d41-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="a8d41-139">Em seguida, adicione a interface ou interfaces correspondentes ao seu script.</span><span class="sxs-lookup"><span data-stu-id="a8d41-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="a8d41-140">O script a seguir usa `fetch` para recuperar dados JSON do servidor de teste na URL dada.</span><span class="sxs-lookup"><span data-stu-id="a8d41-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="a8d41-141">Observe a `JSONData` interface para armazenar os dados como um tipo de correspondência.</span><span class="sxs-lookup"><span data-stu-id="a8d41-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

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

### <a name="other-fetch-samples"></a><span data-ttu-id="a8d41-142">Outras `fetch` amostras</span><span class="sxs-lookup"><span data-stu-id="a8d41-142">Other `fetch` samples</span></span>

* <span data-ttu-id="a8d41-143">A amostra [de busca externa use chamadas de Office Scripts](../resources/samples/external-fetch-calls.md) mostra como obter informações básicas sobre os repositórios de GitHub do usuário.</span><span class="sxs-lookup"><span data-stu-id="a8d41-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="a8d41-144">O [cenário amostral de Office Scripts: Graph dados de nível de água da NOAA](../resources/scenarios/noaa-data-fetch.md) demonstram o comando de busca que está sendo usado para recuperar registros do banco de dados de Marés e Correntes da Administração Oceânica Nacional Oceânica e Atmosférica.</span><span class="sxs-lookup"><span data-stu-id="a8d41-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="a8d41-145">Chamadas externas de Power Automate</span><span class="sxs-lookup"><span data-stu-id="a8d41-145">External calls from Power Automate</span></span>

<span data-ttu-id="a8d41-146">Qualquer chamada de API externa falha quando um script é executado com Power Automate.</span><span class="sxs-lookup"><span data-stu-id="a8d41-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="a8d41-147">Esta é uma diferença comportamental entre executar um script através do aplicativo Excel e através de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="a8d41-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="a8d41-148">Certifique-se de verificar seus scripts para obter tais referências antes de construí-las em um fluxo.</span><span class="sxs-lookup"><span data-stu-id="a8d41-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="a8d41-149">Você terá que usar [HTTP com o Azure AD](/connectors/webcontents/) ou outras ações equivalentes para extrair dados ou empurrá-los para um serviço externo.</span><span class="sxs-lookup"><span data-stu-id="a8d41-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="a8d41-150">Chamadas externas feitas através do Power Automate [Excel conector Online](/connectors/excelonlinebusiness) falham para ajudar a manter as políticas de prevenção de perda de dados existentes.</span><span class="sxs-lookup"><span data-stu-id="a8d41-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="a8d41-151">No entanto, os scripts que são executados através Power Automate são feitos fora da sua organização, e fora dos firewalls da sua organização.</span><span class="sxs-lookup"><span data-stu-id="a8d41-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="a8d41-152">Para proteção adicional contra usuários mal-intencionados neste ambiente externo, o administrador pode controlar o uso de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="a8d41-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="a8d41-153">O administrador pode desativar o conector online Excel em Power Automate ou desativar Office Scripts para Excel na Web através dos [controles de administrador de scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="a8d41-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="a8d41-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="a8d41-154">See also</span></span>

* [<span data-ttu-id="a8d41-155">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="a8d41-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="a8d41-156">Usar chamadas de busca externa em Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="a8d41-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="a8d41-157">Office Cenário da amostra de scripts: Graph dados de nível de água da NOAA</span><span class="sxs-lookup"><span data-stu-id="a8d41-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
