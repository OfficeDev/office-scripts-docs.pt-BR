---
title: Chamada de API externa nos scripts do Office
description: Suporte e orientação para fazer chamadas de API externas em Office Script.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 5d768b53112473c1774f8fe8257b197ffead4a63
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631640"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="7f361-103">Chamada de API externa nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7f361-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="7f361-104">Scripts suportam chamadas para serviços externos.</span><span class="sxs-lookup"><span data-stu-id="7f361-104">Scripts support calls to external services.</span></span> <span data-ttu-id="7f361-105">Use esses serviços para fornecer dados e outras informações à sua workbook.</span><span class="sxs-lookup"><span data-stu-id="7f361-105">Use these services to supply data and other information to your workbook.</span></span>

> [!CAUTION]
> <span data-ttu-id="7f361-106">Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis.</span><span class="sxs-lookup"><span data-stu-id="7f361-106">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="7f361-107">O administrador pode estabelecer proteção de firewall contra essas chamadas.</span><span class="sxs-lookup"><span data-stu-id="7f361-107">Your admin can establish firewall protection against such calls.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7f361-108">As chamadas para APIs externas só podem ser feitas por meio do aplicativo Excel, não por meio Power Automate [em circunstâncias normais.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="7f361-108">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="7f361-109">Configurar seu script para chamadas externas</span><span class="sxs-lookup"><span data-stu-id="7f361-109">Configure your script for external calls</span></span>

<span data-ttu-id="7f361-110">Chamadas externas [são assíncronas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) e exigem que seu script seja marcado como `async` .</span><span class="sxs-lookup"><span data-stu-id="7f361-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="7f361-111">Adicione o `async` prefixo à `main` sua função e retorne um , conforme mostrado `Promise` aqui:</span><span class="sxs-lookup"><span data-stu-id="7f361-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="7f361-112">Scripts que retornam outras informações podem retornar `Promise` um desse tipo.</span><span class="sxs-lookup"><span data-stu-id="7f361-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="7f361-113">Por exemplo, se o script precisar retornar um `Employee` objeto, a assinatura de retorno será `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="7f361-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="7f361-114">Você precisará aprender as interfaces do serviço externo para fazer chamadas para esse serviço.</span><span class="sxs-lookup"><span data-stu-id="7f361-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="7f361-115">Se você estiver usando ou APIs REST , será necessário determinar `fetch` a estrutura JSON dos dados retornados. [](https://wikipedia.org/wiki/Representational_state_transfer)</span><span class="sxs-lookup"><span data-stu-id="7f361-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="7f361-116">Para entrada e saída do script, considere fazer uma para corresponder às `interface` estruturas JSON necessárias.</span><span class="sxs-lookup"><span data-stu-id="7f361-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="7f361-117">Isso oferece ao script mais segurança de tipo.</span><span class="sxs-lookup"><span data-stu-id="7f361-117">This gives the script more type safety.</span></span> <span data-ttu-id="7f361-118">Você pode ver um exemplo disso em [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span><span class="sxs-lookup"><span data-stu-id="7f361-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="7f361-119">Limitações com chamadas externas de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="7f361-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="7f361-120">Não há nenhuma maneira de entrar ou usar o tipo OAuth2 de fluxos de autenticação.</span><span class="sxs-lookup"><span data-stu-id="7f361-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="7f361-121">Todas as chaves e credenciais devem ser codificadas (ou leitura de outra fonte).</span><span class="sxs-lookup"><span data-stu-id="7f361-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="7f361-122">Não há infraestrutura para armazenar credenciais e chaves da API.</span><span class="sxs-lookup"><span data-stu-id="7f361-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="7f361-123">Isso terá que ser gerenciado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="7f361-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="7f361-124">Cookies de documento `localStorage` e objetos não são `sessionStorage` suportados.</span><span class="sxs-lookup"><span data-stu-id="7f361-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span>
* <span data-ttu-id="7f361-125">Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis ou dados externos a serem trazidos para as guias de trabalho internas.</span><span class="sxs-lookup"><span data-stu-id="7f361-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="7f361-126">O administrador pode estabelecer proteção de firewall contra essas chamadas.</span><span class="sxs-lookup"><span data-stu-id="7f361-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="7f361-127">Verifique as políticas locais antes de confiar em chamadas externas.</span><span class="sxs-lookup"><span data-stu-id="7f361-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="7f361-128">Verifique a quantidade de transferência de dados antes de assumir uma dependência.</span><span class="sxs-lookup"><span data-stu-id="7f361-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="7f361-129">Por exemplo, retirar todo o conjuntos de dados externos pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em partes.</span><span class="sxs-lookup"><span data-stu-id="7f361-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="7f361-130">Recuperar informações com `fetch`</span><span class="sxs-lookup"><span data-stu-id="7f361-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="7f361-131">A [API de busca](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera informações de serviços externos.</span><span class="sxs-lookup"><span data-stu-id="7f361-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="7f361-132">É uma `async` API, portanto, você precisa ajustar a `main` assinatura do seu script.</span><span class="sxs-lookup"><span data-stu-id="7f361-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="7f361-133">Faça a `main` função e faça com que ela retorne um `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="7f361-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="7f361-134">Você também deve se certificar `await` da `fetch` chamada e `json` recuperação.</span><span class="sxs-lookup"><span data-stu-id="7f361-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="7f361-135">Isso garante que essas operações terminem antes do script terminar.</span><span class="sxs-lookup"><span data-stu-id="7f361-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="7f361-136">Todos os dados JSON recuperados por `fetch` devem corresponder a uma interface definida no script.</span><span class="sxs-lookup"><span data-stu-id="7f361-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="7f361-137">O valor retornado deve ser atribuído a um tipo específico porque [Office scripts não suportam o `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts).</span><span class="sxs-lookup"><span data-stu-id="7f361-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="7f361-138">Você deve consultar a documentação do seu serviço para ver quais são os nomes e tipos das propriedades retornadas.</span><span class="sxs-lookup"><span data-stu-id="7f361-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="7f361-139">Em seguida, adicione a interface ou interfaces correspondentes ao seu script.</span><span class="sxs-lookup"><span data-stu-id="7f361-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="7f361-140">O script a seguir `fetch` usa para recuperar dados JSON do servidor de teste na URL determinada.</span><span class="sxs-lookup"><span data-stu-id="7f361-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="7f361-141">Observe a `JSONData` interface para armazenar os dados como um tipo correspondente.</span><span class="sxs-lookup"><span data-stu-id="7f361-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

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

### <a name="other-fetch-samples"></a><span data-ttu-id="7f361-142">Outros `fetch` exemplos</span><span class="sxs-lookup"><span data-stu-id="7f361-142">Other `fetch` samples</span></span>

* <span data-ttu-id="7f361-143">O exemplo Usar chamadas de busca externa [Office scripts](../resources/samples/external-fetch-calls.md) mostra como obter informações básicas sobre os repositórios de GitHub do usuário.</span><span class="sxs-lookup"><span data-stu-id="7f361-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="7f361-144">O Office de exemplo scripts: Graph dados de nível de água do [NOAA](../resources/scenarios/noaa-data-fetch.md) demonstra o comando fetch que está sendo usado para recuperar registros do banco de dados De onda e currents da Administração Nacional Oceânica e Atmosférico.</span><span class="sxs-lookup"><span data-stu-id="7f361-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="7f361-145">Chamadas externas de Power Automate</span><span class="sxs-lookup"><span data-stu-id="7f361-145">External calls from Power Automate</span></span>

<span data-ttu-id="7f361-146">Qualquer chamada de API externa falha quando um script é executado com Power Automate.</span><span class="sxs-lookup"><span data-stu-id="7f361-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="7f361-147">Essa é uma diferença comportamental entre executar um script por meio do aplicativo Excel e por meio Power Automate.</span><span class="sxs-lookup"><span data-stu-id="7f361-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="7f361-148">Certifique-se de verificar seus scripts para essas referências antes de ad construi-las em um fluxo.</span><span class="sxs-lookup"><span data-stu-id="7f361-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="7f361-149">Você terá que usar HTTP com o [Azure AD](/connectors/webcontents/) ou outras ações equivalentes para puxar dados ou pressioná-los para um serviço externo.</span><span class="sxs-lookup"><span data-stu-id="7f361-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="7f361-150">As chamadas externas feitas por meio do conector Power Automate [Excel Online](/connectors/excelonlinebusiness) falham para ajudar a manter políticas de prevenção contra perda de dados existentes.</span><span class="sxs-lookup"><span data-stu-id="7f361-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="7f361-151">No entanto, os scripts que são executados por Power Automate são feitos fora da sua organização e fora dos firewalls da sua organização.</span><span class="sxs-lookup"><span data-stu-id="7f361-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="7f361-152">Para obter proteção adicional contra usuários mal-intencionados nesse ambiente externo, o administrador pode controlar o uso de Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="7f361-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="7f361-153">O administrador pode desabilitar o conector Excel Online no Power Automate ou desativar scripts do Office para Excel na Web por meio dos controles de administrador Office [Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="7f361-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="7f361-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="7f361-154">See also</span></span>

* [<span data-ttu-id="7f361-155">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7f361-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="7f361-156">Usar chamadas de busca externa em Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7f361-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="7f361-157">Office Cenário de exemplo de scripts: Graph dados de nível de água do NOAA</span><span class="sxs-lookup"><span data-stu-id="7f361-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
