---
title: Fazer chamadas de API externas em Scripts do Office
description: Saiba como fazer chamadas de API externas em Scripts do Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0ed57ed3b97309dbb7ea196695dcc347e133b3cf
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754800"
---
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="84ea6-103">Chamadas de API externas de Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="84ea6-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="84ea6-104">Scripts do Office permitem [suporte limitado a chamada de API externa](../../develop/external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="84ea6-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="84ea6-105">Não há nenhuma maneira de entrar ou usar o tipo OAuth2 de fluxos de autenticação.</span><span class="sxs-lookup"><span data-stu-id="84ea6-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="84ea6-106">Todas as chaves e credenciais devem ser codificadas (ou leitura de outra fonte).</span><span class="sxs-lookup"><span data-stu-id="84ea6-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="84ea6-107">Não há infraestrutura para armazenar credenciais e chaves da API.</span><span class="sxs-lookup"><span data-stu-id="84ea6-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="84ea6-108">Isso terá que ser gerenciado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="84ea6-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="84ea6-109">Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis ou dados externos a serem trazidos para as guias de trabalho internas.</span><span class="sxs-lookup"><span data-stu-id="84ea6-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="84ea6-110">O administrador pode estabelecer proteção de firewall contra essas chamadas.</span><span class="sxs-lookup"><span data-stu-id="84ea6-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="84ea6-111">Verifique as políticas locais antes de confiar em chamadas externas.</span><span class="sxs-lookup"><span data-stu-id="84ea6-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="84ea6-112">Se um script usar uma chamada de API, ele não funcionará em um cenário do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="84ea6-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="84ea6-113">Você terá que usar a ação HTTP do Power Automate ou ações equivalentes para puxar dados ou pressioná-los para um serviço externo.</span><span class="sxs-lookup"><span data-stu-id="84ea6-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="84ea6-114">Uma chamada de API externa envolve uma sintaxe assíncrona da API e requer um conhecimento ligeiramente avançado da maneira como a comunicação assíncrona funciona.</span><span class="sxs-lookup"><span data-stu-id="84ea6-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="84ea6-115">Verifique a quantidade de transferência de dados antes de assumir uma dependência.</span><span class="sxs-lookup"><span data-stu-id="84ea6-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="84ea6-116">Por exemplo, retirar todo o conjuntos de dados externos pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em partes.</span><span class="sxs-lookup"><span data-stu-id="84ea6-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="84ea6-117">Conhecimento e recursos úteis</span><span class="sxs-lookup"><span data-stu-id="84ea6-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="84ea6-118">[API REST](https://en.wikipedia.org/wiki/Representational_state_transfer): A maneira mais provável é usar a chamada da API.</span><span class="sxs-lookup"><span data-stu-id="84ea6-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="84ea6-119">[ `async` : Entenda como isso funciona. `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)</span><span class="sxs-lookup"><span data-stu-id="84ea6-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="84ea6-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Entenda como isso funciona.</span><span class="sxs-lookup"><span data-stu-id="84ea6-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="84ea6-121">Etapas</span><span class="sxs-lookup"><span data-stu-id="84ea6-121">Steps</span></span>

1. <span data-ttu-id="84ea6-122">Marque sua `main` função como uma função assíncrona adicionando `async` prefixo.</span><span class="sxs-lookup"><span data-stu-id="84ea6-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="84ea6-123">Por exemplo, `async function main(workbook: ExcelScript.Workbook)`.</span><span class="sxs-lookup"><span data-stu-id="84ea6-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="84ea6-124">Qual tipo de chamada de API você está fazendo?</span><span class="sxs-lookup"><span data-stu-id="84ea6-124">Which type of API call are you making?</span></span> <span data-ttu-id="84ea6-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="84ea6-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="84ea6-126">Consulte o material da API REST para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="84ea6-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="84ea6-127">Obtenha o ponto de extremidade da API de serviço, requisitos de autenticação, headers, etc.</span><span class="sxs-lookup"><span data-stu-id="84ea6-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="84ea6-128">Defina a entrada ou saída `interface` para ajudar na conclusão do código e na verificação do tempo de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="84ea6-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="84ea6-129">Consulte [vídeo](#training-video-how-to-make-external-api-calls) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="84ea6-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="84ea6-130">Código, teste, otimizar.</span><span class="sxs-lookup"><span data-stu-id="84ea6-130">Code, test, optimize.</span></span> <span data-ttu-id="84ea6-131">Você pode criar uma função para sua rotina de chamada de API para torná-la reutilizável de outras partes do seu script ou para reutilização em um script diferente (copiar colar se torna muito mais fácil dessa maneira).</span><span class="sxs-lookup"><span data-stu-id="84ea6-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="84ea6-132">Cenário</span><span class="sxs-lookup"><span data-stu-id="84ea6-132">Scenario</span></span>

<span data-ttu-id="84ea6-133">Este script obtém informações básicas sobre repositórios do GitHub do usuário.</span><span class="sxs-lookup"><span data-stu-id="84ea6-133">This script gets basic information about the user's GitHub repositories.</span></span>

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="84ea6-134">Recursos usados no exemplo</span><span class="sxs-lookup"><span data-stu-id="84ea6-134">Resources used in the sample</span></span>

1. [<span data-ttu-id="84ea6-135">Obter referência da API Github de repositórios.</span><span class="sxs-lookup"><span data-stu-id="84ea6-135">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="84ea6-136">Saída de chamada da API: vá para um navegador da Web ou qualquer interface HTTP e digite , substituindo o espaço reservado `https://api.github.com/users/{USERNAME}/repos` {USERNAME} pela ID do Github.</span><span class="sxs-lookup"><span data-stu-id="84ea6-136">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="84ea6-137">Informações buscadas: repo.name, repo.size, repo.owner.id, repo.license?. name</span><span class="sxs-lookup"><span data-stu-id="84ea6-137">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="84ea6-138">Código de exemplo: obter informações básicas sobre repositórios do GitHub do usuário</span><span class="sxs-lookup"><span data-stu-id="84ea6-138">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="84ea6-139">Vídeo de treinamento: como fazer chamadas de API externas</span><span class="sxs-lookup"><span data-stu-id="84ea6-139">Training video: How to make external API calls</span></span>

<span data-ttu-id="84ea6-140">[![Assista a um vídeo sobre como fazer chamadas de API externas](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vídeo sobre como fazer chamadas de API externas")</span><span class="sxs-lookup"><span data-stu-id="84ea6-140">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
