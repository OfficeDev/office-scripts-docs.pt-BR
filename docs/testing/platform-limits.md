---
title: Limites e requisitos de plataforma com scripts do Office
description: Limites de recursos e suporte ao navegador para Scripts do Office quando usados com o Excel na Web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 93307b6204f409f26c77b5ead33188205d5c4b4d
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837262"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="7e3ea-103">Limites e requisitos de plataforma com scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7e3ea-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="7e3ea-104">Há algumas limitações de plataforma que você deve estar ciente ao desenvolver scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="7e3ea-105">Este artigo detalha o suporte ao navegador e os limites de dados para Scripts do Office para Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="7e3ea-106">Suporte do navegador</span><span class="sxs-lookup"><span data-stu-id="7e3ea-106">Browser support</span></span>

<span data-ttu-id="7e3ea-107">Os Scripts do Office funcionam em qualquer navegador [compatível com o Office para a Web.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="7e3ea-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="7e3ea-108">No entanto, alguns recursos JavaScript não são suportados no Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="7e3ea-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="7e3ea-109">Quaisquer recursos introduzidos [no ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="7e3ea-110">Se as pessoas em sua organização ainda usarem esse navegador, teste seus scripts nesse ambiente ao compartilhar.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="7e3ea-111">Cookies de terceiros</span><span class="sxs-lookup"><span data-stu-id="7e3ea-111">Third-party cookies</span></span>

<span data-ttu-id="7e3ea-112">Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **Automatizar** no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="7e3ea-113">Verifique as configurações do navegador se a guia não está sendo exibida.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="7e3ea-114">Se você estiver usando uma sessão privada do navegador, talvez seja necessário reabilitar essa configuração sempre.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="7e3ea-115">Alguns navegadores referem-se a essa configuração como "todos os cookies", em vez de "cookies de terceiros".</span><span class="sxs-lookup"><span data-stu-id="7e3ea-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="7e3ea-116">Instruções para ajustar configurações de cookie em navegadores populares</span><span class="sxs-lookup"><span data-stu-id="7e3ea-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="7e3ea-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="7e3ea-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="7e3ea-118">Borda</span><span class="sxs-lookup"><span data-stu-id="7e3ea-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="7e3ea-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="7e3ea-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="7e3ea-120">Safari</span><span class="sxs-lookup"><span data-stu-id="7e3ea-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="7e3ea-121">Limites de dados</span><span class="sxs-lookup"><span data-stu-id="7e3ea-121">Data limits</span></span>

<span data-ttu-id="7e3ea-122">Há limites sobre a quantidade de dados do Excel que podem ser transferidos de uma só vez e quantas transações individuais do Power Automate podem ser conduzidas.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="7e3ea-123">Excel</span><span class="sxs-lookup"><span data-stu-id="7e3ea-123">Excel</span></span>

<span data-ttu-id="7e3ea-124">O Excel para a Web tem as seguintes limitações ao fazer chamadas para a planilha por meio de um script:</span><span class="sxs-lookup"><span data-stu-id="7e3ea-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="7e3ea-125">Solicitações e respostas são limitadas a **5 MB**.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="7e3ea-126">Um intervalo é limitado a **cinco milhões de células**.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="7e3ea-127">Se você estiver encontrando erros ao lidar com conjuntos de dados grandes, tente usar vários intervalos menores em vez de intervalos maiores.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="7e3ea-128">Você também pode APIs como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para direcionar células específicas em vez de intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="7e3ea-129">Ligar/Desligar Automação</span><span class="sxs-lookup"><span data-stu-id="7e3ea-129">Power Automate</span></span>

<span data-ttu-id="7e3ea-130">Ao usar scripts do Office com o Power Automate, cada usuário é limitado a **200 chamadas por dia.**</span><span class="sxs-lookup"><span data-stu-id="7e3ea-130">When using Office Scripts with Power Automate, each user is limited to **200 calls per day**.</span></span> <span data-ttu-id="7e3ea-131">Esse limite é redefinido às 00:00 UTC.</span><span class="sxs-lookup"><span data-stu-id="7e3ea-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="7e3ea-132">A plataforma Power Automate também tem limitações de uso, que podem ser encontradas nos seguintes artigos:</span><span class="sxs-lookup"><span data-stu-id="7e3ea-132">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="7e3ea-133">Limites e configuração no Power Automate</span><span class="sxs-lookup"><span data-stu-id="7e3ea-133">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="7e3ea-134">Problemas conhecidos e limitações para o conector do Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="7e3ea-134">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="7e3ea-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="7e3ea-135">See also</span></span>

- [<span data-ttu-id="7e3ea-136">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7e3ea-136">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="7e3ea-137">Desfazer os efeitos de um script do Office</span><span class="sxs-lookup"><span data-stu-id="7e3ea-137">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="7e3ea-138">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7e3ea-138">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="7e3ea-139">Scripts básicos para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="7e3ea-139">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
