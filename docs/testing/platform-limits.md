---
title: Limites e requisitos da plataforma com Office Scripts
description: Limites de recursos e suporte ao navegador para Office scripts quando usados com Excel na Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545578"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="37a26-103">Limites e requisitos da plataforma com Office Scripts</span><span class="sxs-lookup"><span data-stu-id="37a26-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="37a26-104">Há algumas limitações de plataforma das quais você deve estar ciente ao desenvolver Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="37a26-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="37a26-105">Este artigo detalha o suporte do navegador e os limites de dados Office scripts para Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="37a26-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="37a26-106">Suporte do navegador</span><span class="sxs-lookup"><span data-stu-id="37a26-106">Browser support</span></span>

<span data-ttu-id="37a26-107">Office Os scripts funcionam em qualquer navegador que [oferece suporte Office para a Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="37a26-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="37a26-108">No entanto, alguns recursos JavaScript não são suportados no Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="37a26-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="37a26-109">Quaisquer recursos introduzidos [no ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11.</span><span class="sxs-lookup"><span data-stu-id="37a26-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="37a26-110">Se as pessoas em sua organização ainda usarem esse navegador, teste seus scripts nesse ambiente ao compartilhar.</span><span class="sxs-lookup"><span data-stu-id="37a26-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="37a26-111">Cookies de terceiros</span><span class="sxs-lookup"><span data-stu-id="37a26-111">Third-party cookies</span></span>

<span data-ttu-id="37a26-112">Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **Automatizar** no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="37a26-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="37a26-113">Verifique as configurações do navegador se a guia não está sendo exibida.</span><span class="sxs-lookup"><span data-stu-id="37a26-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="37a26-114">Se você estiver usando uma sessão privada do navegador, talvez seja necessário reabilitar essa configuração sempre.</span><span class="sxs-lookup"><span data-stu-id="37a26-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="37a26-115">Alguns navegadores referem-se a essa configuração como "todos os cookies", em vez de "cookies de terceiros".</span><span class="sxs-lookup"><span data-stu-id="37a26-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="37a26-116">Instruções para ajustar configurações de cookie em navegadores populares</span><span class="sxs-lookup"><span data-stu-id="37a26-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="37a26-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="37a26-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="37a26-118">Borda</span><span class="sxs-lookup"><span data-stu-id="37a26-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="37a26-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="37a26-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="37a26-120">Safari</span><span class="sxs-lookup"><span data-stu-id="37a26-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="37a26-121">Limites de dados</span><span class="sxs-lookup"><span data-stu-id="37a26-121">Data limits</span></span>

<span data-ttu-id="37a26-122">Há limites sobre a quantidade Excel dados podem ser transferidos de uma só vez e quantas transações individuais Power Automate podem ser conduzidas.</span><span class="sxs-lookup"><span data-stu-id="37a26-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="37a26-123">Excel</span><span class="sxs-lookup"><span data-stu-id="37a26-123">Excel</span></span>

<span data-ttu-id="37a26-124">Excel para a Web tem as seguintes limitações ao fazer chamadas para a lista de trabalho por meio de um script:</span><span class="sxs-lookup"><span data-stu-id="37a26-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="37a26-125">Solicitações e respostas são limitadas a **5 MB**.</span><span class="sxs-lookup"><span data-stu-id="37a26-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="37a26-126">Um intervalo é limitado a **cinco milhões de células**.</span><span class="sxs-lookup"><span data-stu-id="37a26-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="37a26-127">Se você estiver encontrando erros ao lidar com conjuntos de dados grandes, tente usar vários intervalos menores em vez de intervalos maiores.</span><span class="sxs-lookup"><span data-stu-id="37a26-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="37a26-128">Por exemplo, consulte o exemplo [Gravar um grande conjuntos de](../resources/samples/write-large-dataset.md) dados.</span><span class="sxs-lookup"><span data-stu-id="37a26-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="37a26-129">Você também pode usar APIs como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para direcionar células específicas em vez de intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="37a26-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="37a26-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="37a26-130">Power Automate</span></span>

<span data-ttu-id="37a26-131">Ao usar Office scripts com Power Automate, cada usuário é limitado a **400** chamadas para a ação Executar Script por dia .</span><span class="sxs-lookup"><span data-stu-id="37a26-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="37a26-132">Esse limite é redefinido às 00:00 UTC.</span><span class="sxs-lookup"><span data-stu-id="37a26-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="37a26-133">A Power Automate plataforma também tem limitações de uso, que podem ser encontradas nos seguintes artigos:</span><span class="sxs-lookup"><span data-stu-id="37a26-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="37a26-134">Limites e configuração no Power Automate</span><span class="sxs-lookup"><span data-stu-id="37a26-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="37a26-135">Problemas conhecidos e limitações para o conector Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="37a26-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="37a26-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="37a26-136">See also</span></span>

- [<span data-ttu-id="37a26-137">Solucionar Office scripts</span><span class="sxs-lookup"><span data-stu-id="37a26-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="37a26-138">Desfazer os efeitos do Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="37a26-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="37a26-139">Melhorar o desempenho de seus Office Scripts</span><span class="sxs-lookup"><span data-stu-id="37a26-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="37a26-140">Fundamentos de script para Office scripts no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="37a26-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
