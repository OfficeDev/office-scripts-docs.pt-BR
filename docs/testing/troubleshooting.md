---
title: Solução de problemas dos scripts do Office
description: Dicas e técnicas de depuração para scripts do Office, bem como recursos da ajuda.
ms.date: 10/30/2020
localization_priority: Normal
ms.openlocfilehash: b45957bd336edce527397253cacec8cb09df715a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342875"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="c50e9-103">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="c50e9-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="c50e9-104">Ao desenvolver scripts do Office, você pode cometer erros.</span><span class="sxs-lookup"><span data-stu-id="c50e9-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="c50e9-105">Não há problema.</span><span class="sxs-lookup"><span data-stu-id="c50e9-105">It's okay.</span></span> <span data-ttu-id="c50e9-106">Temos ferramentas que ajudam a encontrar os problemas e a fazer com que seus scripts funcionem perfeitamente.</span><span class="sxs-lookup"><span data-stu-id="c50e9-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="c50e9-107">Logs do console</span><span class="sxs-lookup"><span data-stu-id="c50e9-107">Console logs</span></span>

<span data-ttu-id="c50e9-108">Às vezes, durante a solução de problemas, convém imprimir mensagens na tela.</span><span class="sxs-lookup"><span data-stu-id="c50e9-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="c50e9-109">Eles podem mostrar o valor atual de variáveis ou quais caminhos de código estão sendo disparados.</span><span class="sxs-lookup"><span data-stu-id="c50e9-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="c50e9-110">Para fazer isso, faça o log do texto no console.</span><span class="sxs-lookup"><span data-stu-id="c50e9-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="c50e9-111">As cadeias de caracteres passadas para `console.log` serão exibidas no console de registro em log do editor de código.</span><span class="sxs-lookup"><span data-stu-id="c50e9-111">Strings passed to `console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="c50e9-112">Para ativar o console, pressione o botão **reticências** e selecione **logs...**</span><span class="sxs-lookup"><span data-stu-id="c50e9-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="c50e9-113">Os logs não afetam a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c50e9-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="c50e9-114">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="c50e9-114">Error messages</span></span>

<span data-ttu-id="c50e9-115">Quando o script do Excel encontra um problema em execução, ele produz um erro.</span><span class="sxs-lookup"><span data-stu-id="c50e9-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="c50e9-116">Você verá um pop-up de solicitação perguntando se deseja **exibir os logs**.</span><span class="sxs-lookup"><span data-stu-id="c50e9-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="c50e9-117">Pressione esse botão para abrir o console e exibir quaisquer erros.</span><span class="sxs-lookup"><span data-stu-id="c50e9-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="c50e9-118">Guia automatizar não aparecendo ou os scripts do Office não estão disponíveis</span><span class="sxs-lookup"><span data-stu-id="c50e9-118">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="c50e9-119">As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **automatizada** não aparecendo no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="c50e9-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="c50e9-120">[Verifique se a licença do Microsoft 365 inclui scripts do Office](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="c50e9-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="c50e9-121">[Verifique se há suporte para o seu navegador](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="c50e9-121">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="c50e9-122">[Verifique se os cookies de terceiros estão habilitados](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="c50e9-122">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="c50e9-123">[Certifique-se de que o administrador não desabilitou scripts do Office no centro de administração do Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="c50e9-123">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a><span data-ttu-id="c50e9-124">Recursos de ajuda</span><span class="sxs-lookup"><span data-stu-id="c50e9-124">Help resources</span></span>

<span data-ttu-id="c50e9-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores que desejam ajudar com problemas de codificação.</span><span class="sxs-lookup"><span data-stu-id="c50e9-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="c50e9-126">Muitas vezes, você poderá encontrar a solução para o problema por meio de uma pesquisa rápida de estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="c50e9-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="c50e9-127">Caso contrário, faça a pergunta e marque-a com a marca "Office-scripts".</span><span class="sxs-lookup"><span data-stu-id="c50e9-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="c50e9-128">Não deixe de mencionar que você está criando um *script* do Office, não um *suplemento* do Office.</span><span class="sxs-lookup"><span data-stu-id="c50e9-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="c50e9-129">Se você encontrar um problema com a API JavaScript do Office, crie um problema no repositório do GitHub do [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="c50e9-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="c50e9-130">Os membros da equipe do produto responderão a problemas e fornecerão mais assistência.</span><span class="sxs-lookup"><span data-stu-id="c50e9-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="c50e9-131">A criação de um problema no repositório **OfficeDev/Office-js** indica que você encontrou uma falha na biblioteca da API JavaScript do Office para a qual a equipe de produto deve tratar.</span><span class="sxs-lookup"><span data-stu-id="c50e9-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="c50e9-132">Se houver um problema com o gravador de ação ou editor, envie comentários através do botão **ajuda > comentários** no Excel.</span><span class="sxs-lookup"><span data-stu-id="c50e9-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="c50e9-133">Também consulte</span><span class="sxs-lookup"><span data-stu-id="c50e9-133">See also</span></span>

- [<span data-ttu-id="c50e9-134">Scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="c50e9-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="c50e9-135">Conceitos básicos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="c50e9-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="c50e9-136">Limites de plataforma com scripts do Office</span><span class="sxs-lookup"><span data-stu-id="c50e9-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="c50e9-137">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="c50e9-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="c50e9-138">Desfazer os efeitos de um script do Office</span><span class="sxs-lookup"><span data-stu-id="c50e9-138">Undo the effects of an Office Script</span></span>](undo.md)
