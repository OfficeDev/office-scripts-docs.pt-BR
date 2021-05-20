---
title: Solução de problemas Office Scripts
description: Depuração de dicas e técnicas para Office Scripts, bem como recursos de ajuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545550"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="05e35-103">Solução de problemas Office Scripts</span><span class="sxs-lookup"><span data-stu-id="05e35-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="05e35-104">À medida que você desenvolve Office Scripts, você pode cometer erros.</span><span class="sxs-lookup"><span data-stu-id="05e35-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="05e35-105">Está tudo bem, está tudo bem.</span><span class="sxs-lookup"><span data-stu-id="05e35-105">It's okay.</span></span> <span data-ttu-id="05e35-106">Você tem as ferramentas para ajudar a encontrar os problemas e fazer seus roteiros funcionarem perfeitamente.</span><span class="sxs-lookup"><span data-stu-id="05e35-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="05e35-107">Tipos de erros</span><span class="sxs-lookup"><span data-stu-id="05e35-107">Types of errors</span></span>

<span data-ttu-id="05e35-108">Office Os erros de scripts caem em uma das duas categorias:</span><span class="sxs-lookup"><span data-stu-id="05e35-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="05e35-109">Compilar erros ou avisos em tempo de compilação</span><span class="sxs-lookup"><span data-stu-id="05e35-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="05e35-110">Erros de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="05e35-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="05e35-111">Erros de tempo de compilação</span><span class="sxs-lookup"><span data-stu-id="05e35-111">Compile-time errors</span></span>

<span data-ttu-id="05e35-112">Erros e avisos de tempo de compilação são mostrados inicialmente no Editor de Código.</span><span class="sxs-lookup"><span data-stu-id="05e35-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="05e35-113">Estes são mostrados pelos sublinhados vermelhos ondulados no editor.</span><span class="sxs-lookup"><span data-stu-id="05e35-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="05e35-114">Eles também são exibidos na guia **Problemas** na parte inferior do painel de tarefas do Editor de Código.</span><span class="sxs-lookup"><span data-stu-id="05e35-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="05e35-115">A seleção do erro dará mais detalhes sobre o problema e sugerirá soluções.</span><span class="sxs-lookup"><span data-stu-id="05e35-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="05e35-116">Os erros de tempo de compilação devem ser resolvidos antes de executar o script.</span><span class="sxs-lookup"><span data-stu-id="05e35-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Um erro do compilador mostrado no texto do hover do Editor de código":::

<span data-ttu-id="05e35-118">Você também pode ver sublinhas de aviso laranja e mensagens informacionais cinzas.</span><span class="sxs-lookup"><span data-stu-id="05e35-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="05e35-119">Isso indica sugestões de desempenho ou outras possibilidades onde o script pode ter efeitos não intencionais.</span><span class="sxs-lookup"><span data-stu-id="05e35-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="05e35-120">Tais avisos devem ser examinados de perto antes de rejeií-los.</span><span class="sxs-lookup"><span data-stu-id="05e35-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="05e35-121">Erros de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="05e35-121">Runtime errors</span></span>

<span data-ttu-id="05e35-122">Erros de tempo de execução acontecem por causa de problemas lógicos no script.</span><span class="sxs-lookup"><span data-stu-id="05e35-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="05e35-123">Isso pode ser porque um objeto usado no script não está na pasta de trabalho, uma tabela é formatada de forma diferente do previsto, ou alguma outra pequena discrepância entre os requisitos do script e a pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="05e35-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="05e35-124">O script a seguir gera um erro quando uma planilha chamada "TestSheet" não está presente.</span><span class="sxs-lookup"><span data-stu-id="05e35-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="05e35-125">Mensagens de console</span><span class="sxs-lookup"><span data-stu-id="05e35-125">Console messages</span></span>

<span data-ttu-id="05e35-126">Ambos os erros de tempo de compilação e tempo de execução exibem mensagens de erro no console quando um script é executado.</span><span class="sxs-lookup"><span data-stu-id="05e35-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="05e35-127">Eles dão um número de linha onde o problema foi encontrado.</span><span class="sxs-lookup"><span data-stu-id="05e35-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="05e35-128">Tenha em mente que a causa raiz de qualquer problema pode ser uma linha de código diferente do indicado no console.</span><span class="sxs-lookup"><span data-stu-id="05e35-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="05e35-129">A imagem a seguir mostra a saída do console para o erro [explícito `any` ](../develop/typescript-restrictions.md) do compilador.</span><span class="sxs-lookup"><span data-stu-id="05e35-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="05e35-130">Observe o texto `[5, 16]` no início da sequência de erros.</span><span class="sxs-lookup"><span data-stu-id="05e35-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="05e35-131">Isso indica que o erro está na linha 5, começando pelo caractere 16.</span><span class="sxs-lookup"><span data-stu-id="05e35-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O console Code Editor exibindo uma mensagem de erro explícita 'qualquer'":::

<span data-ttu-id="05e35-133">A imagem a seguir mostra a saída do console para um erro de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="05e35-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="05e35-134">Aqui, o script tenta adicionar uma planilha com o nome de uma planilha existente.</span><span class="sxs-lookup"><span data-stu-id="05e35-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="05e35-135">Novamente, observe a "Linha 2" que precede o erro para mostrar qual linha investigar.</span><span class="sxs-lookup"><span data-stu-id="05e35-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="O console Code Editor exibindo um erro da chamada 'addWorksheet'":::

## <a name="console-logs"></a><span data-ttu-id="05e35-137">Logs de console</span><span class="sxs-lookup"><span data-stu-id="05e35-137">Console logs</span></span>

<span data-ttu-id="05e35-138">Imprima mensagens na tela com a `console.log` instrução.</span><span class="sxs-lookup"><span data-stu-id="05e35-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="05e35-139">Esses registros podem mostrar o valor atual das variáveis ou quais caminhos de código estão sendo acionados.</span><span class="sxs-lookup"><span data-stu-id="05e35-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="05e35-140">Para fazer isso, chame `console.log` qualquer objeto como parâmetro.</span><span class="sxs-lookup"><span data-stu-id="05e35-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="05e35-141">Normalmente, `string` um é o tipo mais fácil de ler no console.</span><span class="sxs-lookup"><span data-stu-id="05e35-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="05e35-142">As strings passadas `console.log` são exibidas no console de registro do Editor de Código, na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="05e35-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="05e35-143">Os registros são encontrados na guia **Saída,** embora a guia ganhe automaticamente o foco quando um registro é gravado.</span><span class="sxs-lookup"><span data-stu-id="05e35-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="05e35-144">Os registros não afetam a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05e35-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="05e35-145">Automatize a guia não aparecendo ou Office Scripts indisponíveis</span><span class="sxs-lookup"><span data-stu-id="05e35-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="05e35-146">As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **Automate** que não aparece em Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="05e35-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="05e35-147">[Certifique-se de que sua licença de Microsoft 365 inclui scripts Office](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="05e35-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="05e35-148">[Verifique se seu navegador está suportado](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="05e35-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="05e35-149">[Certifique-se de que cookies de terceiros estão ativados](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="05e35-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="05e35-150">[Certifique-se de que o administrador não desabilitou Office Scripts no centro administrativo Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="05e35-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="05e35-151">Solução de problemas de scripts em Power Automate</span><span class="sxs-lookup"><span data-stu-id="05e35-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="05e35-152">Para obter informações específicas para executar scripts através de Power Automate, consulte [Troubleshoot Office Scripts em execução em Power Automate](power-automate-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="05e35-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="05e35-153">Recursos de ajuda</span><span class="sxs-lookup"><span data-stu-id="05e35-153">Help resources</span></span>

<span data-ttu-id="05e35-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores dispostos a ajudar com problemas de codificação.</span><span class="sxs-lookup"><span data-stu-id="05e35-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="05e35-155">Muitas vezes, você será capaz de encontrar a solução para o seu problema através de uma pesquisa rápida stack overflow.</span><span class="sxs-lookup"><span data-stu-id="05e35-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="05e35-156">Se não, faça sua pergunta e marque-a com a tag "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="05e35-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="05e35-157">Não deixe de mencionar que você está criando um *script* Office , não um *complemento* Office .</span><span class="sxs-lookup"><span data-stu-id="05e35-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="05e35-158">Se você encontrar um problema com a API javascript Office, crie um problema no repositório [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub.</span><span class="sxs-lookup"><span data-stu-id="05e35-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="05e35-159">Os membros da equipe de produtos responderão às questões e prestarão assistência adicional.</span><span class="sxs-lookup"><span data-stu-id="05e35-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="05e35-160">Criar um problema no repositório **OfficeDev/office-js** indica que você encontrou uma falha na biblioteca de API JavaScript Office que a equipe do produto deve abordar.</span><span class="sxs-lookup"><span data-stu-id="05e35-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="05e35-161">Se houver algum problema com o Gravador de Ação ou Editor, envie feedback através do botão **Ajuda > Feedback** em Excel.</span><span class="sxs-lookup"><span data-stu-id="05e35-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="05e35-162">Confira também</span><span class="sxs-lookup"><span data-stu-id="05e35-162">See also</span></span>

- [<span data-ttu-id="05e35-163">Práticas recomendadas no Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="05e35-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="05e35-164">Limites de plataforma com scripts Office</span><span class="sxs-lookup"><span data-stu-id="05e35-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="05e35-165">Melhore o desempenho de seus scripts de Office</span><span class="sxs-lookup"><span data-stu-id="05e35-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="05e35-166">Solução de problemas Office Scripts em execução no PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="05e35-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="05e35-167">Desfazer os efeitos do Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="05e35-167">Undo the effects of Office Scripts</span></span>](undo.md)
