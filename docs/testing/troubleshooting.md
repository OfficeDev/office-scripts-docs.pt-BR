---
title: Solucionar Office scripts
description: Dicas e técnicas de depuração para Office scripts, bem como recursos de ajuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074645"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="245c4-103">Solucionar Office scripts</span><span class="sxs-lookup"><span data-stu-id="245c4-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="245c4-104">À medida que você Office scripts, você pode cometer erros.</span><span class="sxs-lookup"><span data-stu-id="245c4-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="245c4-105">Não há problema.</span><span class="sxs-lookup"><span data-stu-id="245c4-105">It's okay.</span></span> <span data-ttu-id="245c4-106">Você tem as ferramentas para ajudar a encontrar os problemas e fazer seus scripts funcionarem perfeitamente.</span><span class="sxs-lookup"><span data-stu-id="245c4-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="245c4-107">Tipos de erros</span><span class="sxs-lookup"><span data-stu-id="245c4-107">Types of errors</span></span>

<span data-ttu-id="245c4-108">Office Os erros de scripts se enquadram em uma das duas categorias:</span><span class="sxs-lookup"><span data-stu-id="245c4-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="245c4-109">Erros ou avisos em tempo de compilação</span><span class="sxs-lookup"><span data-stu-id="245c4-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="245c4-110">Erros de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="245c4-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="245c4-111">Erros em tempo de compilação</span><span class="sxs-lookup"><span data-stu-id="245c4-111">Compile-time errors</span></span>

<span data-ttu-id="245c4-112">Erros e avisos de tempo de compilação são mostrados inicialmente no Editor de Código.</span><span class="sxs-lookup"><span data-stu-id="245c4-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="245c4-113">Eles são mostrados pelos sublinhados vermelho ondulados no editor.</span><span class="sxs-lookup"><span data-stu-id="245c4-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="245c4-114">Eles também são exibidos na guia **Problemas** na parte inferior do painel de tarefas Editor de Código.</span><span class="sxs-lookup"><span data-stu-id="245c4-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="245c4-115">Selecionar o erro dará mais detalhes sobre o problema e sugerirá soluções.</span><span class="sxs-lookup"><span data-stu-id="245c4-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="245c4-116">Erros em tempo de compilação devem ser resolvidos antes de executar o script.</span><span class="sxs-lookup"><span data-stu-id="245c4-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Um erro de compilador mostrado no texto de foco do Editor de Código.":::

<span data-ttu-id="245c4-118">Você também pode ver sublinhados de aviso laranja e mensagens informativas cinzas.</span><span class="sxs-lookup"><span data-stu-id="245c4-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="245c4-119">Elas indicam sugestões de desempenho ou outras possibilidades em que o script pode ter efeitos não intencional.</span><span class="sxs-lookup"><span data-stu-id="245c4-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="245c4-120">Esses avisos devem ser examinados de perto antes de descartá-los.</span><span class="sxs-lookup"><span data-stu-id="245c4-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="245c4-121">Erros de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="245c4-121">Runtime errors</span></span>

<span data-ttu-id="245c4-122">Erros de tempo de execução ocorrem devido a problemas de lógica no script.</span><span class="sxs-lookup"><span data-stu-id="245c4-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="245c4-123">Isso pode ser porque um objeto usado no script não está na guia de trabalho, uma tabela é formatada de forma diferente do previsto ou alguma outra pequena discrepância entre os requisitos do script e a atual.</span><span class="sxs-lookup"><span data-stu-id="245c4-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="245c4-124">O script a seguir gera um erro quando uma planilha chamada "TestSheet" não está presente.</span><span class="sxs-lookup"><span data-stu-id="245c4-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="245c4-125">Mensagens de console</span><span class="sxs-lookup"><span data-stu-id="245c4-125">Console messages</span></span>

<span data-ttu-id="245c4-126">Erros de tempo de compilação e tempo de execução exibem mensagens de erro no console quando um script é executado.</span><span class="sxs-lookup"><span data-stu-id="245c4-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="245c4-127">Eles dão um número de linha onde o problema foi encontrado.</span><span class="sxs-lookup"><span data-stu-id="245c4-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="245c4-128">Lembre-se de que a causa raiz de qualquer problema pode ser uma linha de código diferente da indicada no console.</span><span class="sxs-lookup"><span data-stu-id="245c4-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="245c4-129">A imagem a seguir mostra a saída do console para [o erro explícito `any` ](../develop/typescript-restrictions.md) do compilador.</span><span class="sxs-lookup"><span data-stu-id="245c4-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="245c4-130">Observe o texto `[5, 16]` no início da cadeia de caracteres de erro.</span><span class="sxs-lookup"><span data-stu-id="245c4-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="245c4-131">Isso indica que o erro está na linha 5, começando no caractere 16.</span><span class="sxs-lookup"><span data-stu-id="245c4-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O console do Editor de Código exibindo uma mensagem de erro &quot;qualquer&quot; explícita.":::

<span data-ttu-id="245c4-133">A imagem a seguir mostra a saída do console para um erro de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="245c4-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="245c4-134">Aqui, o script tenta adicionar uma planilha com o nome de uma planilha existente.</span><span class="sxs-lookup"><span data-stu-id="245c4-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="245c4-135">Novamente, observe a "Linha 2" anterior ao erro para mostrar qual linha investigar.</span><span class="sxs-lookup"><span data-stu-id="245c4-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="O console do Editor de Código exibindo um erro da chamada 'addWorksheet'.":::

## <a name="console-logs"></a><span data-ttu-id="245c4-137">Logs de console</span><span class="sxs-lookup"><span data-stu-id="245c4-137">Console logs</span></span>

<span data-ttu-id="245c4-138">Imprimir mensagens na tela com a `console.log` instrução.</span><span class="sxs-lookup"><span data-stu-id="245c4-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="245c4-139">Esses logs podem mostrar o valor atual das variáveis ou quais caminhos de código estão sendo disparados.</span><span class="sxs-lookup"><span data-stu-id="245c4-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="245c4-140">Para fazer isso, chame `console.log` qualquer objeto como parâmetro.</span><span class="sxs-lookup"><span data-stu-id="245c4-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="245c4-141">Normalmente, um `string` é o tipo mais fácil de ler no console.</span><span class="sxs-lookup"><span data-stu-id="245c4-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="245c4-142">As cadeias de caracteres passadas para são exibidas no console de registro em log do Editor de Código, na `console.log` parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="245c4-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="245c4-143">Os logs são encontrados na guia **Saída,** embora a guia automaticamente obtém o foco quando um log é gravado.</span><span class="sxs-lookup"><span data-stu-id="245c4-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="245c4-144">Os logs não afetam a agenda de trabalho.</span><span class="sxs-lookup"><span data-stu-id="245c4-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="245c4-145">Guia Automatizar não aparecendo ou Office Scripts indisponíveis</span><span class="sxs-lookup"><span data-stu-id="245c4-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="245c4-146">As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **Automatizar** que não aparecem no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="245c4-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="245c4-147">[Certifique-se de Microsoft 365 sua licença de Office Scripts](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="245c4-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="245c4-148">[Verifique se o navegador tem suporte](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="245c4-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="245c4-149">[Verifique se os cookies de terceiros estão habilitados.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="245c4-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="245c4-150">[Verifique se o administrador não desabilitou Office scripts no Centro de administração do Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="245c4-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="245c4-151">Solucionar problemas de scripts em Power Automate</span><span class="sxs-lookup"><span data-stu-id="245c4-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="245c4-152">Para obter informações específicas sobre como executar scripts Power Automate, consulte [Troubleshoot Office Scripts em execução em Power Automate](power-automate-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="245c4-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="245c4-153">Recursos de ajuda</span><span class="sxs-lookup"><span data-stu-id="245c4-153">Help resources</span></span>

<span data-ttu-id="245c4-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores dispostos a ajudar com problemas de codificação.</span><span class="sxs-lookup"><span data-stu-id="245c4-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="245c4-155">Muitas vezes, você poderá encontrar a solução para seu problema por meio de uma pesquisa rápida de Estouro de Pilha.</span><span class="sxs-lookup"><span data-stu-id="245c4-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="245c4-156">Se não, faça sua pergunta e marque-a com a marca "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="245c4-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="245c4-157">Certifique-se de mencionar que você está criando um *script* de Office , não um Office *Desem.*</span><span class="sxs-lookup"><span data-stu-id="245c4-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="245c4-158">Para enviar uma solicitação de recurso para Office Scripts, poste sua ideia na página Voz do Usuário [ou](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)se a solicitação de recurso já existir lá, adicione seu voto a ela.</span><span class="sxs-lookup"><span data-stu-id="245c4-158">To submit a feature request for Office Scripts, post your idea to our [User Voice page](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439), or if the feature request already exists there, add your vote for it.</span></span> <span data-ttu-id="245c4-159">Certifique-se de arquivar a solicitação em Excel para a Web na categoria "Macros, Scripts e Complementos".</span><span class="sxs-lookup"><span data-stu-id="245c4-159">Be sure to file the request under Excel for the web in the "Macros, Scripts and Add-ins" category.</span></span>

<span data-ttu-id="245c4-160">Se houver um problema com o Gravador de Ações ou Editor, entre em contato conosco.</span><span class="sxs-lookup"><span data-stu-id="245c4-160">If there is a problem with the Action Recorder or Editor, please let us know.</span></span> <span data-ttu-id="245c4-161">No menu do painel de tarefas do Editor de **Código...** selecione o botão **Enviar comentários** para compartilhar quaisquer problemas.</span><span class="sxs-lookup"><span data-stu-id="245c4-161">In the Code Editor task pane's **...** menu, select the **Send feedback** button to share any issues.</span></span>

:::image type="content" source="../images/code-editor-feedback.png" alt-text="O menu de estouro do Editor de Código com o botão &quot;Enviar comentários&quot;.":::

## <a name="see-also"></a><span data-ttu-id="245c4-163">Confira também</span><span class="sxs-lookup"><span data-stu-id="245c4-163">See also</span></span>

- [<span data-ttu-id="245c4-164">Práticas recomendadas nos Scripts do Office </span><span class="sxs-lookup"><span data-stu-id="245c4-164">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="245c4-165">Limites da plataforma com Office Scripts</span><span class="sxs-lookup"><span data-stu-id="245c4-165">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="245c4-166">Melhorar o desempenho de seus Office Scripts</span><span class="sxs-lookup"><span data-stu-id="245c4-166">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="245c4-167">Solucionar Office scripts em execução no PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="245c4-167">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="245c4-168">Desfazer os efeitos do Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="245c4-168">Undo the effects of Office Scripts</span></span>](undo.md)
