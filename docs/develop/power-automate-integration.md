---
title: Execute Office scripts com Power Automate
description: Como obter Office Scripts para Excel na Web trabalhando com um fluxo de trabalho Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545037"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="e1876-103">Execute Office scripts com Power Automate</span><span class="sxs-lookup"><span data-stu-id="e1876-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="e1876-104">[Power Automate](https://flow.microsoft.com) permite adicionar scripts Office a um fluxo de trabalho maior e automatizado.</span><span class="sxs-lookup"><span data-stu-id="e1876-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="e1876-105">Você pode usar Power Automate fazer coisas como adicionar o conteúdo de um e-mail à mesa de uma planilha ou criar ações em suas ferramentas de gerenciamento de projetos com base em comentários de pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1876-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="get-started"></a><span data-ttu-id="e1876-106">Introdução</span><span class="sxs-lookup"><span data-stu-id="e1876-106">Get started</span></span>

<span data-ttu-id="e1876-107">Se você é novo em Power Automate, recomendamos visitar [Get started with Power Automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="e1876-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="e1876-108">Lá, você pode aprender mais sobre todas as possibilidades de automação disponíveis para você.</span><span class="sxs-lookup"><span data-stu-id="e1876-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="e1876-109">Os documentos aqui se concentram em como Office Scripts trabalham com Power Automate e como isso pode ajudar a melhorar sua experiência Excel.</span><span class="sxs-lookup"><span data-stu-id="e1876-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="e1876-110">Para começar a combinar Power Automate e Office Scripts, siga o tutorial Comece a [usar scripts com Power Automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="e1876-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="e1876-111">Isso vai te ensinar como criar um fluxo que chama de script simples.</span><span class="sxs-lookup"><span data-stu-id="e1876-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="e1876-112">Depois de completar esse tutorial e os dados do Pass para scripts em um tutorial [de fluxo de fluxo Power Automate executado automaticamente,](../tutorials/excel-power-automate-trigger.md) retorne aqui para obter informações detalhadas sobre a conexão Office Scripts para fluxos Power Automate.</span><span class="sxs-lookup"><span data-stu-id="e1876-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="e1876-113">Excel Conector on-line (business)</span><span class="sxs-lookup"><span data-stu-id="e1876-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="e1876-114">[Conectores](/connectors/connectors) são as pontes entre Power Automate e aplicações.</span><span class="sxs-lookup"><span data-stu-id="e1876-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="e1876-115">O [conector Excel Online (Business)](/connectors/excelonlinebusiness) dá aos seus fluxos acesso a Excel livros de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1876-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="e1876-116">A ação "Executar script" permite que você chame qualquer Office Script acessível através da pasta de trabalho selecionada.</span><span class="sxs-lookup"><span data-stu-id="e1876-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="e1876-117">Você também pode fornecer parâmetros de entrada de seus scripts para que os dados possam ser fornecidos pelo fluxo ou ter suas informações de retorno do script para etapas posteriores no fluxo.</span><span class="sxs-lookup"><span data-stu-id="e1876-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e1876-118">A ação "Executar script" dá às pessoas que usam o conector Excel acesso significativo à sua pasta de trabalho e seus dados.</span><span class="sxs-lookup"><span data-stu-id="e1876-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="e1876-119">Além disso, existem riscos de segurança com scripts que fazem chamadas de API externas, como explicado em [chamadas externas de Power Automate](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="e1876-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="e1876-120">Se o administrador estiver preocupado com a exposição de dados altamente [confidenciais,](/microsoft-365/admin/manage/manage-office-scripts-settings)eles podem desligar o conector Excel Online ou restringir o acesso a scripts Office através dos controles de administrador de scripts Office .</span><span class="sxs-lookup"><span data-stu-id="e1876-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="e1876-121">Transferência de dados em fluxos para scripts</span><span class="sxs-lookup"><span data-stu-id="e1876-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="e1876-122">Power Automate permite que você passe pedaços de dados entre etapas do seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="e1876-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="e1876-123">Os scripts podem ser configurados para aceitar qualquer tipo de informação que você precise e retornar qualquer coisa da sua pasta de trabalho que você deseja em seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="e1876-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="e1876-124">A entrada para o seu script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="e1876-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="e1876-125">A saída do script é declarada adicionando um tipo de retorno a `main` .</span><span class="sxs-lookup"><span data-stu-id="e1876-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="e1876-126">Quando você cria um bloco "Executar script" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos.</span><span class="sxs-lookup"><span data-stu-id="e1876-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="e1876-127">Se você alterar os parâmetros ou retornar os tipos do seu script, você precisará refazer o bloco "Executar script" do seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="e1876-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="e1876-128">Isso garante que os dados estão sendo analisados corretamente.</span><span class="sxs-lookup"><span data-stu-id="e1876-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="e1876-129">As seções a seguir cobrem os detalhes de entrada e saída para scripts usados em Power Automate.</span><span class="sxs-lookup"><span data-stu-id="e1876-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="e1876-130">Se você quiser uma abordagem prática para aprender este tópico, experimente os dados do [Pass para scripts em um](../tutorials/excel-power-automate-trigger.md) tutorial de fluxo de Power Automate executado automaticamente ou explore o cenário de amostra [de lembretes de tarefas automatizados.](../resources/scenarios/task-reminders.md)</span><span class="sxs-lookup"><span data-stu-id="e1876-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-pass-data-to-a-script"></a><span data-ttu-id="e1876-131">`main` Parâmetros: Passar dados para um script</span><span class="sxs-lookup"><span data-stu-id="e1876-131">`main` Parameters: Pass data to a script</span></span>

<span data-ttu-id="e1876-132">Toda a entrada do script é especificada como parâmetros adicionais para a `main` função.</span><span class="sxs-lookup"><span data-stu-id="e1876-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="e1876-133">Por exemplo, se você quisesse um script para aceitar um `string` que representa um nome como entrada, você mudaria a assinatura para `main` `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="e1876-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="e1876-134">Quando você está configurando um fluxo em Power Automate, você pode especificar a entrada do script como valores [estáticos, expressões](/power-automate/use-expressions-in-conditions)ou conteúdo dinâmico.</span><span class="sxs-lookup"><span data-stu-id="e1876-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="e1876-135">Detalhes sobre o conector de um serviço individual podem ser encontrados na [documentação do conector Power Automate](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="e1876-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="e1876-136">Ao adicionar parâmetros de entrada à função de um `main` script, considere as seguintes franquias e restrições.</span><span class="sxs-lookup"><span data-stu-id="e1876-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="e1876-137">O primeiro parâmetro deve ser de tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="e1876-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="e1876-138">Seu nome de parâmetro não importa.</span><span class="sxs-lookup"><span data-stu-id="e1876-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="e1876-139">Cada parâmetro deve ter um tipo (como `string` ou `number` ).</span><span class="sxs-lookup"><span data-stu-id="e1876-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="e1876-140">Os tipos `string` `number` básicos, `boolean` , , , e são `unknown` `object` `undefined` suportados.</span><span class="sxs-lookup"><span data-stu-id="e1876-140">The basic types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="e1876-141">Arrays dos tipos básicos listados anteriormente são suportados.</span><span class="sxs-lookup"><span data-stu-id="e1876-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="e1876-142">As matrizes aninhadas são suportadas como parâmetros (mas não como tipos de retorno).</span><span class="sxs-lookup"><span data-stu-id="e1876-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="e1876-143">Os tipos sindicais são permitidos se forem uma união de literais pertencentes a um único tipo `"Left" | "Right"` (como).</span><span class="sxs-lookup"><span data-stu-id="e1876-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="e1876-144">Sindicatos de um tipo apoiado com indefinidos também são apoiados (como `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="e1876-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="e1876-145">Os tipos de objetos são permitidos se contiverem propriedades do `string` `number` `boolean` tipo, matrizes suportadas ou outros objetos suportados.</span><span class="sxs-lookup"><span data-stu-id="e1876-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="e1876-146">O exemplo a seguir mostra objetos aninhados que são suportados como tipos de parâmetro:</span><span class="sxs-lookup"><span data-stu-id="e1876-146">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="e1876-147">Os objetos devem ter sua interface ou definição de classe definida no script.</span><span class="sxs-lookup"><span data-stu-id="e1876-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="e1876-148">Um objeto também pode ser definido anonimamente inline, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="e1876-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="e1876-149">Parâmetros opcionais são permitidos e podem ser denotados como tal usando o modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="e1876-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="e1876-150">Valores de parâmetro padrão são permitidos (por exemplo `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="e1876-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="return-data-from-a-script"></a><span data-ttu-id="e1876-151">Retornar dados de um script</span><span class="sxs-lookup"><span data-stu-id="e1876-151">Return data from a script</span></span>

<span data-ttu-id="e1876-152">Os scripts podem retornar dados da pasta de trabalho para serem usados como conteúdo dinâmico em um fluxo Power Automate.</span><span class="sxs-lookup"><span data-stu-id="e1876-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="e1876-153">Assim como os parâmetros de entrada, Power Automate coloca algumas restrições no tipo de retorno.</span><span class="sxs-lookup"><span data-stu-id="e1876-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="e1876-154">Os tipos `string` `number` básicos, `boolean` , , , , e são `void` `undefined` suportados.</span><span class="sxs-lookup"><span data-stu-id="e1876-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="e1876-155">Os tipos de união usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.</span><span class="sxs-lookup"><span data-stu-id="e1876-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="e1876-156">Tipos de matriz são permitidos se forem do tipo `string` `number` , ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="e1876-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="e1876-157">Eles também são permitidos se o tipo for um tipo de união apoiada ou apoiado literal.</span><span class="sxs-lookup"><span data-stu-id="e1876-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="e1876-158">Os tipos de objetos usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.</span><span class="sxs-lookup"><span data-stu-id="e1876-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="e1876-159">A digitação implícita é suportada, embora deva seguir as mesmas regras de um tipo definido.</span><span class="sxs-lookup"><span data-stu-id="e1876-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="e1876-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1876-160">Example</span></span>

<span data-ttu-id="e1876-161">A captura de tela a seguir mostra um fluxo de Power Automate que é acionado sempre que um problema [de GitHub](https://github.com/) é atribuído a você.</span><span class="sxs-lookup"><span data-stu-id="e1876-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="e1876-162">O fluxo executa um script que adiciona o problema a uma tabela em uma Excel livro de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1876-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="e1876-163">Se houver cinco ou mais problemas nessa tabela, o fluxo envia um lembrete de e-mail.</span><span class="sxs-lookup"><span data-stu-id="e1876-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="O editor de fluxo Power Automate mostrando o fluxo de exemplo":::

<span data-ttu-id="e1876-165">A `main` função do script especifica o ID de edição e o título de emissão como parâmetros de entrada, e o script retorna o número de linhas na tabela de emissão.</span><span class="sxs-lookup"><span data-stu-id="e1876-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="e1876-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="e1876-166">See also</span></span>

- [<span data-ttu-id="e1876-167">Execute Office scripts em Excel na Web com Power Automate</span><span class="sxs-lookup"><span data-stu-id="e1876-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="e1876-168">Passar dados para scripts numa execução automática do fluxo no Power Automate.</span><span class="sxs-lookup"><span data-stu-id="e1876-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="e1876-169">Retorna dados de um script para um fluxo do Power Automate executado automaticamente</span><span class="sxs-lookup"><span data-stu-id="e1876-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="e1876-170">Solução de problemas para Power Automate com scripts Office</span><span class="sxs-lookup"><span data-stu-id="e1876-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="e1876-171">Começar a usar o Power Automate</span><span class="sxs-lookup"><span data-stu-id="e1876-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="e1876-172">Excel Documentação de referência do conector online (business)</span><span class="sxs-lookup"><span data-stu-id="e1876-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
