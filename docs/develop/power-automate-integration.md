---
title: Executar scripts do Office com o Power Automate
description: Como obter scripts do Office para Excel na Web trabalhando com um fluxo de trabalho do Power Automate.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755067"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="bb4d5-103">Executar scripts do Office com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="bb4d5-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="bb4d5-104">[O Power Automate](https://flow.microsoft.com) permite adicionar Scripts do Office a um fluxo de trabalho maior e automatizado.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="bb4d5-105">Você pode usar o Power Automate para fazer coisas como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projeto com base nos comentários da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="bb4d5-106">Introdução</span><span class="sxs-lookup"><span data-stu-id="bb4d5-106">Getting started</span></span>

<span data-ttu-id="bb4d5-107">Se você for novo no Power Automate, recomendamos visitar [Começar a usar o Power Automate.](/power-automate/getting-started)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="bb4d5-108">Lá, você pode saber mais sobre todas as possibilidades de automação disponíveis para você.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="bb4d5-109">Os documentos aqui se concentram em como os Scripts do Office funcionam com o Power Automate e como isso pode ajudar a melhorar sua experiência do Excel.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="bb4d5-110">Para começar a combinar Power Automate e Scripts do Office, siga o tutorial [Iniciar usando scripts com o Power Automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="bb4d5-111">Isso ensinará como criar um fluxo que chama um script simples.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="bb4d5-112">Depois de concluir esse tutorial e o Passar dados para scripts em um tutorial de fluxo do Power Automate executado automaticamente, retorne aqui para obter informações detalhadas sobre como conectar scripts do Office aos fluxos do Power [Automate.](../tutorials/excel-power-automate-trigger.md)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="bb4d5-113">Conector do Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="bb4d5-114">[Conectores](/connectors/connectors) são as pontes entre o Power Automate e os aplicativos.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="bb4d5-115">O [conector do Excel Online (Business)](/connectors/excelonlinebusiness) fornece aos fluxos acesso às planilhas do Excel.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="bb4d5-116">A ação "Executar script" permite chamar qualquer Script do Office acessível por meio da workbook selecionada.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="bb4d5-117">Você também pode dar aos seus scripts parâmetros de entrada para que os dados possam ser fornecidos pelo fluxo ou fazer com que seu script retorne informações para etapas posteriores no fluxo.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bb4d5-118">A ação "Executar script" oferece às pessoas que usam o conector do Excel acesso significativo à sua planilha e seus dados.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="bb4d5-119">Além disso, há riscos de segurança com scripts que fazem chamadas de API externas, conforme explicado em [Chamadas externas do Power Automate](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="bb4d5-120">Se o administrador estiver preocupado com a exposição de dados altamente confidenciais, ele poderá desativar o conector do Excel Online ou restringir o acesso aos Scripts do Office por meio dos controles de administrador [de Scripts do Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="bb4d5-121">Transferência de dados em fluxos para scripts</span><span class="sxs-lookup"><span data-stu-id="bb4d5-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="bb4d5-122">O Power Automate permite que você passe partes de dados entre etapas do seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="bb4d5-123">Os scripts podem ser configurados para aceitar qualquer tipo de informação que você precisa e retornar qualquer coisa da sua workbook que você deseja em seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="bb4d5-124">A entrada para o script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="bb4d5-125">A saída do script é declarada adicionando um tipo de retorno a `main` .</span><span class="sxs-lookup"><span data-stu-id="bb4d5-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="bb4d5-126">Quando você cria um bloco "Executar Script" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="bb4d5-127">Se você alterar os parâmetros ou retornar tipos de script, precisará refazer o bloco "Executar script" do seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="bb4d5-128">Isso garante que os dados estão sendo analisados corretamente.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="bb4d5-129">As seções a seguir abrangem os detalhes de entrada e saída para scripts usados no Power Automate.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="bb4d5-130">Se você quiser uma abordagem prática para aprender este tópico, experimente o passar dados para scripts em um tutorial de fluxo do [Power Automate](../tutorials/excel-power-automate-trigger.md) executado automaticamente ou explore o cenário de exemplo lembretes de tarefas [automatizados.](../resources/scenarios/task-reminders.md)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="bb4d5-131">`main` Parâmetros: passar dados para um script</span><span class="sxs-lookup"><span data-stu-id="bb4d5-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="bb4d5-132">Todas as entradas de script são especificadas como parâmetros adicionais para a `main` função.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="bb4d5-133">Por exemplo, se você quisesse que um script aceitasse um nome que representasse um nome como entrada, você `string` alteraria a `main` assinatura para `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="bb4d5-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="bb4d5-134">Ao configurar um fluxo no Power Automate, você pode especificar a entrada de script como valores [estáticos, expressões](/power-automate/use-expressions-in-conditions)ou conteúdo dinâmico.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="bb4d5-135">Os detalhes sobre o conector de um serviço individual podem ser encontrados na documentação [do Power Automate Connector.](/connectors/)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="bb4d5-136">Ao adicionar parâmetros de entrada à função de um `main` script, considere as seguintes restrições e concessões.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="bb4d5-137">O primeiro parâmetro deve ser do tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="bb4d5-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="bb4d5-138">Seu nome de parâmetro não importa.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="bb4d5-139">Cada parâmetro deve ter um tipo (como `string` ou `number` ).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="bb4d5-140">Os tipos `string` `number` básicos , `boolean` , , , , e são `any` `unknown` `object` `undefined` suportados.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="bb4d5-141">Há suporte para matrizes dos tipos básicos listados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="bb4d5-142">As matrizes aninhadas são suportadas como parâmetros (mas não como tipos de retorno).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="bb4d5-143">Os tipos de união são permitidos se eles são uma união de literais pertencentes a um único tipo (como `"Left" | "Right"` ).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="bb4d5-144">Também há suporte para uniões de um tipo com suporte indefinido (como `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="bb4d5-145">Os tipos de objeto são permitidos se eles contêm propriedades do tipo , , matrizes com `string` suporte ou outros objetos com `number` `boolean` suporte.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="bb4d5-146">O exemplo a seguir mostra objetos aninhados com suporte como tipos de parâmetro:</span><span class="sxs-lookup"><span data-stu-id="bb4d5-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="bb4d5-147">Os objetos devem ter sua interface ou definição de classe definida no script.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="bb4d5-148">Um objeto também pode ser definido anonimamente em linha, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="bb4d5-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="bb4d5-149">Parâmetros opcionais são permitidos e podem ser denodos como tal usando o modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="bb4d5-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="bb4d5-150">Os valores de parâmetro padrão são permitidos (por `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` exemplo.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="bb4d5-151">Retornando dados de um script</span><span class="sxs-lookup"><span data-stu-id="bb4d5-151">Returning data from a script</span></span>

<span data-ttu-id="bb4d5-152">Os scripts podem retornar dados da caixa de trabalho a serem usados como conteúdo dinâmico em um fluxo do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="bb4d5-153">Assim como nos parâmetros de entrada, o Power Automate coloca algumas restrições no tipo de retorno.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="bb4d5-154">Os tipos `string` básicos `number` , , e são `boolean` `void` `undefined` suportados.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="bb4d5-155">Os tipos de união usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="bb4d5-156">Os tipos de matriz são permitidos se eles são do tipo `string` `number` , ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="bb4d5-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="bb4d5-157">Eles também são permitidos se o tipo for uma união com suporte ou um tipo literal com suporte.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="bb4d5-158">Os tipos de objeto usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="bb4d5-159">A digitação implícita é suportada, embora ela deve seguir as mesmas regras de um tipo definido.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="bb4d5-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bb4d5-160">Example</span></span>

<span data-ttu-id="bb4d5-161">A captura de tela a seguir mostra um fluxo do Power Automate que é acionado sempre que um problema [do GitHub](https://github.com/) é atribuído a você.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="bb4d5-162">O fluxo executa um script que adiciona o problema a uma tabela em uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="bb4d5-163">Se houver cinco ou mais problemas nessa tabela, o fluxo enviará um lembrete de email.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="O editor de fluxo do Power Automate mostrando o fluxo de exemplo.":::

<span data-ttu-id="bb4d5-165">A função do script especifica a ID do problema e o título do problema como parâmetros de entrada, e o script retorna o número de linhas `main` na tabela de problemas.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bb4d5-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="bb4d5-166">See also</span></span>

- [<span data-ttu-id="bb4d5-167">Executar scripts do Office no Excel na Web com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="bb4d5-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="bb4d5-168">Passar dados para scripts numa execução automática do fluxo no Power Automate.</span><span class="sxs-lookup"><span data-stu-id="bb4d5-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="bb4d5-169">Retorna dados de um script para um fluxo do Power Automate executado automaticamente</span><span class="sxs-lookup"><span data-stu-id="bb4d5-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="bb4d5-170">Solução de problemas de informações para o Power Automate com scripts do Office</span><span class="sxs-lookup"><span data-stu-id="bb4d5-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="bb4d5-171">Começar a usar o Power Automate</span><span class="sxs-lookup"><span data-stu-id="bb4d5-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="bb4d5-172">Documentação de referência do conector do Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="bb4d5-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
