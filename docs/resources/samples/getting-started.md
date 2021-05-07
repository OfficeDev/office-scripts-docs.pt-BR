---
title: Iniciando com Office Scripts
description: Noções básicas Office scripts, incluindo acesso, ambiente e padrões de script.
ms.date: 04/01/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: d30c4fb4523c49b559e057eede4d5de162b74f9c
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232757"
---
# <a name="getting-started"></a><span data-ttu-id="d97a8-103">Introdução</span><span class="sxs-lookup"><span data-stu-id="d97a8-103">Getting started</span></span>

<span data-ttu-id="d97a8-104">Esta seção fornece detalhes sobre os conceitos básicos Office scripts, incluindo acesso, ambiente, fundamentos de script e alguns padrões básicos de script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-104">This section provides details about the basics of Office Scripts including access, environment, script fundamentals, and few basic script patterns.</span></span>

## <a name="environment-setup"></a><span data-ttu-id="d97a8-105">Configuração de ambiente</span><span class="sxs-lookup"><span data-stu-id="d97a8-105">Environment setup</span></span>

<span data-ttu-id="d97a8-106">Saiba mais sobre os conceitos básicos de acesso, ambiente e editor de scripts.</span><span class="sxs-lookup"><span data-stu-id="d97a8-106">Learn about the basics of access, environment, and script editor.</span></span>

<span data-ttu-id="d97a8-107">[![Noções básicas do Office scripts](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Noções básicas do Office scripts")</span><span class="sxs-lookup"><span data-stu-id="d97a8-107">[![Basics of Office Scripts application](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Basics of Office Scripts application")</span></span>

### <a name="access"></a><span data-ttu-id="d97a8-108">Acesso</span><span class="sxs-lookup"><span data-stu-id="d97a8-108">Access</span></span>

<span data-ttu-id="d97a8-109">Office Os scripts exigem configurações de administrador disponíveis para Microsoft 365 administrador **em Configurações**  >  **org**  >  **Office Scripts**.</span><span class="sxs-lookup"><span data-stu-id="d97a8-109">Office Scripts requires admin settings available for Microsoft 365 administrator under **Settings** > **Org settings** > **Office Scripts**.</span></span> <span data-ttu-id="d97a8-110">Por padrão, ele está ligado para todos os usuários.</span><span class="sxs-lookup"><span data-stu-id="d97a8-110">By default, it's turned on for all users.</span></span> <span data-ttu-id="d97a8-111">Há duas subconjunções, que o administrador pode ativar e desativar.</span><span class="sxs-lookup"><span data-stu-id="d97a8-111">There are two sub-settings, which the admin can turn on and off.</span></span>

* <span data-ttu-id="d97a8-112">Capacidade de compartilhar scripts dentro da organização</span><span class="sxs-lookup"><span data-stu-id="d97a8-112">Ability to share scripts within the organization</span></span>
* <span data-ttu-id="d97a8-113">Capacidade de usar scripts em Power Automate</span><span class="sxs-lookup"><span data-stu-id="d97a8-113">Ability to use scripts in Power Automate</span></span>

<span data-ttu-id="d97a8-114">Você pode saber se você tem acesso Office scripts abrindo um arquivo no Excel na Web (navegador) e vendo se a guia Automatizar aparece na faixa de opções Excel ou não. </span><span class="sxs-lookup"><span data-stu-id="d97a8-114">You can tell if you have access to Office Scripts by opening a file in Excel on the web (browser) and seeing if the **Automate** tab appears in the Excel ribbon or not.</span></span>
<span data-ttu-id="d97a8-115">Se você ainda não conseguir ver a guia **Automatizar,** verifique [esta seção de solução de problemas](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span><span class="sxs-lookup"><span data-stu-id="d97a8-115">If you still can't see the **Automate** tab, check [this troubleshooting section](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>

### <a name="availability"></a><span data-ttu-id="d97a8-116">Disponibilidade</span><span class="sxs-lookup"><span data-stu-id="d97a8-116">Availability</span></span>

<span data-ttu-id="d97a8-117">Office Os scripts estão disponíveis apenas no Excel na Web para licenças Enterprise E3+ (contas de consumidor e E1 não são suportadas).</span><span class="sxs-lookup"><span data-stu-id="d97a8-117">Office Scripts is available only in the Excel on the web for Enterprise E3+ licenses (Consumer and E1 accounts are not supported).</span></span> <span data-ttu-id="d97a8-118">Office Os scripts ainda não são suportados no Excel no Windows e no Mac.</span><span class="sxs-lookup"><span data-stu-id="d97a8-118">Office Scripts is not yet supported in Excel on Windows and Mac.</span></span>

### <a name="scripts-and-editor"></a><span data-ttu-id="d97a8-119">Scripts e editor</span><span class="sxs-lookup"><span data-stu-id="d97a8-119">Scripts and editor</span></span>

<span data-ttu-id="d97a8-120">O editor de código é integrado Excel na Web (versão online).</span><span class="sxs-lookup"><span data-stu-id="d97a8-120">The code editor is built right into Excel on the web (online version).</span></span> <span data-ttu-id="d97a8-121">Se você tiver usado editores como Visual Studio Code ou Sublime, essa experiência de edição será bastante semelhante.</span><span class="sxs-lookup"><span data-stu-id="d97a8-121">If you have used editors like Visual Studio Code or Sublime, this editing experience will be quite similar.</span></span>
<span data-ttu-id="d97a8-122">A maioria das teclas de atalho que Visual Studio Code editor usa funciona na experiência de edição Office scripts também.</span><span class="sxs-lookup"><span data-stu-id="d97a8-122">Most of the shortcut keys that Visual Studio Code editor uses work in the Office Scripts editing experience as well.</span></span> <span data-ttu-id="d97a8-123">Confira os seguintes apostilas de teclas de atalho.</span><span class="sxs-lookup"><span data-stu-id="d97a8-123">Check out the following shortcut keys handouts.</span></span>

* [<span data-ttu-id="d97a8-124">macOS</span><span class="sxs-lookup"><span data-stu-id="d97a8-124">macOS</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)
* [<span data-ttu-id="d97a8-125">Windows</span><span class="sxs-lookup"><span data-stu-id="d97a8-125">Windows</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

#### <a name="key-things-to-note"></a><span data-ttu-id="d97a8-126">Principais coisas a observar</span><span class="sxs-lookup"><span data-stu-id="d97a8-126">Key things to note</span></span>

* <span data-ttu-id="d97a8-127">Office Os scripts só estão disponíveis para arquivos armazenados em OneDrive for Business, SharePoint sites e sites de equipe.</span><span class="sxs-lookup"><span data-stu-id="d97a8-127">Office Scripts is only available for files stored in OneDrive for Business, SharePoint sites, and Team sites.</span></span>
* <span data-ttu-id="d97a8-128">O editor não mostra a extensão do script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-128">The editor doesn't show the script's extension.</span></span> <span data-ttu-id="d97a8-129">Na verdade, esses são arquivos TypeScript, mas eles são armazenados com uma extensão personalizada chamada `.osts` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-129">In reality, these are TypeScript files but they are stored with a custom extension called `.osts`.</span></span>
* <span data-ttu-id="d97a8-130">Os scripts são armazenados em sua própria pasta OneDrive for Business `My Files/Documents/OfficeScripts` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-130">The scripts are stored in your own OneDrive for Business folder `My Files/Documents/OfficeScripts`.</span></span> <span data-ttu-id="d97a8-131">Você não precisará gerenciar essa pasta.</span><span class="sxs-lookup"><span data-stu-id="d97a8-131">You won't need to manage this folder.</span></span> <span data-ttu-id="d97a8-132">Por sua parte, você pode ignorar esse aspecto enquanto o editor gerencia a experiência de exibição/edição.</span><span class="sxs-lookup"><span data-stu-id="d97a8-132">For your part, you can ignore this aspect as the editor manages the viewing/editing experience.</span></span>
* <span data-ttu-id="d97a8-133">Os scripts não são armazenados como parte Excel arquivos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-133">Scripts are not stored as part of Excel files.</span></span> <span data-ttu-id="d97a8-134">Eles são armazenados separadamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-134">They are stored separately.</span></span>
* <span data-ttu-id="d97a8-135">Você pode compartilhar o script com um arquivo Excel que, na verdade, significa que você está vinculando o script com o arquivo, não anexando-o.</span><span class="sxs-lookup"><span data-stu-id="d97a8-135">You can share the script with an Excel file which in effect means you are linking the script with the file, not attaching it.</span></span> <span data-ttu-id="d97a8-136">Quem tiver acesso ao arquivo Excel também poderá **exibir,** executar **ou** fazer uma **cópia** do script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-136">Whoever has access to the Excel file will also be able to **view**, **run**, or **make a copy** of the script.</span></span> <span data-ttu-id="d97a8-137">Essa é uma diferença importante em comparação com macros VBA.</span><span class="sxs-lookup"><span data-stu-id="d97a8-137">This is a key difference compared to VBA macros.</span></span>
* <span data-ttu-id="d97a8-138">A menos que você compartilhe seus scripts, ninguém mais poderá acessá-los como ele reside em sua própria biblioteca.</span><span class="sxs-lookup"><span data-stu-id="d97a8-138">Unless you share your scripts, no one else can access it as it resides in your own library.</span></span>
* <span data-ttu-id="d97a8-139">Os scripts não podem ser vinculados a partir de um disco local ou locais de nuvem personalizados.</span><span class="sxs-lookup"><span data-stu-id="d97a8-139">Scripts can't be linked from a local disk or custom cloud locations.</span></span> <span data-ttu-id="d97a8-140">Office Os scripts só reconhecem e executam um script que está em local predefinido (sua pasta OneDrive acima) ou scripts compartilhados.</span><span class="sxs-lookup"><span data-stu-id="d97a8-140">Office Scripts only recognizes and runs a script that is on predefined location (your OneDrive folder mentioned above) or shared scripts.</span></span>
* <span data-ttu-id="d97a8-141">Durante a edição, os arquivos são temporariamente salvos no navegador, mas você terá que salvar o script antes de fechar a janela Excel para salvá-lo no local OneDrive local.</span><span class="sxs-lookup"><span data-stu-id="d97a8-141">During editing, files are temporarily saved in the browser but you'll have to save the script before closing the Excel window to save it to the OneDrive location.</span></span> <span data-ttu-id="d97a8-142">Não se esqueça de salvar o arquivo após as edições.</span><span class="sxs-lookup"><span data-stu-id="d97a8-142">Don't forget to save the file after edits.</span></span>

## <a name="gentle-introduction-to-scripting"></a><span data-ttu-id="d97a8-143">Introdução suave ao script</span><span class="sxs-lookup"><span data-stu-id="d97a8-143">Gentle introduction to scripting</span></span>

<span data-ttu-id="d97a8-144">Office Scripts são scripts autônomos escritos no idioma TypeScript que contêm instruções para executar alguma automação em relação à Excel de trabalho selecionada.</span><span class="sxs-lookup"><span data-stu-id="d97a8-144">Office Scripts are standalone scripts written in the TypeScript language that contain instructions to perform some automation against the selected Excel workbook.</span></span> <span data-ttu-id="d97a8-145">Todas as instruções de automação são autoconstrutivas em um script e os scripts não podem invocar ou chamar outros scripts.</span><span class="sxs-lookup"><span data-stu-id="d97a8-145">All automation instructions are self-contained within a script and scripts can't invoke or call other scripts.</span></span> <span data-ttu-id="d97a8-146">Todos os scripts são armazenados em arquivos autônomos e armazenados na pasta de OneDrive do usuário.</span><span class="sxs-lookup"><span data-stu-id="d97a8-146">All scripts are stored in standalone files and stored on the user's OneDrive folder.</span></span> <span data-ttu-id="d97a8-147">Você pode gravar um novo script, editar um script gravado ou gravar todo um novo script do zero, tudo dentro de uma interface de editor integrado.</span><span class="sxs-lookup"><span data-stu-id="d97a8-147">You can record a new script, edit a recorded script, or write a whole new script from scratch, all within a built-in editor interface.</span></span> <span data-ttu-id="d97a8-148">A melhor parte Office scripts é que eles não precisam de mais configuração dos usuários.</span><span class="sxs-lookup"><span data-stu-id="d97a8-148">The best part of Office Scripts is that they don't need any further setup from users.</span></span> <span data-ttu-id="d97a8-149">Sem bibliotecas externas, páginas da Web ou elementos de interface do usuário, configuração etc. Toda a configuração do ambiente é manipulada por Office Scripts e permite acesso fácil e rápido à automação por meio de uma interface de API simples.</span><span class="sxs-lookup"><span data-stu-id="d97a8-149">No external libraries, web pages, or UI elements, setup, etc. All the environment setup is handled by Office Scripts and it allows easy and fast access to automation through a simple API interface.</span></span>

<span data-ttu-id="d97a8-150">Alguns dos conceitos básicos úteis para entender como editar e navegar em torno de scripts incluem:</span><span class="sxs-lookup"><span data-stu-id="d97a8-150">Some of the basic concepts helpful to understand how to edit and navigate around scripts include:</span></span>

* <span data-ttu-id="d97a8-151">Sintaxe básica da linguagem TypeScript</span><span class="sxs-lookup"><span data-stu-id="d97a8-151">Basic TypeScript language syntax</span></span>
* <span data-ttu-id="d97a8-152">Noções `main` básicas sobre função e argumentos</span><span class="sxs-lookup"><span data-stu-id="d97a8-152">Understanding of `main` function and arguments</span></span>
* <span data-ttu-id="d97a8-153">Objetos e hierarquia, métodos, propriedades</span><span class="sxs-lookup"><span data-stu-id="d97a8-153">Objects and hierarchy, methods, properties</span></span>
* <span data-ttu-id="d97a8-154">Coleção (matriz): navegação e operações</span><span class="sxs-lookup"><span data-stu-id="d97a8-154">Collection (array): navigation and operations</span></span>
* <span data-ttu-id="d97a8-155">Definições de tipo</span><span class="sxs-lookup"><span data-stu-id="d97a8-155">Type definitions</span></span>
* <span data-ttu-id="d97a8-156">Ambiente: registro/edição, executar, examinar resultados, compartilhar</span><span class="sxs-lookup"><span data-stu-id="d97a8-156">Environment: record/edit, run, examine results, share</span></span>

<span data-ttu-id="d97a8-157">Este vídeo e seção explicam alguns desses conceitos em detalhes.</span><span class="sxs-lookup"><span data-stu-id="d97a8-157">This video and section explain some of these concepts in detail.</span></span>

<span data-ttu-id="d97a8-158">[![Noções básicas Office scripts](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Noções básicas de scripts")</span><span class="sxs-lookup"><span data-stu-id="d97a8-158">[![Basics of Office Scripts](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Basics of Scripts")</span></span>

### <a name="language-typescript"></a><span data-ttu-id="d97a8-159">Idioma: TypeScript</span><span class="sxs-lookup"><span data-stu-id="d97a8-159">Language: TypeScript</span></span>

<span data-ttu-id="d97a8-160">[Office Scripts](../../index.md) é escrito usando a linguagem [TypeScript](https://www.typescriptlang.org/), que é uma linguagem de código aberto que se cria em JavaScript (uma das mais usadas do mundo) adicionando definições de tipo estático.</span><span class="sxs-lookup"><span data-stu-id="d97a8-160">[Office Scripts](../../index.md) is written using the [TypeScript language](https://www.typescriptlang.org/), which is an open-source language that builds on JavaScript (one of the world's most used) by adding static type definitions.</span></span> <span data-ttu-id="d97a8-161">Como diz o site, forneça uma maneira de descrever a forma de um objeto, fornecendo uma documentação melhor e permitindo que TypeScript valide que seu `Types` código está funcionando corretamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-161">As the website says, `Types` provide a way to describe the shape of an object, providing better documentation, and allowing TypeScript to validate that your code is working correctly.</span></span>

<span data-ttu-id="d97a8-162">A sintaxe de idioma em si é escrita usando [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) com tipificações adicionais definidas no script usando convenções TypeScript.</span><span class="sxs-lookup"><span data-stu-id="d97a8-162">The language syntax itself is written using [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) with additional typings defined in the script using TypeScript conventions.</span></span> <span data-ttu-id="d97a8-163">Na maioria das vezes, você pode pensar em Office scripts como escritos em JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d97a8-163">For the most part, you can think of Office Scripts as written in JavaScript.</span></span> <span data-ttu-id="d97a8-164">É essencial que você entenda as noções básicas da linguagem JavaScript para iniciar sua jornada Office Scripts; embora você não precise ser proficiente para começar sua jornada de automação.</span><span class="sxs-lookup"><span data-stu-id="d97a8-164">It is essential that you understand the basics of JavaScript language to begin your Office Scripts journey; though you don't need to be proficient at it to begin your automation journey.</span></span> <span data-ttu-id="d97a8-165">Com o Office de ação scripts, você pode entender as instruções de script porque os comentários de código estão incluídos e você pode acompanhar e fazer pequenas edições.</span><span class="sxs-lookup"><span data-stu-id="d97a8-165">With the Office Scripts' action recorder, you can understand the script statements because code comments are included and you can follow along and make small edits.</span></span>

<span data-ttu-id="d97a8-166">Office As APIs de scripts, que permitem que o script interaja com Excel, são projetadas para usuários finais que podem não ter muito plano de fundo de codificação.</span><span class="sxs-lookup"><span data-stu-id="d97a8-166">Office Scripts APIs, which allow the script to interact with Excel, are designed for end-users who may not have much coding background.</span></span> <span data-ttu-id="d97a8-167">AS APIs podem ser invocadas de forma síncrona e você não precisa conhecer tópicos avançados, como promessas ou retornos de chamada.</span><span class="sxs-lookup"><span data-stu-id="d97a8-167">APIs can be invoked synchronously and you don't need to know advanced topics such as promises or callbacks.</span></span> <span data-ttu-id="d97a8-168">Office O design da API de scripts fornece:</span><span class="sxs-lookup"><span data-stu-id="d97a8-168">Office Scripts API design provides:</span></span>

* <span data-ttu-id="d97a8-169">Modelo de objeto simples com métodos, getters/setters.</span><span class="sxs-lookup"><span data-stu-id="d97a8-169">Simple object model with methods, getters/setters.</span></span>
* <span data-ttu-id="d97a8-170">Coleções de objetos de fácil acesso como matrizes regulares.</span><span class="sxs-lookup"><span data-stu-id="d97a8-170">Easy-to-access object collections as regular arrays.</span></span>
* <span data-ttu-id="d97a8-171">Opções simples de tratamento de erros.</span><span class="sxs-lookup"><span data-stu-id="d97a8-171">Simple error handling options.</span></span>
* <span data-ttu-id="d97a8-172">Desempenho otimizado para cenários selecionados que ajudam os usuários a se concentrarem no cenário em mãos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-172">Optimized performance for select scenarios helping users to focus on the scenario at hand.</span></span>

### <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="d97a8-173">`main` função: o ponto de partida do script</span><span class="sxs-lookup"><span data-stu-id="d97a8-173">`main` function: The script's starting point</span></span>

<span data-ttu-id="d97a8-174">Office A execução de scripts começa na `main` função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-174">Office Scripts' execution begins at the `main` function.</span></span> <span data-ttu-id="d97a8-175">Um script é um único arquivo que contém uma ou muitas funções, juntamente com declarações de tipos, interfaces, variáveis, etc. Para seguir junto com o script, comece com a função como Excel sempre invoca a função primeiro `main` quando você executa qualquer `main` script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-175">A script is a single file containing one or many functions along with declarations of types, interfaces, variables, etc. To follow along with the script, begin with the `main` function as Excel always first invokes the `main` function when you execute any script.</span></span> <span data-ttu-id="d97a8-176">A função sempre terá pelo menos um argumento (ou parâmetro) chamado , que é apenas um nome de variável que identifica a agenda de trabalho atual com a qual o `main` `workbook` script está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="d97a8-176">The `main` function will always have at least one argument (or parameter) named `workbook`, which is just a variable name identifying the current workbook against which the script is running.</span></span> <span data-ttu-id="d97a8-177">Você pode definir argumentos adicionais para uso com Power Automate execução (offline).</span><span class="sxs-lookup"><span data-stu-id="d97a8-177">You can define additional arguments for usage with Power Automate (offline) execution.</span></span>

* `function main(workbook: ExcelScript.Workbook)`

<span data-ttu-id="d97a8-178">Um script pode ser organizado em funções menores para ajudar na reutilização de código, clareza, etc. Outras funções podem estar dentro ou fora da função principal, mas sempre no mesmo arquivo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-178">A script can be organized into smaller functions to aid with code reusability, clarity, etc. Other functions can be inside or outside of the main function but always in the same file.</span></span> <span data-ttu-id="d97a8-179">Um script é autoconstruido e só pode usar funções definidas no mesmo arquivo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-179">A script is self-contained and can only use functions defined in the same file.</span></span> <span data-ttu-id="d97a8-180">Scripts não podem invocar ou chamar outro Office Script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-180">Scripts cannot invoke or call another Office Script.</span></span>

<span data-ttu-id="d97a8-181">Portanto, em resumo:</span><span class="sxs-lookup"><span data-stu-id="d97a8-181">So, in summary:</span></span>

* <span data-ttu-id="d97a8-182">A `main` função é o ponto de entrada para qualquer script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-182">The `main` function is the entry point for any script.</span></span> <span data-ttu-id="d97a8-183">Quando a função é executada, o aplicativo Excel invoca essa função principal fornecendo a workbook como seu primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-183">When the function is executed, the Excel application invokes this main function by providing the workbook as its first parameter.</span></span>
* <span data-ttu-id="d97a8-184">É importante manter o primeiro argumento e `workbook` sua declaração de tipo como ele aparece.</span><span class="sxs-lookup"><span data-stu-id="d97a8-184">It's important to keep the first argument `workbook` and its type declaration as it appears.</span></span> <span data-ttu-id="d97a8-185">Você pode adicionar novos argumentos à função (consulte a próxima seção), mas mantenha o `main` primeiro argumento como está.</span><span class="sxs-lookup"><span data-stu-id="d97a8-185">You can add new arguments to the `main` function (see the next section) but do keep the first argument as is.</span></span>

:::image type="content" source="../../images/getting-started-main-introduction.png" alt-text="A função principal é o ponto de entrada do script":::

#### <a name="send-or-receive-data-from-other-apps"></a><span data-ttu-id="d97a8-187">Enviar ou receber dados de outros aplicativos</span><span class="sxs-lookup"><span data-stu-id="d97a8-187">Send or receive data from other apps</span></span>

<span data-ttu-id="d97a8-188">Você pode conectar Excel outras partes da sua organização executando scripts em [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="d97a8-188">You can connect Excel to other parts of your organization by running scripts in [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="d97a8-189">Saiba mais sobre [como executar Office scripts em Power Automate fluxos](../../develop/power-automate-integration.md).</span><span class="sxs-lookup"><span data-stu-id="d97a8-189">Learn more about [running Office Scripts in Power Automate flows](../../develop/power-automate-integration.md).</span></span>

<span data-ttu-id="d97a8-190">A maneira de receber ou enviar dados de e para Excel é por meio da `main` função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-190">The way to receive or send data from and to Excel is through the `main` function.</span></span> <span data-ttu-id="d97a8-191">Pense nele como o gateway de informações que permite que os dados de entrada e de saída sejam descritos e usados no script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-191">Think of it as the information gateway that allows incoming and outgoing data to be described and used in the script.</span></span> <span data-ttu-id="d97a8-192">Você pode receber dados de fora do script usando o tipo de dados e retornar quaisquer dados reconhecidos pelo TypeScript, como , , ou quaisquer objetos na forma de interfaces que você `string` `string` definir no `number` `boolean` script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-192">You can receive data from outside the script using the `string` data type and return any TypeScript-recognized data such as `string`, `number`, `boolean`, or any objects in the form of interfaces you define in the script.</span></span>

:::image type="content" source="../../images/getting-started-data-in-out.png" alt-text="As entradas e saídas de um script":::

#### <a name="use-functions-to-organize-and-reuse-code"></a><span data-ttu-id="d97a8-194">Usar funções para organizar e reutilizar código</span><span class="sxs-lookup"><span data-stu-id="d97a8-194">Use functions to organize and reuse code</span></span>

<span data-ttu-id="d97a8-195">Você pode usar funções para organizar e reutilizar código em seu script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-195">You can use functions to organize and reuse code within your script.</span></span>

:::image type="content" source="../../images/getting-started-use-functions.png" alt-text="Usando funções em um script":::

### <a name="objects-hierarchy-methods-properties-collections"></a><span data-ttu-id="d97a8-197">Objetos, hierarquia, métodos, propriedades, coleções</span><span class="sxs-lookup"><span data-stu-id="d97a8-197">Objects, hierarchy, methods, properties, collections</span></span>

<span data-ttu-id="d97a8-198">Todo o Excel de objeto de Excel é definido em uma estrutura hierárquica de objetos, começando com o objeto workbook do tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-198">All of Excel's object model is defined in a hierarchical structure of objects, beginning with the workbook object of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="d97a8-199">Um objeto pode conter métodos, propriedades e outros objetos dentro dele.</span><span class="sxs-lookup"><span data-stu-id="d97a8-199">An object can contain methods, properties, and other objects within it.</span></span> <span data-ttu-id="d97a8-200">Os objetos são vinculados uns aos outros usando os métodos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-200">Objects are linked to each other using the methods.</span></span> <span data-ttu-id="d97a8-201">O método de um objeto pode retornar outro objeto ou coleção de objetos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-201">An object's method can return another object or collection of objects.</span></span> <span data-ttu-id="d97a8-202">Usar o recurso de IntelliSense (conclusão de código) do editor de código é uma ótima maneira de explorar a hierarquia de objetos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-202">Using the code editor's IntelliSense (code completion) feature is a great way to explore the object hierarchy.</span></span> <span data-ttu-id="d97a8-203">Você também pode usar o [site de documentação de referência oficial](/javascript/api/office-scripts/overview) para acompanhar as relações entre objetos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-203">You can also use the [official reference documentation site](/javascript/api/office-scripts/overview) to follow along with the relationships among objects.</span></span>

<span data-ttu-id="d97a8-204">Um [objeto](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) é uma coleção de propriedades e uma propriedade é uma associação entre um nome (ou chave) e um valor.</span><span class="sxs-lookup"><span data-stu-id="d97a8-204">An [object](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) is a collection of properties, and a property is an association between a name (or key) and a value.</span></span> <span data-ttu-id="d97a8-205">O valor de uma propriedade pode ser uma função, nesse caso, a propriedade é conhecida como um método.</span><span class="sxs-lookup"><span data-stu-id="d97a8-205">A property's value can be a function, in which case the property is known as a method.</span></span> <span data-ttu-id="d97a8-206">No caso do modelo de objeto Office Scripts, um objeto representa uma coisa no arquivo Excel que os usuários interagem com um gráfico, hiperlink, tabela dinâmica etc. Ele também pode representar o comportamento de um objeto, como os atributos de proteção de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="d97a8-206">In the case of the Office Scripts object model, an object represents a thing in the Excel file that users interact with such as a chart, hyperlink, pivot-table, etc. It can also represent the behavior of an object such as the protection attributes of a worksheet.</span></span>

<span data-ttu-id="d97a8-207">O tópico de objetos TypeScript e propriedades vs métodos é bastante profundo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-207">The topic of TypeScript objects and properties vs methods is quite deep.</span></span> <span data-ttu-id="d97a8-208">Para começar a usar o script e ser produtivo, você pode se lembrar de algumas coisas básicas:</span><span class="sxs-lookup"><span data-stu-id="d97a8-208">In order to get started with the script and be productive, you can remember a few basic things:</span></span>

* <span data-ttu-id="d97a8-209">Ambos os objetos e propriedades são acessados usando notação (ponto), com o objeto no lado esquerdo do e a propriedade ou método `.` `.` no lado direito.</span><span class="sxs-lookup"><span data-stu-id="d97a8-209">Both objects and properties are accessed using `.` (dot) notation, with the object on the left side of the `.` and the property or method on the right side.</span></span> <span data-ttu-id="d97a8-210">Exemplos: `hyperlink.address` , `range.getAddress()` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-210">Examples: `hyperlink.address`, `range.getAddress()`.</span></span>
* <span data-ttu-id="d97a8-211">As propriedades são escalares na natureza (cadeias de caracteres, booleanos, números).</span><span class="sxs-lookup"><span data-stu-id="d97a8-211">Properties are scalar in nature (strings, booleans, numbers).</span></span> <span data-ttu-id="d97a8-212">Por exemplo, nome de uma pasta de trabalho, posição de uma planilha, o valor de se a tabela tem um rodapé ou não.</span><span class="sxs-lookup"><span data-stu-id="d97a8-212">For example, name of a workbook, position of a worksheet, the value of whether the table has a footer or not.</span></span>
* <span data-ttu-id="d97a8-213">Os métodos são 'invocados' ou 'executados' usando os parênteses de fechamento aberto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-213">Methods are 'invoked' or 'executed' using the open-close parentheses.</span></span> <span data-ttu-id="d97a8-214">Exemplo: `table.delete()`.</span><span class="sxs-lookup"><span data-stu-id="d97a8-214">Example: `table.delete()`.</span></span> <span data-ttu-id="d97a8-215">Às vezes, um argumento é passado para uma função incluindo-os entre parênteses de fechamento aberto: `range.setValue('Hello')` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-215">Sometimes an argument is passed to a function by including them between open-close parentheses: `range.setValue('Hello')`.</span></span> <span data-ttu-id="d97a8-216">Você pode passar muitos argumentos para uma função (conforme definido por seu contrato/assinatura) e separá-los usando `,` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-216">You can pass many arguments to a function (as defined by its contract/signature) and separate them using `,`.</span></span>  <span data-ttu-id="d97a8-217">Por exemplo: `worksheet.addTable('A1:D6', true)`.</span><span class="sxs-lookup"><span data-stu-id="d97a8-217">For example: `worksheet.addTable('A1:D6', true)`.</span></span> <span data-ttu-id="d97a8-218">Você pode passar argumentos de qualquer tipo conforme exigido pelo método, como cadeias de caracteres, número, booleano ou até mesmo outros objetos, por exemplo, , onde é um objeto criado em outro lugar no `worksheet.addTable(targetRange, true)` `targetRange` script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-218">You can pass arguments of any type as required by the method such as strings, number, boolean, or even other objects, for example, `worksheet.addTable(targetRange, true)`, where `targetRange` is an object created elsewhere in the script.</span></span>
* <span data-ttu-id="d97a8-219">Os métodos podem retornar uma coisa como uma propriedade escalar (nome, endereço, etc.) ou outro objeto (intervalo, gráfico) ou não retornar nada (como o caso com `delete` métodos).</span><span class="sxs-lookup"><span data-stu-id="d97a8-219">Methods can return a thing such as a scalar property (name, address, etc.) or another object (range, chart), or not return anything at all (such as the case with `delete` methods).</span></span> <span data-ttu-id="d97a8-220">Você recebe o que o método retorna declarando uma variável ou atribuindo a uma variável existente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-220">You receive what the method returns by declaring a variable or assigning to an existing variable.</span></span> <span data-ttu-id="d97a8-221">Você pode ver isso no lado esquerdo da instrução, como `const table = worksheet.addTable('A1:D6', true)` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-221">You can see that on the left hand side of statement such as `const table = worksheet.addTable('A1:D6', true)`.</span></span>
* <span data-ttu-id="d97a8-222">Na maior parte, o modelo de objeto Office Scripts consiste em objetos com métodos que vinculam várias partes do Excel modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-222">For the most part, the Office Scripts object model consists of objects with methods that link various parts of the Excel object model.</span></span> <span data-ttu-id="d97a8-223">Muito raramente você se depara com propriedades que são de valores escalares ou de objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-223">Very rarely you'll come across properties that are of scalar or object values.</span></span>
* <span data-ttu-id="d97a8-224">Em Office Scripts, um método Excel modelo de objeto deve conter parênteses abertos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-224">In Office Scripts, an Excel object model method has to contain open-close parentheses.</span></span> <span data-ttu-id="d97a8-225">O uso de métodos sem eles não é permitido (como atribuir um método a uma variável).</span><span class="sxs-lookup"><span data-stu-id="d97a8-225">Using methods without them is not allowed (such as assigning a method to a variable).</span></span>

<span data-ttu-id="d97a8-226">Vamos ver alguns métodos no `workbook` objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-226">Let's look at a few methods on the `workbook` object.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Return a boolean (true or false) setting of whether the workbook is set to auto-save or not. 
    const autoSave = workbook.getAutoSave(); 
    // Get workbook name.
    const name = workbook.getName();
    // Get active cell range object.
    const cell = workbook.getActiveCell();
    // Get table named SALES.
    const cell = workbook.getTable('SALES');
    // Get all slicer objects.
    const slicers = workbook.getSlicers();
}
```

<span data-ttu-id="d97a8-227">Neste exemplo:</span><span class="sxs-lookup"><span data-stu-id="d97a8-227">In this example:</span></span>

* <span data-ttu-id="d97a8-228">Os métodos do `workbook` objeto, como e retornam uma propriedade `getAutoSave()` escalar `getName()` (cadeia de caracteres, número, booleano).</span><span class="sxs-lookup"><span data-stu-id="d97a8-228">The methods of the `workbook` object such as `getAutoSave()` and `getName()` return a scalar property (string, number, boolean).</span></span>
* <span data-ttu-id="d97a8-229">Métodos como retornar `getActiveCell()` outro objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-229">Methods such as `getActiveCell()` return another object.</span></span>
* <span data-ttu-id="d97a8-230">O `getTable()` método aceita um argumento (nome da tabela neste caso) e retorna uma tabela específica na caixa de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d97a8-230">The `getTable()` method accepts an argument (table name in this case) and returns a specific table in the workbook.</span></span>
* <span data-ttu-id="d97a8-231">O método retorna uma matriz (referida em muitos lugares como uma coleção) de todos os objetos `getSlicers()` slicer dentro da lista de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d97a8-231">The `getSlicers()` method returns an array (referred to in many places as a collection) of all slicer objects within the workbook.</span></span>

<span data-ttu-id="d97a8-232">Você observará que todos esses métodos têm um prefixo, que é apenas uma convenção usada no modelo de objeto Office Scripts para transmitir que o método está retornando `get` algo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-232">You'll notice that all of these methods have a `get` prefix, which is just a convention used in the Office Scripts object model to convey that the method is returning something.</span></span> <span data-ttu-id="d97a8-233">Eles também são comumente chamados de "getters".</span><span class="sxs-lookup"><span data-stu-id="d97a8-233">They are also commonly referred to as 'getters'.</span></span>

<span data-ttu-id="d97a8-234">Há dois outros tipos de métodos que veremos agora no próximo exemplo:</span><span class="sxs-lookup"><span data-stu-id="d97a8-234">There are two other types of methods that we'll now see in the next example:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get a worksheet named 'Sheet1.
    const sheet = workbook.getWorksheet('Sheet1'); 
    // Set name to SALES.
    sheet.setName('SALES');
    // Position the worksheet at the beginning.
    sheet.setPosition(0);
}
```

<span data-ttu-id="d97a8-235">Neste exemplo:</span><span class="sxs-lookup"><span data-stu-id="d97a8-235">In this example:</span></span>

* <span data-ttu-id="d97a8-236">O `setName()` método define um novo nome para a planilha.</span><span class="sxs-lookup"><span data-stu-id="d97a8-236">The `setName()` method sets a new name to the worksheet.</span></span> <span data-ttu-id="d97a8-237">`setPosition()` define a posição como a primeira célula.</span><span class="sxs-lookup"><span data-stu-id="d97a8-237">`setPosition()` sets the position to the first cell.</span></span>
* <span data-ttu-id="d97a8-238">Esses métodos modificam o arquivo Excel configurando uma propriedade ou comportamento da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d97a8-238">Such methods modify the Excel file by setting a property or behavior of the workbook.</span></span> <span data-ttu-id="d97a8-239">Esses métodos são chamados de "setters".</span><span class="sxs-lookup"><span data-stu-id="d97a8-239">These methods are called 'setters'.</span></span>
* <span data-ttu-id="d97a8-240">Normalmente, os "setters" têm um "getter" de parceiro, por exemplo, e , ambos `worksheet.getPosition` `worksheet.setPosition` são métodos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-240">Typically 'setters' have a companion 'getter', for example, `worksheet.getPosition` and `worksheet.setPosition`, both of which are methods.</span></span>

#### <a name="undefined-and-null-primitive-types"></a><span data-ttu-id="d97a8-241">`undefined` e `null` tipos primitivos</span><span class="sxs-lookup"><span data-stu-id="d97a8-241">`undefined` and `null` primitive types</span></span>

<span data-ttu-id="d97a8-242">Veja a seguir dois tipos de dados primitivos que você deve estar ciente:</span><span class="sxs-lookup"><span data-stu-id="d97a8-242">The following are two primitive data types that you must be aware of:</span></span>

1. <span data-ttu-id="d97a8-243">O valor [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) representa a ausência intencional de qualquer valor de objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-243">The value [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) represents the intentional absence of any object value.</span></span> <span data-ttu-id="d97a8-244">É um dos valores primitivos do JavaScript e é usado para indicar que uma variável não tem valor.</span><span class="sxs-lookup"><span data-stu-id="d97a8-244">It is one of JavaScript's primitive values and is used to indicate that a variable has no value.</span></span>
1. <span data-ttu-id="d97a8-245">Uma variável que não foi atribuída a um valor é do tipo [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined) .</span><span class="sxs-lookup"><span data-stu-id="d97a8-245">A variable that has not been assigned a value is of type [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined).</span></span> <span data-ttu-id="d97a8-246">Um método ou instrução também pode retornar se a variável `undefined` avaliada não tiver um valor atribuído.</span><span class="sxs-lookup"><span data-stu-id="d97a8-246">A method or statement can also return `undefined` if the variable that's being evaluated doesn't have an assigned value.</span></span>

<span data-ttu-id="d97a8-247">Esses dois tipos são recortados como parte do tratamento de erros e podem causar bastante dor de cabeça se não for tratado corretamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-247">These two types crop up as part of error handling and can cause quite a bit of headache if not handled properly.</span></span> <span data-ttu-id="d97a8-248">Felizmente, TypeScript/JavaScript oferece uma maneira de verificar se uma variável é do tipo `undefined` ou `null` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-248">Fortunately, TypeScript/JavaScript offers a way to check if a variable is of type `undefined` or `null`.</span></span> <span data-ttu-id="d97a8-249">Vamos falar sobre algumas dessas verificações em seções posteriores, incluindo o tratamento de erros.</span><span class="sxs-lookup"><span data-stu-id="d97a8-249">We will talk about some of those checks in later sections, including error handling.</span></span>

#### <a name="method-chaining"></a><span data-ttu-id="d97a8-250">Encadeamento de método</span><span class="sxs-lookup"><span data-stu-id="d97a8-250">Method chaining</span></span>

<span data-ttu-id="d97a8-251">Você pode usar a notação de ponto para conectar objetos que estão sendo retornados de um método para reduzir seu código.</span><span class="sxs-lookup"><span data-stu-id="d97a8-251">You can use dot notation to connect objects being returned from a method to shorten your code.</span></span> <span data-ttu-id="d97a8-252">Às vezes, essa técnica torna o código fácil de ler e gerenciar.</span><span class="sxs-lookup"><span data-stu-id="d97a8-252">Sometimes this technique makes the code easy to read and manage.</span></span> <span data-ttu-id="d97a8-253">No entanto, há poucas coisas a serem cientes.</span><span class="sxs-lookup"><span data-stu-id="d97a8-253">However, there are few things to be aware of.</span></span> <span data-ttu-id="d97a8-254">Vejamos os exemplos a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-254">Let's look at the following examples.</span></span>

<span data-ttu-id="d97a8-255">O código a seguir obtém a célula ativa e a próxima célula e define o valor.</span><span class="sxs-lookup"><span data-stu-id="d97a8-255">The following code gets the active cell and the next cell, then sets the value.</span></span> <span data-ttu-id="d97a8-256">Esse é um bom candidato para usar encadeamento, pois esse código terá êxito o tempo todo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-256">This is a good candidate to use chaining as this code will succeed all the time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getActiveCell().getOffsetRange(0,1).setValue('Next cell');
}
```

<span data-ttu-id="d97a8-257">No entanto, o código a seguir (que obtém uma tabela chamada **SALES** e liga seu estilo de coluna em faixa) tem um problema.</span><span class="sxs-lookup"><span data-stu-id="d97a8-257">However, the following code (which gets a table named **SALES** and turns on its banded column style) has an issue.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  workbook.getTable('SALES').setShowBandedColumns(true);
}
```

<span data-ttu-id="d97a8-258">E se a **tabela SALES** não existir?</span><span class="sxs-lookup"><span data-stu-id="d97a8-258">What if the **SALES** table doesn't exist?</span></span> <span data-ttu-id="d97a8-259">O script falhará com um erro (mostrado a seguir) porque retorna (que é um `getTable('SALES')` tipo JavaScript indicando que não há tabela como `undefined` **SALES**).</span><span class="sxs-lookup"><span data-stu-id="d97a8-259">The script will fail with an error (shown next) because `getTable('SALES')` returns `undefined` (which is a JavaScript type indicating that there is no table such as **SALES**).</span></span> <span data-ttu-id="d97a8-260">Chamar o `setShowBandedColumns` método em não faz `undefined` sentido, ou seja, `undefined.setShowBandedColumns(true)` e, portanto, o script termina em um erro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-260">Calling the `setShowBandedColumns` method on `undefined` makes no sense, that is, `undefined.setShowBandedColumns(true)`, and hence the script ends in an error.</span></span>

```text
Line 2: Cannot read property 'setShowBandedColumns' of undefined
```

<span data-ttu-id="d97a8-261">Você pode [](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) usar o operador de encadeamento opcional que fornece uma maneira de simplificar o acesso a valores por meio de objetos conectados quando for possível que uma referência ou método seja ou (que é a maneira do JavaScript indicar um objeto ou resultado não atribuído ou inexistente) para lidar com essa `undefined` `null` condição.</span><span class="sxs-lookup"><span data-stu-id="d97a8-261">You could use the [optional chaining operator](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) that provides a way to simplify accessing values through connected objects when it's possible that a reference or method may be `undefined` or `null` (which is JavaScript's way of indicating an unassigned or nonexistent object or result) to handle this condition.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // This line will not fail as the setShowBandedColumns method is executed only if the SALES table is present.
    workbook.getTable('SALES')?.setShowBandedColumns(true); 
}
```

<span data-ttu-id="d97a8-262">Se você deseja manipular condições de objeto inexistentes ou tipo que está sendo retornado por um método, é melhor atribuir o valor de retorno do método e lidar com isso `undefined` separadamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-262">If you wish to handle nonexistent object conditions or `undefined` type being returned by a method, then it is better to assign the return value from the method and handle that separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const salesTable = workbook.getTable('SALES');
    if (salesTable) {
        salesTable.setShowBandedColumns(true);
    } else { 
        // Handle this condition.
    }
}
```

#### <a name="get-object-reference"></a><span data-ttu-id="d97a8-263">Obter referência de objeto</span><span class="sxs-lookup"><span data-stu-id="d97a8-263">Get object reference</span></span>

<span data-ttu-id="d97a8-264">O `workbook` objeto é dado a você na `main` função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-264">The `workbook` object is given to you in the `main` function.</span></span> <span data-ttu-id="d97a8-265">Você pode começar a usar o `workbook` objeto e acessar seus métodos diretamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-265">You can begin to use the `workbook` object and access its methods directly.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get workbook name.
    const name = workbook.getName();
    // Display name to console.
    console.log(name);
}
```

<span data-ttu-id="d97a8-266">Para usar todos os outros objetos dentro da lista de trabalho, comece com o objeto e vá para baixo da hierarquia até chegar ao `workbook` objeto que você está procurando.</span><span class="sxs-lookup"><span data-stu-id="d97a8-266">For using all other objects within the workbook, begin with `workbook` object and go down the hierarchy until you get to the object you are looking for.</span></span> <span data-ttu-id="d97a8-267">Você pode obter a referência do objeto buscando o objeto usando seu método ou recuperando a coleção de `get` objetos, conforme mostrado abaixo:</span><span class="sxs-lookup"><span data-stu-id="d97a8-267">You can get the object reference by fetching the object using its `get` method or by retrieving the collection of objects as shown below:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    const sheet = workbook.getActiveWorksheet();
    // Fetch using an ID or key.
    const sheet = workbook.getWorksheet('SomeSheetName');
    // Invoke methods on the object.
    sheet.setPosition(0); 
    
    // Get collection of methods.
    const tables = sheet.getTables();
    console.log('Total tables in this sheet: ' + tables.length);
}
```

#### <a name="check-if-an-object-exists-then-delete-and-add"></a><span data-ttu-id="d97a8-268">Verifique se existe um objeto, exclua e adicione</span><span class="sxs-lookup"><span data-stu-id="d97a8-268">Check if an object exists, then delete, and add</span></span>

<span data-ttu-id="d97a8-269">Para criar um objeto, digamos com um nome predefinido, é sempre melhor remover um objeto semelhante que pode existir e, em seguida, adicioná-lo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-269">For creating an object, say with a predefined name, it is always better to remove a similar object that may exist and then add it.</span></span> <span data-ttu-id="d97a8-270">Você pode fazer isso usando o padrão a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-270">You can do that using the following pattern.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added. 
  let name = "Index";
  // Check if the worksheet already exists. If not, add the worksheet.
  let sheet = workbook.getWorksheet('Index');
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Call the delete method on the object to remove it. 
    sheet.delete();
  } 
    // Add a blank worksheet. 
  console.log(`Adding the worksheet named  ${name}.`)
  const indexSheet = workbook.addWorksheet("Index");
}

```

<span data-ttu-id="d97a8-271">Como alternativa, para excluir um objeto que pode ou não existir, use o padrão a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-271">Alternatively, for deleting an object that may or may not exist, use the following pattern.</span></span>

```TypeScript
    // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
    workbook.getWorksheet('Index')?.delete(); 
```

#### <a name="note-about-adding-an-object"></a><span data-ttu-id="d97a8-272">Observação sobre a adição de um objeto</span><span class="sxs-lookup"><span data-stu-id="d97a8-272">Note about adding an object</span></span>

<span data-ttu-id="d97a8-273">Para criar, inserir ou adicionar um objeto como uma slicer, tabela dinâmica, planilha etc., use o **método** add_Object_ correspondente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-273">To create, insert, or add an object such as a slicer, pivot table, worksheet, etc., use the corresponding **add_Object_** method.</span></span> <span data-ttu-id="d97a8-274">Esse método está disponível em seu objeto pai.</span><span class="sxs-lookup"><span data-stu-id="d97a8-274">Such a method is available on its parent object.</span></span> <span data-ttu-id="d97a8-275">Por exemplo, o `addChart()` método está disponível no `worksheet` objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-275">For example, the `addChart()` method is available on `worksheet` object.</span></span> <span data-ttu-id="d97a8-276">O **add_Object_** retorna o objeto que ele cria.</span><span class="sxs-lookup"><span data-stu-id="d97a8-276">The **add_Object_** method returns the object it creates.</span></span> <span data-ttu-id="d97a8-277">Receba o valor retornado e use-o posteriormente em seu script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-277">Receive the returned value and use it later in your script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Add object and get a reference to it. 
  const indexSheet = workbook.addWorksheet("Index");
  // Use it elsewhere in the script 
  console.log(indexSheet.getPosition());
}

```

<span data-ttu-id="d97a8-278">Como alternativa, para excluir um objeto que pode ou não existir, use este padrão:</span><span class="sxs-lookup"><span data-stu-id="d97a8-278">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
    workbook.getWorksheet('Index')?.delete(); // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
```

#### <a name="collections"></a><span data-ttu-id="d97a8-279">Coleções</span><span class="sxs-lookup"><span data-stu-id="d97a8-279">Collections</span></span>

<span data-ttu-id="d97a8-280">Coleções são objetos como tabelas, gráficos, colunas etc. que podem ser recuperados como uma matriz e iterados para processamento.</span><span class="sxs-lookup"><span data-stu-id="d97a8-280">Collections are objects such as tables, charts, columns, etc. that can be retrieved as an array and iterated over for processing.</span></span> <span data-ttu-id="d97a8-281">Você pode recuperar uma coleção usando o método correspondente e processar os dados em um loop usando uma das muitas técnicas de transição da matriz `get` TypeScript, como:</span><span class="sxs-lookup"><span data-stu-id="d97a8-281">You can retrieve a collection using the corresponding `get` method and process the data in a loop using one of many TypeScript array traversal techniques such as:</span></span>

* [<span data-ttu-id="d97a8-282">`for` ou `while`</span><span class="sxs-lookup"><span data-stu-id="d97a8-282">`for` or `while`</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
* [`for..of`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/for...of)
* [`forEach`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach)

* [<span data-ttu-id="d97a8-283">Noções básicas de idiomas de matrizes</span><span class="sxs-lookup"><span data-stu-id="d97a8-283">Language basics of arrays</span></span>](https://developer.mozilla.org//docs/Learn/JavaScript/First_steps/Arrays)

<span data-ttu-id="d97a8-284">Este script demonstra como usar coleções com suporte em Office SCRIPTs.</span><span class="sxs-lookup"><span data-stu-id="d97a8-284">This script demonstrates how to use collections supported in Office Scripts APIs.</span></span> <span data-ttu-id="d97a8-285">Ele colore cada guia de planilha no arquivo com uma cor aleatória.</span><span class="sxs-lookup"><span data-stu-id="d97a8-285">It colors each worksheet tab in the file with a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get all sheets as a collection.
  const sheets = workbook.getWorksheets();
  const names = sheets.map ((sheet) => sheet.getName());
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  // Get information from specific sheets within the collection.
  console.log(`First sheet name is: ${names[0]}`);
  if (sheets.length > 1) {
    console.log(`Last sheet's Id is: ${sheets[sheets.length -1].getId()}`);
  }
  // Color each worksheet with random color.
  for (const sheet of sheets) {
    sheet.setTabColor(`#${Math.random().toString(16).substr(-6)}`);
  }
}
```

## <a name="type-declarations"></a><span data-ttu-id="d97a8-286">Declarações de tipo</span><span class="sxs-lookup"><span data-stu-id="d97a8-286">Type declarations</span></span>

<span data-ttu-id="d97a8-287">Declarações de tipo ajudam os usuários a entender o tipo de variável com a qual estão lidando.</span><span class="sxs-lookup"><span data-stu-id="d97a8-287">Type declarations help users understand the type of variable they are dealing with.</span></span> <span data-ttu-id="d97a8-288">Ele ajuda na conclusão automática de métodos e ajuda nas verificações de qualidade do tempo de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="d97a8-288">It helps with auto-completion of methods and assists in development time quality checks.</span></span>

<span data-ttu-id="d97a8-289">Você pode encontrar declarações de tipo no script em vários locais, incluindo declaração de função, declaração de variável, IntelliSense definições, etc.</span><span class="sxs-lookup"><span data-stu-id="d97a8-289">You can find type declarations in the script in various places including function declaration, variable declaration, IntelliSense definitions, etc.</span></span>

<span data-ttu-id="d97a8-290">Exemplos:</span><span class="sxs-lookup"><span data-stu-id="d97a8-290">Examples:</span></span>

* `function main(workbook: ExcelScript.Workbook)`
* `let myRange: ExcelScript.Range;`
* `function getMaxAmount(range: ExcelScript.Range): number`

<span data-ttu-id="d97a8-291">Você pode identificar os tipos facilmente no editor de código, pois ele geralmente aparece distintamente em uma cor diferente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-291">You can identify the types easily in the code editor as it usually appears distinctly in a different color.</span></span> <span data-ttu-id="d97a8-292">Um dois `:` pontos geralmente precede a declaração de tipo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-292">A colon `:` usually precedes the type declaration.</span></span>  

<span data-ttu-id="d97a8-293">Os tipos de escrita podem ser opcionais em TypeScript porque a inferência de tipo permite que você receba muita energia sem escrever código adicional.</span><span class="sxs-lookup"><span data-stu-id="d97a8-293">Writing types can be optional in TypeScript because type inference allows you to get a lot of power without writing additional code.</span></span> <span data-ttu-id="d97a8-294">Na maior parte, o idioma TypeScript é bom para inferir os tipos de variáveis.</span><span class="sxs-lookup"><span data-stu-id="d97a8-294">For the most part, the TypeScript language is good at inferring the types of variables.</span></span> <span data-ttu-id="d97a8-295">No entanto, em certos casos, Office scripts exigem que as declarações de tipo sejam explicitamente definidas se o idioma não puder identificar claramente o tipo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-295">However, in certain cases, Office Scripts require the type declarations to be explicitly defined if the language is unable to clearly identify the type.</span></span> <span data-ttu-id="d97a8-296">Além disso, explícito ou `any` implícito não é permitido Office Script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-296">Also, explicit or implicit `any` is not allowed in Office Script.</span></span> <span data-ttu-id="d97a8-297">Mais sobre esse assunto adiante.</span><span class="sxs-lookup"><span data-stu-id="d97a8-297">More on that later.</span></span>

### <a name="excelscript-types"></a><span data-ttu-id="d97a8-298">`ExcelScript` types</span><span class="sxs-lookup"><span data-stu-id="d97a8-298">`ExcelScript` types</span></span>

<span data-ttu-id="d97a8-299">Em Office Scripts, você usará os seguintes tipos de tipos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-299">In Office Scripts, you will use the following kinds of types.</span></span>

* <span data-ttu-id="d97a8-300">Tipos de idioma `number` nativo, como `string` , , , , , `object` `boolean` `null` etc.</span><span class="sxs-lookup"><span data-stu-id="d97a8-300">Native language types such as `number`, `string`, `object`, `boolean`, `null`, etc.</span></span>
* <span data-ttu-id="d97a8-301">Excel Tipos de API.</span><span class="sxs-lookup"><span data-stu-id="d97a8-301">Excel API types.</span></span> <span data-ttu-id="d97a8-302">Eles começam com `ExcelScript` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-302">They begin with `ExcelScript`.</span></span> <span data-ttu-id="d97a8-303">Por exemplo, `ExcelScript.Range` , `ExcelScript.Table` , etc.</span><span class="sxs-lookup"><span data-stu-id="d97a8-303">For example, `ExcelScript.Range`, `ExcelScript.Table`, etc.</span></span>
* <span data-ttu-id="d97a8-304">Quaisquer interfaces personalizadas que você possa ter definido no script usando `interface` instruções.</span><span class="sxs-lookup"><span data-stu-id="d97a8-304">Any custom interfaces you may have defined in the script using `interface` statements.</span></span>

<span data-ttu-id="d97a8-305">Consulte exemplos de cada um desses grupos em seguida.</span><span class="sxs-lookup"><span data-stu-id="d97a8-305">See examples of each of these groups next.</span></span>

<span data-ttu-id="d97a8-306">**_Tipos de idioma nativo_**</span><span class="sxs-lookup"><span data-stu-id="d97a8-306">**_Native language types_**</span></span>

<span data-ttu-id="d97a8-307">No exemplo a seguir, observe locais `string` onde , e foram `number` `boolean` usados.</span><span class="sxs-lookup"><span data-stu-id="d97a8-307">In the following example, notice places where `string`, `number`, and `boolean` have been used.</span></span> <span data-ttu-id="d97a8-308">Esses são tipos **de idioma TypeScript** nativos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-308">These are native **TypeScript** language types.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook)
{
  const table = workbook.getActiveWorksheet().getTables()[0];
  const sales = table.getColumnByName('Sales').getRange().getValues();
  console.log(sales);
  // Add 100 to each value.
  const revisedSales = salesAs1DArray.map(data => data as number + 100);
  // Add a column.
  table.addColumn(-1, revisedSales);  
}
/**
 * Extract a column from 2D array and return result.
 */
function extractColumn(data: (string | number | boolean)[][], index: number): (string | number | boolean)[] {

  const column = data.map((row) => {
    return row[index];
  })
  return column;
}
/**
 * Convert a flat array into a 2D array that can be used as range column.
 */
function convertColumnTo2D(data: (string | number | boolean)[]): (string | number | boolean)[][] {

  const columnAs2D = data.map((row) => {
    return [row];
  })
  return columnAs2D;
}
```

<span data-ttu-id="d97a8-309">**_Tipos do ExcelScript_**</span><span class="sxs-lookup"><span data-stu-id="d97a8-309">**_ExcelScript types_**</span></span>

<span data-ttu-id="d97a8-310">No exemplo a seguir, uma função auxiliar tem dois argumentos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-310">In the following example, a helper function takes two arguments.</span></span> <span data-ttu-id="d97a8-311">O primeiro é a `sheet` variável que é do `ExcelScript.Worksheet` tipo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-311">The first one is the `sheet` variable which is of type `ExcelScript.Worksheet` type.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for update.
    if (usedRange) { 
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);      
    targetRange.setValues([data]);
    return;
}
```

<span data-ttu-id="d97a8-312">**_Tipos personalizados_**</span><span class="sxs-lookup"><span data-stu-id="d97a8-312">**_Custom types_**</span></span>

<span data-ttu-id="d97a8-313">A interface personalizada `ReportImages` é usada para retornar imagens para outra ação de fluxo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-313">The custom interface `ReportImages` is used to return images to another flow action.</span></span> <span data-ttu-id="d97a8-314">A `main` declaração de função inclui instruções para dizer a TypeScript que um objeto desse tipo está sendo `: ReportImages` retornado.</span><span class="sxs-lookup"><span data-stu-id="d97a8-314">The `main` function declaration includes `: ReportImages` instruction to tell TypeScript that an object of that type is being returned.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  let chart = workbook.getWorksheet("Sheet1").getCharts()[0];
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  
  const chartImage = chart.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

### <a name="type-assertion-overriding-the-type"></a><span data-ttu-id="d97a8-315">Tipo de afirmação (substituindo o tipo)</span><span class="sxs-lookup"><span data-stu-id="d97a8-315">Type assertion (overriding the type)</span></span>

<span data-ttu-id="d97a8-316">Como diz a documentação [TypeScript,](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) "Às vezes, você terminará em uma situação em que você conhecerá mais sobre um valor do que TypeScript.</span><span class="sxs-lookup"><span data-stu-id="d97a8-316">As the TypeScript [documentation](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) states, "Sometimes you'll end up in a situation where you'll know more about a value than TypeScript does.</span></span> <span data-ttu-id="d97a8-317">Normalmente, isso acontecerá quando você sabe que o tipo de alguma entidade pode ser mais específico do que seu tipo atual.</span><span class="sxs-lookup"><span data-stu-id="d97a8-317">Usually, this will happen when you know the type of some entity could be more specific than its current type.</span></span> <span data-ttu-id="d97a8-318">As declarações de tipo são uma maneira de dizer ao compilador "confie em mim, eu sei o que estou fazendo".</span><span class="sxs-lookup"><span data-stu-id="d97a8-318">Type assertions are a way to tell the compiler “trust me, I know what I'm doing.”</span></span> <span data-ttu-id="d97a8-319">Uma afirmação de tipo é como um tipo lançado em outros idiomas, mas não executa nenhuma verificação especial ou reestruturação de dados.</span><span class="sxs-lookup"><span data-stu-id="d97a8-319">A type assertion is like a type cast in other languages, but it performs no special checking or restructuring of data.</span></span> <span data-ttu-id="d97a8-320">Ele não tem impacto no tempo de execução e é usado puramente pelo compilador."</span><span class="sxs-lookup"><span data-stu-id="d97a8-320">It has no runtime impact and is used purely by the compiler."</span></span>

<span data-ttu-id="d97a8-321">Você pode afirmar o tipo usando a `as` palavra-chave ou usando colchetes angulares, conforme mostrado no código a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-321">You can assert the type using the `as` keyword or using angle brackets as shown in following code.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let data = workbook.getActiveCell().getValue();
  // Since the add10 function only accepts number, assert data's type as number, otherwise the script cannot be run.
  const answer1 = add10(data as number);
  const answer2 = add10(<number> data);
}

function add10(data: number) { 
  return data + 10;
}
```

#### <a name="any-type-in-the-script"></a><span data-ttu-id="d97a8-322">Tipo 'any' no script</span><span class="sxs-lookup"><span data-stu-id="d97a8-322">'any' type in the script</span></span>

<span data-ttu-id="d97a8-323">O [site TypeScript afirma](https://www.typescriptlang.org/docs/handbook/basic-types.html#any):</span><span class="sxs-lookup"><span data-stu-id="d97a8-323">The [TypeScript website states](https://www.typescriptlang.org/docs/handbook/basic-types.html#any):</span></span>

  <span data-ttu-id="d97a8-324">Em algumas situações, nem todas as informações de tipo estão disponíveis ou sua declaração levaria uma quantidade inadequada de esforço.</span><span class="sxs-lookup"><span data-stu-id="d97a8-324">In some situations, not all type information is available or its declaration would take an inappropriate amount of effort.</span></span> <span data-ttu-id="d97a8-325">Isso pode ocorrer para valores de código que foi gravado sem TypeScript ou uma biblioteca de terceiros.</span><span class="sxs-lookup"><span data-stu-id="d97a8-325">These may occur for values from code that has been written without TypeScript or a 3rd party library.</span></span> <span data-ttu-id="d97a8-326">Nesses casos, podemos optar por não fazer a verificação de tipos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-326">In these cases, we might want to opt-out of type checking.</span></span> <span data-ttu-id="d97a8-327">Para fazer isso, rotularemos esses valores com o `any` tipo:</span><span class="sxs-lookup"><span data-stu-id="d97a8-327">To do so, we label these values with the `any` type:</span></span>

  ```TypeScript
  declare function getValue(key: string): any;
  // OK, return value of 'getValue' is not checked
  const str: string = getValue("myString");
  ```

<span data-ttu-id="d97a8-328">**Não `any` é permitido explicitar**</span><span class="sxs-lookup"><span data-stu-id="d97a8-328">**Explicit `any` is NOT allowed**</span></span>

```TypeScript
// This is not allowed
let someVariable: any; 
```

<span data-ttu-id="d97a8-329">O `any` tipo apresenta desafios para a maneira como Office Scripts processa as APIs Excel.</span><span class="sxs-lookup"><span data-stu-id="d97a8-329">The `any` type presents challenges to the way Office Scripts processes the Excel APIs.</span></span> <span data-ttu-id="d97a8-330">Ele causa problemas quando as variáveis são enviadas Excel APIs para processamento.</span><span class="sxs-lookup"><span data-stu-id="d97a8-330">It causes issues when the variables are sent to Excel APIs for processing.</span></span> <span data-ttu-id="d97a8-331">Conhecer o tipo de variáveis usadas no script é essencial para o processamento do script e, portanto, a definição explícita de qualquer variável com `any` tipo é proibida.</span><span class="sxs-lookup"><span data-stu-id="d97a8-331">Knowing the type of variables used in the script is essential to the processing of script and hence explicit definition of any variable with `any` type is prohibited.</span></span> <span data-ttu-id="d97a8-332">Você receberá um erro de tempo de compilação (erro antes de executar o script) se houver qualquer variável com o tipo `any` declarado no script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-332">You will receive a compile-time error (error prior to running the script) if there is any variable with `any` type declared in the script.</span></span> <span data-ttu-id="d97a8-333">Você também verá um erro no editor.</span><span class="sxs-lookup"><span data-stu-id="d97a8-333">You will see an error in the editor as well.</span></span>

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="Erro explícito de &quot;qualquer&quot;":::

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Erro explícito 'qualquer' mostrado em Output":::

<span data-ttu-id="d97a8-336">No código exibido na imagem anterior, indica que a linha 5 coluna `[5, 16] Explicit Any is not allowed` 16 declara o `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-336">In the code displayed in the previous image, `[5, 16] Explicit Any is not allowed` indicates that line 5 column 16 declares the `any` type.</span></span> <span data-ttu-id="d97a8-337">Isso ajuda você a localizar a linha de código que contém o erro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-337">This helps you locate the line of code that contains the error.</span></span>

<span data-ttu-id="d97a8-338">Para se livrar desse problema, declare sempre o tipo da variável.</span><span class="sxs-lookup"><span data-stu-id="d97a8-338">To get around this issue, always declare the type of the variable.</span></span>

<span data-ttu-id="d97a8-339">Se você não tiver certeza sobre o tipo de variável, um truque legal no TypeScript permite definir tipos [de união.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="d97a8-339">If you are uncertain about the type of a variable, one cool trick in TypeScript allows you to define [union types](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="d97a8-340">Isso pode ser usado para variáveis para manter um intervalo de valores, que pode ser de vários tipos.</span><span class="sxs-lookup"><span data-stu-id="d97a8-340">This can be used for variables to hold a range values, which can be of many types.</span></span>

```TypeScript
// Define value as a union type rather than 'any' type.
let value: (string | number | boolean);
value = someValue_from_another_source;
//...
someRange.setValue(value);
```

### <a name="type-inference"></a><span data-ttu-id="d97a8-341">Digite inferência</span><span class="sxs-lookup"><span data-stu-id="d97a8-341">Type inference</span></span>

<span data-ttu-id="d97a8-342">No TypeScript, há vários locais onde a [inferência](https://www.typescriptlang.org/docs/handbook/type-inference.html) de tipo é usada para fornecer informações de tipo quando não há anotação de tipo explícito.</span><span class="sxs-lookup"><span data-stu-id="d97a8-342">In TypeScript, there are several places where [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html) is used to provide type information when there is no explicit type annotation.</span></span> <span data-ttu-id="d97a8-343">Por exemplo, o tipo da variável x é inferido como um número no código a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-343">For example, the type of the x variable is inferred to be a number in the following code.</span></span>

```TypeScript
let x = 3;
//  ^ = let x: number
```

<span data-ttu-id="d97a8-344">Esse tipo de inferência ocorre ao inicializar variáveis e membros, definir valores padrão do parâmetro e determinar tipos de retorno de função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-344">This kind of inference takes place when initializing variables and members, setting parameter default values, and determining function return types.</span></span>

### <a name="no-implicit-any-rule"></a><span data-ttu-id="d97a8-345">no-implicit-any rule</span><span class="sxs-lookup"><span data-stu-id="d97a8-345">no-implicit-any rule</span></span>

<span data-ttu-id="d97a8-346">Um script requer os tipos das variáveis usadas para serem declaradas explicitamente ou implicitamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-346">A script requires the types of the variables used to be explicitly or implicitly declared.</span></span> <span data-ttu-id="d97a8-347">Se o compilador TypeScript não conseguir determinar o tipo de uma variável (porque o tipo não é declarado explicitamente ou a inferência de tipo não é possível), você receberá um erro de tempo de compilação (erro antes de executar o script).</span><span class="sxs-lookup"><span data-stu-id="d97a8-347">If the TypeScript compiler is unable to determine the type of a variable (either because type is not declared explicitly or type inference is not possible), then you will receive a compilation time error (error prior to running the script).</span></span> <span data-ttu-id="d97a8-348">Você também verá um erro no editor.</span><span class="sxs-lookup"><span data-stu-id="d97a8-348">You will see an error in the editor as well.</span></span>

:::image type="content" source="../../images/getting-started-iany.png" alt-text="O erro implícito &quot;qualquer&quot; mostrado no editor":::

<span data-ttu-id="d97a8-350">Os scripts a seguir têm erros de tempo de compilação porque as variáveis são declaradas sem tipos e TypeScript não pode determinar o tipo no momento da declaração.</span><span class="sxs-lookup"><span data-stu-id="d97a8-350">The following scripts have compilation time errors because variables are declared without types and TypeScript cannot determine the type at the time of declaration.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'value' gets 'any' type
    // because no type is declared.
    let value; 
    // Even when a number type is assigned,
    // the type of 'value' remains any.
    value = 10; 
    // The following statement fails because
    // Office Scripts can't send an argument
    // of type 'any' to Excel for processing.
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'cell' gets 'any' type
    // because no type is defined.
    let cell; 
    cell = workbook.getActiveCell().getValue();
    // Office Scripts can't assign Range type object
    // to a variable of 'any' type.
    console.log(cell.getValue());
    return;
}
```

<span data-ttu-id="d97a8-351">Para evitar esse erro, use os padrões a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-351">To avoid this error, use the following patterns instead.</span></span> <span data-ttu-id="d97a8-352">Em cada caso, a variável e seu tipo são declaradas ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="d97a8-352">In each case, the variable and its type are declared at the same time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const value: number = 10; 
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const cell: ExcelScript.Range = workbook.getActiveCell().getValue();
    console.log(cell.getValue()); 
    return;
}
```

## <a name="error-handling"></a><span data-ttu-id="d97a8-353">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="d97a8-353">Error handling</span></span>

<span data-ttu-id="d97a8-354">Office O erro de scripts pode ser classificado em uma das seguintes categorias.</span><span class="sxs-lookup"><span data-stu-id="d97a8-354">Office Scripts error can be classified into one of the following categories.</span></span>

1. <span data-ttu-id="d97a8-355">Aviso de tempo de compilação mostrado no editor</span><span class="sxs-lookup"><span data-stu-id="d97a8-355">Compile-time warning shown in the editor</span></span>
1. <span data-ttu-id="d97a8-356">Erro de tempo de compilação que aparece quando você é executado, mas ocorre antes do início da execução</span><span class="sxs-lookup"><span data-stu-id="d97a8-356">Compile-time error that appears when you run but occurs before execution begins</span></span>
1. <span data-ttu-id="d97a8-357">Erro de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="d97a8-357">Runtime error</span></span>

<span data-ttu-id="d97a8-358">Os avisos do editor podem ser identificados usando os sublinhados vermelho ondulados no editor:</span><span class="sxs-lookup"><span data-stu-id="d97a8-358">Editor warnings can be identified using the wavy red underlines in the editor:</span></span>

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="Aviso de tempo de compilação mostrado no editor":::

<span data-ttu-id="d97a8-360">Às vezes, você também pode ver sublinhados de aviso laranja e mensagens informativas cinza.</span><span class="sxs-lookup"><span data-stu-id="d97a8-360">At times, you may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="d97a8-361">Eles devem ser examinados de perto, embora não causem erros.</span><span class="sxs-lookup"><span data-stu-id="d97a8-361">They should be examined closely though they are not going to cause errors.</span></span>

<span data-ttu-id="d97a8-362">Não é possível distinguir entre erros de tempo de compilação e tempo de execução, pois ambas as mensagens de erro são idênticas.</span><span class="sxs-lookup"><span data-stu-id="d97a8-362">It isn't possible to distinguish between compile-time and runtime errors as both error messages look identical.</span></span> <span data-ttu-id="d97a8-363">Ambos ocorrem quando você realmente executa o script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-363">They both occur when you actually execute the script.</span></span> <span data-ttu-id="d97a8-364">As imagens a seguir mostram exemplos de um erro de tempo de compilação e um erro de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d97a8-364">The following images show examples of a compile-time error and a runtime error.</span></span>

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Exemplo de um erro em tempo de compilação":::

:::image type="content" source="../../images/getting-started-error-basic.png" alt-text="Exemplo de um erro de tempo de execução":::

<span data-ttu-id="d97a8-367">Em ambos os casos, você verá o número da linha onde ocorreu o erro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-367">In both cases, you will see the line number where the error occurred.</span></span> <span data-ttu-id="d97a8-368">Em seguida, você pode examinar o código, corrigir o problema e executar novamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-368">You can then examine the code, fix the issue, and run again.</span></span>

<span data-ttu-id="d97a8-369">A seguir estão algumas práticas recomendadas para evitar erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d97a8-369">Following are a few best practices to avoid runtime errors.</span></span>

### <a name="check-for-object-existence-before-deletion"></a><span data-ttu-id="d97a8-370">Verificar a existência do objeto antes da exclusão</span><span class="sxs-lookup"><span data-stu-id="d97a8-370">Check for object existence before deletion</span></span>

<span data-ttu-id="d97a8-371">Como alternativa, para excluir um objeto que pode ou não existir, use este padrão:</span><span class="sxs-lookup"><span data-stu-id="d97a8-371">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
// The ? ensures that the delete() API is only invoked if the object exists.
workbook.getWorksheet('Index')?.delete();

// Alternative:
const indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
    indexSheet.delete();
}
```

### <a name="do-pre-checks-at-the-beginning-of-the-script"></a><span data-ttu-id="d97a8-372">Fazer verificações prévias no início do script</span><span class="sxs-lookup"><span data-stu-id="d97a8-372">Do pre-checks at the beginning of the script</span></span>

<span data-ttu-id="d97a8-373">Como prática prática, sempre certifique-se de que todas as suas entradas estão presentes no arquivo Excel antes de executar seu script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-373">As a best practice, always ensure that all your inputs are present in the Excel file prior to running your script.</span></span> <span data-ttu-id="d97a8-374">Você pode ter feito algumas suposições sobre os objetos que estão presentes na workbook.</span><span class="sxs-lookup"><span data-stu-id="d97a8-374">You may have made certain assumptions about objects being present in the workbook.</span></span> <span data-ttu-id="d97a8-375">Se esses objetos não existirem, seu script poderá encontrar um erro ao ler o objeto ou seus dados.</span><span class="sxs-lookup"><span data-stu-id="d97a8-375">If those objects don't exist, your script may encounter an error when you read the object or its data.</span></span> <span data-ttu-id="d97a8-376">Em vez de começar o processamento e o erro no meio depois que parte das atualizações ou processamento já tiver terminado, é melhor fazer todas as pré-verificações no início do script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-376">Rather than beginning the processing and erroring in the middle after part of the updates or processing has already finished, it is better to do all pre-checks at the start of the script.</span></span>

<span data-ttu-id="d97a8-377">Por exemplo, o script a seguir requer que duas tabelas chamadas Table1 e Table2 estão presentes.</span><span class="sxs-lookup"><span data-stu-id="d97a8-377">For example, the following script requires two tables named Table1 and Table2 to be present.</span></span> <span data-ttu-id="d97a8-378">Portanto, o script verifica sua presença e termina com a instrução e `return` uma mensagem apropriada se eles não estão presentes.</span><span class="sxs-lookup"><span data-stu-id="d97a8-378">Hence the script checks for their presence and ends with the `return` statement and an appropriate message if they are not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="d97a8-379">Se a verificação para garantir que a presença de dados de entrada está acontecendo em uma função separada, é importante encerrar o script em emissão da instrução `return` da `main` função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-379">If the verification to ensure the presence of input data is happening in a separate function, it's important to end the script by issuing the `return` statement from the `main` function.</span></span>

<span data-ttu-id="d97a8-380">No exemplo a seguir, `main` a função chama a função para fazer as `inputPresent` verificações prévias.</span><span class="sxs-lookup"><span data-stu-id="d97a8-380">In the following example, the `main` function calls the `inputPresent` function to do the pre-checks.</span></span> <span data-ttu-id="d97a8-381">`inputPresent` retorna um booleano ( ou ) indicando se todas `true` as entradas necessárias estão presentes ou `false` não.</span><span class="sxs-lookup"><span data-stu-id="d97a8-381">`inputPresent` returns a boolean (`true` or `false`) indicating whether all required inputs are present or not.</span></span> <span data-ttu-id="d97a8-382">Em seguida, é responsabilidade da função emitir a instrução (ou seja, de dentro `main` `return` da `main` função) para encerrar o script imediatamente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-382">It's then the responsibility of the `main` function to issue the `return` statement (that is, from within the `main` function) to end the script immediately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }
  return true;
}
```

### <a name="when-to-abort-throw-the-script"></a><span data-ttu-id="d97a8-383">Quando abortar ( `throw` ) o script</span><span class="sxs-lookup"><span data-stu-id="d97a8-383">When to abort (`throw`) the script</span></span>  

<span data-ttu-id="d97a8-384">Na maior parte, você não precisa abortar ( `throw` ) do seu script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-384">For the most part, you don't need to abort (`throw`) from your script.</span></span> <span data-ttu-id="d97a8-385">Isso ocorre porque o script geralmente informa ao usuário que o script não foi executado devido a um problema.</span><span class="sxs-lookup"><span data-stu-id="d97a8-385">This is because the script's usually informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="d97a8-386">Na maioria dos casos, é suficiente terminar o script com uma mensagem de erro e `return` uma instrução da `main` função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-386">In most case, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="d97a8-387">No entanto, se o script estiver sendo executado como parte da Power Automate, talvez você queira cancelar o fluxo se determinadas condições não são atendidas.</span><span class="sxs-lookup"><span data-stu-id="d97a8-387">However, if your script is running as part of Power Automate, you may want to abort the flow if certain conditions are not met.</span></span> <span data-ttu-id="d97a8-388">Portanto, é importante não sobre um erro, mas sim emitir uma instrução para abortar o script para que quaisquer instruções de código subsequentes `return` `throw` não são executados.</span><span class="sxs-lookup"><span data-stu-id="d97a8-388">It's therefore important to not `return` upon an error but rather issue a `throw` statement to abort the script so that any subsequent code statements don't run.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    // Abort script.
    throw `Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

<span data-ttu-id="d97a8-389">Conforme mencionado na seção a seguir, outro cenário é quando você tem várias funções envolvidas ( chamadas que chamam `main` , etc.) o que dificulta a `functionX` propagação do `functionY` erro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-389">As mentioned in the following section, another scenario is when you have several functions involved (`main` calls `functionX` which calls `functionY`, etc.) which makes it hard to propagate the error.</span></span> <span data-ttu-id="d97a8-390">A anulação/lançamento da função aninhada com uma mensagem pode ser mais fácil do que retornar um erro até e retornar com `main` `main` uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-390">Aborting/throwing from the nested function with a message may be easier than returning an error all the way up to `main` and returning from `main` with an error message.</span></span>

### <a name="when-to-use-trycatch-throw-exception"></a><span data-ttu-id="d97a8-391">Quando usar try.. catch (exceção de lançamento)</span><span class="sxs-lookup"><span data-stu-id="d97a8-391">When to use try..catch (throw exception)</span></span>

<span data-ttu-id="d97a8-392">A [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) técnica é uma maneira de detectar se uma chamada de API falhou e lidar com esse erro no script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-392">The [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) technique is a way to detect if an API call failed and handle that error in your script.</span></span> <span data-ttu-id="d97a8-393">Talvez seja importante verificar o valor de retorno de uma API para verificar se ele foi concluído com êxito.</span><span class="sxs-lookup"><span data-stu-id="d97a8-393">It may be important to check the return value of an API to verify that it was completed successfully.</span></span>

<span data-ttu-id="d97a8-394">Considere o trecho de exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-394">Consider the following example snippet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Somewhere in the script, perform a large data update.
  range.setValues(someLargeValues);

}
```

<span data-ttu-id="d97a8-395">A `setValues()` chamada pode falhar e resultar na falha do script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-395">The `setValues()` call may fail and result in the script failure.</span></span> <span data-ttu-id="d97a8-396">Talvez você queira lidar com essa condição em seu código e, talvez, personalizar a mensagem de erro ou separar a atualização em unidades menores, etc. Nesse caso, é importante saber que a API retornou um erro e interpretar ou manipular esse erro.</span><span class="sxs-lookup"><span data-stu-id="d97a8-396">You may wish to handle this condition in your code and perhaps customize the error message or break up the update into smaller units, etc. In that case, it's important to know that the API returned an error and interpret or handle that error.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Please inspect and run again.`);
    console.log(error);
    return; // End script (assuming this is in main function).
}

// OR...

try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Trying a different approach`);
    handleUpdatesInSmallerChunks(someLargeValues);
}

// Continue...
}
```

<span data-ttu-id="d97a8-397">Outro cenário é quando a função principal chama outra função, que, por sua vez, chama outra função (e assim por diante).), e a chamada da API que você se importa acontece na função inferior.</span><span class="sxs-lookup"><span data-stu-id="d97a8-397">Another scenario is when main function calls another function, which in turn calls another function (and so on..), and the API call that you care about happens down in the bottom function.</span></span> <span data-ttu-id="d97a8-398">Propagar o erro até pode não ser `main` viável ou conveniente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-398">Propagating the error all the way up to `main` may not be feasible or convenient.</span></span> <span data-ttu-id="d97a8-399">Nesse caso, lançar um erro na função inferior será mais conveniente.</span><span class="sxs-lookup"><span data-stu-id="d97a8-399">In that case, throwing an error in the bottom function will be most convenient.</span></span>

```TypeScript

function main(workbook: ExcelScript.Workbook) {
    ...
    updateRangeInChunks(sheet.getRange("B1"), data);
    ...
}

function updateRangeInChunks(
    ...
    updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
    ...
}

function updateTargetRange(
      targetCell: ExcelScript.Range,
      values: (string | boolean | number)[][]
    ) {
    const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
    console.log(`Updating the range: ${targetRange.getAddress()}`);
    try {
      targetRange.setValues(values);
    } catch (e) {
      throw `Error while updating the whole range: ${JSON.stringify(e)}`;
    }
    return;
}
```

<span data-ttu-id="d97a8-400">*Aviso:* o `try..catch` uso dentro de um loop diminuirá o script.</span><span class="sxs-lookup"><span data-stu-id="d97a8-400">*Warning*: Using `try..catch` inside of a loop will slow down your script.</span></span> <span data-ttu-id="d97a8-401">Evite usar isso dentro ou ao redor de loops.</span><span class="sxs-lookup"><span data-stu-id="d97a8-401">Avoid using this inside of or around loops.</span></span>

## <a name="basic-performance-considerations"></a><span data-ttu-id="d97a8-402">Considerações básicas sobre o desempenho</span><span class="sxs-lookup"><span data-stu-id="d97a8-402">Basic performance considerations</span></span>

### <a name="avoid-slow-operations-in-the-loop"></a><span data-ttu-id="d97a8-403">Evitar operações lentas no loop</span><span class="sxs-lookup"><span data-stu-id="d97a8-403">Avoid slow operations in the loop</span></span>

<span data-ttu-id="d97a8-404">Determinadas operações quando realizadas dentro/ao redor das instruções de loop, como `for` , , , , `for..of` `map` `forEach` etc. podem levar a um desempenho lento.</span><span class="sxs-lookup"><span data-stu-id="d97a8-404">Certain operations when done inside/around the loop statements such as `for`, `for..of`, `map`, `forEach`, etc. can lead to slow performance.</span></span> <span data-ttu-id="d97a8-405">Evite as seguintes categorias de API.</span><span class="sxs-lookup"><span data-stu-id="d97a8-405">Avoid the following API categories.</span></span>

* <span data-ttu-id="d97a8-406">`get*` APIs</span><span class="sxs-lookup"><span data-stu-id="d97a8-406">`get*` APIs</span></span>

<span data-ttu-id="d97a8-407">Leia todos os dados necessários fora do loop em vez de lê-los dentro do loop.</span><span class="sxs-lookup"><span data-stu-id="d97a8-407">Read all the data you need outside of the loop rather than reading it inside of the loop.</span></span> <span data-ttu-id="d97a8-408">Às vezes, é difícil evitar a leitura dentro de loops; nesse caso, certifique-se de que suas contagens de loop não sejam muito grandes ou gerencie-as em lotes para evitar ter que fazer loop por meio de uma estrutura de dados grande.</span><span class="sxs-lookup"><span data-stu-id="d97a8-408">At times, it is hard to avoid reading inside of loops; in such a case, make sure your loop counts are not too large or manage them in batches to avoid having to loop through a large data structure.</span></span>

<span data-ttu-id="d97a8-409">**Observação:** se o intervalo/dados com que você está lidando for bastante grande (digamos que células de 100 mil >, talvez seja necessário usar técnicas avançadas como separar suas leituras/gravações em várias partes.</span><span class="sxs-lookup"><span data-stu-id="d97a8-409">**Note**: If the range/data you are dealing with is quite large (say >100K cells), you may need to use advanced techniques like breaking up your read/writes into multiple chunks.</span></span> <span data-ttu-id="d97a8-410">O vídeo a seguir é realmente para uma configuração de dados de pequeno porte.</span><span class="sxs-lookup"><span data-stu-id="d97a8-410">The following video is really for a small-mid sized data setup.</span></span> <span data-ttu-id="d97a8-411">Para um grande conjuntos de dados, consulte o [cenário avançado de gravação de dados.](write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="d97a8-411">For a large dataset, refer to [advanced data write scenario](write-large-dataset.md).</span></span>

<span data-ttu-id="d97a8-412">[![Vídeo fornecendo uma dica de otimização de leitura e gravação](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Vídeo mostrando dica de otimização de leitura e gravação")</span><span class="sxs-lookup"><span data-stu-id="d97a8-412">[![Video providing a read-and-write optimization tip](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Video showing read-and-write optimization tip")</span></span>

* <span data-ttu-id="d97a8-413">`console.log` instrução (consulte o exemplo a seguir)</span><span class="sxs-lookup"><span data-stu-id="d97a8-413">`console.log` statement (see the following example)</span></span>

```TypeScript
// Color each cell with random color.
for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
        range
            .getCell(row, col)
            .getFormat()
            .getFill()
            .setColor(`#${Math.random().toString(16).substr(-6)}`);
        /* Avoid such console.log inside loop */
        // console.log("Updating" + range.getCell(row, col).getAddress());
    }
}
```

* <span data-ttu-id="d97a8-414">`try {} catch ()` instrução</span><span class="sxs-lookup"><span data-stu-id="d97a8-414">`try {} catch ()` statement</span></span>

<span data-ttu-id="d97a8-415">Evite loops de `for` tratamento de exceção.</span><span class="sxs-lookup"><span data-stu-id="d97a8-415">Avoid exception handling `for` loops.</span></span> <span data-ttu-id="d97a8-416">Loops dentro e fora.</span><span class="sxs-lookup"><span data-stu-id="d97a8-416">Both inside and outside loops.</span></span>

## <a name="note-to-vba-developers"></a><span data-ttu-id="d97a8-417">Observação para desenvolvedores do VBA</span><span class="sxs-lookup"><span data-stu-id="d97a8-417">Note to VBA developers</span></span>

<span data-ttu-id="d97a8-418">O idioma TypeScript difere do VBA tanto de forma síncte quanto em convenções de nomenal.</span><span class="sxs-lookup"><span data-stu-id="d97a8-418">The TypeScript language differs from VBA both syntactically as well as in naming conventions.</span></span>

<span data-ttu-id="d97a8-419">Confira os trechos equivalentes a seguir.</span><span class="sxs-lookup"><span data-stu-id="d97a8-419">Check out the following equivalent snippets.</span></span>

```vba
Worksheets("Sheet1").Range("A1:G37").Clear
```

```TypeScript
workbook.getWorksheet('Sheet1').getRange('A1:G37').clear(ExcelScript.ClearApplyTo.all);
```

<span data-ttu-id="d97a8-420">Algumas coisas a ser chamada sobre TypeScript:</span><span class="sxs-lookup"><span data-stu-id="d97a8-420">A few things to call out about TypeScript:</span></span>

* <span data-ttu-id="d97a8-421">Você pode observar que todos os métodos precisam ter parênteses abertos para execução.</span><span class="sxs-lookup"><span data-stu-id="d97a8-421">You may notice that all methods need to have open-close parentheses to execute.</span></span> <span data-ttu-id="d97a8-422">Os argumentos são passados de forma idêntica, mas alguns argumentos podem ser necessários para execução (ou seja, obrigatório versus opcional).</span><span class="sxs-lookup"><span data-stu-id="d97a8-422">Arguments are passed identically but some arguments may be required for execution (that is, required vs optional).</span></span>
* <span data-ttu-id="d97a8-423">A convenção de nomenisagem segue camelCase em vez da convenção PascalCase.</span><span class="sxs-lookup"><span data-stu-id="d97a8-423">The naming convention follows camelCase instead of PascalCase convention.</span></span>
* <span data-ttu-id="d97a8-424">Os métodos geralmente têm `get` `set` ou prefixos indicando se ele está lendo ou escrevendo membros do objeto.</span><span class="sxs-lookup"><span data-stu-id="d97a8-424">Methods usually have `get` or `set` prefixes indicating whether it is reading or writing object members.</span></span>
* <span data-ttu-id="d97a8-425">Os blocos de código são definidos e identificados por chaves abertas: `{` `}` .</span><span class="sxs-lookup"><span data-stu-id="d97a8-425">The code blocks are defined and identified by open-close curly braces: `{` `}`.</span></span> <span data-ttu-id="d97a8-426">Os blocos são necessários para `if` condições, `while` instruções, `for` loops, definições de função, etc.</span><span class="sxs-lookup"><span data-stu-id="d97a8-426">Blocks are required for `if` conditions, `while` statements, `for` loops, function definitions, etc.</span></span>
* <span data-ttu-id="d97a8-427">As funções podem chamar outras funções e você pode até definir funções dentro de uma função.</span><span class="sxs-lookup"><span data-stu-id="d97a8-427">Functions can call other functions and you can even define functions within a function.</span></span>

<span data-ttu-id="d97a8-428">Em geral, TypeScript é um idioma diferente e há poucas semelhanças entre eles.</span><span class="sxs-lookup"><span data-stu-id="d97a8-428">Overall, TypeScript is a different language and there are few similarities between them.</span></span> <span data-ttu-id="d97a8-429">No entanto, Office PRÓPRIA API de Scripts usa terminologia semelhante e hierarquia de modelo de dados (modelo de objeto) como APIs VBA e isso deve ajudá-lo a navegar por aí.</span><span class="sxs-lookup"><span data-stu-id="d97a8-429">However, the Office Scripts API themselves use similar terminology and data-model (object model) hierarchy as VBA APIs and that should help you navigate around.</span></span>
