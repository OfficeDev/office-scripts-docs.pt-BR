---
title: Comece a usar scripts de um fluxo manual do Power Automate
description: Um tutorial sobre o uso de Scripts do Office no Power Automate por meio de um acionamento manual.
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 1a8b9659ec6f6354d583496ba0f3e94d4a13c01b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313992"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a><span data-ttu-id="d6f9b-103">Scripts de chamada a partir de um fluxo manual do Power Automate</span><span class="sxs-lookup"><span data-stu-id="d6f9b-103">Call scripts from a manual Power Automate flow</span></span>

<span data-ttu-id="d6f9b-104">Este tutorial ensina como executar um Script do Office para o Excel na web por meio do [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="d6f9b-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="d6f9b-105">Você fará um script que atualizará os valores de duas células com a hora atual.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-105">You'll make a script that updates the values of two cells with the current time.</span></span> <span data-ttu-id="d6f9b-106">Depois, você fará a conexão desse script a um fluxo do Power Automate acionado manualmente, para que o script seja executado sempre que um botão no Power Automate for selecionado.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-106">You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is selected.</span></span> <span data-ttu-id="d6f9b-107">Depois de entender o padrão básico, você pode expandir o fluxo para incluir outros aplicativos e automatizar ainda mais o seu fluxo de trabalho diário.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-107">Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.</span></span>

> [!TIP]
> <span data-ttu-id="d6f9b-108">Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="d6f9b-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="d6f9b-109">[Os Scripts do Office usam TypeScript](../overview/code-editor-environment.md) e este tutorial se destina a pessoas com conhecimento de nível iniciante a intermediário em JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="d6f9b-110">Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="d6f9b-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="d6f9b-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="d6f9b-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="d6f9b-112">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="d6f9b-112">Prepare the workbook</span></span>

<span data-ttu-id="d6f9b-113">O Power Automate não pode usar[referências relativas](../testing/power-automate-troubleshooting.md#avoid-relative-references)como`Workbook.getActiveWorksheet`acessar componentes da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="d6f9b-114">Portanto, precisamos de uma pasta de trabalho e de uma planilha com nomes consistentes que o Power Automate consiga consultar.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-114">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="d6f9b-115">Crie uma pasta de trabalho intitulada **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="d6f9b-116">Na pasta de trabalho **MyWorkbook**, crie uma planilha intitulada **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-116">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="d6f9b-117">Criar um Script do Office</span><span class="sxs-lookup"><span data-stu-id="d6f9b-117">Create an Office Script</span></span>

1. <span data-ttu-id="d6f9b-118">Vá até a guia **Automatizar** e selecione **Todos os Scripts**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-118">Go to the **Automate** tab and select **All Scripts**.</span></span>

2. <span data-ttu-id="d6f9b-119">Selecione **Novo Script**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-119">Select **New Script**.</span></span>

3. <span data-ttu-id="d6f9b-120">Substitua o script padrão pelo script abaixo.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-120">Replace the default script with the following script.</span></span> <span data-ttu-id="d6f9b-121">Esse script adiciona a data e hora atuais às duas primeiras células da planilha **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-121">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="d6f9b-122">Renomeie o script como **Definir data e hora**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-122">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="d6f9b-123">Selecione o nome do script para alterá-lo.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-123">Select the script name to change it.</span></span>

5. <span data-ttu-id="d6f9b-124">Salve o script selecionando **Salvar script**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-124">Save the script by selecting **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="d6f9b-125">Criar um fluxo de trabalho automatizado com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="d6f9b-125">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="d6f9b-126">Entre no [site do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="d6f9b-126">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="d6f9b-127">No menu exibido no lado esquerdo da tela, selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-127">In the menu that's displayed on the left side of the screen, select **Create**.</span></span> <span data-ttu-id="d6f9b-128">Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-128">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="O botão &quot;Criar&quot; do Power Automate.":::

3. <span data-ttu-id="d6f9b-130">Na seção **Começar no espaço em branco**, selecione **Fluxo instantâneo**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-130">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="d6f9b-131">Isso irá criar um fluxo de trabalho ativado manualmente.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-131">This creates a manually activated workflow.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="A opção de fluxo instantâneo do Power Automate para criar um novo fluxo de trabalho.":::

4. <span data-ttu-id="d6f9b-133">Na janela da caixa de diálogo que aparece, insira um nome para seu fluxo na caixa de texto **Nome do fluxo**, selecione **Acionar um fluxo manualmente** na lista de opções em **Escolher como acionar o fluxo**, e em seguida, selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-133">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and then select **Create**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="A opção &quot;Acionar um fluxo manualmente&quot; do Power Automate.":::

    <span data-ttu-id="d6f9b-135">Observe que o fluxo acionado manualmente é apenas um entre os diversos tipos de fluxo.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-135">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="d6f9b-136">No tutorial a seguir, você criará um fluxo que é executado automaticamente quando você recebe um email.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-136">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="d6f9b-137">Selecione **Nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-137">Select **New step**.</span></span>

6. <span data-ttu-id="d6f9b-138">Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-138">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Opção do Excel Online (Business) no Power Automate.":::

7. <span data-ttu-id="d6f9b-140">Em **Ações**, selecione **Executar script**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-140">Under **Actions**, select **Run script**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Executar a opção de ação de script no Power Automate.":::

8. <span data-ttu-id="d6f9b-142">Depois, você selecionará a pasta de trabalho e o script que será utilizado na etapa do fluxo.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-142">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="d6f9b-143">Para o tutorial, você fará o uso da pasta de trabalho criada no seu OneDrive, mas é possível usar qualquer pasta de trabalho em um site OneDrive ou no Microsoft Office SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-143">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="d6f9b-144">Especifique as seguintes configurações para o conector **Executar Script**:</span><span class="sxs-lookup"><span data-stu-id="d6f9b-144">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="d6f9b-145">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="d6f9b-145">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="d6f9b-146">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="d6f9b-146">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="d6f9b-147">**Arquivo**: MyWorkbook.xlsx *(Escolhido por meio do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="d6f9b-147">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="d6f9b-148">**Script**: Definir data e hora</span><span class="sxs-lookup"><span data-stu-id="d6f9b-148">**Script**: Set date and time</span></span>

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="Configurações do conector para executar um script no Power Automate.":::

9. <span data-ttu-id="d6f9b-150">Selecione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-150">Select **Save**.</span></span>

<span data-ttu-id="d6f9b-151">Seu fluxo agora está pronto para ser executado por meio do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-151">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="d6f9b-152">Você pode testá-lo usando o botão **Testar** no editor de fluxo ou seguir as etapas restantes do tutorial para executar o fluxo a partir da sua coleção de fluxos.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-152">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="d6f9b-153">Executar o script por meio da automação</span><span class="sxs-lookup"><span data-stu-id="d6f9b-153">Run the script through Power Automate</span></span>

1. <span data-ttu-id="d6f9b-154">Na página principal do Power Automate, selecione **Meus fluxos**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-154">From the main Power Automate page, select **My flows**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Botão Meus fluxos no Power Automate.":::

2. <span data-ttu-id="d6f9b-156">Selecione **Fluxo do meu tutorial** na lista de fluxos exibida na guia **Meus fluxos**. Isso irá lhe mostrar os detalhes do fluxo que criamos anteriormente.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-156">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="d6f9b-157">Selecione **Executar**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-157">Select **Run**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="Botão Executar no Power Automate.":::

4. <span data-ttu-id="d6f9b-159">Um painel de tarefas irá aparecer para executar o fluxo.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-159">A task pane will appear for running the flow.</span></span> <span data-ttu-id="d6f9b-160">Se você for solicitado a **Entrar** no Excel Online, entre selecionando **Continuar**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-160">If you are asked to **Sign in** to Excel Online, do so by selecting **Continue**.</span></span>

5. <span data-ttu-id="d6f9b-161">Selecione **Executar fluxo**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-161">Select **Run flow**.</span></span> <span data-ttu-id="d6f9b-162">Isso executará o fluxo, que, por sua vez, executará o Script do Office relacionado.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-162">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="d6f9b-163">Selecione **Concluído**.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-163">Select **Done**.</span></span> <span data-ttu-id="d6f9b-164">Você deverá ver a seção **Executar** ser atualizada de acordo.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-164">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="d6f9b-165">Atualize a página para ver os resultados do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-165">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="d6f9b-166">Se o script tiver sido bem-sucedido, vá para a pasta de trabalho para ver as células atualizadas.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-166">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="d6f9b-167">Se tiver falhado, verifique as configurações do fluxo e execute-o novamente.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-167">If it failed, verify the flow's settings and run it a second time.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="Resultado do Power Automate mostrando um fluxo executado com sucesso.":::

## <a name="next-steps"></a><span data-ttu-id="d6f9b-169">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="d6f9b-169">Next steps</span></span>

<span data-ttu-id="d6f9b-170">Faça o tutorial [Transferir dados para scripts em um fluxo executado automaticamente pelo Power Automate](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="d6f9b-170">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="d6f9b-171">O tutorial ensinará como transferir dados de um serviço de fluxo de trabalho para o seu Script do Office e executar o fluxo do Power Automate quando certos eventos ocorrerem.</span><span class="sxs-lookup"><span data-stu-id="d6f9b-171">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
