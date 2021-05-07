---
title: Retorna dados de um script para um fluxo do Power Automate executado automaticamente
description: Um tutorial que mostra como enviar emails de lembrete executando Scripts do Office para o Excel na web através do Power Automate.
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: 54fcfc773d4d2a8d352f7bd22593ac817e7ded0e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232876"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="18a39-103">Retorna dados de um script para um fluxo do Power Automate executado automaticamente (visualização)</span><span class="sxs-lookup"><span data-stu-id="18a39-103">Return data from a script to an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="18a39-104">Este tutorial ensina como retornar informações de um Script do Office para o Excel na web como parte de um fluxo de trabalho automatizado do [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="18a39-104">This tutorial teaches you how to return information from an Office Script for Excel on the web as part of an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="18a39-105">Você fará um script que olha através de uma programação e trabalha com um fluxo para enviar emails de lembrete.</span><span class="sxs-lookup"><span data-stu-id="18a39-105">You'll make a script that looks through a schedule and works with a flow to send reminder emails.</span></span> <span data-ttu-id="18a39-106">Esse fluxo será executado em uma programação regular, fornecendo esses lembretes em seu nome.</span><span class="sxs-lookup"><span data-stu-id="18a39-106">This flow will run on a regular schedule, providing these reminders on your behalf.</span></span>

> [!TIP]
> <span data-ttu-id="18a39-107">Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="18a39-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>
>
> <span data-ttu-id="18a39-108">Se você é novo no Power Automate, recomendamos começar com os tutoriais [ Chamada de scripts de um fluxo manual do Power Automate](excel-power-automate-manual.md) e [Passar dados para scripts em um fluxo automático do Power Automate](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="18a39-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) and [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorials.</span></span>
>
> <span data-ttu-id="18a39-109">[Os Scripts do Office usam TypeScript](../overview/code-editor-environment.md) e este tutorial se destina a pessoas com conhecimento de nível iniciante a intermediário em JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="18a39-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="18a39-110">Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="18a39-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="18a39-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="18a39-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="18a39-112">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="18a39-112">Prepare the workbook</span></span>

1. <span data-ttu-id="18a39-113">Baixe a pasta de trabalho <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> para o seu OneDrive.</span><span class="sxs-lookup"><span data-stu-id="18a39-113">Download the workbook <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="18a39-114">Abra **on-call-rotation.xlsx** no Excel na web.</span><span class="sxs-lookup"><span data-stu-id="18a39-114">Open **on-call-rotation.xlsx** in Excel on the web.</span></span>

1. <span data-ttu-id="18a39-115">Adicione uma linha à tabela com seu nome, endereço de email e datas de início e fim que coincidam com a data atual.</span><span class="sxs-lookup"><span data-stu-id="18a39-115">Add a row to the table with your name, email address, and start and end dates that overlap with the current date.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="18a39-116">O roteiro que você vai escrever utiliza a primeira entrada correspondente na tabela, portanto, certifique-se de que seu nome esteja acima de qualquer linha com a semana atual.</span><span class="sxs-lookup"><span data-stu-id="18a39-116">The script you'll write uses the first matching entry in the table, so make sure your name is above any row with the current week.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="Uma planilha contendo os dados da tabela de rotação de plantão":::

## <a name="create-an-office-script"></a><span data-ttu-id="18a39-118">Criar um Script do Office</span><span class="sxs-lookup"><span data-stu-id="18a39-118">Create an Office Script</span></span>

1. <span data-ttu-id="18a39-119">Vá até a guia **Automatizar** e selecione **Todos os Scripts**.</span><span class="sxs-lookup"><span data-stu-id="18a39-119">Go to the **Automate** tab and select **All Scripts**.</span></span>

1. <span data-ttu-id="18a39-120">Selecione **Novo Script**.</span><span class="sxs-lookup"><span data-stu-id="18a39-120">Select **New Script**.</span></span>

1. <span data-ttu-id="18a39-121">Nomeie o script **Obter uma Pessoa de Plantão**.</span><span class="sxs-lookup"><span data-stu-id="18a39-121">Name the script **Get On-Call Person**.</span></span>

1. <span data-ttu-id="18a39-122">Agora você deve ter um script vazio.</span><span class="sxs-lookup"><span data-stu-id="18a39-122">You should now have an empty script.</span></span> <span data-ttu-id="18a39-123">Queremos usar o roteiro para obter um endereço de email da planilha.</span><span class="sxs-lookup"><span data-stu-id="18a39-123">We want to use the script to get an email address from the spreadsheet.</span></span> <span data-ttu-id="18a39-124">Altere `main` para retornar uma cadeia de caracteres, como esta:</span><span class="sxs-lookup"><span data-stu-id="18a39-124">Change `main` to return a string, like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. <span data-ttu-id="18a39-125">Em seguida, precisamos obter todos os dados da tabela.</span><span class="sxs-lookup"><span data-stu-id="18a39-125">Next, we need to get all the data from the table.</span></span> <span data-ttu-id="18a39-126">Isso nos permite examinar cada linha com o script.</span><span class="sxs-lookup"><span data-stu-id="18a39-126">That lets us look through each row with the script.</span></span> <span data-ttu-id="18a39-127">Adicione o seguinte código dentro da função `main`.</span><span class="sxs-lookup"><span data-stu-id="18a39-127">Add the following code inside the `main` function.</span></span>

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. <span data-ttu-id="18a39-128">As datas na tabela são armazenadas usando o [Número de série da data do Excel ](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span><span class="sxs-lookup"><span data-stu-id="18a39-128">The dates in the table are stored using [Excel's date serial number](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span></span> <span data-ttu-id="18a39-129">Precisamos converter essas datas para datas JavaScript a fim de compará-las.</span><span class="sxs-lookup"><span data-stu-id="18a39-129">We need to convert those dates to JavaScript dates in order to compare them.</span></span> <span data-ttu-id="18a39-130">Adicionaremos uma função auxiliar ao nosso script.</span><span class="sxs-lookup"><span data-stu-id="18a39-130">We'll add a helper function to our script.</span></span> <span data-ttu-id="18a39-131">Adicione o seguinte código fora da função `main`:</span><span class="sxs-lookup"><span data-stu-id="18a39-131">Add the following code outside of the `main` function:</span></span>

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. <span data-ttu-id="18a39-132">Agora, precisamos descobrir qual pessoa está de plantão agora.</span><span class="sxs-lookup"><span data-stu-id="18a39-132">Now, we need to figure out which person is on call right now.</span></span> <span data-ttu-id="18a39-133">A linha deles terá uma data de início e de término em torno da data atual.</span><span class="sxs-lookup"><span data-stu-id="18a39-133">Their row will have a start and end date surrounding the current date.</span></span> <span data-ttu-id="18a39-134">Escreveremos um script para assumir que apenas uma pessoa está de plantão por vez.</span><span class="sxs-lookup"><span data-stu-id="18a39-134">We'll write the script to assume only one person is on call at a time.</span></span> <span data-ttu-id="18a39-135">Os scripts podem retornar matrizes para lidar com múltiplos valores, mas por enquanto retornaremos o primeiro endereço de email correspondente.</span><span class="sxs-lookup"><span data-stu-id="18a39-135">Scripts can return arrays to handle multiple values, but for now we'll return the first matching email address.</span></span> <span data-ttu-id="18a39-136">Adicione o seguinte código ao final da função `main`.</span><span class="sxs-lookup"><span data-stu-id="18a39-136">Add the following code to the end of the `main` function.</span></span>

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. <span data-ttu-id="18a39-137">O script final deve ser semelhante a este:</span><span class="sxs-lookup"><span data-stu-id="18a39-137">The final script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="18a39-138">Criar um fluxo de trabalho automatizado com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="18a39-138">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="18a39-139">Entre no [site do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="18a39-139">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

1. <span data-ttu-id="18a39-140">No menu exibido do lado esquerdo da tela, pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="18a39-140">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="18a39-141">Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="18a39-141">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Botão Criar no Power Automate":::

1. <span data-ttu-id="18a39-143">Na seção **Iniciar de Modelo em Branco**, selecione **Fluxo de nuvem agendado**.</span><span class="sxs-lookup"><span data-stu-id="18a39-143">Under the **Start from blank** section, select **Scheduled cloud flow**.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="O botão Fluxo de nuvem agendado no Power Automate":::

1. <span data-ttu-id="18a39-145">Agora precisamos definir o cronograma para esse fluxo.</span><span class="sxs-lookup"><span data-stu-id="18a39-145">Now we need to set the schedule for this flow.</span></span> <span data-ttu-id="18a39-146">Nossa planilha tem uma nova atribuição de plantão começando toda segunda-feira no primeiro semestre de 2021.</span><span class="sxs-lookup"><span data-stu-id="18a39-146">Our spreadsheet has a new on-call assignment starting every Monday in the first half of 2021.</span></span> <span data-ttu-id="18a39-147">Vamos definir o fluxo para começar nas manhãs de segunda-feira.</span><span class="sxs-lookup"><span data-stu-id="18a39-147">Let's set the flow to run first thing Monday mornings.</span></span> <span data-ttu-id="18a39-148">Use as seguintes opções para configurar o fluxo a ser executado na segunda-feira de cada semana.</span><span class="sxs-lookup"><span data-stu-id="18a39-148">Use the following options to configure the flow to run on Monday each week.</span></span>

    - <span data-ttu-id="18a39-149">**Nome do fluxo**: Notificar a Pessoa de Plantão</span><span class="sxs-lookup"><span data-stu-id="18a39-149">**Flow name**: Notify On-Call Person</span></span>
    - <span data-ttu-id="18a39-150">**Iniciando em**: 4/1/21 à 1h00</span><span class="sxs-lookup"><span data-stu-id="18a39-150">**Starting**: 1/4/21 at 1:00am</span></span>
    - <span data-ttu-id="18a39-151">**Repetir a cada**: 1 Semana</span><span class="sxs-lookup"><span data-stu-id="18a39-151">**Repeat every**: 1 Week</span></span>
    - <span data-ttu-id="18a39-152">**Nesses dias**: M</span><span class="sxs-lookup"><span data-stu-id="18a39-152">**On these days**: M</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="O Diálogo &quot;Construa um fluxo de nuvens programado&quot; do Power Automate mostrando opções. As opções incluem nome do fluxo, hora para começar, quantas vezes repetir e qual dia da semana executar o fluxo":::

1. <span data-ttu-id="18a39-154">Pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="18a39-154">Press **Create**.</span></span>

1. <span data-ttu-id="18a39-155">Pressione **Nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="18a39-155">Press **New step**.</span></span>

1. <span data-ttu-id="18a39-156">Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="18a39-156">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Opção do Excel Online (Business) no Power Automate":::

1. <span data-ttu-id="18a39-158">Em **Ações**, selecione **Executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="18a39-158">Under **Actions**, select **Run script (preview)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Execute a opção de ação de script (visualização) no Power Automate":::

1. <span data-ttu-id="18a39-160">Em seguida, você selecionará a pasta de trabalho e o script que será utilizado na etapa do fluxo.</span><span class="sxs-lookup"><span data-stu-id="18a39-160">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="18a39-161">Use a pasta de trabalho **on-call-rotation.xlsx** que você criou em seu OneDrive.</span><span class="sxs-lookup"><span data-stu-id="18a39-161">Use the **on-call-rotation.xlsx** workbook you created in your OneDrive.</span></span> <span data-ttu-id="18a39-162">Especifique as seguintes configurações para o conector **Executar Script**:</span><span class="sxs-lookup"><span data-stu-id="18a39-162">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="18a39-163">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="18a39-163">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="18a39-164">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="18a39-164">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="18a39-165">**Arquivo**: on-call-rotation.xlsx *(Escolhido através do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="18a39-165">**File**: on-call-rotation.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="18a39-166">**Script**: Obter uma Pessoa de Plantão</span><span class="sxs-lookup"><span data-stu-id="18a39-166">**Script**: Get On-Call Person</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="As configurações do conector do Power Automate para executar um script":::

1. <span data-ttu-id="18a39-168">Pressione **Nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="18a39-168">Press **New step**.</span></span>

1. <span data-ttu-id="18a39-169">Terminaremos o fluxo enviando o email de lembrete.</span><span class="sxs-lookup"><span data-stu-id="18a39-169">We'll end the flow by sending the reminder email.</span></span> <span data-ttu-id="18a39-170">Selecione **Enviar um email (V2)** usando a barra de pesquisa do conector.</span><span class="sxs-lookup"><span data-stu-id="18a39-170">Select **Send an email (V2)** by using the connector's search bar.</span></span> <span data-ttu-id="18a39-171">Use o controle **Adicionar conteúdo dinâmico** para adicionar o endereço de email retornado pelo script.</span><span class="sxs-lookup"><span data-stu-id="18a39-171">Use the **Add dynamic content** control to add the email address returned by the script.</span></span> <span data-ttu-id="18a39-172">Ele será rotulado como **resultado** com o ícone do Excel próximo a ele.</span><span class="sxs-lookup"><span data-stu-id="18a39-172">This will be labelled **result** with the Excel icon next to it.</span></span> <span data-ttu-id="18a39-173">Você pode fornecer qualquer assunto e corpo de texto que desejar.</span><span class="sxs-lookup"><span data-stu-id="18a39-173">You can provide whatever subject and body text you'd like.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="As configurações do conector Power Automate Outlook para enviar um e-mail. As opções incluem o arquivo a ser enviado, o assunto do e-mail e o corpo do e-mail, assim como opções avançadas"::: 

    > [!NOTE]
    > <span data-ttu-id="18a39-p111">Este tutorial usa o Outlook. Sinta-se à vontade para usar seu serviço de e-mail preferido, embora algumas opções possam ser diferentes.</span><span class="sxs-lookup"><span data-stu-id="18a39-p111">This tutorial uses Outlook. Feel free to use your preferred email service instead, though some options may be different.</span></span>

1. <span data-ttu-id="18a39-177">Pressione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="18a39-177">Press **Save**.</span></span>

## <a name="test-the-script-in-power-automate"></a><span data-ttu-id="18a39-178">Teste o script no Power Automate</span><span class="sxs-lookup"><span data-stu-id="18a39-178">Test the script in Power Automate</span></span>

<span data-ttu-id="18a39-179">Seu fluxo funcionará todas as segundas-feiras de manhã.</span><span class="sxs-lookup"><span data-stu-id="18a39-179">Your flow will run every Monday morning.</span></span> <span data-ttu-id="18a39-180">Você pode testar o script agora pressionando o botão **Testar** no canto superior direito da tela.</span><span class="sxs-lookup"><span data-stu-id="18a39-180">You can test the script now by pressing the **Test** button in the upper-right corner of the screen.</span></span> <span data-ttu-id="18a39-181">Selecione **Manualmente** e pressione **Executar Teste** para executar o fluxo agora e testar o comportamento.</span><span class="sxs-lookup"><span data-stu-id="18a39-181">Select **Manually** and press **Run Test** to run the flow now and test the behavior.</span></span> <span data-ttu-id="18a39-182">Pode ser necessário conceder permissões ao Excel e Outlook para continuar.</span><span class="sxs-lookup"><span data-stu-id="18a39-182">You may need to grant permissions to Excel and Outlook to continue.</span></span>

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="O botão Teste do Power Automate":::

> [!TIP]
> <span data-ttu-id="18a39-184">Se o seu fluxo não enviar um email, verifique na planilha se um email válido está listado para o intervalo de datas atual na parte superior da tabela.</span><span class="sxs-lookup"><span data-stu-id="18a39-184">If your flow fails to send an email, double-check in the spreadsheet that a valid email is listed for the current date range at the top of the table.</span></span>

## <a name="next-steps"></a><span data-ttu-id="18a39-185">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="18a39-185">Next steps</span></span>

<span data-ttu-id="18a39-186">Visite [executar os Scripts do Office com o Power Automate](../develop/power-automate-integration.md) para saber mais sobre como conectar Scripts do Office com o Power Automate.</span><span class="sxs-lookup"><span data-stu-id="18a39-186">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="18a39-187">Você também pode conferir o exemplo de [lembretes automáticos de tarefas](../resources/scenarios/task-reminders.md) para aprender a combinar os Scripts do Office e Power Automate com as placas adaptáveis de equipes.</span><span class="sxs-lookup"><span data-stu-id="18a39-187">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
