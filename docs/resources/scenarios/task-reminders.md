---
title: 'Office Cenário de exemplo de scripts: lembretes de tarefas automatizados'
description: Um exemplo que usa Power Automate e Cartões Adaptáveis automatizam lembretes de tarefas em uma planilha de gerenciamento de projeto.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c5515abb1e36d1bf588ab034f62dfda2625c65dc
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232855"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="6dd8d-103">Office Cenário de exemplo de scripts: lembretes de tarefas automatizados</span><span class="sxs-lookup"><span data-stu-id="6dd8d-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="6dd8d-104">Nesse cenário, você está gerenciando um projeto.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="6dd8d-105">Você usa uma planilha Excel para acompanhar o status de seus funcionários todos os meses.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="6dd8d-106">Muitas vezes, você precisa lembrar as pessoas para preencher seu status, então você decidiu automatizar esse processo de lembrete.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="6dd8d-107">Você criará um fluxo Power Automate mensagens para pessoas com campos de status ausentes e aplicará suas respostas à planilha.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="6dd8d-108">Para fazer isso, você desenvolverá um par de scripts para lidar com o trabalho com a workbook.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="6dd8d-109">O primeiro script obtém uma lista de pessoas com status em branco e o segundo script adiciona uma cadeia de caracteres de status à linha direita.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="6dd8d-110">Você também usará cartões [](/microsoftteams/platform/task-modules-and-cards/what-are-cards) adaptáveis Teams para que os funcionários insiram o status diretamente da notificação.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="6dd8d-111">Habilidades de script abordadas</span><span class="sxs-lookup"><span data-stu-id="6dd8d-111">Scripting skills covered</span></span>

- <span data-ttu-id="6dd8d-112">Criar fluxos em Power Automate</span><span class="sxs-lookup"><span data-stu-id="6dd8d-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="6dd8d-113">Passar dados para scripts</span><span class="sxs-lookup"><span data-stu-id="6dd8d-113">Pass data to scripts</span></span>
- <span data-ttu-id="6dd8d-114">Retornar dados de scripts</span><span class="sxs-lookup"><span data-stu-id="6dd8d-114">Return data from scripts</span></span>
- <span data-ttu-id="6dd8d-115">Teams Cartões adaptáveis</span><span class="sxs-lookup"><span data-stu-id="6dd8d-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="6dd8d-116">Tabelas</span><span class="sxs-lookup"><span data-stu-id="6dd8d-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6dd8d-117">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6dd8d-117">Prerequisites</span></span>

<span data-ttu-id="6dd8d-118">Este cenário usa [Power Automate](https://flow.microsoft.com) e [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="6dd8d-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="6dd8d-119">Você precisará de ambos associados à conta que você usa para desenvolver Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="6dd8d-120">Para ter acesso gratuito a uma assinatura do Microsoft Developer para saber mais sobre e trabalhar com esses aplicativos, considere ingressar no programa Microsoft 365 [desenvolvedor.](https://developer.microsoft.com/microsoft-365/dev-program)</span><span class="sxs-lookup"><span data-stu-id="6dd8d-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="6dd8d-121">Instruções de instalação</span><span class="sxs-lookup"><span data-stu-id="6dd8d-121">Setup instructions</span></span>

1. <span data-ttu-id="6dd8d-122">Baixe <a href="task-reminders.xlsx">task-reminders.xlsx</a> para seu OneDrive.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="6dd8d-123">Abra a Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="6dd8d-124">Na guia **Automatizar,** abra **Todos os Scripts.**</span><span class="sxs-lookup"><span data-stu-id="6dd8d-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="6dd8d-125">Primeiro, precisamos de um script para obter todos os funcionários com relatórios de status ausentes na planilha.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="6dd8d-126">No painel de tarefas Editor de **Código,** pressione **Novo Script** e colar o seguinte script no editor.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

5. <span data-ttu-id="6dd8d-127">Salve o script com o nome **Get People**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="6dd8d-128">Em seguida, precisamos de um segundo script para processar os cartões de relatório de status e colocar as novas informações na planilha.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="6dd8d-129">No painel de tarefas Editor de **Código,** pressione **Novo Script** e colar o seguinte script no editor.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

7. <span data-ttu-id="6dd8d-130">Salve o script com o nome **Salvar Status**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="6dd8d-131">Agora, precisamos criar o fluxo.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-131">Now, we need to create the flow.</span></span> <span data-ttu-id="6dd8d-132">Abra [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="6dd8d-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="6dd8d-133">Se você não tiver criado um fluxo antes, confira nosso tutorial Comece a usar [scripts](../../tutorials/excel-power-automate-manual.md) com Power Automate para aprender o básico.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="6dd8d-134">Criar um novo **fluxo instantâneo.**</span><span class="sxs-lookup"><span data-stu-id="6dd8d-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="6dd8d-135">Escolha **Disparar manualmente um fluxo** das opções e pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="6dd8d-136">O fluxo precisa chamar o script **Obter Pessoas** para obter todos os funcionários com campos de status vazios.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="6dd8d-137">Pressione **Nova etapa** e selecione Excel Online **(Business)**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="6dd8d-138">Em **Ações**, selecione **Executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="6dd8d-139">Forneça as seguintes entradas para a etapa de fluxo:</span><span class="sxs-lookup"><span data-stu-id="6dd8d-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="6dd8d-140">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="6dd8d-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="6dd8d-141">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="6dd8d-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="6dd8d-142">**Arquivo**: task-reminders.xlsx *(Escolhido por meio do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="6dd8d-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="6dd8d-143">**Script**: Obter pessoas</span><span class="sxs-lookup"><span data-stu-id="6dd8d-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="O Power Automate que mostra a primeira etapa executar fluxo de script":::

12. <span data-ttu-id="6dd8d-145">Em seguida, o fluxo precisa processar cada Funcionário na matriz retornada pelo script.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="6dd8d-146">Pressione **Nova etapa** e selecione Postar um Cartão **Adaptável para um** Teams usuário e aguarde uma resposta .</span><span class="sxs-lookup"><span data-stu-id="6dd8d-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="6dd8d-147">Para o **campo Destinatário,** adicione **email** do conteúdo dinâmico (a seleção terá o logotipo Excel por ele).</span><span class="sxs-lookup"><span data-stu-id="6dd8d-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="6dd8d-148">Adicionar **email** faz com que a etapa de fluxo seja cercada por um **Apply a cada** bloco.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="6dd8d-149">Isso significa que a matriz será iterada por Power Automate.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="6dd8d-150">O envio de um Cartão Adaptável exige que o JSON do cartão seja fornecido como **a Mensagem**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="6dd8d-151">Você pode usar o [Designer de Cartão Adaptável](https://adaptivecards.io/designer/) para criar cartões personalizados.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="6dd8d-152">Para este exemplo, use o seguinte JSON.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-152">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

15. <span data-ttu-id="6dd8d-153">Preencha os campos restantes da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="6dd8d-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="6dd8d-154">**Mensagem de atualização**: Obrigado por enviar seu relatório de status.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="6dd8d-155">Sua resposta foi adicionada com êxito à planilha.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="6dd8d-156">**Deve atualizar o cartão**: Sim</span><span class="sxs-lookup"><span data-stu-id="6dd8d-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="6dd8d-157">No bloco **Aplicar a cada** bloco, após o Post an **Adaptive Card** to a Teams user and wait for a response , pressione **Adicionar uma ação**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="6dd8d-158">Selecione **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="6dd8d-159">Em **Ações**, selecione **Executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="6dd8d-160">Forneça as seguintes entradas para a etapa de fluxo:</span><span class="sxs-lookup"><span data-stu-id="6dd8d-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="6dd8d-161">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="6dd8d-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="6dd8d-162">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="6dd8d-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="6dd8d-163">**Arquivo**: task-reminders.xlsx *(Escolhido por meio do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="6dd8d-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="6dd8d-164">**Script**: Salvar Status</span><span class="sxs-lookup"><span data-stu-id="6dd8d-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="6dd8d-165">**senderEmail**: email *(conteúdo dinâmico do Excel)*</span><span class="sxs-lookup"><span data-stu-id="6dd8d-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="6dd8d-166">**statusReportResponse**: resposta *(conteúdo dinâmico de Teams)*</span><span class="sxs-lookup"><span data-stu-id="6dd8d-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="O Power Automate que mostra a etapa aplicar a cada etapa":::

17. <span data-ttu-id="6dd8d-168">Salve o fluxo.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="6dd8d-169">Executando o fluxo</span><span class="sxs-lookup"><span data-stu-id="6dd8d-169">Running the flow</span></span>

<span data-ttu-id="6dd8d-170">Para testar o fluxo, certifique-se de que quaisquer linhas de tabela com status em branco usem um endereço de email vinculado a uma conta Teams cliente (você provavelmente deve usar seu próprio endereço de email durante o teste).</span><span class="sxs-lookup"><span data-stu-id="6dd8d-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="6dd8d-171">Você pode selecionar **Testar no** designer de fluxo ou executar o fluxo na página **Meus fluxos.**</span><span class="sxs-lookup"><span data-stu-id="6dd8d-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="6dd8d-172">Depois de iniciar o fluxo e aceitar o uso das conexões necessárias, você deve receber um Cartão Adaptável de Power Automate até Teams.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="6dd8d-173">Depois de preencher o campo de status no cartão, o fluxo continuará e atualizará a planilha com o status que você fornece.</span><span class="sxs-lookup"><span data-stu-id="6dd8d-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="6dd8d-174">Antes de executar o fluxo</span><span class="sxs-lookup"><span data-stu-id="6dd8d-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Uma planilha com um relatório de status contendo uma entrada de status ausente":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="6dd8d-176">Receber o Cartão Adaptável</span><span class="sxs-lookup"><span data-stu-id="6dd8d-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Um Cartão Adaptável em Teams solicitando ao funcionário uma atualização de status":::

### <a name="after-running-the-flow"></a><span data-ttu-id="6dd8d-178">Depois de executar o fluxo</span><span class="sxs-lookup"><span data-stu-id="6dd8d-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Uma planilha com um relatório de status com uma entrada de status agora preenchida":::
