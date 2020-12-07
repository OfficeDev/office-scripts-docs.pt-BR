---
title: Passar dados para scripts numa execução automática do fluxo no Power Automate.
description: Tutorial sobre executar os Scripts do Office para Excel na Web por meio do Power Automate quando emails são recebidos e transmitidos para o script.
ms.date: 11/30/2020
localization_priority: Priority
ms.openlocfilehash: b73f40c70669fedbe8a0adcf346995cb20b62d37
ms.sourcegitcommit: af487756dffea0f8f0cd62710c586842cb08073c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49571476"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="89848-103">Passar dados para scripts em modo de execução automático no fluxo do Power Automate (visualização)</span><span class="sxs-lookup"><span data-stu-id="89848-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="89848-104">Este tutorial ensina como usar um script do Office para Excel na Web fluxo automatizado[ do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="89848-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="89848-105">Seu script irá automaticamente ser executado toda vez que você receber um email, gravando informações do email em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="89848-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span> <span data-ttu-id="89848-106">Ser capaz de passar os dados de outros aplicativos para um Script do Office oferece a você uma grande flexibilidade e liberdade nos seus processos automatizados.</span><span class="sxs-lookup"><span data-stu-id="89848-106">Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.</span></span>

> [!TIP]
> <span data-ttu-id="89848-107">Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="89848-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="89848-108">Se você for novo no Power Automate, recomendamos começar com o [tutorial Chamar scripts do manual de fluxo do Power Automate](excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="89848-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial.</span></span> <span data-ttu-id="89848-109">[Os Scripts do Office usam TypeScript](../overview/code-editor-environment.md) e este tutorial se destina a pessoas com conhecimento de nível iniciante a intermediário em JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="89848-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="89848-110">Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="89848-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="89848-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="89848-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="89848-112">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="89848-112">Prepare the workbook</span></span>

<span data-ttu-id="89848-113">O Power Automate não pode usar[referências relativas](../develop/power-automate-integration.md#avoid-using-relative-references)como`Workbook.getActiveWorksheet`acessar componentes da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="89848-113">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="89848-114">Portanto, precisamos de uma pasta de trabalho e planilha com nomes consistentes para que o Power Automate possa consultar.</span><span class="sxs-lookup"><span data-stu-id="89848-114">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="89848-115">Criar um nome para a pasta de trabalho **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="89848-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="89848-116">Vá para a guia **Automatizar** e selecione **Editor de Códigos**.</span><span class="sxs-lookup"><span data-stu-id="89848-116">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="89848-117">Selecione **Novo script**.</span><span class="sxs-lookup"><span data-stu-id="89848-117">Select **New Script**.</span></span>

4. <span data-ttu-id="89848-118">Substitua o código existente pelo seguinte script e pressione **Executar**.</span><span class="sxs-lookup"><span data-stu-id="89848-118">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="89848-119">Isso instalará a pasta de trabalho com nomes consistentes de planilhas, tabela e tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="89848-119">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script"></a><span data-ttu-id="89848-120">Criar um Script do Office</span><span class="sxs-lookup"><span data-stu-id="89848-120">Create an Office Script</span></span>

<span data-ttu-id="89848-121">Vamos criar um script que registre as informações de um email.</span><span class="sxs-lookup"><span data-stu-id="89848-121">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="89848-122">Gostaríamos de saber como quais dias da semana recebemos a maioria dos emails e quantos remetentes exclusivos estão enviando esses emails.</span><span class="sxs-lookup"><span data-stu-id="89848-122">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="89848-123">Nossa pasta de trabalho tem uma tabela com **Data**, **Dia da semana**, **Endereços de email** e **Colunas de assunto**.</span><span class="sxs-lookup"><span data-stu-id="89848-123">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="89848-124">Nossa planilha também tem uma tabela dinâmica que está sendo dinamizada no **Dia da semana** e **Endereços de email**(essas são as hierarquias de linha).</span><span class="sxs-lookup"><span data-stu-id="89848-124">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="89848-125">A contagem de **assuntos exclusivos** são as informações agregadas que estão sendo exibidas (a hierarquia de dados).</span><span class="sxs-lookup"><span data-stu-id="89848-125">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="89848-126">Faremos com que o nosso script atualize essa tabela dinâmica depois de atualizar a tabela de email.</span><span class="sxs-lookup"><span data-stu-id="89848-126">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="89848-127">Do **Editor de Código**, selecione **Novo Script**.</span><span class="sxs-lookup"><span data-stu-id="89848-127">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="89848-128">O fluxo que criaremos depois no tutorial enviará a informação do nosso script sobre cada email recebido.</span><span class="sxs-lookup"><span data-stu-id="89848-128">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="89848-129">O script precisa aceitar essa entrada pelos parâmetros na `main`função.</span><span class="sxs-lookup"><span data-stu-id="89848-129">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="89848-130">Substitua o script padrão com o script seguinte:</span><span class="sxs-lookup"><span data-stu-id="89848-130">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="89848-131">O script precisa acessar a tabela e a tabela dinâmica da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="89848-131">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="89848-132">Adicione o seguinte código ao corpo do script após a abertura`{`:</span><span class="sxs-lookup"><span data-stu-id="89848-132">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="89848-133">O `dateReceived`parâmetro é do tipo`string`.</span><span class="sxs-lookup"><span data-stu-id="89848-133">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="89848-134">Vamos convertê-la em um[`Date`objeto](../develop/javascript-objects.md#date)para que possamos facilmente obter o dia da semana.</span><span class="sxs-lookup"><span data-stu-id="89848-134">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="89848-135">Depois de fazer isso, será necessário mapear o valor numérico do dia para uma versão mais legível.</span><span class="sxs-lookup"><span data-stu-id="89848-135">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="89848-136">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="89848-136">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. <span data-ttu-id="89848-137">A cadeia`subject` pode incluir a marca de resposta "RE:".</span><span class="sxs-lookup"><span data-stu-id="89848-137">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="89848-138">Vamos remover isso da cadeia de caracteres para que os emails no mesmo thread tenham o mesmo assunto para a tabela.</span><span class="sxs-lookup"><span data-stu-id="89848-138">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="89848-139">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="89848-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="89848-140">Agora que os dados de email foram formatados da nossa preferência, vamos adicionar uma linha à tabela de email.</span><span class="sxs-lookup"><span data-stu-id="89848-140">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="89848-141">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="89848-141">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. <span data-ttu-id="89848-142">Por fim, vamos verificar se a tabela dinâmica está atualizada.</span><span class="sxs-lookup"><span data-stu-id="89848-142">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="89848-143">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="89848-143">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="89848-144">Renomeie seu script **Registre o email** e pressione **Salvar script**.</span><span class="sxs-lookup"><span data-stu-id="89848-144">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="89848-145">O seu script já está pronto para um fluxo de trabalho automatizado.</span><span class="sxs-lookup"><span data-stu-id="89848-145">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="89848-146">Ele deverá ser semelhante ao script a seguir:</span><span class="sxs-lookup"><span data-stu-id="89848-146">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="89848-147">Criar um fluxo de trabalho automatizado com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="89848-147">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="89848-148">Entre no [site do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="89848-148">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="89848-149">No menu exibido do lado esquerdo da tela, pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="89848-149">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="89848-150">Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="89848-150">This brings you to list of ways to create new workflows.</span></span>

    ![O botão Criar no Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="89848-152">Na seção **Começar no espaço em branco**, selecione **Fluxo automático**.</span><span class="sxs-lookup"><span data-stu-id="89848-152">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="89848-153">Isso cria um fluxo de trabalho iniciado por um evento, como o recebimento de emails.</span><span class="sxs-lookup"><span data-stu-id="89848-153">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![A opção de fluxo automatizado em Power Automate.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="89848-155">Na caixa de diálogo exibida, insira o nome para seu fluxo na **caixa de texto** Nome de Fluxo.</span><span class="sxs-lookup"><span data-stu-id="89848-155">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="89848-156">Em seguida, selecione **Quando um novo email chegar** da lista de opções em **escolha o gatilho de fluxo**.</span><span class="sxs-lookup"><span data-stu-id="89848-156">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="89848-157">Talvez seja necessário procurar pela opção usando a caixa de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="89848-157">You may need to search for the option using the search box.</span></span> <span data-ttu-id="89848-158">Por fim, pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="89848-158">Finally, press **Create**.</span></span>

    ![Parte da janela Criar Uma Fluxo Automatizado no Power Automate, que mostra a opção "novo email recebido".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="89848-160">Este tutorial usa o Outlook.</span><span class="sxs-lookup"><span data-stu-id="89848-160">This tutorial uses Outlook.</span></span> <span data-ttu-id="89848-161">Sinta-se à vontade para usar o seu serviço de email preferido, embora algumas opções possam ser diferentes.</span><span class="sxs-lookup"><span data-stu-id="89848-161">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="89848-162">Pressione **Nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="89848-162">Press **New step**.</span></span>

6. <span data-ttu-id="89848-163">Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="89848-163">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![A opção do Power Automate para Excel Online (Business)](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="89848-165">Em **Ações**, selecione **executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="89848-165">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Opção de ação do Power Automate para Executar script (visualização).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="89848-167">Depois, você selecionará a pasta de trabalho, o script e os argumentos de entrada do script para usar na etapa do fluxo.</span><span class="sxs-lookup"><span data-stu-id="89848-167">Next, you'll select the workbook, script, and script input arguments to use in the flow step.</span></span> <span data-ttu-id="89848-168">Para o tutorial, você fará o uso da pasta de trabalho criada no seu OneDrive, mas é possível usar qualquer pasta de trabalho em um site OneDrive ou no Microsoft Office SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="89848-168">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="89848-169">Especifique as seguintes configurações para o conector **Executar Script**:</span><span class="sxs-lookup"><span data-stu-id="89848-169">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="89848-170">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="89848-170">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="89848-171">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="89848-171">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="89848-172">**Arquivo**: MyWorkbook.xlsx *(Escolhido por meio do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="89848-172">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="89848-173">**Script**: Gravar Email</span><span class="sxs-lookup"><span data-stu-id="89848-173">**Script**: Record Email</span></span>
    - <span data-ttu-id="89848-174">**De**: De *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="89848-174">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="89848-175">**DateReceived**: Hora Recebida *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="89848-175">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="89848-176">**assunto**: Assunto *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="89848-176">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="89848-177">*Observe que os parâmetros para o script só aparecerão quando o script for selecionado.*</span><span class="sxs-lookup"><span data-stu-id="89848-177">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![Opção de ação dos parâmetros do Power Automate para Executar script (visualização).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="89848-179">Pressione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="89848-179">Press **Save**.</span></span>

<span data-ttu-id="89848-180">Agora, o fluxo está habilitado.</span><span class="sxs-lookup"><span data-stu-id="89848-180">Your flow is now enabled.</span></span> <span data-ttu-id="89848-181">O seu script será automaticamente executado sempre que você receber um email por meio do Outlook.</span><span class="sxs-lookup"><span data-stu-id="89848-181">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="89848-182">Gerenciar o script no Power Automate</span><span class="sxs-lookup"><span data-stu-id="89848-182">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="89848-183">Na página principal do Power Automate, selecione **Meus fluxos**.</span><span class="sxs-lookup"><span data-stu-id="89848-183">From the main Power Automate page, select **My flows**.</span></span>

    ![Botão Meus fluxos no Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="89848-185">Selecione o seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="89848-185">Select your flow.</span></span> <span data-ttu-id="89848-186">Aqui você pode ver o histórico de execução.</span><span class="sxs-lookup"><span data-stu-id="89848-186">Here you can see the run history.</span></span> <span data-ttu-id="89848-187">Você pode atualizar a página ou pressionar o botão atualizar **Executar Todos** para atualizar o histórico.</span><span class="sxs-lookup"><span data-stu-id="89848-187">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="89848-188">O fluxo será disparado logo após o recebimento de um email.</span><span class="sxs-lookup"><span data-stu-id="89848-188">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="89848-189">Testar o fluxo enviando a si mesmo um email.</span><span class="sxs-lookup"><span data-stu-id="89848-189">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="89848-190">Quando o fluxo é acionado e executa seu script com sucesso, você deverá ver as atualizações da planilha na pasta de trabalho e da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="89848-190">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![A tabela de email após o fluxo ter sido executado algumas vezes.](../images/power-automate-params-tutorial-4.png)

![A tabela dinâmica após o fluxo ter sido executado algumas vezes.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="89848-193">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="89848-193">Next steps</span></span>

<span data-ttu-id="89848-194">Visite [executar os Scripts do Office com o Power Automate](../develop/power-automate-integration.md) para saber mais sobre como conectar Scripts do Office com o Power Automate.</span><span class="sxs-lookup"><span data-stu-id="89848-194">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="89848-195">Você também pode conferir o exemplo de [lembretes automáticos de tarefas](../resources/scenarios/task-reminders.md) para aprender a combinar os Scripts do Office e Power Automate com as placas adaptáveis de equipes.</span><span class="sxs-lookup"><span data-stu-id="89848-195">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
