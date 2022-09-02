---
title: 'Cenário de exemplo de Scripts do Office: lembretes de tarefas automatizados'
description: Um exemplo que usa o Power Automate e os Cartões Adaptáveis automatizam lembretes de tarefas em uma planilha de gerenciamento de projetos.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 148ee32c9f5b95c8a97e39822e278ebf819c460e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572588"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Cenário de exemplo de Scripts do Office: lembretes de tarefas automatizados

Nesse cenário, você está gerenciando um projeto. Você usa uma planilha do Excel para acompanhar o status de seus funcionários todos os meses. Muitas vezes, você precisa lembrar as pessoas de preencher seu status, portanto, decidiu automatizar esse processo de lembrete.

Você criará um fluxo do Power Automate para enviar mensagens a pessoas com campos de status ausentes e aplicará suas respostas à planilha. Para fazer isso, você desenvolverá um par de scripts para lidar com o trabalho com a pasta de trabalho. O primeiro script obtém uma lista de pessoas com status em branco e o segundo script adiciona uma cadeia de caracteres de status à linha direita. Você também usará os Cartões [Adaptáveis do Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que os funcionários insiram seu status diretamente na notificação.

## <a name="scripting-skills-covered"></a>Habilidades de script cobertas

- Criar fluxos no Power Automate
- Passar dados para scripts
- Retornar dados de scripts
- Cartões Adaptáveis do Teams
- Tabelas

## <a name="prerequisites"></a>Pré-requisitos

Esse cenário usa [o Power Automate](https://flow.microsoft.com) e o [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Você precisará de ambos associados à conta que usa para desenvolver Scripts do Office. Para obter acesso gratuito a uma assinatura do Desenvolvedor da Microsoft para saber mais e trabalhar com esses aplicativos, considere ingressar no Programa de Desenvolvedores do [Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instruções de instalação

1. Baixe [task-reminders.xlsx](task-reminders.xlsx) para o OneDrive.

1. Abra a pasta de trabalho Excel na Web.

1. Primeiro, precisamos de um script para obter todos os funcionários com relatórios de status ausentes da planilha. Na guia **Automatizar** , selecione **Novo Script** e cole o script a seguir no editor.

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

1. Salve o script com o nome **Get Pessoas**.

1. Em seguida, precisamos de um segundo script para processar os cartões de relatório de status e colocar as novas informações na planilha. No painel de tarefas editor de código, selecione **Novo Script** e cole o script a seguir no editor.

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

1. Salve o script com o nome **Salvar Status**.

1. Agora, precisamos criar o fluxo. Abra [o Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Se você ainda não criou um fluxo, confira nosso tutorial Começar a usar [scripts com o Power Automate](../../tutorials/excel-power-automate-manual.md) para aprender os conceitos básicos.

1. Crie um novo **fluxo instantâneo**.

1. Escolha **Disparar manualmente um fluxo nas** opções e selecione **Criar**.

1. O fluxo precisa chamar o script **Get Pessoas** para obter todos os funcionários com campos de status vazios. Selecione **Nova etapa** e, em seguida, **selecione Excel Online (Business)**. Em **Ações**, selecione **Executar script**. Forneça as seguintes entradas para a etapa de fluxo:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: task-reminders.xlsx *(escolhido por meio do navegador de arquivos)*
    - **Script**: Obter Pessoas

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="O fluxo do Power Automate mostrando a primeira etapa de fluxo de script de execução.":::

1. Em seguida, o fluxo precisa processar cada Funcionário na matriz retornada pelo script. Selecione **Nova etapa** e, em seguida, **escolha Postar um Cartão Adaptável para um usuário do Teams e aguarde uma resposta**.

1. Para o **campo Destinatário** , adicione **email** do conteúdo dinâmico (a seleção terá o logotipo do Excel por ele). Adicionar **email** faz com que a etapa de fluxo seja circundada por um **Aplicar a cada** bloco. Isso significa que a matriz será iterada pelo Power Automate.

1. O envio de um Cartão Adaptável exige que o [JSON](https://www.w3schools.com/whatis/whatis_json.asp) do cartão seja fornecido como a **Mensagem**. Você pode usar o [Designer de Cartão Adaptável](https://adaptivecards.io/designer/) para criar cartões personalizados. Para este exemplo, use o JSON a seguir.  

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

1. Preencha os campos restantes da seguinte maneira:

    - **Mensagem de** atualização: Obrigado por enviar seu relatório de status. Sua resposta foi adicionada com êxito à planilha.
    - **Deve atualizar o cartão**: Sim

1. Em Aplicar **a cada bloco** , após a postagem de um Cartão **Adaptável** para um usuário do Teams e aguardar uma resposta, selecione **Adicionar uma ação**. Selecione **Excel Online (Business)**. Em **Ações**, selecione **Executar script**. Forneça as seguintes entradas para a etapa de fluxo:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: task-reminders.xlsx *(escolhido por meio do navegador de arquivos)*
    - **Script**: Salvar Status
    - **senderEmail**: email *(conteúdo dinâmico do Excel)*
    - **statusReportResponse**: resposta *(conteúdo dinâmico do Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="O fluxo do Power Automate mostrando a etapa aplicar a cada etapa.":::

1. Salve o fluxo.

## <a name="running-the-flow"></a>Executando o fluxo

Para testar o fluxo, verifique se as linhas da tabela com status em branco usam um endereço de email vinculado a uma conta do Teams (você provavelmente deve usar seu próprio endereço de email durante o teste). Use o **botão Testar** na página do editor de fluxo ou execute o fluxo por meio da **guia Meus fluxos** . Certifique-se de permitir o acesso quando solicitado.

Você deve receber um Cartão Adaptável do Power Automate por meio do Teams. Depois de preencher o campo de status no cartão, o fluxo continuará e atualizará a planilha com o status fornecido.

### <a name="before-running-the-flow"></a>Antes de executar o fluxo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Uma planilha com um relatório de status que contém uma entrada de status ausente.":::

### <a name="receiving-the-adaptive-card"></a>Recebendo o Cartão Adaptável

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Um Cartão Adaptável no Teams solicitando ao funcionário uma atualização de status.":::

### <a name="after-running-the-flow"></a>Depois de executar o fluxo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Uma planilha com um relatório de status com uma entrada de status agora preenchida.":::
