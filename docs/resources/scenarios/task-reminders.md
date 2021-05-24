---
title: 'Office Cenário de exemplo de scripts: lembretes de tarefas automatizados'
description: Um exemplo que usa Power Automate e Cartões Adaptáveis automatizam lembretes de tarefas em uma planilha de gerenciamento de projeto.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545593"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Cenário de exemplo de scripts: lembretes de tarefas automatizados

Nesse cenário, você está gerenciando um projeto. Você usa uma planilha Excel para acompanhar o status de seus funcionários todos os meses. Muitas vezes, você precisa lembrar as pessoas para preencher seu status, então você decidiu automatizar esse processo de lembrete.

Você criará um fluxo Power Automate mensagens para pessoas com campos de status ausentes e aplicará suas respostas à planilha. Para fazer isso, você desenvolverá um par de scripts para lidar com o trabalho com a workbook. O primeiro script obtém uma lista de pessoas com status em branco e o segundo script adiciona uma cadeia de caracteres de status à linha direita. Você também usará cartões [](/microsoftteams/platform/task-modules-and-cards/what-are-cards) adaptáveis Teams para que os funcionários insiram o status diretamente da notificação.

## <a name="scripting-skills-covered"></a>Habilidades de script abordadas

- Criar fluxos em Power Automate
- Passar dados para scripts
- Retornar dados de scripts
- Teams Cartões adaptáveis
- Tabelas

## <a name="prerequisites"></a>Pré-requisitos

Este cenário usa [Power Automate](https://flow.microsoft.com) e [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Você precisará de ambos associados à conta que você usa para desenvolver Office Scripts. Para ter acesso gratuito a uma assinatura do Microsoft Developer para saber mais sobre e trabalhar com esses aplicativos, considere ingressar no programa Microsoft 365 [desenvolvedor.](https://developer.microsoft.com/microsoft-365/dev-program)

## <a name="setup-instructions"></a>Instruções de instalação

1. Baixe <a href="task-reminders.xlsx">task-reminders.xlsx</a> para seu OneDrive.

2. Abra a Excel na Web.

3. Na guia **Automatizar,** abra **Todos os Scripts.**

4. Primeiro, precisamos de um script para obter todos os funcionários com relatórios de status ausentes na planilha. No painel de tarefas Editor de **Código,** pressione **Novo Script** e colar o seguinte script no editor.

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

5. Salve o script com o nome **Get People**.

6. Em seguida, precisamos de um segundo script para processar os cartões de relatório de status e colocar as novas informações na planilha. No painel de tarefas Editor de **Código,** pressione **Novo Script** e colar o seguinte script no editor.

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

7. Salve o script com o nome **Salvar Status**.

8. Agora, precisamos criar o fluxo. Abra [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Se você não tiver criado um fluxo antes, confira nosso tutorial Comece a usar [scripts](../../tutorials/excel-power-automate-manual.md) com Power Automate para aprender o básico.

9. Criar um novo **fluxo instantâneo.**

10. Escolha **Disparar manualmente um fluxo** das opções e pressione **Criar**.

11. O fluxo precisa chamar o script **Obter Pessoas** para obter todos os funcionários com campos de status vazios. Pressione **Nova etapa** e selecione Excel Online **(Business)**. Em **Ações**, selecione **Executar script**. Forneça as seguintes entradas para a etapa de fluxo:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: task-reminders.xlsx *(Escolhido por meio do navegador de arquivos)*
    - **Script**: Obter pessoas

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="O Power Automate que mostra a primeira etapa executar fluxo de script":::

12. Em seguida, o fluxo precisa processar cada Funcionário na matriz retornada pelo script. Pressione **Nova etapa** e selecione Postar um Cartão **Adaptável para um** Teams usuário e aguarde uma resposta .

13. Para o **campo Destinatário,** adicione **email** do conteúdo dinâmico (a seleção terá o logotipo Excel por ele). Adicionar **email** faz com que a etapa de fluxo seja cercada por um **Apply a cada** bloco. Isso significa que a matriz será iterada por Power Automate.

14. O envio de um Cartão Adaptável exige que o JSON do cartão seja fornecido como **a Mensagem**. Você pode usar o [Designer de Cartão Adaptável](https://adaptivecards.io/designer/) para criar cartões personalizados. Para este exemplo, use o seguinte JSON.  

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

15. Preencha os campos restantes da seguinte forma:

    - **Mensagem de atualização**: Obrigado por enviar seu relatório de status. Sua resposta foi adicionada com êxito à planilha.
    - **Deve atualizar o cartão**: Sim

16. No bloco **Aplicar a cada** bloco, após o Post an **Adaptive Card** to a Teams user and wait for a response , pressione **Adicionar uma ação**. Selecione **Excel Online (Business)**. Em **Ações**, selecione **Executar script**. Forneça as seguintes entradas para a etapa de fluxo:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: task-reminders.xlsx *(Escolhido por meio do navegador de arquivos)*
    - **Script**: Salvar Status
    - **senderEmail**: email *(conteúdo dinâmico do Excel)*
    - **statusReportResponse**: resposta *(conteúdo dinâmico de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="O Power Automate que mostra a etapa aplicar a cada etapa":::

17. Salve o fluxo.

## <a name="running-the-flow"></a>Executando o fluxo

Para testar o fluxo, certifique-se de que quaisquer linhas de tabela com status em branco usem um endereço de email vinculado a uma conta Teams cliente (você provavelmente deve usar seu próprio endereço de email durante o teste).

Você pode selecionar **Testar no** designer de fluxo ou executar o fluxo na página **Meus fluxos.** Depois de iniciar o fluxo e aceitar o uso das conexões necessárias, você deve receber um Cartão Adaptável de Power Automate até Teams. Depois de preencher o campo de status no cartão, o fluxo continuará e atualizará a planilha com o status que você fornece.

### <a name="before-running-the-flow"></a>Antes de executar o fluxo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Uma planilha com um relatório de status contendo uma entrada de status ausente":::

### <a name="receiving-the-adaptive-card"></a>Receber o Cartão Adaptável

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Um Cartão Adaptável em Teams solicitando ao funcionário uma atualização de status":::

### <a name="after-running-the-flow"></a>Depois de executar o fluxo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Uma planilha com um relatório de status com uma entrada de status agora preenchida":::
