---
title: 'Office Cenário de exemplo de scripts: Lembretes automatizados de tarefas'
description: Uma amostra que usa Power Automate e Cartões Adaptativos automatiza lembretes de tarefas em uma planilha de gerenciamento de projetos.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545593"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Cenário de exemplo de scripts: Lembretes automatizados de tarefas

Neste cenário você está gerenciando um projeto. Você usa uma planilha Excel para acompanhar o status de seus funcionários todos os meses. Muitas vezes você precisa lembrar as pessoas para preencher seu status, então você decidiu automatizar esse processo de lembrete.

Você criará um fluxo de Power Automate para enviar mensagens às pessoas com campos de status ausentes e aplicar suas respostas à planilha. Para fazer isso, você desenvolverá um par de scripts para lidar com o trabalho com a pasta de trabalho. O primeiro script recebe uma lista de pessoas com status em branco e o segundo script adiciona uma sequência de status à linha direita. Você também fará uso de [Teams Cartões Adaptativos](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que os funcionários insiram seu status diretamente a partir da notificação.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cobertas

- Criar fluxos em Power Automate
- Passe dados para scripts
- Retornar dados de scripts
- Teams Cartões adaptativos
- Tabelas

## <a name="prerequisites"></a>Pré-requisitos

Este cenário utiliza [Power Automate](https://flow.microsoft.com) e [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Você precisará de ambos associados à conta que você usa para desenvolver scripts Office. Para ter acesso gratuito a uma assinatura do Microsoft Developer para aprender e trabalhar com esses aplicativos, considere participar do [programa de desenvolvedores Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instruções de configuração

1. Baixe <a href="task-reminders.xlsx">task-reminders.xlsx</a> para sua OneDrive.

2. Abra a pasta de trabalho em Excel na Web.

3. Na guia **Automate,** abra **todos os scripts**.

4. Primeiro, precisamos de um script para obter todos os funcionários com relatórios de status que estão faltando na planilha. No painel de tarefas do **Editor de Código,** **pressione o Novo Script** e cole o seguinte script no editor.

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

6. Em seguida, precisamos de um segundo script para processar os boletins de status e colocar as novas informações na planilha. No painel de tarefas do **Editor de Código,** **pressione o Novo Script** e cole o seguinte script no editor.

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

7. Salve o script com o nome **Save Status**.

8. Agora, precisamos criar o fluxo. Aberto [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Se você ainda não criou um fluxo antes, confira nosso tutorial [Comece a usar scripts com Power Automate](../../tutorials/excel-power-automate-manual.md) para aprender o básico.

9. Crie um novo **fluxo instantâneo**.

10. Escolha **acionar manualmente um fluxo** das opções e **pressione Criar**.

11. O fluxo precisa chamar o script **Get People** para obter todos os funcionários com campos de status vazios. Pressione **nova etapa** e selecione Excel **Online (Business)**. Em **Ações**, selecione **Executar script**. Forneça as seguintes entradas para a etapa de fluxo:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: task-reminders.xlsx *(Escolhido através do navegador de arquivo)*
    - **Script**: Get People

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="O fluxo Power Automate mostrando a primeira etapa de fluxo de script run":::

12. Em seguida, o fluxo precisa processar cada Funcionário na matriz retornada pelo script. Pressione **nova etapa** e selecione Postar um **cartão adaptativo para um usuário Teams e esperar por uma resposta**.

13. Para o campo **Destinatário,** adicione **e-mail** do conteúdo dinâmico (a seleção terá o logotipo Excel por ele). A adição **de e-mails** faz com que a etapa de fluxo seja cercada por um **Aplicar a cada** bloco. Isso significa que a matriz será iterada por Power Automate.

14. O envio de um Cartão Adaptativo requer que o JSON do cartão seja fornecido como **a Mensagem**. Você pode usar o [Designer de Cartões Adaptativos](https://adaptivecards.io/designer/) para criar cartões personalizados. Para esta amostra, use o seguinte JSON.  

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

    - **Mensagem de atualização**: Obrigado por enviar seu relatório de status. Sua resposta foi adicionada com sucesso à planilha.
    - **Deve atualizar cartão**: Sim

16. No **Aplicar a cada** bloco, seguindo o **Post a Adaptive Card para um usuário Teams e esperar por uma resposta,** **pressione Adicionar uma ação**. Selecione **Excel Online (Negócios)**. Em **Ações**, selecione **Executar script**. Forneça as seguintes entradas para a etapa de fluxo:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: task-reminders.xlsx *(Escolhido através do navegador de arquivo)*
    - **Script**: Salvar status
    - **e-mail**: e-mail *(conteúdo dinâmico de Excel)*
    - **statusReportReponse**: resposta *(conteúdo dinâmico de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="O fluxo Power Automate mostrando a etapa de aplicação a cada passo":::

17. Guarde o fluxo.

## <a name="running-the-flow"></a>Executando o fluxo

Para testar o fluxo, certifique-se de que quaisquer linhas de tabela com status em branco usem um endereço de e-mail vinculado a uma conta Teams (você provavelmente deve usar seu próprio endereço de e-mail durante o teste).

Você pode selecionar **Teste** no designer de fluxo ou executar o fluxo da página **Meus fluxos.** Depois de iniciar o fluxo e aceitar o uso das conexões necessárias, você deve receber um Cartão Adaptive de Power Automate até Teams. Uma vez preenchido o campo de status no cartão, o fluxo continuará e atualizará a planilha com o status que você fornece.

### <a name="before-running-the-flow"></a>Antes de executar o fluxo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Uma planilha com um relatório de status contendo uma entrada de status faltando":::

### <a name="receiving-the-adaptive-card"></a>Recebendo o Cartão Adaptativo

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Um Cartão Adaptativo em Teams pedindo ao funcionário uma atualização de status":::

### <a name="after-running-the-flow"></a>Depois de executar o fluxo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Uma planilha com um relatório de status com uma entrada de status agora preenchida":::
