---
title: Passar dados para scripts numa execução automática do fluxo no Power Automate.
description: Tutorial sobre executar os Scripts do Office para Excel na Web por meio do Power Automate quando emails são recebidos e transmitidos para o script.
ms.date: 06/29/2021
ms.localizationpriority: high
ms.openlocfilehash: 333ccfc753da067111ca4dc0c3e59ce9db360e0a
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326867"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow"></a>Passar dados para scripts numa execução automática do fluxo no Power Automate.

Este tutorial ensina como usar um script do Office para Excel na Web fluxo automatizado[ do Power Automate](https://flow.microsoft.com). Seu script irá automaticamente ser executado toda vez que você receber um email, gravando informações do email em uma pasta de trabalho do Excel. Ser capaz de passar os dados de outros aplicativos para um Script do Office oferece a você uma grande flexibilidade e liberdade nos seus processos automatizados.

> [!TIP]
> Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md). Se você for novo no Power Automate, recomendamos começar com o [tutorial Chamar scripts do manual de fluxo do Power Automate](excel-power-automate-manual.md). [Os Scripts do Office usam TypeScript](../overview/code-editor-environment.md) e este tutorial se destina a pessoas com conhecimento de nível iniciante a intermediário em JavaScript ou TypeScript. Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Preparar a pasta de trabalho

O Power Automate não pode usar[referências relativas](../testing/power-automate-troubleshooting.md#avoid-relative-references)como`Workbook.getActiveWorksheet`acessar componentes da pasta de trabalho. Portanto, precisamos de uma pasta de trabalho e planilha com nomes consistentes para que o Power Automate possa consultar.

1. Criar um nome para a pasta de trabalho **MyWorkbook**.

2. Vá até a guia **Automatizar** e selecione **Todos os Scripts**.

3. Selecione **Novo Script**.

4. Substitua o código existente pelo seguinte script e selecione **Executar**. Isso instalará a pasta de trabalho com nomes consistentes de planilhas, tabela e tabela dinâmica.

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

## <a name="create-an-office-script"></a>Criar um Script do Office

Vamos criar um script que registre as informações de um email. Gostaríamos de saber em quais dias da semana recebemos mais emails e quantos remetentes únicos estão enviando esses emails. Nossa pasta de trabalho tem uma tabela com **Data**, **Dia da semana**, **Endereços de email** e **Colunas de assunto**. Nossa planilha também tem uma tabela dinâmica que está sendo dinamizada no **Dia da semana** e **Endereços de email**(essas são as hierarquias de linha). A contagem de **assuntos exclusivos** são as informações agregadas que estão sendo exibidas (a hierarquia de dados). Faremos com que o nosso script atualize essa tabela dinâmica depois de atualizar a tabela de email.

1. No painel de tarefas do Editor de código, selecione **Novo script**.

2. O fluxo que criaremos depois no tutorial enviará a informação do nosso script sobre cada email recebido. O script precisa aceitar essa entrada pelos parâmetros na `main`função. Substitua o script padrão com o script seguinte:

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. O script precisa acessar a tabela e a tabela dinâmica da pasta de trabalho. Adicione o seguinte código ao corpo do script após a abertura`{`:

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. O `dateReceived`parâmetro é do tipo`string`. Vamos convertê-la em um[`Date`objeto](../develop/javascript-objects.md#date)para que possamos facilmente obter o dia da semana. Depois de fazer isso, será necessário mapear o valor numérico do dia para uma versão mais legível. Adicione o seguinte código no final do script (antes do encerramento `}`):

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. A cadeia`subject` pode incluir a marca de resposta "RE:". Vamos remover isso da cadeia de caracteres para que os emails no mesmo thread tenham o mesmo assunto para a tabela. Adicione o seguinte código no final do script (antes do encerramento `}`):

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Agora que os dados de email foram formatados da nossa preferência, vamos adicionar uma linha à tabela de email. Adicione o seguinte código no final do script (antes do encerramento `}`):

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. Por fim, vamos verificar se a tabela dinâmica está atualizada. Adicione o seguinte código no final do script (antes do encerramento `}`):

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Renomeie seu script **Gravar Email** e selecione **Salvar script**.

O seu script já está pronto para um fluxo de trabalho automatizado. Ele deverá ser semelhante ao script a seguir:

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

## <a name="create-an-automated-workflow-with-power-automate"></a>Criar um fluxo de trabalho automatizado com o Power Automate

1. Entre no [site do Power Automate](https://flow.microsoft.com).

2. No menu exibido no lado esquerdo da tela, selecione **Criar**. Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Botão Criar do Power Automate.":::

3. Na seção **Começar no espaço em branco**, selecione **Fluxo automático**. Isso cria um fluxo de trabalho iniciado por um evento, como o recebimento de emails.

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="A opção de Fluxo automatizado no Power Automate":::

4. Na caixa de diálogo exibida, insira o nome para seu fluxo na **caixa de texto** Nome de Fluxo. Em seguida, selecione **Quando um novo email chegar** da lista de opções em **escolha o gatilho de fluxo**. Talvez seja necessário procurar pela opção usando a caixa de pesquisa. Por fim, selecione **Criar**.

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="Parte do fluxo do Power Automate mostrando o &quot;nome do fluxo&quot; e as opções de &quot;escolher gatilho de fluxo&quot;. O nome do fluxo é &quot;Gravar Fluxo de Emails&quot; e o gatilho é a opção para &quot;Quando um novo email chegar no Outlook&quot;.":::

    > [!NOTE]
    > Este tutorial usa o Outlook. Sinta-se à vontade para usar seu serviço de e-mail preferido, embora algumas opções possam ser diferentes.

5. Selecione **Nova etapa**.

6. Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Opção do Excel Online (Business) no Power Automate.":::

7. Em **Ações**, selecione **Executar script**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Executar a opção de ação do script no Power Automate":::

8. Depois, você selecionará a pasta de trabalho, o script e os argumentos de entrada do script para usar na etapa do fluxo. Para o tutorial, você fará o uso da pasta de trabalho criada no seu OneDrive, mas é possível usar qualquer pasta de trabalho em um site OneDrive ou no Microsoft Office SharePoint Online. Especifique as seguintes configurações para o conector **Executar Script**:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: MyWorkbook.xlsx *(Escolhido por meio do navegador de arquivos)*
    - **Script**: Gravar Email
    - **De**: De *(conteúdo dinâmico do Outlook)*
    - **DateReceived**: Hora Recebida *(conteúdo dinâmico do Outlook)*
    - **assunto**: Assunto *(conteúdo dinâmico do Outlook)*

    *Observe que os parâmetros para o script só aparecerão quando o script for selecionado.*

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="A ação de script de execução do Power Automate mostrando as opções que aparecem depois que o script é selecionado.":::

9. Selecione **Salvar**.

Seu fluxo já está habilitado. Ele executará automaticamente seu script sempre que você receber um e-mail pelo Outlook.

## <a name="manage-the-script-in-power-automate"></a>Gerenciar o script no Power Automate

1. Na página principal do Power Automate, selecione **Meus fluxos**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Botão Meus fluxos no Power Automate":::

2. Selecione o seu fluxo. Aqui você pode ver o histórico de execução. Você pode atualizar a página ou selecionar o botão atualizar **Executar Todos** para atualizar o histórico. O fluxo será disparado logo após o recebimento de um email. Testar o fluxo enviando a si mesmo um email.

Quando o fluxo é acionado e executa seu script com sucesso, você deverá ver as atualizações da planilha na pasta de trabalho e da tabela dinâmica.

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="Uma planilha mostrando a tabela de email depois que o fluxo foi executado três vezes.":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="Uma planilha mostrando o PivotTable depois que o fluxo foi executado três vezes.":::

## <a name="next-steps"></a>Próximas etapas

Complete o tutorial [Retornar dados de um script para um fluxo de execução automática do Power Automate](excel-power-automate-returns.md). Ele ensina como retornar dados de um script para o fluxo.

Você também pode conferir o exemplo de [lembretes automáticos de tarefas](../resources/scenarios/task-reminders.md) para aprender a combinar os Scripts do Office e Power Automate com as placas adaptáveis de equipes.
