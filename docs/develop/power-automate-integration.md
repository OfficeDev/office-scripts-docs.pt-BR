---
title: Executar scripts do Office com o Power Automate
description: Como obter scripts do Office para Excel na Web trabalhar com um fluxo de trabalho do Power Automate.
ms.date: 06/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 61e51861bd2c987c25d40e9ac6d2247122256918
ms.sourcegitcommit: c5ffe0a95b962936ee92e7ffe17388bef6d4fad8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241836"
---
# <a name="run-office-scripts-with-power-automate"></a>Executar scripts do Office com o Power Automate

[O Power Automate](https://flow.microsoft.com) permite adicionar Scripts do Office a um fluxo de trabalho maior e automatizado. Você pode usar o Power Automate para fazer coisas como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projetos com base nos comentários da pasta de trabalho.

## <a name="get-started"></a>Introdução

Se você for novo no Power Automate, recomendamos visitar a [Introdução ao Power Automate](/power-automate/getting-started). Lá, você pode saber mais sobre todas as possibilidades de automação disponíveis para você. Os documentos aqui se concentram em como os Scripts do Office funcionam com o Power Automate e como isso pode ajudar a melhorar sua experiência do Excel.

### <a name="step-by-step-tutorials"></a>Tutoriais passo a passo

Há três tutoriais passo a passo para o Power Automate e scripts do Office. Eles mostram como combinar os serviços automatizados e passar dados entre uma pasta de trabalho e um fluxo.

- [Ligue scripts de um fluxo manual do Power Automate](../tutorials/excel-power-automate-manual.md)
- [Passar dados para scripts numa execução automática do fluxo no Power Automate.](../tutorials/excel-power-automate-trigger.md)
- [Retornar os dados de um script para um fluxo do Power Automate executado automaticamente](../tutorials//excel-power-automate-returns.md)

### <a name="create-a-flow-from-excel"></a>Criar um fluxo do Excel

Você pode começar a usar o Power Automate no Excel com uma variedade de modelos de fluxo. Na guia **Automatizar** , selecione **Automatizar uma Tarefa**.

:::image type="content" source="../images/automate-a-task-button.png" alt-text="O botão &quot;Automatizar uma Tarefa&quot; na faixa de opções.":::

Isso abre um painel de tarefas com várias opções para começar a conectar seus Scripts do Office a soluções automatizadas maiores. Selecione qualquer opção para começar. Seu fluxo é fornecido com a pasta de trabalho atual.

:::image type="content" source="../images/automate-a-task-choices.png" alt-text="Um painel de tarefas mostrando opções de modelo de fluxo, como &quot;Agendar um Script do Office para ser executado no Excel e, em seguida, enviar um email&quot; e &quot;Executar um Script do Office no Excel quando uma resposta Microsoft Forms for recebida&quot;.":::

> [!TIP]
> Você também pode começar a criar um fluxo no menu **Mais opções (...)** em um script individual.

## <a name="excel-online-business-connector"></a>Conector do Excel Online (Business)

[Conectores](/connectors/connectors) são as pontes entre o Power Automate e os aplicativos. O [conector do Excel Online (Business)](/connectors/excelonlinebusiness) fornece aos seus fluxos acesso às pastas de trabalho do Excel. A ação "Executar script" permite chamar qualquer Script do Office acessível por meio da pasta de trabalho selecionada. Você também pode fornecer parâmetros de entrada de scripts para que os dados possam ser fornecidos pelo fluxo ou fazer com que o script retorne informações para etapas posteriores no fluxo.

> [!IMPORTANT]
> A ação "Executar script" fornece às pessoas que usam o conector do Excel acesso significativo à sua pasta de trabalho e seus dados. Além disso, há riscos de segurança com scripts que fazem chamadas de API externas, conforme explicado em [chamadas externas do Power Automate](external-calls.md). Se o administrador estiver preocupado com a exposição de dados altamente confidenciais, ele poderá desativar o conector do Excel Online ou restringir o acesso aos Scripts do Office por meio dos [controles de administrador de Scripts do Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

> [!IMPORTANT]
> O Power Automate **não dá** suporte a scripts armazenados no SharePoint no momento.

## <a name="data-transfer-in-flows-for-scripts"></a>Transferência de dados em fluxos para scripts

O Power Automate permite que você passe partes de dados entre as etapas do fluxo. Os scripts podem ser configurados para aceitar os tipos de informações de que você precisa e retornar qualquer coisa da pasta de trabalho desejada em seu fluxo. A entrada para o script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook`). A saída do script é declarada adicionando um tipo de retorno a `main`.

> [!NOTE]
> Quando você cria um bloco "Executar Script" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos. Se você alterar os parâmetros ou os tipos de retorno do script, precisará refazer o bloco "Executar script" do fluxo. Isso garante que os dados estão sendo analisados corretamente.

As seções a seguir abrangem os detalhes de entrada e saída para scripts usados no Power Automate. Se você quiser uma abordagem prática para aprender este tópico, experimente passar dados para scripts em um tutorial de fluxo do [Power Automate](../tutorials/excel-power-automate-trigger.md) de execução automática ou explore o cenário de exemplo de [lembretes de tarefas automatizadas](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parâmetros: passar dados para um script

Toda a entrada de script é especificada como parâmetros adicionais para a `main` função. Por exemplo, se você quisesse que um script aceita um `string` que representa um nome como entrada, você alteraria `main` a assinatura para `function main(workbook: ExcelScript.Workbook, name: string)`.

Ao configurar um fluxo no Power Automate, você pode especificar a entrada de script como valores estáticos, [expressões](/power-automate/use-expressions-in-conditions) ou conteúdo dinâmico. Detalhes sobre o conector de um serviço individual podem ser encontrados na documentação [do Conector do Power Automate](/connectors/).

#### <a name="type-restrictions"></a>Restrições de tipo

Ao adicionar parâmetros de entrada à função de um `main` script, considere as restrições e as concessões a seguir. Elas também se aplicam ao tipo de retorno do script.

1. O primeiro parâmetro deve ser do tipo `ExcelScript.Workbook`. O nome do parâmetro não importa.

1. Os tipos `string`, `number``boolean`, `unknown`, , `object`e `undefined` têm suporte.

1. Há suporte para matrizes `[]` (ambos `Array<T>` e estilos) dos tipos listados anteriormente. Também há suporte para matrizes aninhadas.

1. Os tipos de união serão permitidos se forem uma união de literais pertencentes a um único tipo (como `"Left" | "Right"`, não `"Left", 5`). Também há suporte para uniões de um tipo com suporte indefinido (como `string | undefined`).

1. Os tipos de objeto serão permitidos se contiverem propriedades do tipo `string`, `number`, `boolean`matrizes com suporte ou outros objetos com suporte. O exemplo a seguir mostra objetos aninhados com suporte como tipos de parâmetro.

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. Os objetos devem ter sua interface ou definição de classe definida no script. Um objeto também pode ser definido anonimamente embutido, como no exemplo a seguir.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Parâmetros opcionais e padrão

1. Parâmetros opcionais são permitidos e são indicados com o modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

1. Os valores de parâmetro padrão são permitidos (por exemplo `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.

### <a name="return-data-from-a-script"></a>Retornar dados de um script

Os scripts podem retornar dados da pasta de trabalho a serem usados como conteúdo dinâmico em um fluxo do Power Automate. As [mesmas restrições de tipo listadas anteriormente](#type-restrictions) se aplicam ao tipo de retorno. Para retornar um objeto, adicione a sintaxe de tipo de retorno à `main` função. Por exemplo, se você quisesse retornar um valor `string` do script, sua `main` assinatura seria `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Exemplo

A captura de tela a seguir mostra um fluxo do Power Automate que é disparado sempre que um problema do [GitHub](https://github.com/) é atribuído a você. O fluxo executa um script que adiciona o problema a uma tabela em uma pasta de trabalho do Excel. Se houver cinco ou mais problemas nessa tabela, o fluxo enviará um lembrete por email.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="O editor de fluxo do Power Automate mostrando o fluxo de exemplo.":::

A `main` função do script especifica a ID do problema e o título do problema como parâmetros de entrada, e o script retorna o número de linhas na tabela de problemas.

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

## <a name="see-also"></a>Confira também

- [Ligue scripts de um fluxo manual do Power Automate](../tutorials/excel-power-automate-manual.md)
- [Passar dados para scripts numa execução automática do fluxo no Power Automate.](../tutorials/excel-power-automate-trigger.md)
- [Retornar os dados de um script para um fluxo do Power Automate executado automaticamente](../tutorials/excel-power-automate-returns.md)
- [Informações de solução de problemas do Power Automate com scripts do Office](../testing/power-automate-troubleshooting.md)
- [Começar a usar o Power Automate](/power-automate/getting-started)
- [Documentação de referência do conector do Excel Online (Business)](/connectors/excelonlinebusiness/)
