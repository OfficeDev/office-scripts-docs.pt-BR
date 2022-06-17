---
title: Executar Office scripts com Power Automate
description: Como obter scripts Office para Excel na Web trabalho com um fluxo Power Automate trabalho.
ms.date: 05/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85c335eeb736ec544eccb2fbdbe819bdbef6848c
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128227"
---
# <a name="run-office-scripts-with-power-automate"></a>Executar Office scripts com Power Automate

[Power Automate](https://flow.microsoft.com) permite adicionar scripts Office a um fluxo de trabalho maior e automatizado. Você pode usar Power Automate fazer coisas como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projetos com base nos comentários da pasta de trabalho.

## <a name="get-started"></a>Introdução

Se você for novo no Power Automate, recomendamos visitar Introdução [com Power Automate](/power-automate/getting-started). Lá, você pode saber mais sobre todas as possibilidades de automação disponíveis para você. Os documentos aqui se concentram em como Office scripts funcionam com Power Automate e como isso pode ajudar a melhorar sua Excel experiência.

Para começar a combinar scripts Power Automate e Office, siga o tutorial Começar a usar [scripts com Power Automate](../tutorials/excel-power-automate-manual.md). Isso ensinará você a criar um fluxo que chama um script simples. Depois de concluir esse tutorial e passar dados para [scripts](../tutorials/excel-power-automate-trigger.md) em um tutorial de fluxo de Power Automate executado automaticamente, retorne aqui para obter informações detalhadas sobre como conectar scripts do Office a fluxos Power Automate.

## <a name="excel-online-business-connector"></a>Excel online (Business)

[Conectores](/connectors/connectors) são as pontes entre Power Automate e aplicativos. O [Excel online (Business)](/connectors/excelonlinebusiness) fornece aos seus fluxos acesso Excel pastas de trabalho. A ação "Executar script" permite chamar qualquer Office script acessível por meio da pasta de trabalho selecionada. Você também pode fornecer parâmetros de entrada de scripts para que os dados possam ser fornecidos pelo fluxo ou fazer com que o script retorne informações para etapas posteriores no fluxo.

> [!IMPORTANT]
> A ação "Executar script" fornece às pessoas que usam o conector Excel acesso significativo à sua pasta de trabalho e seus dados. Além disso, há riscos de segurança com scripts que fazem chamadas de API externas, conforme explicado em chamadas externas [de Power Automate](external-calls.md). Se o administrador estiver preocupado com a exposição de dados altamente confidenciais, ele poderá desativar o conector do Excel Online ou restringir o acesso aos Scripts do Office por meio dos controles de administrador do [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

> [!IMPORTANT]
> Power Automate **não dá** suporte a scripts armazenados SharePoint no momento.

## <a name="data-transfer-in-flows-for-scripts"></a>Transferência de dados em fluxos para scripts

Power Automate permite que você passe partes de dados entre as etapas do fluxo. Os scripts podem ser configurados para aceitar os tipos de informações de que você precisa e retornar qualquer coisa da pasta de trabalho desejada em seu fluxo. A entrada para o script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook`). A saída do script é declarada adicionando um tipo de retorno a `main`.

> [!NOTE]
> Quando você cria um bloco "Executar Script" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos. Se você alterar os parâmetros ou os tipos de retorno do script, precisará refazer o bloco "Executar script" do fluxo. Isso garante que os dados estão sendo analisados corretamente.

As seções a seguir abrangem os detalhes de entrada e saída para scripts usados Power Automate. Se você quiser uma abordagem prática para aprender este tópico, experimente passar dados para [scripts](../tutorials/excel-power-automate-trigger.md) em um tutorial de fluxo de Power Automate executado automaticamente ou explore o cenário de exemplo de [lembretes de tarefas automatizados](../resources/scenarios/task-reminders.md).

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parâmetros: passar dados para um script

Toda a entrada de script é especificada como parâmetros adicionais para a `main` função. Por exemplo, se você quisesse que um script aceita um `string` que representa um nome como entrada, você alteraria `main` a assinatura para `function main(workbook: ExcelScript.Workbook, name: string)`.

Ao configurar um fluxo no Power Automate, você pode especificar a entrada de script como valores [estáticos, expressões](/power-automate/use-expressions-in-conditions) ou conteúdo dinâmico. Detalhes sobre o conector de um serviço individual podem ser encontrados na documentação [do Power Automate Connector](/connectors/).

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

Os scripts podem retornar dados da pasta de trabalho a serem usados como conteúdo dinâmico em um Power Automate fluxo. As [mesmas restrições de tipo listadas anteriormente](#type-restrictions) se aplicam ao tipo de retorno. Para retornar um objeto, adicione a sintaxe de tipo de retorno à `main` função. Por exemplo, se você quisesse retornar um valor `string` do script, sua `main` assinatura seria `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Exemplo

A captura de tela a seguir mostra Power Automate um fluxo que é disparado [sempre](https://github.com/) que um GitHub problema é atribuído a você. O fluxo executa um script que adiciona o problema a uma tabela em uma Excel de trabalho. Se houver cinco ou mais problemas nessa tabela, o fluxo enviará um lembrete por email.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="O Power Automate de fluxo mostrando o fluxo de exemplo.":::

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
- [Informações de solução de problemas para Power Automate com Office Scripts](../testing/power-automate-troubleshooting.md)
- [Começar a usar o Power Automate](/power-automate/getting-started)
- [documentação de referência do conector do Excel Online (Business)](/connectors/excelonlinebusiness/)
