---
title: Executar scripts do Office com o Power Automate
description: Como obter scripts do Office para Excel na Web trabalhando com um fluxo de trabalho do Power Automate.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755067"
---
# <a name="run-office-scripts-with-power-automate"></a>Executar scripts do Office com o Power Automate

[O Power Automate](https://flow.microsoft.com) permite adicionar Scripts do Office a um fluxo de trabalho maior e automatizado. Você pode usar o Power Automate para fazer coisas como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projeto com base nos comentários da pasta de trabalho.

## <a name="getting-started"></a>Introdução

Se você for novo no Power Automate, recomendamos visitar [Começar a usar o Power Automate.](/power-automate/getting-started) Lá, você pode saber mais sobre todas as possibilidades de automação disponíveis para você. Os documentos aqui se concentram em como os Scripts do Office funcionam com o Power Automate e como isso pode ajudar a melhorar sua experiência do Excel.

Para começar a combinar Power Automate e Scripts do Office, siga o tutorial [Iniciar usando scripts com o Power Automate](../tutorials/excel-power-automate-manual.md). Isso ensinará como criar um fluxo que chama um script simples. Depois de concluir esse tutorial e o Passar dados para scripts em um tutorial de fluxo do Power Automate executado automaticamente, retorne aqui para obter informações detalhadas sobre como conectar scripts do Office aos fluxos do Power [Automate.](../tutorials/excel-power-automate-trigger.md)

## <a name="excel-online-business-connector"></a>Conector do Excel Online (Business)

[Conectores](/connectors/connectors) são as pontes entre o Power Automate e os aplicativos. O [conector do Excel Online (Business)](/connectors/excelonlinebusiness) fornece aos fluxos acesso às planilhas do Excel. A ação "Executar script" permite chamar qualquer Script do Office acessível por meio da workbook selecionada. Você também pode dar aos seus scripts parâmetros de entrada para que os dados possam ser fornecidos pelo fluxo ou fazer com que seu script retorne informações para etapas posteriores no fluxo.

> [!IMPORTANT]
> A ação "Executar script" oferece às pessoas que usam o conector do Excel acesso significativo à sua planilha e seus dados. Além disso, há riscos de segurança com scripts que fazem chamadas de API externas, conforme explicado em [Chamadas externas do Power Automate](external-calls.md). Se o administrador estiver preocupado com a exposição de dados altamente confidenciais, ele poderá desativar o conector do Excel Online ou restringir o acesso aos Scripts do Office por meio dos controles de administrador [de Scripts do Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Transferência de dados em fluxos para scripts

O Power Automate permite que você passe partes de dados entre etapas do seu fluxo. Os scripts podem ser configurados para aceitar qualquer tipo de informação que você precisa e retornar qualquer coisa da sua workbook que você deseja em seu fluxo. A entrada para o script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook` ). A saída do script é declarada adicionando um tipo de retorno a `main` .

> [!NOTE]
> Quando você cria um bloco "Executar Script" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos. Se você alterar os parâmetros ou retornar tipos de script, precisará refazer o bloco "Executar script" do seu fluxo. Isso garante que os dados estão sendo analisados corretamente.

As seções a seguir abrangem os detalhes de entrada e saída para scripts usados no Power Automate. Se você quiser uma abordagem prática para aprender este tópico, experimente o passar dados para scripts em um tutorial de fluxo do [Power Automate](../tutorials/excel-power-automate-trigger.md) executado automaticamente ou explore o cenário de exemplo lembretes de tarefas [automatizados.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Parâmetros: passar dados para um script

Todas as entradas de script são especificadas como parâmetros adicionais para a `main` função. Por exemplo, se você quisesse que um script aceitasse um nome que representasse um nome como entrada, você `string` alteraria a `main` assinatura para `function main(workbook: ExcelScript.Workbook, name: string)` .

Ao configurar um fluxo no Power Automate, você pode especificar a entrada de script como valores [estáticos, expressões](/power-automate/use-expressions-in-conditions)ou conteúdo dinâmico. Os detalhes sobre o conector de um serviço individual podem ser encontrados na documentação [do Power Automate Connector.](/connectors/)

Ao adicionar parâmetros de entrada à função de um `main` script, considere as seguintes restrições e concessões.

1. O primeiro parâmetro deve ser do tipo `ExcelScript.Workbook` . Seu nome de parâmetro não importa.

2. Cada parâmetro deve ter um tipo (como `string` ou `number` ).

3. Os tipos `string` `number` básicos , `boolean` , , , , e são `any` `unknown` `object` `undefined` suportados.

4. Há suporte para matrizes dos tipos básicos listados anteriormente.

5. As matrizes aninhadas são suportadas como parâmetros (mas não como tipos de retorno).

6. Os tipos de união são permitidos se eles são uma união de literais pertencentes a um único tipo (como `"Left" | "Right"` ). Também há suporte para uniões de um tipo com suporte indefinido (como `string | undefined` ).

7. Os tipos de objeto são permitidos se eles contêm propriedades do tipo , , matrizes com `string` suporte ou outros objetos com `number` `boolean` suporte. O exemplo a seguir mostra objetos aninhados com suporte como tipos de parâmetro:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Os objetos devem ter sua interface ou definição de classe definida no script. Um objeto também pode ser definido anonimamente em linha, como no exemplo a seguir:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Parâmetros opcionais são permitidos e podem ser denodos como tal usando o modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Os valores de parâmetro padrão são permitidos (por `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` exemplo.

### <a name="returning-data-from-a-script"></a>Retornando dados de um script

Os scripts podem retornar dados da caixa de trabalho a serem usados como conteúdo dinâmico em um fluxo do Power Automate. Assim como nos parâmetros de entrada, o Power Automate coloca algumas restrições no tipo de retorno.

1. Os tipos `string` básicos `number` , , e são `boolean` `void` `undefined` suportados.

2. Os tipos de união usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.

3. Os tipos de matriz são permitidos se eles são do tipo `string` `number` , ou `boolean` . Eles também são permitidos se o tipo for uma união com suporte ou um tipo literal com suporte.

4. Os tipos de objeto usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.

5. A digitação implícita é suportada, embora ela deve seguir as mesmas regras de um tipo definido.

## <a name="example"></a>Exemplo

A captura de tela a seguir mostra um fluxo do Power Automate que é acionado sempre que um problema [do GitHub](https://github.com/) é atribuído a você. O fluxo executa um script que adiciona o problema a uma tabela em uma planilha do Excel. Se houver cinco ou mais problemas nessa tabela, o fluxo enviará um lembrete de email.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="O editor de fluxo do Power Automate mostrando o fluxo de exemplo.":::

A função do script especifica a ID do problema e o título do problema como parâmetros de entrada, e o script retorna o número de linhas `main` na tabela de problemas.

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

- [Executar scripts do Office no Excel na Web com o Power Automate](../tutorials/excel-power-automate-manual.md)
- [Passar dados para scripts numa execução automática do fluxo no Power Automate.](../tutorials/excel-power-automate-trigger.md)
- [Retorna dados de um script para um fluxo do Power Automate executado automaticamente](../tutorials/excel-power-automate-returns.md)
- [Solução de problemas de informações para o Power Automate com scripts do Office](../testing/power-automate-troubleshooting.md)
- [Começar a usar o Power Automate](/power-automate/getting-started)
- [Documentação de referência do conector do Excel Online (Business)](/connectors/excelonlinebusiness/)
