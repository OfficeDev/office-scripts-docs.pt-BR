---
title: Execute Office scripts com Power Automate
description: Como obter Office Scripts para Excel na Web trabalhando com um fluxo de trabalho Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545037"
---
# <a name="run-office-scripts-with-power-automate"></a>Execute Office scripts com Power Automate

[Power Automate](https://flow.microsoft.com) permite adicionar scripts Office a um fluxo de trabalho maior e automatizado. Você pode usar Power Automate fazer coisas como adicionar o conteúdo de um e-mail à mesa de uma planilha ou criar ações em suas ferramentas de gerenciamento de projetos com base em comentários de pasta de trabalho.

## <a name="get-started"></a>Introdução

Se você é novo em Power Automate, recomendamos visitar [Get started with Power Automate](/power-automate/getting-started). Lá, você pode aprender mais sobre todas as possibilidades de automação disponíveis para você. Os documentos aqui se concentram em como Office Scripts trabalham com Power Automate e como isso pode ajudar a melhorar sua experiência Excel.

Para começar a combinar Power Automate e Office Scripts, siga o tutorial Comece a [usar scripts com Power Automate](../tutorials/excel-power-automate-manual.md). Isso vai te ensinar como criar um fluxo que chama de script simples. Depois de completar esse tutorial e os dados do Pass para scripts em um tutorial [de fluxo de fluxo Power Automate executado automaticamente,](../tutorials/excel-power-automate-trigger.md) retorne aqui para obter informações detalhadas sobre a conexão Office Scripts para fluxos Power Automate.

## <a name="excel-online-business-connector"></a>Excel Conector on-line (business)

[Conectores](/connectors/connectors) são as pontes entre Power Automate e aplicações. O [conector Excel Online (Business)](/connectors/excelonlinebusiness) dá aos seus fluxos acesso a Excel livros de trabalho. A ação "Executar script" permite que você chame qualquer Office Script acessível através da pasta de trabalho selecionada. Você também pode fornecer parâmetros de entrada de seus scripts para que os dados possam ser fornecidos pelo fluxo ou ter suas informações de retorno do script para etapas posteriores no fluxo.

> [!IMPORTANT]
> A ação "Executar script" dá às pessoas que usam o conector Excel acesso significativo à sua pasta de trabalho e seus dados. Além disso, existem riscos de segurança com scripts que fazem chamadas de API externas, como explicado em [chamadas externas de Power Automate](external-calls.md). Se o administrador estiver preocupado com a exposição de dados altamente [confidenciais,](/microsoft-365/admin/manage/manage-office-scripts-settings)eles podem desligar o conector Excel Online ou restringir o acesso a scripts Office através dos controles de administrador de scripts Office .

## <a name="data-transfer-in-flows-for-scripts"></a>Transferência de dados em fluxos para scripts

Power Automate permite que você passe pedaços de dados entre etapas do seu fluxo. Os scripts podem ser configurados para aceitar qualquer tipo de informação que você precise e retornar qualquer coisa da sua pasta de trabalho que você deseja em seu fluxo. A entrada para o seu script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook` ). A saída do script é declarada adicionando um tipo de retorno a `main` .

> [!NOTE]
> Quando você cria um bloco "Executar script" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos. Se você alterar os parâmetros ou retornar os tipos do seu script, você precisará refazer o bloco "Executar script" do seu fluxo. Isso garante que os dados estão sendo analisados corretamente.

As seções a seguir cobrem os detalhes de entrada e saída para scripts usados em Power Automate. Se você quiser uma abordagem prática para aprender este tópico, experimente os dados do [Pass para scripts em um](../tutorials/excel-power-automate-trigger.md) tutorial de fluxo de Power Automate executado automaticamente ou explore o cenário de amostra [de lembretes de tarefas automatizados.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parâmetros: Passar dados para um script

Toda a entrada do script é especificada como parâmetros adicionais para a `main` função. Por exemplo, se você quisesse um script para aceitar um `string` que representa um nome como entrada, você mudaria a assinatura para `main` `function main(workbook: ExcelScript.Workbook, name: string)` .

Quando você está configurando um fluxo em Power Automate, você pode especificar a entrada do script como valores [estáticos, expressões](/power-automate/use-expressions-in-conditions)ou conteúdo dinâmico. Detalhes sobre o conector de um serviço individual podem ser encontrados na [documentação do conector Power Automate](/connectors/).

Ao adicionar parâmetros de entrada à função de um `main` script, considere as seguintes franquias e restrições.

1. O primeiro parâmetro deve ser de tipo `ExcelScript.Workbook` . Seu nome de parâmetro não importa.

2. Cada parâmetro deve ter um tipo (como `string` ou `number` ).

3. Os tipos `string` `number` básicos, `boolean` , , , e são `unknown` `object` `undefined` suportados.

4. Arrays dos tipos básicos listados anteriormente são suportados.

5. As matrizes aninhadas são suportadas como parâmetros (mas não como tipos de retorno).

6. Os tipos sindicais são permitidos se forem uma união de literais pertencentes a um único tipo `"Left" | "Right"` (como). Sindicatos de um tipo apoiado com indefinidos também são apoiados (como `string | undefined` ).

7. Os tipos de objetos são permitidos se contiverem propriedades do `string` `number` `boolean` tipo, matrizes suportadas ou outros objetos suportados. O exemplo a seguir mostra objetos aninhados que são suportados como tipos de parâmetro:

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

8. Os objetos devem ter sua interface ou definição de classe definida no script. Um objeto também pode ser definido anonimamente inline, como no exemplo a seguir:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Parâmetros opcionais são permitidos e podem ser denotados como tal usando o modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Valores de parâmetro padrão são permitidos (por exemplo `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

### <a name="return-data-from-a-script"></a>Retornar dados de um script

Os scripts podem retornar dados da pasta de trabalho para serem usados como conteúdo dinâmico em um fluxo Power Automate. Assim como os parâmetros de entrada, Power Automate coloca algumas restrições no tipo de retorno.

1. Os tipos `string` `number` básicos, `boolean` , , , , e são `void` `undefined` suportados.

2. Os tipos de união usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.

3. Tipos de matriz são permitidos se forem do tipo `string` `number` , ou `boolean` . Eles também são permitidos se o tipo for um tipo de união apoiada ou apoiado literal.

4. Os tipos de objetos usados como tipos de retorno seguem as mesmas restrições que fazem quando usados como parâmetros de script.

5. A digitação implícita é suportada, embora deva seguir as mesmas regras de um tipo definido.

## <a name="example"></a>Exemplo

A captura de tela a seguir mostra um fluxo de Power Automate que é acionado sempre que um problema [de GitHub](https://github.com/) é atribuído a você. O fluxo executa um script que adiciona o problema a uma tabela em uma Excel livro de trabalho. Se houver cinco ou mais problemas nessa tabela, o fluxo envia um lembrete de e-mail.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="O editor de fluxo Power Automate mostrando o fluxo de exemplo":::

A `main` função do script especifica o ID de edição e o título de emissão como parâmetros de entrada, e o script retorna o número de linhas na tabela de emissão.

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

- [Execute Office scripts em Excel na Web com Power Automate](../tutorials/excel-power-automate-manual.md)
- [Passar dados para scripts numa execução automática do fluxo no Power Automate.](../tutorials/excel-power-automate-trigger.md)
- [Retorna dados de um script para um fluxo do Power Automate executado automaticamente](../tutorials/excel-power-automate-returns.md)
- [Solução de problemas para Power Automate com scripts Office](../testing/power-automate-troubleshooting.md)
- [Começar a usar o Power Automate](/power-automate/getting-started)
- [Excel Documentação de referência do conector online (business)](/connectors/excelonlinebusiness/)
