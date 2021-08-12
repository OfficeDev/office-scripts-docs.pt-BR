---
title: Grave, edite e crie scripts do Office no Excel na Web
description: Um tutorial sobre o básico dos scripts do Office, incluindo a gravação de scripts com o Gravador de ações e a gravação de dados em uma pasta de trabalho.
ms.date: 05/23/2021
localization_priority: Priority
ms.openlocfilehash: b29d9a5e95f510f63c2c71fc10bb68bc7b5430077a0be09327fc07675bb41f94
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847307"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a>Grave, edite e crie scripts do Office no Excel na Web

Este tutorial ensina os fundamentos da gravação, edição e escrita de um Script do para o Excel na web. Você gravará um script que aplicará uma determinada formatação a uma planilha de registro de vendas. Depois, você editará o script gravado para aplicar outras formatações, criar e classificar uma tabela. Este padrão de registro e edição é uma importante ferramenta para ver como suas ações no Excel são parecidas com um código.

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript. Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Visite o [ambiente do Editor de Código do Scripts do Office](../overview/code-editor-environment.md) para saber mais sobre o ambiente de script.

## <a name="add-data-and-record-a-basic-script"></a>Adicione dados e grave um script básico

Primeiro, precisaremos de alguns dados e um pequeno script inicial.

1. Crie uma nova pasta de trabalho no Excel para a Web.
2. Copie os seguintes dados de vendas de frutas e cole-os na planilha, começando na célula **A1**.

    |Fruta |2018 |2019 |
    |:---|:---|:---|
    |Laranjas |1.000 |1.200 |
    |Limões |800 |900 |
    |Limões-galego |600 |500 |
    |Toranjas |900 |700 |

3. Abra a guia **Automação**. Se você não vir a guia **Automação**, verifique o excedente da faixa de opções selecionando a seta suspensa. Se ainda não estiver lá, siga o conselho do artigo [Solução de Problemas de Scripts do Office ](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).
4. Selecione o botão **Gravar Ações**.
5. Selecione as células **A2:C2** (a linha "Laranjas") e defina a cor de preenchimento como laranja.
6. Interrompa a gravação selecionando o botão **Parar**.

    Sua planilha deve ficar assim (não se preocupe se a cor for diferente):

    :::image type="content" source="../images/tutorial-1.png" alt-text="Uma planilha mostrando a linha de dados das vendas de frutas com a linha contendo &quot;Laranjas&quot; realçada na cor laranja.":::

## <a name="edit-an-existing-script"></a>Edite um script existente

O script anterior coloriu a linha "Laranjas" para ficar laranja. Vamos adicionar uma linha amarela aos "Limões".

1. No painel, agora aberto, **Detalhes**, selecione o botão **Editar**.
2. Você deve ver algo semelhante a este código:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    Este código recebe a planilha atual da pasta de trabalho. Depois, defina a cor de preenchimento do intervalo **A2:C2**.

    Os intervalos são parte fundamental dos scripts do Office no Excel na Web. Um intervalo é um bloco retangular e contíguo de células que contém valores, fórmula e formatação. Eles são a estrutura básica das células através da qual você executará a maioria das tarefas de script.

3. Adicione a seguinte linha no final do script (entre onde `color` está definido e o encerramento `}`):

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. Teste o script selecionando **Executar**. Sua pasta de trabalho já deve ter esta aparência:

    :::image type="content" source="../images/tutorial-2.png" alt-text="Uma planilha mostrando a linha de dados das vendas de frutas com a linha &quot;Laranjas&quot; realçada na cor laranja, e a linha &quot;Limões&quot; realçada na cor amarela.":::

## <a name="create-a-table"></a>Crie uma tabela

Vamos converter esses dados de vendas de frutas em uma tabela. Usaremos nosso script em todo o processo.

1. Adicione a seguinte linha no final do script (antes do encerramento `}`):

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. Essa chamada retorna um `Table` objeto. Vamos usar essa tabela para classificar os dados. Classificaremos os dados em ordem crescente com base nos valores na coluna "Frutas". Adicione a seguinte linha assim que criar a tabela:

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    Seu script deve ter esta aparência:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    As tabelas possuem um objeto`TableSort`, acessado por meio do método `Table.getSort`. Você pode aplicar critérios de classificação a esse objeto. O `apply` método utiliza uma matriz de `SortField` objetos. Nesse caso, só temos um critério de classificação, por isso só usamos um. `SortField`. `key: 0` define a coluna com os valores que determinam a classificação como "0" (que nesse caso, é a primeira coluna na tabela **A** ). `ascending: true` classifica os dados em ordem crescente (em vez de ordem decrescente).

3. Execute o script. Você deverá ver uma tabela como essa:

    :::image type="content" source="../images/tutorial-3.png" alt-text="Uma planilha mostrando a tabela ordenada de vendas de frutas.":::

    > [!NOTE]
    > Se você executar novamente o script, receberá um erro. Isso ocorre porque você não pode criar uma tabela em cima de outra tabela. No entanto, você pode executar o script em uma planilha ou pasta de trabalho diferente.

### <a name="re-run-the-script"></a>Reexecute o script

1. Crie uma nova planilha na pasta de trabalho atual.
2. Copie os dados das frutas do início do tutorial e cole-os na nova planilha, começando na célula **A1**.
3. Execute o script.

## <a name="next-steps"></a>Próximas etapas

Conclua o tutorial [Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.](excel-read-tutorial.md). Ele ensina como ler dados de uma pasta de trabalho com um script do Office.
