---
title: Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.
description: Um tutorial de scripts do Office sobre a leitura de dados de pastas de trabalho e avaliação desses dados no script.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: 0848a24e7333842b5b3b1f82ec8f270514c34d2f
ms.sourcegitcommit: 9df67e007ddbfec79a7360df9f4ea5ac6c86fb08
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49772965"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.

Esse tutorial ensina a ler dados de uma pasta de trabalho com scripts do Office para o Excel na Web. Você estará escrevendo um novo script que formatará um extrato bancário e normalizará os dados desse extrato. Como parte desta limpeza de dados, seu script lerá os valores das células de transação, aplicará uma fórmula simples a cada valor e gravará a resposta resultante na pasta de trabalho. A leitura os dados da pasta de trabalho permite a automatização de alguns dos seus processos de tomada de decisão no script.

> [!TIP]
> Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md). [Os Scripts do Office usam TypeScript](../overview/code-editor-environment.md) e este tutorial se destina a pessoas com conhecimento de nível iniciante a intermediário em JavaScript ou TypeScript. Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a>Ler uma célula

Os scripts feitos com o Gravador de Ação só podem gravar informações na pasta de trabalho. Com o Editor de Códigos, é possível editar e criar scripts que também leem dados de uma pasta de trabalho.

Vamos criar um script que leia dados e atue com base no que foi lido. Vamos usar um exemplo de um extrato bancário. Essa instrução é um relatório combinado de verificação de crédito. Infelizmente, eles relatam alterações no balanço de forma diferente. A declaração de verificação exibe o rendimento como crédito positivo e custos como débito negativo. O demonstrativo de crédito faz o oposto.

No resto do tutorial, normalizaremos os dados usando um script. Primeiro, vamos aprender a ler os dados da pasta de trabalho.

1. Crie uma nova planilha na pasta de trabalho usada para o resto do tutorial.
2. Copie os seguintes dados e cole-os na nova planilha, começando na célula **A1**.

    |Data |Conta |Descrição |Débito |Crédito |
    |:--|:--|:--|:--|:--|
    |10/10/2019 |Verificando |Vinícola Coho |-20.05 | |
    |11/10/2019 |Crédito |A Companhia Telefônica |99.95 | |
    |13/10/2019 |Crédito |Vinícola Coho |154.43 | |
    |15/10/2019 |Verificando |Depósito externo | |1000 |
    |20/10/2019 |Crédito |Vinícola Coho – Reembolso | |-35.45 |
    |25/10/2019 |Verificando |Ideal para sua empresa de produtos orgânicos | -85.64 | |
    |01/11/2019 |Verificando |Depósito externo | |1000 |

3. Abra **Todos os Scripts** e selecione **Novo Script**.
4. Vamos limpar a formatação. Este é um documento financeiro, iremos alterar a formatação dos números nas colunas **Débito** e **Crédito** para mostrar os valores em dólares. Também iremos ajustar a largura da coluna para os dados.

    Substitua o conteúdo do script pelo código a seguir:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. Agora, leremos um valor de uma das colunas de número. Adicione o seguinte código no final do script (antes do encerramento `}`):

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. Execute o script.
7. Você deverá ver `[Array[1]]` no console. Isso não é um número porque os intervalos são matrizes bidimensionais de dados. Esse intervalo bidimensional está sendo registrado diretamente no console. Felizmente, o Editor de Códigos permite visualizar o conteúdo da matriz.
8. Quando uma matriz bidimensional é registrada no console, ela agrupa os valores de coluna em cada linha. Expanda o log de matriz pressionando o triângulo azul.
9. Expanda o segundo nível da matriz, pressionando o triângulo azul exibido recentemente. Agora, você deverá ver isto:

    ![O log do console exibindo a saída "20.05", aninhada em duas matrizes](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>Modificar o valor de uma célula

Agora que podemos ler os dados, usaremos eles para modificar a pasta de trabalho. Deixaremos o valor da célula **D2** positivo com a função `Math.abs`. O objeto [Matemática](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contém várias funções às quais seus scripts têm acesso. É possível encontrar mais informações sobre `Math` e outros objetos internos [Usando objetos JavaScript internos nos scripts do Office](../develop/javascript-objects.md).

1. Usaremos os métodos `getValue` e `setValue` para alterar o valor da célula. Esses métodos funcionam em uma única célula. Ao lidar com intervalos de várias células, use `getValues` e `setValues`. Adicione o seguinte código ao final do script:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > Estamos [lançando](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) o valor retornado de `range.getValue()` para um `number` usando a palavra-chave `as`. Isso é necessário porque um intervalo pode ser cadeias de caracteres, números ou booleanas. Nesta instância, precisamos explicitamente de um número.

2. O valor da célula **D2** agora deverá ser positivo.

## <a name="modify-the-values-of-a-column"></a>Modificar os valores de uma coluna

Agora que sabemos ler e escrever em uma única célula, vamos generalizar o script para trabalhar em todas as colunas de **Débito** e **Crédito**.

1. Remova o código que afeta apenas uma única célula (o código de valor absoluto anterior), de modo que o script agora se pareça com este:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. Adicione um loop que percorra as linhas nas duas últimas colunas. Para cada célula, o script define o valor para o valor absoluto do valor atual.

    Observe que a matriz que define a localização das células é baseada em zero. Isso significa que a célula **A1** é `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    Essa parte do script faz várias tarefas importantes. Primeiro, ela obtém os valores e a contagem de linhas do intervalo usado. Isso nos permite ver os valores e saber quando parar. Segundo, ela reitera através do intervalo usado, verificando cada célula nas colunas **Débito** ou **Crédito**. Por fim, se o valor na célula não for 0, ele será substituído pelo valor absoluto. Estamos evitando zeros, para que possamos deixar as células em branco.

3. Execute o script.

    Seu extrato bancário agora deverá ter a seguinte aparência:

    ![O extrato bancário como uma tabela formatada apenas com valores positivos.](../images/tutorial-5.png)

## <a name="next-steps"></a>Próximas etapas

Abra o Editor de códigos e experimente alguns dos [Scripts de exemplo para scripts do Office no Excel na Web](../resources/excel-samples.md). Visite também [Fundamentos de Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md) para saber mais sobre como criar scripts do Office.

A próxima série de tutoriais de Scripts do Office tem foco na utilização de Scripts do Office com o Power Automate. Saiba mais sobre as vantagens da combinação das duas plataformas em [Executar Scripts do Office com o Power Automate](../develop/power-automate-integration.md) ou tente o tutorial [Chamar Scripts no manual de fluxo do Power Automate](excel-power-automate-manual.md) para criar um fluxo no Power Automate que utiliza um Script do Office.
