---
title: Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.
description: Um tutorial de scripts do Office sobre a leitura de dados de pastas de trabalho e avaliação desses dados no script.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700041"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.

Esse tutorial ensina a ler dados de uma pasta de trabalho com scripts do Office para o Excel na Web. Em seguida, edite os dados lidos e coloque-os de volta na pasta de trabalho.

> [!TIP]
> Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Preview note](../includes/preview-note.md)]

Antes de iniciar este tutorial, você precisará acessar os scripts do Office, que exigem o seguinte:

- [Excel na Web](https://www.office.com/launch/excel).
- Peça para o administrador [habilitar os scripts do Office da sua organização](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), o que adiciona a guia **Automação** à faixa de opções.

> [!IMPORTANT]
> Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript. Se você não conhece o JavaScript, recomendamos que revise o [tutorial do Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Visite [Scripts do Office no Excel na Web](../overview/excel.md) para saber mais sobre o ambiente de scripts.

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

3. Abra o **Editor de códigos** e escolha **Novo script**.
4. Vamos limpar a formatação. Este é um documento financeiro, iremos alterar a formatação dos números nas colunas **Débito** e **Crédito** para mostrar os valores em dólares. Também iremos ajustar a largura da coluna para os dados.

    Substitua o conteúdo do script pelo código a seguir:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. Agora, leremos um valor de uma das colunas de número. Adicione o seguinte código ao final do script:

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    Anote as chamadas para `load` e `sync`. Aprenda mais detalhes desses métodos em [Fundamentos de Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md#sync-and-load). Por enquanto, solicite que os dados sejam lidos e sincronize seu script com a pasta de trabalho para lê-los.

6. Execute o script.
7. Abra o console. Vá para o menu **Reticências** e pressione **Logs...**.
8. Você deverá ver `[Array[1]]` no console. Isso não é um número porque os intervalos são matrizes bidimensionais de dados. Esse intervalo bidimensional está sendo registrado diretamente no console. Felizmente, o Editor de códigos permite visualizar o conteúdo da matriz.
9. Quando uma matriz bidimensional é registrada no console, ela agrupa os valores de coluna em cada linha. Expanda o log de matriz pressionando o triângulo azul.
10. Expanda o segundo nível da matriz, pressionando o triângulo azul exibido recentemente. Agora, você deverá ver isto:

    ![O log do console mostrando a saída "-20.05", aninhada sob duas matrizes.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>Modificar o valor de uma célula

Agora que podemos ler os dados, usaremos eles para modificar a pasta de trabalho. Deixaremos o valor da célula **D2** positivo com a função `Math.abs`. O objeto [Matemática](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contém várias funções às quais seus scripts têm acesso. É possível encontrar mais informações sobre `Math` e outros objetos internos [Usando objetos JavaScript internos nos scripts do Office](../develop/javascript-objects.md).

1. Adicione o seguinte código ao final do script:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. O valor da célula **D2** agora deverá ser positivo.

## <a name="modify-the-values-of-a-column"></a>Modificar os valores de uma coluna

Agora que sabemos ler e escrever em uma única célula, vamos generalizar o script para trabalhar em todas as colunas de **Débito** e **Crédito**.

1. Remova o código que afeta apenas uma única célula (o código de valor absoluto anterior), de modo que o script agora se pareça com este:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. Adicione um loop que percorra as linhas nas duas últimas colunas. Para cada célula, o script define o valor para o valor absoluto do valor atual.

    Observe que a matriz que define a localização das células é baseada em zero. Isso significa que a célula **A1** é `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    Essa parte do script faz várias tarefas importantes. Primeiro, ela carrega os valores e a contagem de linhas do intervalo usado. Isso nos permite ver os valores e saber quando parar. Segundo, ela reitera através do intervalo usado, verificando cada célula nas colunas **Débito** ou **Crédito**. Por fim, se o valor na célula não for 0, ele será substituído pelo valor absoluto. Estamos evitando zeros, para que possamos deixar as células em branco.

3. Execute o script.

    Seu extrato bancário agora deverá ter a seguinte aparência:

    ![O extrato bancário como uma tabela formatada apenas com valores positivos.](../images/tutorial-5.png)

## <a name="next-steps"></a>Próximas etapas

Abra o Editor de códigos e experimente alguns dos [Scripts de exemplo para scripts do Office no Excel na Web](../resources/excel-samples.md). Visite também [Fundamentos de Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md) para saber mais sobre como criar scripts do Office.