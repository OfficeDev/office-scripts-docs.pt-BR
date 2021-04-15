---
title: Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.
description: Um tutorial de scripts do Office sobre a leitura de dados de pastas de trabalho e avaliação desses dados no script.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: d6321cb91a425da3fd45329d5171f1d5694b2b99
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754850"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="dd472-103">Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="dd472-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="dd472-104">Esse tutorial ensina a ler dados de uma pasta de trabalho com scripts do Office para o Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="dd472-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="dd472-105">Você estará escrevendo um novo script que formatará um extrato bancário e normalizará os dados desse extrato.</span><span class="sxs-lookup"><span data-stu-id="dd472-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="dd472-106">Como parte desta limpeza de dados, seu script lerá os valores das células de transação, aplicará uma fórmula simples a cada valor e gravará a resposta resultante na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="dd472-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="dd472-107">A leitura os dados da pasta de trabalho permite a automatização de alguns dos seus processos de tomada de decisão no script.</span><span class="sxs-lookup"><span data-stu-id="dd472-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="dd472-108">Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="dd472-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="dd472-109">[Os Scripts do Office usam TypeScript](../overview/code-editor-environment.md) e este tutorial se destina a pessoas com conhecimento de nível iniciante a intermediário em JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="dd472-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="dd472-110">Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="dd472-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dd472-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="dd472-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="dd472-112">Ler uma célula</span><span class="sxs-lookup"><span data-stu-id="dd472-112">Read a cell</span></span>

<span data-ttu-id="dd472-113">Os scripts feitos com o Gravador de Ação só podem gravar informações na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="dd472-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="dd472-114">Com o Editor de Códigos, é possível editar e criar scripts que também leem dados de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="dd472-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="dd472-115">Vamos criar um script que leia dados e atue com base no que foi lido.</span><span class="sxs-lookup"><span data-stu-id="dd472-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="dd472-116">Vamos usar um exemplo de um extrato bancário.</span><span class="sxs-lookup"><span data-stu-id="dd472-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="dd472-117">Essa instrução é um relatório combinado de verificação de crédito.</span><span class="sxs-lookup"><span data-stu-id="dd472-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="dd472-118">Infelizmente, eles relatam alterações no balanço de forma diferente.</span><span class="sxs-lookup"><span data-stu-id="dd472-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="dd472-119">A declaração de verificação exibe o rendimento como crédito positivo e custos como débito negativo.</span><span class="sxs-lookup"><span data-stu-id="dd472-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="dd472-120">O demonstrativo de crédito faz o oposto.</span><span class="sxs-lookup"><span data-stu-id="dd472-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="dd472-121">No resto do tutorial, normalizaremos os dados usando um script.</span><span class="sxs-lookup"><span data-stu-id="dd472-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="dd472-122">Primeiro, vamos aprender a ler os dados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="dd472-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="dd472-123">Crie uma nova planilha na pasta de trabalho usada para o resto do tutorial.</span><span class="sxs-lookup"><span data-stu-id="dd472-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="dd472-124">Copie os seguintes dados e cole-os na nova planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="dd472-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="dd472-125">Data</span><span class="sxs-lookup"><span data-stu-id="dd472-125">Date</span></span> |<span data-ttu-id="dd472-126">Conta</span><span class="sxs-lookup"><span data-stu-id="dd472-126">Account</span></span> |<span data-ttu-id="dd472-127">Descrição</span><span class="sxs-lookup"><span data-stu-id="dd472-127">Description</span></span> |<span data-ttu-id="dd472-128">Débito</span><span class="sxs-lookup"><span data-stu-id="dd472-128">Debit</span></span> |<span data-ttu-id="dd472-129">Crédito</span><span class="sxs-lookup"><span data-stu-id="dd472-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="dd472-130">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-130">10/10/2019</span></span> |<span data-ttu-id="dd472-131">Verificando</span><span class="sxs-lookup"><span data-stu-id="dd472-131">Checking</span></span> |<span data-ttu-id="dd472-132">Vinícola Coho</span><span class="sxs-lookup"><span data-stu-id="dd472-132">Coho Vineyard</span></span> |<span data-ttu-id="dd472-133">-20.05</span><span class="sxs-lookup"><span data-stu-id="dd472-133">-20.05</span></span> | |
    |<span data-ttu-id="dd472-134">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-134">10/11/2019</span></span> |<span data-ttu-id="dd472-135">Crédito</span><span class="sxs-lookup"><span data-stu-id="dd472-135">Credit</span></span> |<span data-ttu-id="dd472-136">A Companhia Telefônica</span><span class="sxs-lookup"><span data-stu-id="dd472-136">The Phone Company</span></span> |<span data-ttu-id="dd472-137">99.95</span><span class="sxs-lookup"><span data-stu-id="dd472-137">99.95</span></span> | |
    |<span data-ttu-id="dd472-138">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-138">10/13/2019</span></span> |<span data-ttu-id="dd472-139">Crédito</span><span class="sxs-lookup"><span data-stu-id="dd472-139">Credit</span></span> |<span data-ttu-id="dd472-140">Vinícola Coho</span><span class="sxs-lookup"><span data-stu-id="dd472-140">Coho Vineyard</span></span> |<span data-ttu-id="dd472-141">154.43</span><span class="sxs-lookup"><span data-stu-id="dd472-141">154.43</span></span> | |
    |<span data-ttu-id="dd472-142">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-142">10/15/2019</span></span> |<span data-ttu-id="dd472-143">Verificando</span><span class="sxs-lookup"><span data-stu-id="dd472-143">Checking</span></span> |<span data-ttu-id="dd472-144">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="dd472-144">External Deposit</span></span> | |<span data-ttu-id="dd472-145">1000</span><span class="sxs-lookup"><span data-stu-id="dd472-145">1000</span></span> |
    |<span data-ttu-id="dd472-146">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-146">10/20/2019</span></span> |<span data-ttu-id="dd472-147">Crédito</span><span class="sxs-lookup"><span data-stu-id="dd472-147">Credit</span></span> |<span data-ttu-id="dd472-148">Vinícola Coho – Reembolso</span><span class="sxs-lookup"><span data-stu-id="dd472-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="dd472-149">-35.45</span><span class="sxs-lookup"><span data-stu-id="dd472-149">-35.45</span></span> |
    |<span data-ttu-id="dd472-150">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-150">10/25/2019</span></span> |<span data-ttu-id="dd472-151">Verificando</span><span class="sxs-lookup"><span data-stu-id="dd472-151">Checking</span></span> |<span data-ttu-id="dd472-152">Ideal para sua empresa de produtos orgânicos</span><span class="sxs-lookup"><span data-stu-id="dd472-152">Best For You Organics Company</span></span> | <span data-ttu-id="dd472-153">-85.64</span><span class="sxs-lookup"><span data-stu-id="dd472-153">-85.64</span></span> | |
    |<span data-ttu-id="dd472-154">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="dd472-154">11/01/2019</span></span> |<span data-ttu-id="dd472-155">Verificando</span><span class="sxs-lookup"><span data-stu-id="dd472-155">Checking</span></span> |<span data-ttu-id="dd472-156">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="dd472-156">External Deposit</span></span> | |<span data-ttu-id="dd472-157">1000</span><span class="sxs-lookup"><span data-stu-id="dd472-157">1000</span></span> |

3. <span data-ttu-id="dd472-158">Abra **Todos os Scripts** e selecione **Novo Script**.</span><span class="sxs-lookup"><span data-stu-id="dd472-158">Open **All Scripts** and select **New Script**.</span></span>
4. <span data-ttu-id="dd472-159">Vamos limpar a formatação.</span><span class="sxs-lookup"><span data-stu-id="dd472-159">Let's clean up the formatting.</span></span> <span data-ttu-id="dd472-160">Este é um documento financeiro, iremos alterar a formatação dos números nas colunas **Débito** e **Crédito** para mostrar os valores em dólares.</span><span class="sxs-lookup"><span data-stu-id="dd472-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="dd472-161">Também iremos ajustar a largura da coluna para os dados.</span><span class="sxs-lookup"><span data-stu-id="dd472-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="dd472-162">Substitua o conteúdo do script pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="dd472-162">Replace the script contents with the following code:</span></span>

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

5. <span data-ttu-id="dd472-163">Agora, leremos um valor de uma das colunas de número.</span><span class="sxs-lookup"><span data-stu-id="dd472-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="dd472-164">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="dd472-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="dd472-165">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="dd472-165">Run the script.</span></span>
7. <span data-ttu-id="dd472-166">Você deverá ver `[Array[1]]` no console.</span><span class="sxs-lookup"><span data-stu-id="dd472-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="dd472-167">Isso não é um número porque os intervalos são matrizes bidimensionais de dados.</span><span class="sxs-lookup"><span data-stu-id="dd472-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="dd472-168">Esse intervalo bidimensional está sendo registrado diretamente no console.</span><span class="sxs-lookup"><span data-stu-id="dd472-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="dd472-169">Felizmente, o Editor de Códigos permite visualizar o conteúdo da matriz.</span><span class="sxs-lookup"><span data-stu-id="dd472-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="dd472-170">Quando uma matriz bidimensional é registrada no console, ela agrupa os valores de coluna em cada linha.</span><span class="sxs-lookup"><span data-stu-id="dd472-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="dd472-171">Expanda o log de matriz pressionando o triângulo azul.</span><span class="sxs-lookup"><span data-stu-id="dd472-171">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="dd472-172">Expanda o segundo nível da matriz, pressionando o triângulo azul exibido recentemente.</span><span class="sxs-lookup"><span data-stu-id="dd472-172">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="dd472-173">Agora, você deverá ver isto:</span><span class="sxs-lookup"><span data-stu-id="dd472-173">You should now see this:</span></span>

    :::image type="content" source="../images/tutorial-4.png" alt-text="O log do console exibindo a saída &quot;-20.05&quot;, aninhada em duas matrizes":::

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="dd472-175">Modificar o valor de uma célula</span><span class="sxs-lookup"><span data-stu-id="dd472-175">Modify the value of a cell</span></span>

<span data-ttu-id="dd472-176">Agora que podemos ler os dados, usaremos eles para modificar a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="dd472-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="dd472-177">Deixaremos o valor da célula **D2** positivo com a função `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="dd472-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="dd472-178">O objeto [Matemática](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contém várias funções às quais seus scripts têm acesso.</span><span class="sxs-lookup"><span data-stu-id="dd472-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="dd472-179">É possível encontrar mais informações sobre `Math` e outros objetos internos [Usando objetos JavaScript internos nos scripts do Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="dd472-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="dd472-180">Usaremos os métodos `getValue` e `setValue` para alterar o valor da célula.</span><span class="sxs-lookup"><span data-stu-id="dd472-180">We'll use `getValue` and `setValue` methods to change the value of the cell.</span></span> <span data-ttu-id="dd472-181">Esses métodos funcionam em uma única célula.</span><span class="sxs-lookup"><span data-stu-id="dd472-181">These methods work on a single cell.</span></span> <span data-ttu-id="dd472-182">Ao lidar com intervalos de várias células, use `getValues` e `setValues`.</span><span class="sxs-lookup"><span data-stu-id="dd472-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span> <span data-ttu-id="dd472-183">Adicione o seguinte código ao final do script:</span><span class="sxs-lookup"><span data-stu-id="dd472-183">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > <span data-ttu-id="dd472-184">Estamos [lançando](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) o valor retornado de `range.getValue()` para um `number` usando a palavra-chave `as`.</span><span class="sxs-lookup"><span data-stu-id="dd472-184">We are [casting](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) the returned value of `range.getValue()` to a `number` by using the `as` keyword.</span></span> <span data-ttu-id="dd472-185">Isso é necessário porque um intervalo pode ser cadeias de caracteres, números ou booleanas.</span><span class="sxs-lookup"><span data-stu-id="dd472-185">This is necessary because a range could be strings, numbers, or booleans.</span></span> <span data-ttu-id="dd472-186">Nesta instância, precisamos explicitamente de um número.</span><span class="sxs-lookup"><span data-stu-id="dd472-186">In this instance, we explicitly need a number.</span></span>

2. <span data-ttu-id="dd472-187">O valor da célula **D2** agora deverá ser positivo.</span><span class="sxs-lookup"><span data-stu-id="dd472-187">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="dd472-188">Modificar os valores de uma coluna</span><span class="sxs-lookup"><span data-stu-id="dd472-188">Modify the values of a column</span></span>

<span data-ttu-id="dd472-189">Agora que sabemos ler e escrever em uma única célula, vamos generalizar o script para trabalhar em todas as colunas de **Débito** e **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="dd472-189">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="dd472-190">Remova o código que afeta apenas uma única célula (o código de valor absoluto anterior), de modo que o script agora se pareça com este:</span><span class="sxs-lookup"><span data-stu-id="dd472-190">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

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

2. <span data-ttu-id="dd472-191">Adicione um loop que percorra as linhas nas duas últimas colunas.</span><span class="sxs-lookup"><span data-stu-id="dd472-191">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="dd472-192">Para cada célula, o script define o valor para o valor absoluto do valor atual.</span><span class="sxs-lookup"><span data-stu-id="dd472-192">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="dd472-193">Observe que a matriz que define a localização das células é baseada em zero.</span><span class="sxs-lookup"><span data-stu-id="dd472-193">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="dd472-194">Isso significa que a célula **A1** é `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="dd472-194">That means cell **A1** is `range[0][0]`.</span></span>

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

    <span data-ttu-id="dd472-195">Essa parte do script faz várias tarefas importantes.</span><span class="sxs-lookup"><span data-stu-id="dd472-195">This portion of the script does several important tasks.</span></span> <span data-ttu-id="dd472-196">Primeiro, ela obtém os valores e a contagem de linhas do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="dd472-196">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="dd472-197">Isso nos permite ver os valores e saber quando parar.</span><span class="sxs-lookup"><span data-stu-id="dd472-197">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="dd472-198">Segundo, ela reitera através do intervalo usado, verificando cada célula nas colunas **Débito** ou **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="dd472-198">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="dd472-199">Por fim, se o valor na célula não for 0, ele será substituído pelo valor absoluto.</span><span class="sxs-lookup"><span data-stu-id="dd472-199">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="dd472-200">Estamos evitando zeros, para que possamos deixar as células em branco.</span><span class="sxs-lookup"><span data-stu-id="dd472-200">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="dd472-201">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="dd472-201">Run the script.</span></span>

    <span data-ttu-id="dd472-202">Seu extrato bancário agora deverá ter a seguinte aparência:</span><span class="sxs-lookup"><span data-stu-id="dd472-202">Your banking statement should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-5.png" alt-text="Uma planilha mostrando o extrato bancário como uma tabela formatada apenas com valores positivos.":::

## <a name="next-steps"></a><span data-ttu-id="dd472-204">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="dd472-204">Next steps</span></span>

<span data-ttu-id="dd472-205">Abra o Editor de códigos e experimente alguns dos [Scripts de exemplo para scripts do Office no Excel na Web](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="dd472-205">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="dd472-206">Visite também [Fundamentos de Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md) para saber mais sobre como criar scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="dd472-206">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="dd472-207">A próxima série de tutoriais de Scripts do Office tem foco na utilização de Scripts do Office com o Power Automate.</span><span class="sxs-lookup"><span data-stu-id="dd472-207">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="dd472-208">Saiba mais sobre as vantagens da combinação das duas plataformas em [Executar Scripts do Office com o Power Automate](../develop/power-automate-integration.md) ou tente o tutorial [Chamar Scripts no manual de fluxo do Power Automate](excel-power-automate-manual.md) para criar um fluxo no Power Automate que utiliza um Script do Office.</span><span class="sxs-lookup"><span data-stu-id="dd472-208">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
