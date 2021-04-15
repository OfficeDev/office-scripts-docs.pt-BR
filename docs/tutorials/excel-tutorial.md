---
title: Grave, edite e crie scripts do Office no Excel na Web
description: Um tutorial sobre o básico dos scripts do Office, incluindo a gravação de scripts com o Gravador de ações e a gravação de dados em uma pasta de trabalho.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: ae864cc08453a9c8a2538f15ceee1275e131725d
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754842"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="eed31-103">Grave, edite e crie scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="eed31-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="eed31-104">Este tutorial ensina os fundamentos da gravação, edição e escrita de um Script do para o Excel na web.</span><span class="sxs-lookup"><span data-stu-id="eed31-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="eed31-105">Você gravará um script que aplicará uma determinada formatação a uma planilha de registro de vendas.</span><span class="sxs-lookup"><span data-stu-id="eed31-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="eed31-106">Depois, você editará o script gravado para aplicar outras formatações, criar e classificar uma tabela.</span><span class="sxs-lookup"><span data-stu-id="eed31-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="eed31-107">Este padrão de registro e edição é uma importante ferramenta para ver como suas ações no Excel são parecidas com um código.</span><span class="sxs-lookup"><span data-stu-id="eed31-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="eed31-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="eed31-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="eed31-109">Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="eed31-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="eed31-110">Se você é novo no JavaScript, recomendamos começar com o [tutorial da Mozilla sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="eed31-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="eed31-111">Visite o [ambiente do Editor de Código do Scripts do Office](../overview/code-editor-environment.md) para saber mais sobre o ambiente de script.</span><span class="sxs-lookup"><span data-stu-id="eed31-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="eed31-112">Adicione dados e grave um script básico</span><span class="sxs-lookup"><span data-stu-id="eed31-112">Add data and record a basic script</span></span>

<span data-ttu-id="eed31-113">Primeiro, precisaremos de alguns dados e um pequeno script inicial.</span><span class="sxs-lookup"><span data-stu-id="eed31-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="eed31-114">Crie uma nova pasta de trabalho no Excel para a Web.</span><span class="sxs-lookup"><span data-stu-id="eed31-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="eed31-115">Copie os seguintes dados de vendas de frutas e cole-os na planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="eed31-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="eed31-116">Fruta</span><span class="sxs-lookup"><span data-stu-id="eed31-116">Fruit</span></span> |<span data-ttu-id="eed31-117">2018</span><span class="sxs-lookup"><span data-stu-id="eed31-117">2018</span></span> |<span data-ttu-id="eed31-118">2019</span><span class="sxs-lookup"><span data-stu-id="eed31-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="eed31-119">Laranjas</span><span class="sxs-lookup"><span data-stu-id="eed31-119">Oranges</span></span> |<span data-ttu-id="eed31-120">1.000</span><span class="sxs-lookup"><span data-stu-id="eed31-120">1000</span></span> |<span data-ttu-id="eed31-121">1.200</span><span class="sxs-lookup"><span data-stu-id="eed31-121">1200</span></span> |
    |<span data-ttu-id="eed31-122">Limões</span><span class="sxs-lookup"><span data-stu-id="eed31-122">Lemons</span></span> |<span data-ttu-id="eed31-123">800</span><span class="sxs-lookup"><span data-stu-id="eed31-123">800</span></span> |<span data-ttu-id="eed31-124">900</span><span class="sxs-lookup"><span data-stu-id="eed31-124">900</span></span> |
    |<span data-ttu-id="eed31-125">Limões-galego</span><span class="sxs-lookup"><span data-stu-id="eed31-125">Limes</span></span> |<span data-ttu-id="eed31-126">600</span><span class="sxs-lookup"><span data-stu-id="eed31-126">600</span></span> |<span data-ttu-id="eed31-127">500</span><span class="sxs-lookup"><span data-stu-id="eed31-127">500</span></span> |
    |<span data-ttu-id="eed31-128">Toranjas</span><span class="sxs-lookup"><span data-stu-id="eed31-128">Grapefruits</span></span> |<span data-ttu-id="eed31-129">900</span><span class="sxs-lookup"><span data-stu-id="eed31-129">900</span></span> |<span data-ttu-id="eed31-130">700</span><span class="sxs-lookup"><span data-stu-id="eed31-130">700</span></span> |

3. <span data-ttu-id="eed31-131">Abra a guia **Automação**. Se você não vir a guia **Automação**, verifique o extravasamento da fita pressionando a seta suspensa.</span><span class="sxs-lookup"><span data-stu-id="eed31-131">Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span>
4. <span data-ttu-id="eed31-132">Pressione o botão **Ações de registro**.</span><span class="sxs-lookup"><span data-stu-id="eed31-132">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="eed31-133">Clique nas células **A2:C2** (a linha "Laranjas") e defina a cor de preenchimento como laranja.</span><span class="sxs-lookup"><span data-stu-id="eed31-133">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="eed31-134">Pare a gravação pressionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="eed31-134">Stop the recording by pressing the **Stop** button.</span></span>
7. <span data-ttu-id="eed31-135">Preencha o campo **Nome do script** com um nome digno de memória.</span><span class="sxs-lookup"><span data-stu-id="eed31-135">Fill in the **Script Name** field with a memorable name.</span></span>
8. <span data-ttu-id="eed31-136">*Opcional:* Preencha o campo **Descrição** com uma descrição significativa.</span><span class="sxs-lookup"><span data-stu-id="eed31-136">*Optional:* Fill in the **Description** field with a meaningful description.</span></span> <span data-ttu-id="eed31-137">Isso é usado para fornecer contexto sobre o que o script faz.</span><span class="sxs-lookup"><span data-stu-id="eed31-137">This is used to provide context as to what the script does.</span></span> <span data-ttu-id="eed31-138">Para o tutorial, você pode usar "Linhas de códigos de cores de uma tabela".</span><span class="sxs-lookup"><span data-stu-id="eed31-138">For the tutorial, you can use "Color-codes rows of a table".</span></span>

   > [!TIP]
   > <span data-ttu-id="eed31-139">Você pode editar a descrição de um script posteriormente no painel **Detalhes do script**, localizado no menu do Editor de códigos **...**.</span><span class="sxs-lookup"><span data-stu-id="eed31-139">You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.</span></span>

9. <span data-ttu-id="eed31-140">Salve o script pressionando o botão **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="eed31-140">Save the script by pressing the **Save** button.</span></span>

    <span data-ttu-id="eed31-141">Sua planilha deve ficar assim (não se preocupe se a cor for diferente):</span><span class="sxs-lookup"><span data-stu-id="eed31-141">Your worksheet should look like this (don't worry if the color is different):</span></span>

    :::image type="content" source="../images/tutorial-1.png" alt-text="Uma planilha mostrando a linha de dados das vendas de frutas com a linha contendo &quot;Laranjas&quot; realçada na cor laranja.":::

## <a name="edit-an-existing-script"></a><span data-ttu-id="eed31-143">Edite um script existente</span><span class="sxs-lookup"><span data-stu-id="eed31-143">Edit an existing script</span></span>

<span data-ttu-id="eed31-144">O script anterior coloriu a linha "Laranjas" para ficar laranja.</span><span class="sxs-lookup"><span data-stu-id="eed31-144">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="eed31-145">Vamos adicionar uma linha amarela aos "Limões".</span><span class="sxs-lookup"><span data-stu-id="eed31-145">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="eed31-146">A partir do painel, agora aberto em **Detalhes**, pressione o botão **Editar**.</span><span class="sxs-lookup"><span data-stu-id="eed31-146">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="eed31-147">Você deve ver algo semelhante a este código:</span><span class="sxs-lookup"><span data-stu-id="eed31-147">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="eed31-148">Este código recebe a planilha atual da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="eed31-148">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="eed31-149">Depois, defina a cor de preenchimento do intervalo **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="eed31-149">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="eed31-150">Os intervalos são parte fundamental dos scripts do Office no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="eed31-150">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="eed31-151">Um intervalo é um bloco retangular e contíguo de células que contém valores, fórmula e formatação.</span><span class="sxs-lookup"><span data-stu-id="eed31-151">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="eed31-152">Eles são a estrutura básica das células através da qual você executará a maioria das tarefas de script.</span><span class="sxs-lookup"><span data-stu-id="eed31-152">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="eed31-153">Adicione a seguinte linha no final do script (entre onde `color` está definido e o encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="eed31-153">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="eed31-154">Teste o script pressionando **Executar**.</span><span class="sxs-lookup"><span data-stu-id="eed31-154">Test the script by pressing **Run**.</span></span> <span data-ttu-id="eed31-155">Sua pasta de trabalho já deve ter esta aparência:</span><span class="sxs-lookup"><span data-stu-id="eed31-155">Your workbook should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-2.png" alt-text="Uma planilha mostrando a linha de dados das vendas de frutas com a linha &quot;Laranjas&quot; realçada na cor laranja, e a linha &quot;Limões&quot; realçada na cor amarela.":::

## <a name="create-a-table"></a><span data-ttu-id="eed31-157">Crie uma tabela</span><span class="sxs-lookup"><span data-stu-id="eed31-157">Create a table</span></span>

<span data-ttu-id="eed31-158">Vamos converter esses dados de vendas de frutas em uma tabela.</span><span class="sxs-lookup"><span data-stu-id="eed31-158">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="eed31-159">Usaremos nosso script em todo o processo.</span><span class="sxs-lookup"><span data-stu-id="eed31-159">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="eed31-160">Adicione a seguinte linha no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="eed31-160">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="eed31-161">Essa chamada retorna um `Table` objeto.</span><span class="sxs-lookup"><span data-stu-id="eed31-161">That call returns a `Table` object.</span></span> <span data-ttu-id="eed31-162">Vamos usar essa tabela para classificar os dados.</span><span class="sxs-lookup"><span data-stu-id="eed31-162">Let's use that table to sort the data.</span></span> <span data-ttu-id="eed31-163">Classificaremos os dados em ordem crescente com base nos valores na coluna "Frutas".</span><span class="sxs-lookup"><span data-stu-id="eed31-163">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="eed31-164">Adicione a seguinte linha assim que criar a tabela:</span><span class="sxs-lookup"><span data-stu-id="eed31-164">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="eed31-165">Seu script deve ter esta aparência:</span><span class="sxs-lookup"><span data-stu-id="eed31-165">Your script should look like this:</span></span>

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

    <span data-ttu-id="eed31-166">As tabelas possuem um objeto`TableSort`, acessado por meio do método `Table.getSort`.</span><span class="sxs-lookup"><span data-stu-id="eed31-166">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="eed31-167">Você pode aplicar critérios de classificação a esse objeto.</span><span class="sxs-lookup"><span data-stu-id="eed31-167">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="eed31-168">O `apply` método utiliza uma matriz de `SortField` objetos.</span><span class="sxs-lookup"><span data-stu-id="eed31-168">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="eed31-169">Nesse caso, só temos um critério de classificação, por isso só usamos um. `SortField`.</span><span class="sxs-lookup"><span data-stu-id="eed31-169">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="eed31-170">`key: 0` define a coluna com os valores que determinam a classificação como "0" (que nesse caso, é a primeira coluna na tabela **A** ).</span><span class="sxs-lookup"><span data-stu-id="eed31-170">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="eed31-171">`ascending: true` classifica os dados em ordem crescente (em vez de ordem decrescente).</span><span class="sxs-lookup"><span data-stu-id="eed31-171">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="eed31-172">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="eed31-172">Run the script.</span></span> <span data-ttu-id="eed31-173">Você deve visualizar uma tabela como esta:</span><span class="sxs-lookup"><span data-stu-id="eed31-173">You should see a table like this:</span></span>

    :::image type="content" source="../images/tutorial-3.png" alt-text="Uma planilha mostrando a tabela ordenada de vendas de frutas.":::

    > [!NOTE]
    > <span data-ttu-id="eed31-175">Se você executar novamente o script, receberá um erro.</span><span class="sxs-lookup"><span data-stu-id="eed31-175">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="eed31-176">Isso ocorre porque você não pode criar uma tabela em cima de outra tabela.</span><span class="sxs-lookup"><span data-stu-id="eed31-176">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="eed31-177">No entanto, você pode executar o script em uma planilha ou pasta de trabalho diferente.</span><span class="sxs-lookup"><span data-stu-id="eed31-177">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="eed31-178">Reexecute o script</span><span class="sxs-lookup"><span data-stu-id="eed31-178">Re-run the script</span></span>

1. <span data-ttu-id="eed31-179">Crie uma nova planilha na pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="eed31-179">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="eed31-180">Copie os dados das frutas do início do tutorial e cole-os na nova planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="eed31-180">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="eed31-181">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="eed31-181">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="eed31-182">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="eed31-182">Next steps</span></span>

<span data-ttu-id="eed31-183">Conclua o tutorial [Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="eed31-183">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="eed31-184">Ele ensina como ler dados de uma pasta de trabalho com um script do Office.</span><span class="sxs-lookup"><span data-stu-id="eed31-184">It teaches you how to read data from a workbook with an Office Script.</span></span>
