---
title: Grave, edite e crie scripts do Office no Excel na Web
description: Um tutorial sobre o básico dos scripts do Office, incluindo a gravação de scripts com o Gravador de ações e a gravação de dados em uma pasta de trabalho.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 1971ff2ffd80554beb6ac561677ee3384f87ca81
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700037"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="53734-103">Grave, edite e crie scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="53734-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="53734-104">Este tutorial ensinará os conceitos básicos de gravação, edição e escrita de um Script do Office para Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="53734-104">This tutorial will teach you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="53734-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="53734-105">Prerequisites</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="53734-106">Antes de iniciar este tutorial, você precisará acessar os scripts do Office, que exigem o seguinte:</span><span class="sxs-lookup"><span data-stu-id="53734-106">Before starting this tutorial, you'll need access to Office Scripts, which requires the following:</span></span>

- <span data-ttu-id="53734-107">[Excel na Web](https://www.office.com/launch/excel).</span><span class="sxs-lookup"><span data-stu-id="53734-107">[Excel on the web](https://www.office.com/launch/excel).</span></span>
- <span data-ttu-id="53734-108">Peça para o administrador [habilitar os scripts do Office da sua organização](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), o que adiciona a guia **Automação** à faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="53734-108">Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="53734-109">Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="53734-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="53734-110">Se você não conhece o JavaScript, recomendamos que revise o [tutorial do Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="53734-110">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="53734-111">Visite [Scripts do Office no Excel na Web](../overview/excel.md) para saber mais sobre o ambiente de scripts.</span><span class="sxs-lookup"><span data-stu-id="53734-111">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="53734-112">Adicione dados e grave um script básico</span><span class="sxs-lookup"><span data-stu-id="53734-112">Add data and record a basic script</span></span>

<span data-ttu-id="53734-113">Primeiro, precisaremos de alguns dados e um pequeno script inicial.</span><span class="sxs-lookup"><span data-stu-id="53734-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="53734-114">Crie uma nova pasta de trabalho no Excel para a Web.</span><span class="sxs-lookup"><span data-stu-id="53734-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="53734-115">Copie os seguintes dados de vendas de frutas e cole-os na planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="53734-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="53734-116">Fruta</span><span class="sxs-lookup"><span data-stu-id="53734-116">Fruit</span></span> |<span data-ttu-id="53734-117">2018</span><span class="sxs-lookup"><span data-stu-id="53734-117">2018</span></span> |<span data-ttu-id="53734-118">2019</span><span class="sxs-lookup"><span data-stu-id="53734-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="53734-119">Laranjas</span><span class="sxs-lookup"><span data-stu-id="53734-119">Oranges</span></span> |<span data-ttu-id="53734-120">1.000</span><span class="sxs-lookup"><span data-stu-id="53734-120">1000</span></span> |<span data-ttu-id="53734-121">1.200</span><span class="sxs-lookup"><span data-stu-id="53734-121">1200</span></span> |
    |<span data-ttu-id="53734-122">Limões</span><span class="sxs-lookup"><span data-stu-id="53734-122">Lemons</span></span> |<span data-ttu-id="53734-123">800</span><span class="sxs-lookup"><span data-stu-id="53734-123">800</span></span> |<span data-ttu-id="53734-124">900</span><span class="sxs-lookup"><span data-stu-id="53734-124">900</span></span> |
    |<span data-ttu-id="53734-125">Limões-galego</span><span class="sxs-lookup"><span data-stu-id="53734-125">Limes</span></span> |<span data-ttu-id="53734-126">600</span><span class="sxs-lookup"><span data-stu-id="53734-126">600</span></span> |<span data-ttu-id="53734-127">500</span><span class="sxs-lookup"><span data-stu-id="53734-127">500</span></span> |
    |<span data-ttu-id="53734-128">Toranjas</span><span class="sxs-lookup"><span data-stu-id="53734-128">Grapefruits</span></span> |<span data-ttu-id="53734-129">900</span><span class="sxs-lookup"><span data-stu-id="53734-129">900</span></span> |<span data-ttu-id="53734-130">700</span><span class="sxs-lookup"><span data-stu-id="53734-130">700</span></span> |

3. <span data-ttu-id="53734-131">Abra a guia **Automação**. Se você não vir a guia **Automação**, verifique o extravasamento da fita pressionando a seta suspensa.</span><span class="sxs-lookup"><span data-stu-id="53734-131">Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span>
4. <span data-ttu-id="53734-132">Pressione o botão **Ações de registro**.</span><span class="sxs-lookup"><span data-stu-id="53734-132">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="53734-133">Clique nas células **A2:C2** (a linha "Laranjas") e defina a cor de preenchimento como laranja.</span><span class="sxs-lookup"><span data-stu-id="53734-133">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="53734-134">Pare a gravação pressionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="53734-134">Stop the recording by pressing the **Stop** button.</span></span>
7. <span data-ttu-id="53734-135">Preencha o campo **Nome do script** com um nome digno de memória.</span><span class="sxs-lookup"><span data-stu-id="53734-135">Fill in the **Script Name** field with a memorable name.</span></span>
8. <span data-ttu-id="53734-136">*Opcional:* Preencha o campo **Descrição** com uma descrição significativa.</span><span class="sxs-lookup"><span data-stu-id="53734-136">*Optional:* Fill in the **Description** field with a meaningful description.</span></span> <span data-ttu-id="53734-137">Isso é usado para fornecer contexto sobre o que o script faz.</span><span class="sxs-lookup"><span data-stu-id="53734-137">This is used to provide context as to what the script does.</span></span> <span data-ttu-id="53734-138">Para o tutorial, você pode usar "Linhas de códigos de cores de uma tabela".</span><span class="sxs-lookup"><span data-stu-id="53734-138">For the tutorial, you can use "Color-codes rows of a table".</span></span>

   > [!TIP]
   > <span data-ttu-id="53734-139">Você pode editar a descrição de um script posteriormente no painel **Detalhes do script**, localizado no menu do Editor de códigos **...**.</span><span class="sxs-lookup"><span data-stu-id="53734-139">You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.</span></span>

9. <span data-ttu-id="53734-140">Salve o script pressionando o botão **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="53734-140">Save the script by pressing the **Save** button.</span></span>

    <span data-ttu-id="53734-141">Sua planilha deve ficar assim (não se preocupe se a cor for diferente):</span><span class="sxs-lookup"><span data-stu-id="53734-141">Your worksheet should look like this (don't worry if the color is different):</span></span>

    ![Uma linha de dados de vendas de frutas com a linha "Laranjas" destacada em laranja.](../images/tutorial-1.png)

## <a name="edit-an-existing-script"></a><span data-ttu-id="53734-143">Edite um script existente</span><span class="sxs-lookup"><span data-stu-id="53734-143">Edit an existing script</span></span>

<span data-ttu-id="53734-144">O script anterior coloriu a linha "Laranjas" para ficar laranja.</span><span class="sxs-lookup"><span data-stu-id="53734-144">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="53734-145">Vamos adicionar uma linha amarela aos "Limões".</span><span class="sxs-lookup"><span data-stu-id="53734-145">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="53734-146">Abra a guia **Automação**.</span><span class="sxs-lookup"><span data-stu-id="53734-146">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="53734-147">Pressione o botão **Editor de códigos**.</span><span class="sxs-lookup"><span data-stu-id="53734-147">Press the **Code Editor** button.</span></span>
3. <span data-ttu-id="53734-148">Abra o script que você gravou na seção anterior.</span><span class="sxs-lookup"><span data-stu-id="53734-148">Open the script you recorded in the previous section.</span></span> <span data-ttu-id="53734-149">Você deve ver algo semelhante a este código:</span><span class="sxs-lookup"><span data-stu-id="53734-149">You should see something similar to this code:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
    }
    ```

    <span data-ttu-id="53734-150">Esse código obtém a planilha atual acessando primeiro a coleção de planilhas da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="53734-150">This code gets the current worksheet by first accessing the workbook's worksheet collection.</span></span> <span data-ttu-id="53734-151">Depois, defina a cor de preenchimento do intervalo **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="53734-151">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="53734-152">Os intervalos são parte fundamental dos scripts do Office no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="53734-152">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="53734-153">Um intervalo é um bloco retangular e contíguo de células que contém valores, fórmula e formatação.</span><span class="sxs-lookup"><span data-stu-id="53734-153">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="53734-154">Eles são a estrutura básica das células através da qual você executará a maioria das tarefas de script.</span><span class="sxs-lookup"><span data-stu-id="53734-154">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

4. <span data-ttu-id="53734-155">Adicione a seguinte linha no final do script (entre onde `color` está definido e o encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="53734-155">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
    ```

5. <span data-ttu-id="53734-156">Teste o script pressionando **Executar**.</span><span class="sxs-lookup"><span data-stu-id="53734-156">Test the script by pressing **Run**.</span></span> <span data-ttu-id="53734-157">Sua pasta de trabalho já deve ter esta aparência:</span><span class="sxs-lookup"><span data-stu-id="53734-157">Your workbook should now look like this:</span></span>

    ![Uma linha de dados de vendas de frutas com a linha "Laranjas" é realçada em laranja e a linha "Limões" é realçada em amarelo.](../images/tutorial-2.png)

## <a name="create-a-table"></a><span data-ttu-id="53734-159">Crie uma tabela</span><span class="sxs-lookup"><span data-stu-id="53734-159">Create a table</span></span>

<span data-ttu-id="53734-160">Vamos converter esses dados de vendas de frutas em uma tabela.</span><span class="sxs-lookup"><span data-stu-id="53734-160">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="53734-161">Usaremos nosso script em todo o processo.</span><span class="sxs-lookup"><span data-stu-id="53734-161">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="53734-162">Adicione a seguinte linha no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="53734-162">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.tables.add("A1:C5", true);
    ```

2. <span data-ttu-id="53734-163">Essa chamada retorna um `Table` objeto.</span><span class="sxs-lookup"><span data-stu-id="53734-163">That call returns a `Table` object.</span></span> <span data-ttu-id="53734-164">Vamos usar essa tabela para classificar os dados.</span><span class="sxs-lookup"><span data-stu-id="53734-164">Let's use that table to sort the data.</span></span> <span data-ttu-id="53734-165">Classificaremos os dados em ordem crescente com base nos valores na coluna "Frutas".</span><span class="sxs-lookup"><span data-stu-id="53734-165">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="53734-166">Adicione a seguinte linha assim que criar a tabela:</span><span class="sxs-lookup"><span data-stu-id="53734-166">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.sort.apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="53734-167">Seu script deve ter esta aparência:</span><span class="sxs-lookup"><span data-stu-id="53734-167">Your script should look like this:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
      selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
      let table = selectedSheet.tables.add("A1:C5", true);
      table.sort.apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="53734-168">As tabelas têm um objeto `TableSort` acessado através da propriedade `Table.sort`.</span><span class="sxs-lookup"><span data-stu-id="53734-168">Tables have a `TableSort` object, accessed through the `Table.sort` property.</span></span> <span data-ttu-id="53734-169">Você pode aplicar critérios de classificação a esse objeto.</span><span class="sxs-lookup"><span data-stu-id="53734-169">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="53734-170">O `apply` método utiliza uma matriz de `SortField` objetos.</span><span class="sxs-lookup"><span data-stu-id="53734-170">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="53734-171">Nesse caso, só temos um critério de classificação, por isso só usamos um. `SortField`.</span><span class="sxs-lookup"><span data-stu-id="53734-171">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="53734-172">`key: 0` define a coluna com os valores que determinam a classificação como "0" (que nesse caso, é a primeira coluna na tabela **A** ).</span><span class="sxs-lookup"><span data-stu-id="53734-172">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="53734-173">`ascending: true` classifica os dados em ordem crescente (em vez de ordem decrescente).</span><span class="sxs-lookup"><span data-stu-id="53734-173">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="53734-174">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="53734-174">Run the script.</span></span> <span data-ttu-id="53734-175">Você deve visualizar uma tabela como esta:</span><span class="sxs-lookup"><span data-stu-id="53734-175">You should see a table like this:</span></span>

    ![Uma tabela de vendas de frutas sortidas.](../images/tutorial-3.png)

    > [!NOTE]
    > <span data-ttu-id="53734-177">Se você executar novamente o script, receberá um erro.</span><span class="sxs-lookup"><span data-stu-id="53734-177">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="53734-178">Isso ocorre porque você não pode criar uma tabela em cima de outra tabela.</span><span class="sxs-lookup"><span data-stu-id="53734-178">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="53734-179">No entanto, você pode executar o script em uma planilha ou pasta de trabalho diferente.</span><span class="sxs-lookup"><span data-stu-id="53734-179">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="53734-180">Reexecute o script</span><span class="sxs-lookup"><span data-stu-id="53734-180">Re-run the script</span></span>

1. <span data-ttu-id="53734-181">Crie uma nova planilha na pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="53734-181">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="53734-182">Copie os dados das frutas do início do tutorial e cole-os na nova planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="53734-182">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="53734-183">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="53734-183">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="53734-184">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="53734-184">Next steps</span></span>

<span data-ttu-id="53734-185">Conclua o tutorial [Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="53734-185">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="53734-186">Ele ensina como ler dados de uma pasta de trabalho com um script do Office.</span><span class="sxs-lookup"><span data-stu-id="53734-186">It teaches you how to read data from a workbook with an Office Script.</span></span>
