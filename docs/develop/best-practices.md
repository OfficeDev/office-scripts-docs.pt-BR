---
title: Práticas recomendadas no Scripts do Office
description: Como evitar problemas comuns e gravar scripts robustos Office que podem manipular entradas ou dados inesperados.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546018"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="cacd4-103">Práticas recomendadas no Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="cacd4-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="cacd4-104">Esses padrões e práticas são projetados para ajudar seus scripts a executar com êxito sempre.</span><span class="sxs-lookup"><span data-stu-id="cacd4-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="cacd4-105">Use-os para evitar armadilhas comuns à medida que você começa a automatizar seu fluxo de trabalho Excel de usuário.</span><span class="sxs-lookup"><span data-stu-id="cacd4-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="cacd4-106">Verificar se um objeto está presente</span><span class="sxs-lookup"><span data-stu-id="cacd4-106">Verify an object is present</span></span>

<span data-ttu-id="cacd4-107">Os scripts geralmente dependem de uma determinada planilha ou tabela que está presente na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cacd4-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="cacd4-108">No entanto, eles podem ser renomeados ou removidos entre as executações de script.</span><span class="sxs-lookup"><span data-stu-id="cacd4-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="cacd4-109">Verificando se essas tabelas ou planilhas existem antes de chamar métodos neles, você pode garantir que o script não termine abruptamente.</span><span class="sxs-lookup"><span data-stu-id="cacd4-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="cacd4-110">O código de exemplo a seguir verifica se a planilha "Index" está presente na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cacd4-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="cacd4-111">Se a planilha estiver presente, o script obtém um intervalo e continua.</span><span class="sxs-lookup"><span data-stu-id="cacd4-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="cacd4-112">Se não estiver presente, o script registra uma mensagem de erro personalizada.</span><span class="sxs-lookup"><span data-stu-id="cacd4-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="cacd4-113">O operador TypeScript `?` verifica se o objeto existe antes de chamar um método.</span><span class="sxs-lookup"><span data-stu-id="cacd4-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="cacd4-114">Isso pode tornar seu código mais simplificado se você não precisar fazer nada especial quando o objeto não existir.</span><span class="sxs-lookup"><span data-stu-id="cacd4-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="cacd4-115">Validar primeiro os dados e o estado da workbook</span><span class="sxs-lookup"><span data-stu-id="cacd4-115">Validate data and workbook state first</span></span>

<span data-ttu-id="cacd4-116">Certifique-se de que todas as planilhas, tabelas, formas e outros objetos estão presentes antes de trabalhar nos dados.</span><span class="sxs-lookup"><span data-stu-id="cacd4-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="cacd4-117">Usando o padrão anterior, verifique se tudo está na caixa de trabalho e corresponde às suas expectativas.</span><span class="sxs-lookup"><span data-stu-id="cacd4-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="cacd4-118">Fazer isso antes que qualquer dado seja gravado garante que o script não deixe a workbook em um estado parcial.</span><span class="sxs-lookup"><span data-stu-id="cacd4-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="cacd4-119">O script a seguir exige que duas tabelas chamadas "Table1" e "Table2" sejam presentes.</span><span class="sxs-lookup"><span data-stu-id="cacd4-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="cacd4-120">O script primeiro verifica se as tabelas estão presentes e termina com a instrução e `return` uma mensagem apropriada, se não estiver.</span><span class="sxs-lookup"><span data-stu-id="cacd4-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="cacd4-121">Se a verificação estiver ocorrendo em uma função separada, você ainda deve encerrar o script em emissão da `return` instrução da `main` função.</span><span class="sxs-lookup"><span data-stu-id="cacd4-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="cacd4-122">Retornar da subfunção não termina o script.</span><span class="sxs-lookup"><span data-stu-id="cacd4-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="cacd4-123">O script a seguir tem o mesmo comportamento do anterior.</span><span class="sxs-lookup"><span data-stu-id="cacd4-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="cacd4-124">A diferença é que a `main` função chama a função para verificar `inputPresent` tudo.</span><span class="sxs-lookup"><span data-stu-id="cacd4-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="cacd4-125">`inputPresent` retorna um booleano ( `true` ou ) para indicar se todas as entradas necessárias estão `false` presentes.</span><span class="sxs-lookup"><span data-stu-id="cacd4-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="cacd4-126">A `main` função usa esse booleano para decidir sobre continuar ou encerrar o script.</span><span class="sxs-lookup"><span data-stu-id="cacd4-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="cacd4-127">Quando usar uma `throw` instrução</span><span class="sxs-lookup"><span data-stu-id="cacd4-127">When to use a `throw` statement</span></span>

<span data-ttu-id="cacd4-128">Uma [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instrução indica que ocorreu um erro inesperado.</span><span class="sxs-lookup"><span data-stu-id="cacd4-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="cacd4-129">Ele termina o código imediatamente.</span><span class="sxs-lookup"><span data-stu-id="cacd4-129">It ends the code immediately.</span></span> <span data-ttu-id="cacd4-130">Na maior parte, você não precisa fazer `throw` isso no script.</span><span class="sxs-lookup"><span data-stu-id="cacd4-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="cacd4-131">Normalmente, o script informa automaticamente ao usuário que o script não foi executado devido a um problema.</span><span class="sxs-lookup"><span data-stu-id="cacd4-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="cacd4-132">Na maioria dos casos, é suficiente terminar o script com uma mensagem de erro e `return` uma instrução da `main` função.</span><span class="sxs-lookup"><span data-stu-id="cacd4-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="cacd4-133">No entanto, se o script estiver sendo executado como parte de um fluxo Power Automate, talvez você queira impedir que o fluxo continue.</span><span class="sxs-lookup"><span data-stu-id="cacd4-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="cacd4-134">Uma `throw` instrução interrompe o script e diz ao fluxo para parar também.</span><span class="sxs-lookup"><span data-stu-id="cacd4-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="cacd4-135">O script a seguir mostra como usar a `throw` instrução em nosso exemplo de verificação de tabela.</span><span class="sxs-lookup"><span data-stu-id="cacd4-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="cacd4-136">Quando usar uma `try...catch` instrução</span><span class="sxs-lookup"><span data-stu-id="cacd4-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="cacd4-137">A [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrução é uma maneira de detectar se uma chamada de API falhará e continuará executando o script.</span><span class="sxs-lookup"><span data-stu-id="cacd4-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="cacd4-138">Considere o trecho a seguir que executa uma grande atualização de dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="cacd4-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="cacd4-139">Se `someLargeValues` for maior do que Excel a Web pode manipular, a chamada `setValues()` falhará.</span><span class="sxs-lookup"><span data-stu-id="cacd4-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="cacd4-140">Em seguida, o script também falha com um [erro de tempo de execução.](../testing/troubleshooting.md#runtime-errors)</span><span class="sxs-lookup"><span data-stu-id="cacd4-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="cacd4-141">A `try...catch` instrução permite que seu script reconheça essa condição, sem encerrar imediatamente o script e mostrar o erro padrão.</span><span class="sxs-lookup"><span data-stu-id="cacd4-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="cacd4-142">Uma abordagem para dar ao usuário de script uma experiência melhor é apresentar uma mensagem de erro personalizada.</span><span class="sxs-lookup"><span data-stu-id="cacd4-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="cacd4-143">O trecho a seguir mostra uma `try...catch` instrução registrando mais informações de erro para ajudar melhor o leitor.</span><span class="sxs-lookup"><span data-stu-id="cacd4-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="cacd4-144">Outra abordagem para lidar com erros é ter comportamento de fallback que lida com o caso de erro.</span><span class="sxs-lookup"><span data-stu-id="cacd4-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="cacd4-145">O trecho a seguir usa o bloco para tentar um método alternativo separar a atualização `catch` em partes menores e evitar o erro.</span><span class="sxs-lookup"><span data-stu-id="cacd4-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="cacd4-146">Para ver um exemplo completo sobre como atualizar um intervalo grande, consulte [Write a large dataset](../resources/samples/write-large-dataset.md).</span><span class="sxs-lookup"><span data-stu-id="cacd4-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="cacd4-147">Usar `try...catch` dentro ou ao redor de um loop retarda seu script.</span><span class="sxs-lookup"><span data-stu-id="cacd4-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="cacd4-148">Para obter mais informações de desempenho, consulte [Evite usar `try...catch` blocos](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span><span class="sxs-lookup"><span data-stu-id="cacd4-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="cacd4-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="cacd4-149">See also</span></span>

- [<span data-ttu-id="cacd4-150">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="cacd4-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="cacd4-151">Solução de problemas de informações para Power Automate com Office Scripts</span><span class="sxs-lookup"><span data-stu-id="cacd4-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="cacd4-152">Limites da plataforma com Office Scripts</span><span class="sxs-lookup"><span data-stu-id="cacd4-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="cacd4-153">Melhorar o desempenho de seus Office Scripts</span><span class="sxs-lookup"><span data-stu-id="cacd4-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
