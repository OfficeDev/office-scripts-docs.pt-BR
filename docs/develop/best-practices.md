---
title: Práticas recomendadas no Scripts do Office
description: Como prevenir problemas comuns e escrever scripts robustos Office que podem lidar com entradas ou dados inesperados.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546018"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="d919c-103">Práticas recomendadas no Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="d919c-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="d919c-104">Esses padrões e práticas são projetados para ajudar seus scripts a serem executados com sucesso todas as vezes.</span><span class="sxs-lookup"><span data-stu-id="d919c-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="d919c-105">Use-os para evitar armadilhas comuns à medida que você começa a automatizar seu fluxo de trabalho Excel.</span><span class="sxs-lookup"><span data-stu-id="d919c-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="d919c-106">Verifique se um objeto está presente</span><span class="sxs-lookup"><span data-stu-id="d919c-106">Verify an object is present</span></span>

<span data-ttu-id="d919c-107">Os scripts muitas vezes dependem de uma determinada planilha ou tabela presente na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d919c-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="d919c-108">No entanto, eles podem ser renomeados ou removidos entre as corridas de script.</span><span class="sxs-lookup"><span data-stu-id="d919c-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="d919c-109">Ao verificar se essas tabelas ou planilhas existem antes de chamar métodos sobre elas, você pode garantir que o script não termine abruptamente.</span><span class="sxs-lookup"><span data-stu-id="d919c-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="d919c-110">O código de amostra a seguir verifica se a planilha "Índice" está presente na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d919c-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="d919c-111">Se a planilha estiver presente, o script ganha uma faixa e prossegue.</span><span class="sxs-lookup"><span data-stu-id="d919c-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="d919c-112">Se ele não estiver presente, o script registra uma mensagem de erro personalizada.</span><span class="sxs-lookup"><span data-stu-id="d919c-112">If it isn't present, the script logs a custom error message.</span></span>

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

<span data-ttu-id="d919c-113">O operador TypeScript `?` verifica se o objeto existe antes de chamar um método.</span><span class="sxs-lookup"><span data-stu-id="d919c-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="d919c-114">Isso pode tornar seu código mais simplificado se você não precisar fazer nada especial quando o objeto não existe.</span><span class="sxs-lookup"><span data-stu-id="d919c-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="d919c-115">Validar dados e estado da pasta de trabalho primeiro</span><span class="sxs-lookup"><span data-stu-id="d919c-115">Validate data and workbook state first</span></span>

<span data-ttu-id="d919c-116">Certifique-se de que todas as planilhas, tabelas, formas e outros objetos estejam presentes antes de trabalhar nos dados.</span><span class="sxs-lookup"><span data-stu-id="d919c-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="d919c-117">Usando o padrão anterior, verifique se está tudo na pasta de trabalho e corresponda às suas expectativas.</span><span class="sxs-lookup"><span data-stu-id="d919c-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="d919c-118">Fazer isso antes que qualquer dado seja escrito garante que seu script não deixe a pasta de trabalho em um estado parcial.</span><span class="sxs-lookup"><span data-stu-id="d919c-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="d919c-119">O roteiro a seguir requer que duas tabelas chamadas "Tabela1" e "Tabela2" estejam presentes.</span><span class="sxs-lookup"><span data-stu-id="d919c-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="d919c-120">O script primeiro verifica se as tabelas estão presentes e, em seguida, termina com a `return` instrução e uma mensagem apropriada se não estiverem.</span><span class="sxs-lookup"><span data-stu-id="d919c-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

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

<span data-ttu-id="d919c-121">Se a verificação estiver acontecendo em uma função separada, você ainda deve terminar o script emitindo a `return` instrução da `main` função.</span><span class="sxs-lookup"><span data-stu-id="d919c-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="d919c-122">Voltar da subfunção não encerra o roteiro.</span><span class="sxs-lookup"><span data-stu-id="d919c-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="d919c-123">O roteiro a seguir tem o mesmo comportamento do anterior.</span><span class="sxs-lookup"><span data-stu-id="d919c-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="d919c-124">A diferença é que a `main` função chama a função para verificar `inputPresent` tudo.</span><span class="sxs-lookup"><span data-stu-id="d919c-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="d919c-125">`inputPresent` retorna um booleano ( `true` ou ) para indicar se todas as `false` entradas necessárias estão presentes.</span><span class="sxs-lookup"><span data-stu-id="d919c-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="d919c-126">A `main` função usa esse booleano para decidir sobre continuar ou terminar o script.</span><span class="sxs-lookup"><span data-stu-id="d919c-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

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

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="d919c-127">Quando usar uma `throw` declaração</span><span class="sxs-lookup"><span data-stu-id="d919c-127">When to use a `throw` statement</span></span>

<span data-ttu-id="d919c-128">Uma [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) declaração indica que ocorreu um erro inesperado.</span><span class="sxs-lookup"><span data-stu-id="d919c-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="d919c-129">Acaba com o código imediatamente.</span><span class="sxs-lookup"><span data-stu-id="d919c-129">It ends the code immediately.</span></span> <span data-ttu-id="d919c-130">Na maior parte do tempo, você não precisa `throw` do seu roteiro.</span><span class="sxs-lookup"><span data-stu-id="d919c-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="d919c-131">Normalmente, o script informa automaticamente ao usuário que o script não foi executado devido a um problema.</span><span class="sxs-lookup"><span data-stu-id="d919c-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="d919c-132">Na maioria dos casos, é suficiente para terminar o script com uma mensagem de erro e uma `return` instrução da `main` função.</span><span class="sxs-lookup"><span data-stu-id="d919c-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="d919c-133">No entanto, se o seu script estiver sendo executado como parte de um fluxo de Power Automate, você pode querer impedir que o fluxo continue.</span><span class="sxs-lookup"><span data-stu-id="d919c-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="d919c-134">Uma `throw` declaração interrompe o roteiro e diz que o fluxo também pare.</span><span class="sxs-lookup"><span data-stu-id="d919c-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="d919c-135">O script a seguir mostra como usar a `throw` instrução em nosso exemplo de verificação de tabela.</span><span class="sxs-lookup"><span data-stu-id="d919c-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

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

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="d919c-136">Quando usar uma `try...catch` declaração</span><span class="sxs-lookup"><span data-stu-id="d919c-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="d919c-137">A [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrução é uma maneira de detectar se uma chamada de API falha e continuar executando o script.</span><span class="sxs-lookup"><span data-stu-id="d919c-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="d919c-138">Considere o trecho a seguir que realiza uma grande atualização de dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="d919c-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="d919c-139">Se `someLargeValues` for maior do que Excel para a web pode lidar, a `setValues()` chamada falha.</span><span class="sxs-lookup"><span data-stu-id="d919c-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="d919c-140">O script também falha com um [erro de tempo de execução](../testing/troubleshooting.md#runtime-errors).</span><span class="sxs-lookup"><span data-stu-id="d919c-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="d919c-141">A `try...catch` instrução permite que seu script reconheça essa condição, sem encerrar imediatamente o script e mostrar o erro padrão.</span><span class="sxs-lookup"><span data-stu-id="d919c-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="d919c-142">Uma abordagem para dar ao usuário do script uma melhor experiência é apresentar-lhes uma mensagem de erro personalizada.</span><span class="sxs-lookup"><span data-stu-id="d919c-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="d919c-143">O trecho a seguir mostra uma `try...catch` declaração registrando mais informações de erro para ajudar melhor o leitor.</span><span class="sxs-lookup"><span data-stu-id="d919c-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="d919c-144">Outra abordagem para lidar com erros é ter um comportamento de recuo que lida com o caso de erro.</span><span class="sxs-lookup"><span data-stu-id="d919c-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="d919c-145">O trecho a seguir usa o `catch` bloco para tentar um método alternativo dividir a atualização em pedaços menores e evitar o erro.</span><span class="sxs-lookup"><span data-stu-id="d919c-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="d919c-146">Para obter um exemplo completo sobre como atualizar uma grande gama, consulte [Gravar um grande conjunto de dados](../resources/samples/write-large-dataset.md).</span><span class="sxs-lookup"><span data-stu-id="d919c-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

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
> <span data-ttu-id="d919c-147">Usar `try...catch` dentro ou ao redor de um loop retarda seu script.</span><span class="sxs-lookup"><span data-stu-id="d919c-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="d919c-148">Para obter mais informações sobre desempenho, consulte [Evitar o uso de `try...catch` blocos](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span><span class="sxs-lookup"><span data-stu-id="d919c-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="d919c-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="d919c-149">See also</span></span>

- [<span data-ttu-id="d919c-150">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="d919c-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="d919c-151">Solução de problemas para Power Automate com scripts Office</span><span class="sxs-lookup"><span data-stu-id="d919c-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="d919c-152">Limites de plataforma com scripts Office</span><span class="sxs-lookup"><span data-stu-id="d919c-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="d919c-153">Melhore o desempenho de seus scripts de Office</span><span class="sxs-lookup"><span data-stu-id="d919c-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
