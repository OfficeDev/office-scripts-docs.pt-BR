---
title: Melhorar o desempenho dos scripts do Office
description: Crie scripts mais rápidos compreendendo a comunicação entre a planilha do Excel e seu script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867867"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="ddfa0-103">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="ddfa0-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="ddfa0-104">O objetivo dos Scripts do Office é automatizar uma série de tarefas normalmente realizadas para economizar tempo.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="ddfa0-105">Um script lento pode parecer que ele não acelera seu fluxo de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="ddfa0-106">Na maioria das vezes, seu script ficará perfeitamente bem e será executado conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="ddfa0-107">No entanto, há alguns cenários que podem afetar o desempenho.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="ddfa0-108">O motivo mais comum para um script lento é a comunicação excessiva com a agenda.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="ddfa0-109">O script é executado no computador local, enquanto a agenda existe na nuvem.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="ddfa0-110">Em determinados momentos, seu script sincroniza seus dados locais com os da agenda.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="ddfa0-111">Isso significa que todas as operações de gravação (como) serão aplicadas somente à plano de trabalho quando essa sincronização nos `workbook.addWorksheet()` bastidores ocorrer.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="ddfa0-112">Da mesma forma, qualquer operação de leitura (como) só obter dados da área de trabalho `myRange.getValues()` para o script nesses momentos.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="ddfa0-113">Em ambos os casos, o script busca informações antes de agir sobre os dados.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="ddfa0-114">Por exemplo, o código a seguir registrará com precisão o número de linhas no intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="ddfa0-115">As APIs de scripts do Office garantem que todos os dados da lista de trabalho ou script sejam precisos e atualizados quando necessário.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="ddfa0-116">Você não precisa se preocupar com essas sincronizações para que seu script seja executado corretamente.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="ddfa0-117">No entanto, um reconhecimento dessa comunicação entre scripts e nuvem pode ajudá-lo a evitar chamadas de rede não precisas.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="ddfa0-118">Otimizações de desempenho</span><span class="sxs-lookup"><span data-stu-id="ddfa0-118">Performance optimizations</span></span>

<span data-ttu-id="ddfa0-119">Você pode aplicar técnicas simples para ajudar a reduzir a comunicação com a nuvem.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="ddfa0-120">Os seguintes padrões ajudam a acelerar seus scripts.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="ddfa0-121">Ler dados de uma vez em vez de repetidamente em um loop.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="ddfa0-122">Remova instruções `console.log` desnecessárias.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="ddfa0-123">Evite usar blocos try/catch.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="ddfa0-124">Ler dados da área de trabalho fora de um loop</span><span class="sxs-lookup"><span data-stu-id="ddfa0-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="ddfa0-125">Qualquer método que obtém dados da agenda pode disparar uma chamada de rede.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="ddfa0-126">Em vez de fazer repetidamente a mesma chamada, você deve salvar dados localmente sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="ddfa0-127">Isso é especialmente verdadeiro ao lidar com loops.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="ddfa0-128">Considere um script para obter a contagem de números negativos no intervalo usado de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="ddfa0-129">O script precisa iterar em todas as células no intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="ddfa0-130">Para fazer isso, ele precisa do intervalo, do número de linhas e do número de colunas.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="ddfa0-131">Você deve armazená-los como variáveis locais antes de iniciar o loop.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="ddfa0-132">Caso contrário, cada iteração do loop força um retorno à agenda.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="ddfa0-133">Como um experimento, tente substituir `usedRangeValues` no loop por `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="ddfa0-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="ddfa0-134">Você pode notar que o script leva consideravelmente mais tempo para ser executado ao lidar com intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="ddfa0-135">Remover instruções `console.log` desnecessárias</span><span class="sxs-lookup"><span data-stu-id="ddfa0-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="ddfa0-136">O log do console é uma ferramenta vital [para depurar seus scripts.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="ddfa0-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="ddfa0-137">No entanto, ele força o script a sincronizar com a agenda para garantir que as informações registradas estejam atualizadas.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="ddfa0-138">Considere remover instruções de registro em log desnecessárias (como aquelas usadas para teste) antes de compartilhar seu script.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="ddfa0-139">Isso normalmente não causará um problema de desempenho perceptível, a menos que `console.log()` a instrução esteja em um loop.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="ddfa0-140">Evite usar blocos try/catch</span><span class="sxs-lookup"><span data-stu-id="ddfa0-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="ddfa0-141">Não recomendamos o uso de [ `try` / `catch` blocos](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) como parte do fluxo de controle esperado de um script.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="ddfa0-142">A maioria dos erros pode ser evitada verificando objetos retornados da agenda.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="ddfa0-143">Por exemplo, o script a seguir verifica se a tabela retornada pela lista de trabalho existe antes de tentar adicionar uma linha.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a><span data-ttu-id="ddfa0-144">Ajuda caso a caso</span><span class="sxs-lookup"><span data-stu-id="ddfa0-144">Case-by-case help</span></span>

<span data-ttu-id="ddfa0-145">À medida que a plataforma de Scripts do Office se expande para trabalhar com o [Power Automate](https://flow.microsoft.com/), Cartões [Adaptáveis](/adaptive-cards)e outros recursos entre produtos, os detalhes da comunicação entre as guias de trabalho de script se tornam mais complexos.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="ddfa0-146">Se precisar de ajuda para fazer seu script ser executado mais rapidamente, entre em contato com o [Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts)</span><span class="sxs-lookup"><span data-stu-id="ddfa0-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="ddfa0-147">Certifique-se de marcar sua pergunta com "office-scripts" para que os especialistas possam encontrá-la e ajudar.</span><span class="sxs-lookup"><span data-stu-id="ddfa0-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="ddfa0-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="ddfa0-148">See also</span></span>

- [<span data-ttu-id="ddfa0-149">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="ddfa0-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="ddfa0-150">Documentos da Web do MDN: Loops e iteração</span><span class="sxs-lookup"><span data-stu-id="ddfa0-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)