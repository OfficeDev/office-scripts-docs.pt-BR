---
title: Melhore o desempenho de seus scripts de Office
description: Crie scripts mais rápidos entendendo a comunicação entre a Excel pasta de trabalho e seu script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544988"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="6a0b8-103">Melhore o desempenho de seus scripts de Office</span><span class="sxs-lookup"><span data-stu-id="6a0b8-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="6a0b8-104">O objetivo do Office Scripts é automatizar séries comumente executadas de tarefas para economizar tempo.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="6a0b8-105">Um script lento pode parecer que não acelera seu fluxo de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="6a0b8-106">Na maioria das vezes, seu roteiro estará perfeitamente bem e será executado como esperado.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="6a0b8-107">No entanto, existem alguns cenários evitáveis que podem afetar o desempenho.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="6a0b8-108">A razão mais comum para um script lento é a comunicação excessiva com a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="6a0b8-109">Seu script é executado em sua máquina local, enquanto a pasta de trabalho existe na nuvem.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="6a0b8-110">Em certos momentos, seu script sincroniza seus dados locais com os da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="6a0b8-111">Isso significa que qualquer operação de gravação (como `workbook.addWorksheet()` ) só é aplicada à pasta de trabalho quando essa sincronização nos bastidores acontece.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="6a0b8-112">Da mesma forma, qualquer operação de leitura (como `myRange.getValues()` ) só recebe dados da pasta de trabalho para o script nesses momentos.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="6a0b8-113">Em ambos os casos, o script busca informações antes de agir sobre os dados.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="6a0b8-114">Por exemplo, o código a seguir registrará com precisão o número de linhas no intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="6a0b8-115">Office As APIs de scripts garantem que quaisquer dados na pasta de trabalho ou script são precisos e atualizados quando necessário.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="6a0b8-116">Você não precisa se preocupar com essas sincronizações para que seu script seja executado corretamente.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="6a0b8-117">No entanto, uma consciência dessa comunicação script-to-cloud pode ajudá-lo a evitar chamadas de rede não fornecidas.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="6a0b8-118">Otimizações de desempenho</span><span class="sxs-lookup"><span data-stu-id="6a0b8-118">Performance optimizations</span></span>

<span data-ttu-id="6a0b8-119">Você pode aplicar técnicas simples para ajudar a reduzir a comunicação à nuvem.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="6a0b8-120">Os seguintes padrões ajudam a acelerar seus scripts.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="6a0b8-121">Leia dados da pasta de trabalho uma vez em vez de repetidamente em um loop.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="6a0b8-122">Remova `console.log` declarações desnecessárias.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="6a0b8-123">Evite usar blocos de tentativa/captura.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="6a0b8-124">Leia dados da pasta de trabalho fora de um loop</span><span class="sxs-lookup"><span data-stu-id="6a0b8-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="6a0b8-125">Qualquer método que obtenha dados da pasta de trabalho pode acionar uma chamada de rede.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="6a0b8-126">Em vez de fazer repetidamente a mesma chamada, você deve salvar os dados localmente sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="6a0b8-127">Isso é especialmente verdade quando se lida com loops.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="6a0b8-128">Considere um script para obter a contagem de números negativos na faixa usada de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="6a0b8-129">O script precisa iterar sobre cada célula da gama usada.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="6a0b8-130">Para isso, precisa do intervalo, do número de linhas e do número de colunas.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="6a0b8-131">Você deve armazená-los como variáveis locais antes de iniciar o loop.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="6a0b8-132">Caso contrário, cada iteração do loop forçará um retorno à pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="6a0b8-133">Como um experimento, tente substituir `usedRangeValues` no loop com `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="6a0b8-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="6a0b8-134">Você pode notar que o script leva consideravelmente mais tempo para ser executado ao lidar com grandes faixas.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="6a0b8-135">Evite usar `try...catch` blocos em loops ou ao redor</span><span class="sxs-lookup"><span data-stu-id="6a0b8-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="6a0b8-136">Não recomendamos o uso [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) de declarações em loops ou loops circundantes.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="6a0b8-137">Isso é pela mesma razão que você deve evitar ler dados em um loop: cada iteração força o script a sincronizar com a pasta de trabalho para garantir que nenhum erro tenha sido jogado.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="6a0b8-138">A maioria dos erros pode ser evitada verificando objetos retornados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="6a0b8-139">Por exemplo, o script a seguir verifica se a tabela devolvida pela pasta de trabalho existe antes de tentar adicionar uma linha.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="6a0b8-140">Remover `console.log` declarações desnecessárias</span><span class="sxs-lookup"><span data-stu-id="6a0b8-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="6a0b8-141">O registro de consoles é uma ferramenta vital para [depurar seus scripts](../testing/troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="6a0b8-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="6a0b8-142">No entanto, ele força o script a sincronizar com a pasta de trabalho para garantir que as informações registradas estão atualizadas.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="6a0b8-143">Considere remover declarações de registro desnecessárias (como as usadas para testes) antes de compartilhar seu script.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="6a0b8-144">Isso normalmente não causará um problema de desempenho perceptível, a menos que a `console.log()` instrução esteja em um loop.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="6a0b8-145">Ajuda caso a caso</span><span class="sxs-lookup"><span data-stu-id="6a0b8-145">Case-by-case help</span></span>

<span data-ttu-id="6a0b8-146">À medida que a plataforma Office Scripts se expande para trabalhar com [Power Automate,](https://flow.microsoft.com/) [Cartões Adaptativos](/adaptive-cards)e outros recursos entre produtos, os detalhes da comunicação script-workbook se tornam mais complexos.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="6a0b8-147">Se você precisar de ajuda para fazer seu script funcionar mais rápido, entre em contato com [o Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span><span class="sxs-lookup"><span data-stu-id="6a0b8-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="6a0b8-148">Certifique-se de marcar sua pergunta com "office-scripts-dev" para que os especialistas possam encontrá-la e ajudar.</span><span class="sxs-lookup"><span data-stu-id="6a0b8-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="6a0b8-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="6a0b8-149">See also</span></span>

- [<span data-ttu-id="6a0b8-150">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="6a0b8-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="6a0b8-151">MDN web docs: Loops e iteração</span><span class="sxs-lookup"><span data-stu-id="6a0b8-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
