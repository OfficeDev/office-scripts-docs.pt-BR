---
title: Restrições typeScript em Office Scripts
description: Os detalhes do compilador TypeScript e linter usados pelo editor de código Office Scripts.
ms.date: 05/24/2021
localization_priority: Normal
ms.openlocfilehash: 449a8abbcfdcfde53d0c9b96106f73259de368b1
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639852"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="6b8a5-103">Restrições typeScript em Office Scripts</span><span class="sxs-lookup"><span data-stu-id="6b8a5-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="6b8a5-104">Office Os scripts usam o idioma TypeScript.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="6b8a5-105">Na maioria das partes, qualquer código TypeScript ou JavaScript funcionará em Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="6b8a5-106">No entanto, há algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e conforme o pretendido com sua Excel de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="6b8a5-107">Nenhum tipo "qualquer" no Office Scripts</span><span class="sxs-lookup"><span data-stu-id="6b8a5-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="6b8a5-108">Os [tipos de](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) escrita são opcionais no TypeScript, pois os tipos podem ser inferidos.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="6b8a5-109">No entanto, Office scripts exige que uma variável não possa ser [do tipo qualquer](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span><span class="sxs-lookup"><span data-stu-id="6b8a5-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="6b8a5-110">Tanto explícito quanto `any` implícito não são permitidos Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="6b8a5-111">Esses casos são relatados como erros.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="6b8a5-112">Explícito `any`</span><span class="sxs-lookup"><span data-stu-id="6b8a5-112">Explicit `any`</span></span>

<span data-ttu-id="6b8a5-113">Não é possível declarar explicitamente que uma variável seja do tipo `any` Office Scripts (ou seja, `let value: any;` ).</span><span class="sxs-lookup"><span data-stu-id="6b8a5-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let value: any;`).</span></span> <span data-ttu-id="6b8a5-114">O `any` tipo causa problemas quando processado por Excel.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="6b8a5-115">Por exemplo, um `Range` precisa saber que um valor é um , ou `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="6b8a5-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="6b8a5-116">Você receberá um erro em tempo de compilação (um erro antes de executar o script) se qualquer variável for explicitamente definida como o `any` tipo no script.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A mensagem explícita &quot;qualquer&quot; no texto de foco do Editor de Código":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O erro explícito &quot;qualquer&quot; na janela do console":::

<span data-ttu-id="6b8a5-119">Na captura de tela anterior, indica que a linha #2, a coluna `[2, 14] Explicit Any is not allowed` #14 define o `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-119">In the previous screenshot, `[2, 14] Explicit Any is not allowed` indicates that line #2, column #14 defines `any` type.</span></span> <span data-ttu-id="6b8a5-120">Isso ajuda a localizar o erro.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-120">This helps you locate the error.</span></span>

<span data-ttu-id="6b8a5-121">Para se livrar desse problema, sempre defina o tipo da variável.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="6b8a5-122">Se você não tiver certeza sobre o tipo de uma variável, poderá usar um tipo [de união](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span><span class="sxs-lookup"><span data-stu-id="6b8a5-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="6b8a5-123">Isso pode ser útil para variáveis que mantém valores, que podem ser do tipo , ou (o tipo para valores é `Range` `string` uma `number` `boolean` `Range` união dessas: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="6b8a5-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="6b8a5-124">Implícito `any`</span><span class="sxs-lookup"><span data-stu-id="6b8a5-124">Implicit `any`</span></span>

<span data-ttu-id="6b8a5-125">Tipos de variável TypeScript podem ser [definidos implicitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="6b8a5-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="6b8a5-126">Se o compilador TypeScript não conseguir determinar o tipo de uma variável (porque o tipo não é definido explicitamente ou a inferência de tipo não é possível), então é implícito e você receberá um erro de tempo de `any` compilação.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="A mensagem implícita &quot;qualquer&quot; no texto de foco do Editor de Código":::

<span data-ttu-id="6b8a5-128">O caso mais comum em qualquer `any` implícito está em uma declaração variável, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="6b8a5-128">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="6b8a5-129">Há duas maneiras de evitar isso:</span><span class="sxs-lookup"><span data-stu-id="6b8a5-129">There are two ways to avoid this:</span></span>

* <span data-ttu-id="6b8a5-130">Atribua a variável a um tipo implicitamente identificável ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="6b8a5-130">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="6b8a5-131">Digite explicitamente a variável ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="6b8a5-131">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="6b8a5-132">Sem herdar Office ou interfaces de script</span><span class="sxs-lookup"><span data-stu-id="6b8a5-132">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="6b8a5-133">Classes e interfaces criadas em seu Office Script não podem estender ou [implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office scripts ou interfaces.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-133">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="6b8a5-134">Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-134">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="6b8a5-135">Funções TypeScript incompatíveis</span><span class="sxs-lookup"><span data-stu-id="6b8a5-135">Incompatible TypeScript functions</span></span>

<span data-ttu-id="6b8a5-136">Office As APIs de scripts não podem ser usadas no seguinte:</span><span class="sxs-lookup"><span data-stu-id="6b8a5-136">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="6b8a5-137">Funções de gerador</span><span class="sxs-lookup"><span data-stu-id="6b8a5-137">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="6b8a5-138">Array.sort</span><span class="sxs-lookup"><span data-stu-id="6b8a5-138">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="6b8a5-139">`eval` não tem suporte</span><span class="sxs-lookup"><span data-stu-id="6b8a5-139">`eval` is not supported</span></span>

<span data-ttu-id="6b8a5-140">A função [de avaliação](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por motivos de segurança.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-140">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="6b8a5-141">Identifers restritos</span><span class="sxs-lookup"><span data-stu-id="6b8a5-141">Restricted identifers</span></span>

<span data-ttu-id="6b8a5-142">As palavras a seguir não podem ser usadas como identificadores em um script.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-142">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="6b8a5-143">Eles são termos reservados.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-143">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="6b8a5-144">Somente funções de seta em retornos de chamada de matriz</span><span class="sxs-lookup"><span data-stu-id="6b8a5-144">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="6b8a5-145">Seus scripts só podem usar funções [de seta ao](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) fornecer argumentos de retorno de chamada para [métodos Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="6b8a5-145">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="6b8a5-146">Não é possível passar qualquer tipo de identificador ou função "tradicional" para esses métodos.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-146">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="performance-warnings"></a><span data-ttu-id="6b8a5-147">Avisos de desempenho</span><span class="sxs-lookup"><span data-stu-id="6b8a5-147">Performance warnings</span></span>

<span data-ttu-id="6b8a5-148">O linter do Editor de Código dá [avisos](https://wikipedia.org/wiki/Lint_(software)) se o script pode ter problemas de desempenho.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-148">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="6b8a5-149">Os casos e como trabalhar ao redor deles são documentados em Melhorar o desempenho do seu [Office Scripts](web-client-performance.md).</span><span class="sxs-lookup"><span data-stu-id="6b8a5-149">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="6b8a5-150">Chamadas de API externas</span><span class="sxs-lookup"><span data-stu-id="6b8a5-150">External API calls</span></span>

<span data-ttu-id="6b8a5-151">Consulte [Suporte a chamada de API externa Office Scripts](external-calls.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="6b8a5-151">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="6b8a5-152">Confira também</span><span class="sxs-lookup"><span data-stu-id="6b8a5-152">See also</span></span>

* [<span data-ttu-id="6b8a5-153">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="6b8a5-153">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="6b8a5-154">Melhorar o desempenho de seus Office Scripts</span><span class="sxs-lookup"><span data-stu-id="6b8a5-154">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
