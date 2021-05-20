---
title: Restrições do TypeScript em scripts Office
description: As especificidades do compilador TypeScript e do linter usados pelo Office Scripts Code Editor.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545016"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="d3b12-103">Restrições do TypeScript em scripts Office</span><span class="sxs-lookup"><span data-stu-id="d3b12-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="d3b12-104">Office Os scripts usam a linguagem TypeScript.</span><span class="sxs-lookup"><span data-stu-id="d3b12-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="d3b12-105">Na maioria das vezes, qualquer código TypeScript ou JavaScript funcionará em scripts Office.</span><span class="sxs-lookup"><span data-stu-id="d3b12-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="d3b12-106">No entanto, existem algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e como pretendido com sua Excel livro de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d3b12-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="d3b12-107">Nenhum tipo 'qualquer' em scripts Office</span><span class="sxs-lookup"><span data-stu-id="d3b12-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="d3b12-108">Os [tipos de](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) escrita são opcionais no TypeScript, porque os tipos podem ser inferidos.</span><span class="sxs-lookup"><span data-stu-id="d3b12-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="d3b12-109">No entanto, Office Scripts exige que uma variável não possa ser do [tipo qualquer](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span><span class="sxs-lookup"><span data-stu-id="d3b12-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="d3b12-110">Tanto explícito quanto implícito `any` não são permitidos em Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="d3b12-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="d3b12-111">Esses casos são relatados como erros.</span><span class="sxs-lookup"><span data-stu-id="d3b12-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="d3b12-112">explícito `any`</span><span class="sxs-lookup"><span data-stu-id="d3b12-112">Explicit `any`</span></span>

<span data-ttu-id="d3b12-113">Você não pode declarar explicitamente uma variável como sendo de tipo `any` em scripts Office (isto é, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="d3b12-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="d3b12-114">O `any` tipo causa problemas quando processado por Excel.</span><span class="sxs-lookup"><span data-stu-id="d3b12-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="d3b12-115">Por exemplo, `Range` é preciso saber que um valor é um , ou `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="d3b12-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="d3b12-116">Você receberá um erro de tempo de compilação (um erro antes de executar o script) se qualquer variável for explicitamente definida como o `any` tipo no script.</span><span class="sxs-lookup"><span data-stu-id="d3b12-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A mensagem explícita 'qualquer' no texto do hover do Editor de código":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O erro explícito de 'qualquer' na janela do console":::

<span data-ttu-id="d3b12-119">Na captura de tela anterior `[5, 16] Explicit Any is not allowed` indica que a linha #5, a coluna #16 define o `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="d3b12-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="d3b12-120">Isso ajuda a localizar o erro.</span><span class="sxs-lookup"><span data-stu-id="d3b12-120">This helps you locate the error.</span></span>

<span data-ttu-id="d3b12-121">Para contornar essa questão, defina sempre o tipo da variável.</span><span class="sxs-lookup"><span data-stu-id="d3b12-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="d3b12-122">Se você não tem certeza sobre o tipo de variável, você pode usar um [tipo de união](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span><span class="sxs-lookup"><span data-stu-id="d3b12-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="d3b12-123">Isso pode ser útil para variáveis que possuem `Range` valores, que podem ser do tipo `string` `number` , ou `boolean` (o tipo de `Range` valores é uma união dessas: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="d3b12-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="d3b12-124">implícito `any`</span><span class="sxs-lookup"><span data-stu-id="d3b12-124">Implicit `any`</span></span>

<span data-ttu-id="d3b12-125">Os tipos de variável TypeScript podem ser [implicitamente](https://www.typescriptlang.org/docs/handbook/type-inference.html) definidos.</span><span class="sxs-lookup"><span data-stu-id="d3b12-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="d3b12-126">Se o compilador TypeScript não for capaz de determinar o tipo de variável (porque o tipo não é definido explicitamente ou a inferência do tipo não for possível), então é implícito `any` e você receberá um erro de tempo de compilação.</span><span class="sxs-lookup"><span data-stu-id="d3b12-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="d3b12-127">O caso mais comum em qualquer implícito `any` está em uma declaração variável, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="d3b12-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="d3b12-128">Há duas maneiras de evitar isso:</span><span class="sxs-lookup"><span data-stu-id="d3b12-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="d3b12-129">Atribua a variável a um tipo implicitamente identificável ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="d3b12-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="d3b12-130">Digite explicitamente a variável ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="d3b12-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="d3b12-131">Sem herdar Office classes ou interfaces do Script</span><span class="sxs-lookup"><span data-stu-id="d3b12-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="d3b12-132">Classes e interfaces criadas em seu Office Script não podem [estender ou implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office classes ou interfaces de Scripts.</span><span class="sxs-lookup"><span data-stu-id="d3b12-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="d3b12-133">Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="d3b12-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="d3b12-134">Funções typeScript incompatíveis</span><span class="sxs-lookup"><span data-stu-id="d3b12-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="d3b12-135">Office As APIs de scripts não podem ser usadas no seguinte:</span><span class="sxs-lookup"><span data-stu-id="d3b12-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="d3b12-136">Funções do gerador</span><span class="sxs-lookup"><span data-stu-id="d3b12-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="d3b12-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="d3b12-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="d3b12-138">`eval` não é suportado</span><span class="sxs-lookup"><span data-stu-id="d3b12-138">`eval` is not supported</span></span>

<span data-ttu-id="d3b12-139">A [função eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por razões de segurança.</span><span class="sxs-lookup"><span data-stu-id="d3b12-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="d3b12-140">Identificações restritas</span><span class="sxs-lookup"><span data-stu-id="d3b12-140">Restricted identifers</span></span>

<span data-ttu-id="d3b12-141">As seguintes palavras não podem ser usadas como identificadores em um script.</span><span class="sxs-lookup"><span data-stu-id="d3b12-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="d3b12-142">São termos reservados.</span><span class="sxs-lookup"><span data-stu-id="d3b12-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="d3b12-143">Apenas funções de seta em chamadas de matriz</span><span class="sxs-lookup"><span data-stu-id="d3b12-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="d3b12-144">Seus scripts só podem usar [funções de seta](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) ao fornecer argumentos de retorno de chamada para métodos [Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="d3b12-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="d3b12-145">Você não pode passar qualquer tipo de identificador ou função "tradicional" para esses métodos.</span><span class="sxs-lookup"><span data-stu-id="d3b12-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="d3b12-146">Avisos de desempenho</span><span class="sxs-lookup"><span data-stu-id="d3b12-146">Performance warnings</span></span>

<span data-ttu-id="d3b12-147">O [linter](https://wikipedia.org/wiki/Lint_(software)) do Editor de Códigos dá avisos se o script pode ter problemas de desempenho.</span><span class="sxs-lookup"><span data-stu-id="d3b12-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="d3b12-148">Os casos e como contornar eles são documentados em [Melhorar o desempenho de seus scripts Office](web-client-performance.md).</span><span class="sxs-lookup"><span data-stu-id="d3b12-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="d3b12-149">Chamadas de API externas</span><span class="sxs-lookup"><span data-stu-id="d3b12-149">External API calls</span></span>

<span data-ttu-id="d3b12-150">Consulte [suporte de chamada de API externa em scripts Office](external-calls.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="d3b12-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="d3b12-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="d3b12-151">See also</span></span>

* [<span data-ttu-id="d3b12-152">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="d3b12-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="d3b12-153">Melhore o desempenho de seus scripts de Office</span><span class="sxs-lookup"><span data-stu-id="d3b12-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
