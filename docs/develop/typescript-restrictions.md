---
title: Restrições TypeScript em Scripts do Office
description: Os detalhes do compilador TypeScript e linter usados pelo Editor de Código de Scripts do Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 88d0b5873a2f7350f88417d2e340343dbd183606
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755046"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="7c78f-103">Restrições TypeScript em Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7c78f-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="7c78f-104">Scripts do Office usam o idioma TypeScript.</span><span class="sxs-lookup"><span data-stu-id="7c78f-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="7c78f-105">Na maior parte, qualquer código TypeScript ou JavaScript funcionará em um Script do Office.</span><span class="sxs-lookup"><span data-stu-id="7c78f-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="7c78f-106">No entanto, há algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e conforme pretendido com sua planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="7c78f-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="7c78f-107">Nenhum tipo "qualquer" em Scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7c78f-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="7c78f-108">Os [tipos de](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) escrita são opcionais no TypeScript, pois os tipos podem ser inferidos.</span><span class="sxs-lookup"><span data-stu-id="7c78f-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="7c78f-109">No entanto, o Script do Office exige que uma variável não possa ser [do tipo qualquer](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span><span class="sxs-lookup"><span data-stu-id="7c78f-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="7c78f-110">Tanto explícito quanto `any` implícito não são permitidos em um Script do Office.</span><span class="sxs-lookup"><span data-stu-id="7c78f-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="7c78f-111">Esses casos são relatados como erros.</span><span class="sxs-lookup"><span data-stu-id="7c78f-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="7c78f-112">Explícito `any`</span><span class="sxs-lookup"><span data-stu-id="7c78f-112">Explicit `any`</span></span>

<span data-ttu-id="7c78f-113">Não é possível declarar explicitamente que uma variável seja do tipo em Scripts do `any` Office (ou seja, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="7c78f-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="7c78f-114">O `any` tipo causa problemas quando processado pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="7c78f-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="7c78f-115">Por exemplo, um `Range` precisa saber que um valor é um , ou `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="7c78f-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="7c78f-116">Você receberá um erro em tempo de compilação (um erro antes de executar o script) se qualquer variável for explicitamente definida como o `any` tipo no script.</span><span class="sxs-lookup"><span data-stu-id="7c78f-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A mensagem explícita &quot;qualquer&quot; no texto de foco do editor de código":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O erro Explícito Qualquer na janela do console.":::

<span data-ttu-id="7c78f-119">Na captura de tela anterior indica que a linha #5, a coluna `[5, 16] Explicit Any is not allowed` #16 define o `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="7c78f-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="7c78f-120">Isso ajuda a localizar o erro.</span><span class="sxs-lookup"><span data-stu-id="7c78f-120">This helps you locate the error.</span></span>

<span data-ttu-id="7c78f-121">Para se livrar desse problema, sempre defina o tipo da variável.</span><span class="sxs-lookup"><span data-stu-id="7c78f-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="7c78f-122">Se você não tiver certeza sobre o tipo de uma variável, poderá usar um tipo [de união](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span><span class="sxs-lookup"><span data-stu-id="7c78f-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="7c78f-123">Isso pode ser útil para variáveis que mantém valores, que podem ser do tipo , ou (o tipo para valores é `Range` `string` uma `number` `boolean` `Range` união dessas: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="7c78f-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="7c78f-124">Implícito `any`</span><span class="sxs-lookup"><span data-stu-id="7c78f-124">Implicit `any`</span></span>

<span data-ttu-id="7c78f-125">Tipos de variável TypeScript podem ser [definidos implicitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="7c78f-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="7c78f-126">Se o compilador TypeScript não conseguir determinar o tipo de uma variável (porque o tipo não é definido explicitamente ou a inferência de tipo não é possível), então é implícito e você receberá um erro de tempo de `any` compilação.</span><span class="sxs-lookup"><span data-stu-id="7c78f-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="7c78f-127">O caso mais comum em qualquer `any` implícito está em uma declaração variável, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="7c78f-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="7c78f-128">Há duas maneiras de evitar isso:</span><span class="sxs-lookup"><span data-stu-id="7c78f-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="7c78f-129">Atribua a variável a um tipo implicitamente identificável ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="7c78f-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="7c78f-130">Digite explicitamente a variável ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="7c78f-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="7c78f-131">Sem herdar classes ou interfaces do Office Script</span><span class="sxs-lookup"><span data-stu-id="7c78f-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="7c78f-132">Classes e interfaces criadas em seu Script do Office não podem estender ou [implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) classes ou interfaces do Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="7c78f-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="7c78f-133">Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="7c78f-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="7c78f-134">Funções TypeScript incompatíveis</span><span class="sxs-lookup"><span data-stu-id="7c78f-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="7c78f-135">As APIs de Scripts do Office não podem ser usadas no seguinte:</span><span class="sxs-lookup"><span data-stu-id="7c78f-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="7c78f-136">Funções de gerador</span><span class="sxs-lookup"><span data-stu-id="7c78f-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="7c78f-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="7c78f-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="7c78f-138">`eval` não tem suporte</span><span class="sxs-lookup"><span data-stu-id="7c78f-138">`eval` is not supported</span></span>

<span data-ttu-id="7c78f-139">A função [de avaliação](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por motivos de segurança.</span><span class="sxs-lookup"><span data-stu-id="7c78f-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="7c78f-140">Identifers restritos</span><span class="sxs-lookup"><span data-stu-id="7c78f-140">Restricted identifers</span></span>

<span data-ttu-id="7c78f-141">As palavras a seguir não podem ser usadas como identificadores em um script.</span><span class="sxs-lookup"><span data-stu-id="7c78f-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="7c78f-142">Eles são termos reservados.</span><span class="sxs-lookup"><span data-stu-id="7c78f-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="7c78f-143">Somente funções de seta em retornos de chamada de matriz</span><span class="sxs-lookup"><span data-stu-id="7c78f-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="7c78f-144">Seus scripts só podem usar funções [de seta ao](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) fornecer argumentos de retorno de chamada para [métodos Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="7c78f-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="7c78f-145">Não é possível passar qualquer tipo de identificador ou função "tradicional" para esses métodos.</span><span class="sxs-lookup"><span data-stu-id="7c78f-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="7c78f-146">Avisos de desempenho</span><span class="sxs-lookup"><span data-stu-id="7c78f-146">Performance warnings</span></span>

<span data-ttu-id="7c78f-147">O linter do Editor de Código dá [avisos](https://wikipedia.org/wiki/Lint_(software)) se o script pode ter problemas de desempenho.</span><span class="sxs-lookup"><span data-stu-id="7c78f-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="7c78f-148">Os casos e como trabalhar ao redor deles estão documentados em [Melhorar o desempenho dos scripts do Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="7c78f-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="7c78f-149">Chamadas de API externas</span><span class="sxs-lookup"><span data-stu-id="7c78f-149">External API calls</span></span>

<span data-ttu-id="7c78f-150">Confira [Suporte a chamada de API externa em Scripts do Office](external-calls.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="7c78f-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="7c78f-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="7c78f-151">See also</span></span>

* [<span data-ttu-id="7c78f-152">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="7c78f-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="7c78f-153">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="7c78f-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
