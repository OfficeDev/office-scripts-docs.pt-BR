---
title: Restrições de TypeScript em scripts do Office
description: As especificidades do compilador TypeScript e linter usados pelo Editor de Código de Scripts do Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099895"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="66eb0-103">Restrições de TypeScript em scripts do Office</span><span class="sxs-lookup"><span data-stu-id="66eb0-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="66eb0-104">Scripts do Office usam a linguagem TypeScript.</span><span class="sxs-lookup"><span data-stu-id="66eb0-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="66eb0-105">Na maioria das partes, qualquer código TypeScript ou JavaScript funcionará em um Script do Office.</span><span class="sxs-lookup"><span data-stu-id="66eb0-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="66eb0-106">No entanto, há algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e conforme o esperado com a sua planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="66eb0-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="66eb0-107">Nenhum tipo 'any' nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="66eb0-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="66eb0-108">Escrever [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) é opcional no TypeScript, porque os tipos podem ser adiados.</span><span class="sxs-lookup"><span data-stu-id="66eb0-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="66eb0-109">No entanto, o Script do Office exige que uma variável não possa ser [do tipo qualquer.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="66eb0-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="66eb0-110">Tanto explícitas quanto `any` implícitas não são permitidas em um Script do Office.</span><span class="sxs-lookup"><span data-stu-id="66eb0-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="66eb0-111">Esses casos são relatados como erros.</span><span class="sxs-lookup"><span data-stu-id="66eb0-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="66eb0-112">Explícito `any`</span><span class="sxs-lookup"><span data-stu-id="66eb0-112">Explicit `any`</span></span>

<span data-ttu-id="66eb0-113">Você não pode declarar explicitamente uma variável para ser do `any` tipo em Scripts do Office (ou seja, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="66eb0-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="66eb0-114">O `any` tipo causa problemas quando processado pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="66eb0-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="66eb0-115">Por exemplo, `Range` um precisa saber que um valor é um , ou `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="66eb0-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="66eb0-116">Você receberá um erro de tempo de compilação (um erro antes de executar o script) se qualquer variável for explicitamente definida como o tipo `any` no script.</span><span class="sxs-lookup"><span data-stu-id="66eb0-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![A mensagem explícita no texto de foco do editor de código](../images/explicit-any-editor-message.png)

![O erro explícito na janela do console](../images/explicit-any-error-message.png)

<span data-ttu-id="66eb0-119">Na captura de tela acima `[5, 16] Explicit Any is not allowed` indica que a linha #5, coluna #16 define o `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="66eb0-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="66eb0-120">Isso ajuda a localizar o erro.</span><span class="sxs-lookup"><span data-stu-id="66eb0-120">This helps you locate the error.</span></span>

<span data-ttu-id="66eb0-121">Para se livrar desse problema, defina sempre o tipo da variável.</span><span class="sxs-lookup"><span data-stu-id="66eb0-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="66eb0-122">Se você não tiver certeza sobre o tipo de uma variável, poderá usar um tipo [de união.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="66eb0-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="66eb0-123">Isso pode ser útil para variáveis que têm valores, que podem ser do tipo , ou (o tipo de valores é `Range` `string` uma `number` `boolean` `Range` união desses: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="66eb0-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="66eb0-124">Implícito `any`</span><span class="sxs-lookup"><span data-stu-id="66eb0-124">Implicit `any`</span></span>

<span data-ttu-id="66eb0-125">Tipos de variável TypeScript podem ser [implicitamente](( https://www.typescriptlang.org/docs/handbook/type-inference.html) definidos.</span><span class="sxs-lookup"><span data-stu-id="66eb0-125">TypeScript variable types can be [implicitly]((https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="66eb0-126">Se o compilador typeScript não puder determinar o tipo de uma variável (porque o tipo não é definido explicitamente ou a inferência de tipo não é possível), então ele é implícito e você receberá um erro de tempo de `any` compilação.</span><span class="sxs-lookup"><span data-stu-id="66eb0-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="66eb0-127">O caso mais comum em qualquer `any` implícito está em uma declaração de variável, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="66eb0-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="66eb0-128">Há duas maneiras de evitar isso:</span><span class="sxs-lookup"><span data-stu-id="66eb0-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="66eb0-129">Atribuir a variável a um tipo implicitamente identificável ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="66eb0-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="66eb0-130">Digite explicitamente a variável ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="66eb0-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="66eb0-131">Nenhuma interface ou classes de script do Office herdado</span><span class="sxs-lookup"><span data-stu-id="66eb0-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="66eb0-132">Classes e interfaces criadas no script do Office não podem estender ou [implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) classes ou interfaces de Scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="66eb0-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="66eb0-133">Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="66eb0-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="66eb0-134">Funções TypeScript incompatíveis</span><span class="sxs-lookup"><span data-stu-id="66eb0-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="66eb0-135">As APIs de scripts do Office não podem ser usadas no seguinte:</span><span class="sxs-lookup"><span data-stu-id="66eb0-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="66eb0-136">Funções de gerador</span><span class="sxs-lookup"><span data-stu-id="66eb0-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="66eb0-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="66eb0-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="66eb0-138">`eval` não é suportado</span><span class="sxs-lookup"><span data-stu-id="66eb0-138">`eval` is not supported</span></span>

<span data-ttu-id="66eb0-139">A função de [avaliação](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por motivos de segurança.</span><span class="sxs-lookup"><span data-stu-id="66eb0-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="66eb0-140">Identificadors restritos</span><span class="sxs-lookup"><span data-stu-id="66eb0-140">Restricted identifers</span></span>

<span data-ttu-id="66eb0-141">As palavras a seguir não podem ser usadas como identificadores em um script.</span><span class="sxs-lookup"><span data-stu-id="66eb0-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="66eb0-142">Eles são termos reservados.</span><span class="sxs-lookup"><span data-stu-id="66eb0-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="66eb0-143">Avisos de desempenho</span><span class="sxs-lookup"><span data-stu-id="66eb0-143">Performance warnings</span></span>

<span data-ttu-id="66eb0-144">O [linter](https://wikipedia.org/wiki/Lint_(software)) do Editor de Código fornece avisos se o script pode ter problemas de desempenho.</span><span class="sxs-lookup"><span data-stu-id="66eb0-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="66eb0-145">Os casos e como trabalhar em torno deles estão documentados em [Melhorar o desempenho dos scripts do Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="66eb0-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="66eb0-146">Chamadas de API externas</span><span class="sxs-lookup"><span data-stu-id="66eb0-146">External API calls</span></span>

<span data-ttu-id="66eb0-147">Confira [o suporte à chamada da API externa nos Scripts do Office](external-calls.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="66eb0-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="66eb0-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="66eb0-148">See also</span></span>

* [<span data-ttu-id="66eb0-149">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="66eb0-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="66eb0-150">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="66eb0-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
