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
# <a name="typescript-restrictions-in-office-scripts"></a>Restrições do TypeScript em scripts Office

Office Os scripts usam a linguagem TypeScript. Na maioria das vezes, qualquer código TypeScript ou JavaScript funcionará em scripts Office. No entanto, existem algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e como pretendido com sua Excel livro de trabalho.

## <a name="no-any-type-in-office-scripts"></a>Nenhum tipo 'qualquer' em scripts Office

Os [tipos de](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) escrita são opcionais no TypeScript, porque os tipos podem ser inferidos. No entanto, Office Scripts exige que uma variável não possa ser do [tipo qualquer](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Tanto explícito quanto implícito `any` não são permitidos em Office Scripts. Esses casos são relatados como erros.

### <a name="explicit-any"></a>explícito `any`

Você não pode declarar explicitamente uma variável como sendo de tipo `any` em scripts Office (isto é, `let someVariable: any;` ). O `any` tipo causa problemas quando processado por Excel. Por exemplo, `Range` é preciso saber que um valor é um , ou `string` `number` `boolean` . Você receberá um erro de tempo de compilação (um erro antes de executar o script) se qualquer variável for explicitamente definida como o `any` tipo no script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A mensagem explícita 'qualquer' no texto do hover do Editor de código":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O erro explícito de 'qualquer' na janela do console":::

Na captura de tela anterior `[5, 16] Explicit Any is not allowed` indica que a linha #5, a coluna #16 define o `any` tipo. Isso ajuda a localizar o erro.

Para contornar essa questão, defina sempre o tipo da variável. Se você não tem certeza sobre o tipo de variável, você pode usar um [tipo de união](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). Isso pode ser útil para variáveis que possuem `Range` valores, que podem ser do tipo `string` `number` , ou `boolean` (o tipo de `Range` valores é uma união dessas: `string | number | boolean` ).

### <a name="implicit-any"></a>implícito `any`

Os tipos de variável TypeScript podem ser [implicitamente](https://www.typescriptlang.org/docs/handbook/type-inference.html) definidos. Se o compilador TypeScript não for capaz de determinar o tipo de variável (porque o tipo não é definido explicitamente ou a inferência do tipo não for possível), então é implícito `any` e você receberá um erro de tempo de compilação.

O caso mais comum em qualquer implícito `any` está em uma declaração variável, como `let value;` . Há duas maneiras de evitar isso:

* Atribua a variável a um tipo implicitamente identificável ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).
* Digite explicitamente a variável ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Sem herdar Office classes ou interfaces do Script

Classes e interfaces criadas em seu Office Script não podem [estender ou implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office classes ou interfaces de Scripts. Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funções typeScript incompatíveis

Office As APIs de scripts não podem ser usadas no seguinte:

* [Funções do gerador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` não é suportado

A [função eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por razões de segurança.

## <a name="restricted-identifers"></a>Identificações restritas

As seguintes palavras não podem ser usadas como identificadores em um script. São termos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Apenas funções de seta em chamadas de matriz

Seus scripts só podem usar [funções de seta](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) ao fornecer argumentos de retorno de chamada para métodos [Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Você não pode passar qualquer tipo de identificador ou função "tradicional" para esses métodos.

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

## <a name="performance-warnings"></a>Avisos de desempenho

O [linter](https://wikipedia.org/wiki/Lint_(software)) do Editor de Códigos dá avisos se o script pode ter problemas de desempenho. Os casos e como contornar eles são documentados em [Melhorar o desempenho de seus scripts Office](web-client-performance.md).

## <a name="external-api-calls"></a>Chamadas de API externas

Consulte [suporte de chamada de API externa em scripts Office](external-calls.md) para obter mais informações.

## <a name="see-also"></a>Confira também

* [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
* [Melhore o desempenho de seus scripts de Office](web-client-performance.md)
