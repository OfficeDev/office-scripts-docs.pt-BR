---
title: Restrições typeScript em Office Scripts
description: Os detalhes do compilador TypeScript e linter usados pelo editor de código Office Scripts.
ms.date: 05/24/2021
localization_priority: Normal
ms.openlocfilehash: 0bc6b4c0acaf9bb42f8200a0850dd7254632f965
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074442"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restrições typeScript em Office Scripts

Office Os scripts usam o idioma TypeScript. Na maioria das partes, qualquer código TypeScript ou JavaScript funcionará em Office Scripts. No entanto, há algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e conforme o pretendido com sua Excel de trabalho.

## <a name="no-any-type-in-office-scripts"></a>Nenhum tipo "qualquer" no Office Scripts

Os [tipos de](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) escrita são opcionais no TypeScript, pois os tipos podem ser inferidos. No entanto, Office scripts exige que uma variável não possa ser [do tipo qualquer](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Tanto explícito quanto `any` implícito não são permitidos Office Scripts. Esses casos são relatados como erros.

### <a name="explicit-any"></a>Explícito `any`

Não é possível declarar explicitamente que uma variável seja do tipo `any` Office Scripts (ou seja, `let value: any;` ). O `any` tipo causa problemas quando processado por Excel. Por exemplo, um `Range` precisa saber que um valor é um , ou `string` `number` `boolean` . Você receberá um erro em tempo de compilação (um erro antes de executar o script) se qualquer variável for explicitamente definida como o `any` tipo no script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A mensagem explícita &quot;qualquer&quot; no texto de foco do Editor de Código.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O erro explícito &quot;qualquer&quot; na janela do console.":::

Na captura de tela anterior, indica que a linha #2, a coluna `[2, 14] Explicit Any is not allowed` #14 define o `any` tipo. Isso ajuda a localizar o erro.

Para se livrar desse problema, sempre defina o tipo da variável. Se você não tiver certeza sobre o tipo de uma variável, poderá usar um tipo [de união](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). Isso pode ser útil para variáveis que mantém valores, que podem ser do tipo , ou (o tipo para valores é `Range` `string` uma `number` `boolean` `Range` união dessas: `string | number | boolean` ).

### <a name="implicit-any"></a>Implícito `any`

Tipos de variável TypeScript podem ser [definidos implicitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html) Se o compilador TypeScript não conseguir determinar o tipo de uma variável (porque o tipo não é definido explicitamente ou a inferência de tipo não é possível), então é implícito e você receberá um erro de tempo de `any` compilação.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="A mensagem implícita &quot;qualquer&quot; no texto de foco do Editor de Código.":::

O caso mais comum em qualquer `any` implícito está em uma declaração variável, como `let value;` . Há duas maneiras de evitar isso:

* Atribua a variável a um tipo implicitamente identificável ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).
* Digite explicitamente a variável ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Sem herdar Office ou interfaces de script

Classes e interfaces criadas em seu Office Script não podem estender ou [implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office scripts ou interfaces. Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funções TypeScript incompatíveis

Office As APIs de scripts não podem ser usadas no seguinte:

* [Funções de gerador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` não tem suporte

A função [de avaliação](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por motivos de segurança.

## <a name="restricted-identifers"></a>Identifers restritos

As palavras a seguir não podem ser usadas como identificadores em um script. Eles são termos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Somente funções de seta em retornos de chamada de matriz

Seus scripts só podem usar funções [de seta ao](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) fornecer argumentos de retorno de chamada para [métodos Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Não é possível passar qualquer tipo de identificador ou função "tradicional" para esses métodos.

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

O linter do Editor de Código dá [avisos](https://wikipedia.org/wiki/Lint_(software)) se o script pode ter problemas de desempenho. Os casos e como trabalhar ao redor deles são documentados em Melhorar o desempenho do seu [Office Scripts](web-client-performance.md).

## <a name="external-api-calls"></a>Chamadas de API externas

Consulte [Suporte a chamada de API externa Office Scripts](external-calls.md) para obter mais informações.

## <a name="see-also"></a>Confira também

* [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
* [Melhorar o desempenho de seus Office Scripts](web-client-performance.md)
