---
title: Restrições typeScript em Office Scripts
description: As especificações do compilador TypeScript e do linter usados pelo editor de código Office Scripts.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ba0dfe60081a0bb65dec4e694c7d534cb8df63
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585678"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restrições typeScript em Office Scripts

Office scripts usam o idioma TypeScript. Na maior parte, qualquer código TypeScript ou JavaScript funcionará em Office Scripts. No entanto, há algumas restrições impostas pelo Editor de Código para garantir que seu script funcione de forma consistente e conforme o pretendido com sua Excel de trabalho.

## <a name="no-any-type-in-office-scripts"></a>Nenhum tipo "qualquer" no Office Scripts

Os [tipos de](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) escrita são opcionais no TypeScript, pois os tipos podem ser inferidos. No entanto, Office Scripts exige que uma variável não possa ser [do tipo qualquer](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Tanto explícito quanto implícito `any` não são permitidos Office Scripts. Esses casos são relatados como erros.

### <a name="explicit-any"></a>Explícito `any`

Não é possível declarar explicitamente que uma variável seja `any` do tipo Office Scripts (ou seja, `let value: any;`). O `any` tipo causa problemas quando processado por Excel. Por exemplo, um precisa `Range` saber que um valor é `string`um , `number`ou `boolean`. Você receberá um erro em tempo de compilação (um erro antes de executar o script) `any` se qualquer variável for explicitamente definida como o tipo no script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A mensagem explícita &quot;qualquer&quot; no texto de foco do Editor de Código.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O erro explícito &quot;qualquer&quot; na janela do console.":::

Na captura de tela anterior, `[2, 14] Explicit Any is not allowed` indica que a linha #2, a coluna #14 define o `any` tipo. Isso ajuda a localizar o erro.

Para se livrar desse problema, sempre defina o tipo da variável. Se você não tiver certeza sobre o tipo de variável, poderá usar um tipo [de união](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). Isso pode ser útil para `Range` variáveis que mantém valores, `string`que podem ser do tipo , `number`ou `boolean` (o `Range` tipo para valores é uma união dessas: `string | number | boolean`).

### <a name="implicit-any"></a>Implícito `any`

Tipos de variável TypeScript podem ser [definidos implicitamente](https://www.typescriptlang.org/docs/handbook/type-inference.html) . Se o compilador TypeScript não conseguir determinar o tipo de uma variável (porque o tipo não é definido explicitamente ou a inferência de tipo não é possível), `any` então é implícito e você receberá um erro de tempo de compilação.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="A mensagem implícita &quot;qualquer&quot; no texto de foco do Editor de Código.":::

O caso mais comum em qualquer implícito `any` está em uma declaração variável, como `let value;`. Há duas maneiras de evitar isso:

* Atribua a variável a um tipo implicitamente identificável (`let value = 5;` ou `let value = workbook.getWorksheet();`).
* Digite explicitamente a variável (`let value: number;`)

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Sem herdar Office ou interfaces de script

Classes e interfaces criadas em seu Office Script não podem estender ou [implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office scripts ou interfaces. Em outras palavras, nada no `ExcelScript` namespace pode ter subclasses ou subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funções TypeScript incompatíveis

Office APIs de scripts não podem ser usadas no seguinte:

* [Funções de gerador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` não tem suporte

A função [de avaliação](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript não é suportada por motivos de segurança.

## <a name="restricted-identifiers"></a>Identificadores restritos

As palavras a seguir não podem ser usadas como identificadores em um script. Eles são termos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Somente funções de seta em retornos de chamada de matriz

Seus scripts só podem usar funções [de seta ao](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) fornecer argumentos de retorno de chamada para [métodos Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) . Não é possível passar qualquer tipo de identificador ou função "tradicional" para esses métodos.

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

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>Não há `ExcelScript` suporte para união de tipos e tipos definidos pelo usuário

Office scripts são convertidos em tempo de execução de blocos de código síncronos para assíncronos. A comunicação com a workbook por meio [de promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) está oculta do criador do script. Essa conversão não dá suporte a tipos [de](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) união que incluem `ExcelScript` tipos e tipos definidos pelo usuário. Nesse caso, o `Promise` é retornado para o script, mas o compilador Office script não espera e o criador do script não pode interagir com `Promise`o .

O exemplo de código a seguir mostra uma união sem suporte entre `ExcelScript.Table` e uma `MyTable` interface personalizada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="constructors-dont-support-office-scripts-apis-and-console-statements"></a>Os construtores não suportam Office scripts e instruções `console`

`console`as instruções e muitas OFFICE scripts exigem sincronização com a Excel de trabalho. Essas sincronizações usam instruções `await` na versão de tempo de execução compilada do script. `await` não é suportado em [construtores](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Classes/constructor). Se você precisar de classes com construtores, evite usar Office APIs de scripts ou `console` instruções nesses blocos de código.

O exemplo de código a seguir demonstra esse cenário. Ele gera um erro que diz `failed to load [code] [library]`.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  class MyClass {
    constructor() {
      // Console statements and Office Scripts APIs aren't supported in constructors.
      console.log("This won't print.");
    }
  }

  let test = new MyClass();
}
```

## <a name="performance-warnings"></a>Avisos de desempenho

O linter do Editor de [Código dá avisos](https://wikipedia.org/wiki/Lint_(software)) se o script pode ter problemas de desempenho. Os casos e como trabalhar ao redor deles são documentados em [Melhorar o desempenho de Office Scripts](web-client-performance.md).

## <a name="external-api-calls"></a>Chamadas de API externas

Consulte [Suporte a chamada de API externa Office Scripts](external-calls.md) para obter mais informações.

## <a name="see-also"></a>Confira também

* [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
* [Melhorar o desempenho de seus Office Scripts](web-client-performance.md)
