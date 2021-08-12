---
title: Práticas recomendadas no Scripts do Office
description: Como evitar problemas comuns e gravar scripts robustos Office que podem manipular entradas ou dados inesperados.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: cdea3583120109cda05c05cb7c4f908e929bbff0d37e615b1820f67b57fbe24f
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846614"
---
# <a name="best-practices-in-office-scripts"></a>Práticas recomendadas no Scripts do Office

Esses padrões e práticas são projetados para ajudar seus scripts a executar com êxito sempre. Use-os para evitar armadilhas comuns à medida que você começa a automatizar seu fluxo de trabalho Excel de usuário.

## <a name="verify-an-object-is-present"></a>Verificar se um objeto está presente

Os scripts geralmente dependem de uma determinada planilha ou tabela que está presente na pasta de trabalho. No entanto, eles podem ser renomeados ou removidos entre as executações de script. Verificando se essas tabelas ou planilhas existem antes de chamar métodos neles, você pode garantir que o script não termine abruptamente.

O código de exemplo a seguir verifica se a planilha "Index" está presente na pasta de trabalho. Se a planilha estiver presente, o script obtém um intervalo e continua. Se não estiver presente, o script registra uma mensagem de erro personalizada.

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

O operador TypeScript `?` verifica se o objeto existe antes de chamar um método. Isso pode tornar seu código mais simplificado se você não precisar fazer nada especial quando o objeto não existir.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Validar primeiro os dados e o estado da workbook

Certifique-se de que todas as planilhas, tabelas, formas e outros objetos estão presentes antes de trabalhar nos dados. Usando o padrão anterior, verifique se tudo está na caixa de trabalho e corresponde às suas expectativas. Fazer isso antes que qualquer dado seja gravado garante que o script não deixe a workbook em um estado parcial.

O script a seguir exige que duas tabelas chamadas "Table1" e "Table2" sejam presentes. O script primeiro verifica se as tabelas estão presentes e termina com a instrução e `return` uma mensagem apropriada, se não estiver.

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

Se a verificação estiver ocorrendo em uma função separada, você ainda deve encerrar o script em emissão da `return` instrução da `main` função. Retornar da subfunção não termina o script.

O script a seguir tem o mesmo comportamento do anterior. A diferença é que a `main` função chama a função para verificar `inputPresent` tudo. `inputPresent` retorna um booleano ( `true` ou ) para indicar se todas as entradas necessárias estão `false` presentes. A `main` função usa esse booleano para decidir sobre continuar ou encerrar o script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
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

## <a name="when-to-use-a-throw-statement"></a>Quando usar uma `throw` instrução

Uma [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instrução indica que ocorreu um erro inesperado. Ele termina o código imediatamente. Na maior parte, você não precisa fazer `throw` isso no script. Normalmente, o script informa automaticamente ao usuário que o script não foi executado devido a um problema. Na maioria dos casos, é suficiente terminar o script com uma mensagem de erro e `return` uma instrução da `main` função.

No entanto, se o script estiver sendo executado como parte de um fluxo Power Automate, talvez você queira impedir que o fluxo continue. Uma `throw` instrução interrompe o script e diz ao fluxo para parar também.

O script a seguir mostra como usar a `throw` instrução em nosso exemplo de verificação de tabela.

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

## <a name="when-to-use-a-trycatch-statement"></a>Quando usar uma `try...catch` instrução

A [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrução é uma maneira de detectar se uma chamada de API falhará e continuará executando o script.

Considere o trecho a seguir que executa uma grande atualização de dados em um intervalo.

```TypeScript
range.setValues(someLargeValues);
```

Se `someLargeValues` for maior do que Excel para a Web possa manipular, a chamada `setValues()` falhará. Em seguida, o script também falha com um [erro de tempo de execução.](../testing/troubleshooting.md#runtime-errors) A `try...catch` instrução permite que seu script reconheça essa condição, sem encerrar imediatamente o script e mostrar o erro padrão.

Uma abordagem para dar ao usuário de script uma experiência melhor é apresentar uma mensagem de erro personalizada. O trecho a seguir mostra uma `try...catch` instrução registrando mais informações de erro para ajudar melhor o leitor.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Outra abordagem para lidar com erros é ter comportamento de fallback que lida com o caso de erro. O trecho a seguir usa o bloco para tentar um método alternativo separar a atualização `catch` em partes menores e evitar o erro.

> [!TIP]
> Para ver um exemplo completo sobre como atualizar um intervalo grande, consulte [Write a large dataset](../resources/samples/write-large-dataset.md).

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
> Usar `try...catch` dentro ou ao redor de um loop retarda seu script. Para obter mais informações de desempenho, consulte [Evite usar `try...catch` blocos](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Confira também

- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Solução de problemas de informações para Power Automate com Office Scripts](../testing/power-automate-troubleshooting.md)
- [Limites da plataforma com Office Scripts](../testing/platform-limits.md)
- [Melhorar o desempenho de seus Office Scripts](web-client-performance.md)
