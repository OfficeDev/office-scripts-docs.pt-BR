---
title: Práticas recomendadas no Scripts do Office
description: Como prevenir problemas comuns e escrever scripts robustos Office que podem lidar com entradas ou dados inesperados.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546018"
---
# <a name="best-practices-in-office-scripts"></a>Práticas recomendadas no Scripts do Office

Esses padrões e práticas são projetados para ajudar seus scripts a serem executados com sucesso todas as vezes. Use-os para evitar armadilhas comuns à medida que você começa a automatizar seu fluxo de trabalho Excel.

## <a name="verify-an-object-is-present"></a>Verifique se um objeto está presente

Os scripts muitas vezes dependem de uma determinada planilha ou tabela presente na pasta de trabalho. No entanto, eles podem ser renomeados ou removidos entre as corridas de script. Ao verificar se essas tabelas ou planilhas existem antes de chamar métodos sobre elas, você pode garantir que o script não termine abruptamente.

O código de amostra a seguir verifica se a planilha "Índice" está presente na pasta de trabalho. Se a planilha estiver presente, o script ganha uma faixa e prossegue. Se ele não estiver presente, o script registra uma mensagem de erro personalizada.

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

O operador TypeScript `?` verifica se o objeto existe antes de chamar um método. Isso pode tornar seu código mais simplificado se você não precisar fazer nada especial quando o objeto não existe.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Validar dados e estado da pasta de trabalho primeiro

Certifique-se de que todas as planilhas, tabelas, formas e outros objetos estejam presentes antes de trabalhar nos dados. Usando o padrão anterior, verifique se está tudo na pasta de trabalho e corresponda às suas expectativas. Fazer isso antes que qualquer dado seja escrito garante que seu script não deixe a pasta de trabalho em um estado parcial.

O roteiro a seguir requer que duas tabelas chamadas "Tabela1" e "Tabela2" estejam presentes. O script primeiro verifica se as tabelas estão presentes e, em seguida, termina com a `return` instrução e uma mensagem apropriada se não estiverem.

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

Se a verificação estiver acontecendo em uma função separada, você ainda deve terminar o script emitindo a `return` instrução da `main` função. Voltar da subfunção não encerra o roteiro.

O roteiro a seguir tem o mesmo comportamento do anterior. A diferença é que a `main` função chama a função para verificar `inputPresent` tudo. `inputPresent` retorna um booleano ( `true` ou ) para indicar se todas as `false` entradas necessárias estão presentes. A `main` função usa esse booleano para decidir sobre continuar ou terminar o script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
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

## <a name="when-to-use-a-throw-statement"></a>Quando usar uma `throw` declaração

Uma [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) declaração indica que ocorreu um erro inesperado. Acaba com o código imediatamente. Na maior parte do tempo, você não precisa `throw` do seu roteiro. Normalmente, o script informa automaticamente ao usuário que o script não foi executado devido a um problema. Na maioria dos casos, é suficiente para terminar o script com uma mensagem de erro e uma `return` instrução da `main` função.

No entanto, se o seu script estiver sendo executado como parte de um fluxo de Power Automate, você pode querer impedir que o fluxo continue. Uma `throw` declaração interrompe o roteiro e diz que o fluxo também pare.

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

## <a name="when-to-use-a-trycatch-statement"></a>Quando usar uma `try...catch` declaração

A [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrução é uma maneira de detectar se uma chamada de API falha e continuar executando o script.

Considere o trecho a seguir que realiza uma grande atualização de dados em um intervalo.

```TypeScript
range.setValues(someLargeValues);
```

Se `someLargeValues` for maior do que Excel para a web pode lidar, a `setValues()` chamada falha. O script também falha com um [erro de tempo de execução](../testing/troubleshooting.md#runtime-errors). A `try...catch` instrução permite que seu script reconheça essa condição, sem encerrar imediatamente o script e mostrar o erro padrão.

Uma abordagem para dar ao usuário do script uma melhor experiência é apresentar-lhes uma mensagem de erro personalizada. O trecho a seguir mostra uma `try...catch` declaração registrando mais informações de erro para ajudar melhor o leitor.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Outra abordagem para lidar com erros é ter um comportamento de recuo que lida com o caso de erro. O trecho a seguir usa o `catch` bloco para tentar um método alternativo dividir a atualização em pedaços menores e evitar o erro.

> [!TIP]
> Para obter um exemplo completo sobre como atualizar uma grande gama, consulte [Gravar um grande conjunto de dados](../resources/samples/write-large-dataset.md).

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
> Usar `try...catch` dentro ou ao redor de um loop retarda seu script. Para obter mais informações sobre desempenho, consulte [Evitar o uso de `try...catch` blocos](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Confira também

- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Solução de problemas para Power Automate com scripts Office](../testing/power-automate-troubleshooting.md)
- [Limites de plataforma com scripts Office](../testing/platform-limits.md)
- [Melhore o desempenho de seus scripts de Office](web-client-performance.md)
