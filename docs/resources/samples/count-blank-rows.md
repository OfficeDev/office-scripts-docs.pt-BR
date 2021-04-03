---
title: Contar linhas em branco em planilhas
description: Saiba como usar scripts do Office para detectar se há linhas em branco em vez de dados em planilhas e, em seguida, relatar a contagem de linhas em branco a ser usada em um fluxo do Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 1f52b9c4d538d5d3e64dc61dae3e27d046b56862
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571074"
---
# <a name="count-blank-rows-on-sheets"></a>Contar linhas em branco em planilhas

Este projeto inclui dois scripts:

* [Contar linhas em branco em uma determinada planilha:](#sample-code-count-blank-rows-on-a-given-sheet)percorre o intervalo usado em uma determinada planilha e retorna uma contagem de linhas em branco.
* [Contar linhas em branco em todas as planilhas](#sample-code-count-blank-rows-on-all-sheets): percorre o intervalo usado em todas as _planilhas_ e retorna uma contagem de linhas em branco.

> [!NOTE]
> Para nosso script, uma linha em branco é qualquer linha onde não há dados. A linha pode ter formatação.

_Esta planilha retorna a contagem de 4 linhas em branco_

![Dados com linhas em branco](../../images/blank-rows.png)

_Esta planilha retorna a contagem de 0 linhas em branco (todas as linhas têm alguns dados)_

![Dados sem linhas em branco](../../images/no-blank-rows.png)

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Código de exemplo: Contar linhas em branco em uma determinada planilha

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Código de exemplo: Contar linhas em branco em todas as planilhas

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a>Usar com o Power Automate

![Captura de tela mostrando como configurar no Power Automate](../../images/use-in-power-automate.png)
