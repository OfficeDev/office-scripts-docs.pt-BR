---
title: Noções básicas de intervalo em Scripts do Office
description: Saiba noções básicas sobre como usar o objeto Range em Scripts do Office.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571042"
---
# <a name="range-basics"></a>Noções básicas de intervalo

`Range` é o objeto base no modelo de objeto do Excel scripts do Office. [As APIs de](/javascript/api/office-scripts/excelscript/excelscript.range) intervalo permitem acesso aos dados e ao formato disponível na grade e vinculam outros objetos-chave no Excel, como planilhas, tabelas, gráficos, etc.

Um intervalo é identificado usando seu endereço como "A1:B4" ou usando um item nomeado, que é uma chave nomeada para um determinado conjunto de células. No modelo de objeto do Excel, uma célula e um grupo de células são chamados de _intervalo_. `Range` pode conter atributos de nível de célula, como dados em uma célula e também atributos de nível de célula e células, como formato, bordas, etc.

`Range` também pode ser obtido por meio da seleção do usuário que consiste em pelo menos uma célula. À medida que você interage com o intervalo, é importante manter essas relações de célula e intervalo claras.

A seguir estão o conjunto principal de getters, setters e outros métodos úteis mais usados em scripts. Este é um ótimo ponto de partida para sua jornada de API. As seções posteriores agrupam os métodos e ajudam a criar um modelo mental à medida que você começa a desbloquear `Range` as APIs do objeto.

## <a name="example-scripts"></a>Scripts de exemplo

* [Leitura e gravação básicas](#basic-read-and-write)
* [Adicionar linha no final da planilha](#add-row-at-the-end-of-worksheet)
* [Limpar filtro de coluna](clear-table-filter-for-active-cell.md)
* [Colorir cada célula com cor exclusiva](#color-each-cell-with-unique-color)
* [Intervalo de atualizações com valores usando matriz 2D (bidimensional)](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a>Leitura e gravação básicas

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a>Adicionar linha no final da planilha

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a>Colorir cada célula com cor exclusiva

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a>Intervalo de atualizações com valores usando matriz 2D

Calcula dinamicamente a dimensão do intervalo a ser atualizada com base nos valores da matriz 2D.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a>Vídeos de treinamento: Noções básicas de intervalo

_Noções básicas de intervalo_

[![Assista ao vídeo passo a passo em Noções básicas de intervalo](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Vídeo passo a passo sobre noções básicas de intervalo")

_Adicionar linha no final da planilha_

[![Assista ao vídeo passo a passo sobre como adicionar uma linha no final de uma planilha](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Vídeo passo a passo sobre como adicionar uma linha no final de uma planilha")

## <a name="methods-that-return-some-range-metadata"></a>Métodos que retornam alguns metadados de intervalo

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a>Métodos que retornam dados/constantes associadas a um determinado intervalo

### <a name="returned-as-single-cell-value"></a>Retornado como valor de célula única

* getFormula(), getFormulaLocal()
* getFormulaR1C1()
* getNumberFormat(), getNumberFormatLocal()
* getText()
* getValue()
* getValueType()

### <a name="returned-as-2d-arrays-whole-range"></a>Retornado como matrizes 2D (intervalo inteiro)

* getFormulas(), getFormulasLocal()
* getFormulasR1C1()
* getNumberFormatCategories()
* getNumberFormats(), getNumberFormatsLocal()
* getTexts()
* getValues()
* getValueTypes()
* getHidden()
* getIsEntireRow()
* getIsEntireColumn()

## <a name="methods-that-return-other-range-object"></a>Métodos que retornam outro objeto range

* getSurroundingRegion() - semelhante a CurrentRegion no VBA
* getCell(row, column)
* getColumn(column)
* getColumnHidden()
* getColumnsAfter(count)
* getColumnsBefore(count)
* getEntireColumn()
* getEntireRow()
* getLastCell()
* getLastColumn()
* getLastRow()
* getRow(row)
* getRowHidden()
* getRowsAbove(count)
* getRowsBelow(count)

**Importante/Interessante**

* _workbook_.getSelectedRange()
* _workbook_.getActiveCell()
* getUsedRange(valuesOnly)
* getAbsoluteResizedRange(numRows, numColumns)
* getOffsetRange(rowOffset, columnOffset)
* getResizedRange(deltaRows, deltaColumns)

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a>Métodos que retornam um objeto range em relação a outro objeto range

* getBoundingRect(anotherRange)
* getIntersection(anotherRange)

## <a name="methods-that-return-other-objects-non-range-objects"></a>Métodos que retornam outros objetos (objetos que não são de intervalo)

* getDirectPrecedents()
* getWorksheet()
* getTables(fullyContained)
* getPivotTables(fullyContained)
* getDataValidation()
* getPredefinedCellStyle()

## <a name="set-methods"></a>Definir métodos

### <a name="singular-cell-set-methods"></a>Métodos de conjunto de células singulares

* setFormula(formula)
* setFormulaLocal(formulaLocal)
* setFormulaR1C1(formulaR1C1)
* setNumberFormatLocal(numberFormatLocal)
* setValue(value)

### <a name="2d--entire-range-set-methods"></a>Métodos 2D / conjunto de intervalos inteiros

* setFormulas(formulas)
* setFormulasLocal(formulasLocal)
* setFormulasR1C1(formulasR1C1)
* setNumberFormat(numberFormat)
* setNumberFormats(numberFormats)
* setNumberFormatsLocal(numberFormatsLocal)
* setValues(values)

## <a name="other-methods"></a>Outros métodos

* merge(across)
* unmerge()

## <a name="coming-soon"></a>Em breve

* APIs de borda de intervalo
