---
title: Gerenciar o modo de cálculo Excel
description: Saiba como usar Office scripts para gerenciar o modo de cálculo em Excel na Web.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232449"
---
# <a name="manage-calculation-mode-in-excel"></a>Gerenciar o modo de cálculo Excel

Este exemplo mostra como [](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts. Você pode experimentar o script em qualquer arquivo Excel arquivo.

## <a name="scenario"></a>Cenário

No Excel na Web, o modo de cálculo de um arquivo pode ser controlado programaticamente usando APIs. As ações a seguir são possíveis usando Office Scripts.

1. Obter o modo de cálculo.
1. De definir o modo de cálculo.
1. Calcule Excel fórmulas para arquivos que estão definidos para o modo manual (também chamado de recálcula).

## <a name="sample-code-control-calculation-mode"></a>Código de exemplo: Modo de cálculo de controle

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a>Vídeo de treinamento: Gerenciar o modo de cálculo

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/iw6O8QH01CI).
