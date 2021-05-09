---
title: Gerenciar o modo de cálculo Excel
description: Saiba como usar Office scripts para gerenciar o modo de cálculo em Excel na Web.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285910"
---
# <a name="manage-calculation-mode-in-excel"></a>Gerenciar o modo de cálculo Excel

Este exemplo mostra como [](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts. Você pode experimentar o script em qualquer arquivo Excel arquivo.

## <a name="scenario"></a>Cenário

As guias de trabalho com um grande número de fórmulas podem demorar um pouco para recalcular. Em vez de Excel controle quando os cálculos ocorrem, você pode gerenciá-los como parte do seu script. Isso ajudará no desempenho em determinados cenários.

O script de exemplo define o modo de cálculo como manual. Isso significa que a workbook só recalculará fórmulas quando o script diz a ela (ou você [calcula manualmente](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)por meio da interface do usuário ). Em seguida, o script exibe o modo de cálculo atual e recalcula totalmente toda a workbook.

## <a name="sample-code-control-calculation-mode"></a>Código de exemplo: Modo de cálculo de controle

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a>Vídeo de treinamento: Gerenciar o modo de cálculo

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/iw6O8QH01CI).
