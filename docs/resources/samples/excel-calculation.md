---
title: Gerenciar o modo de cálculo Excel
description: Saiba como usar Office Scripts para gerenciar o modo de cálculo em Excel na Web.
ms.date: 05/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: fec88c904d95bfdab1514d44921f7fb1c6e9dd35
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585510"
---
# <a name="manage-calculation-mode-in-excel"></a>Gerenciar o modo de cálculo Excel

Este exemplo mostra como usar o modo [](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) de cálculo e calcular métodos em Excel na Web usando Office Scripts. Você pode experimentar o script em qualquer arquivo Excel arquivo.

## <a name="scenario"></a>Cenário

As guias de trabalho com um grande número de fórmulas podem demorar um pouco para recalcular. Em vez de Excel controle quando os cálculos ocorrem, você pode gerenciá-los como parte do script. Isso ajudará no desempenho em determinados cenários.

O script de exemplo define o modo de cálculo como manual. Isso significa que a workbook só recalculará fórmulas quando o script diz a ela (ou você [calcula manualmente por meio da interface do usuário](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4)). Em seguida, o script exibe o modo de cálculo atual e recalcula totalmente toda a workbook.

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
