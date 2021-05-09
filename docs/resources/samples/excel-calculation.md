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
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="9ab3e-103">Gerenciar o modo de cálculo Excel</span><span class="sxs-lookup"><span data-stu-id="9ab3e-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="9ab3e-104">Este exemplo mostra como [](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="9ab3e-105">Você pode experimentar o script em qualquer arquivo Excel arquivo.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="9ab3e-106">Cenário</span><span class="sxs-lookup"><span data-stu-id="9ab3e-106">Scenario</span></span>

<span data-ttu-id="9ab3e-107">As guias de trabalho com um grande número de fórmulas podem demorar um pouco para recalcular.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="9ab3e-108">Em vez de Excel controle quando os cálculos ocorrem, você pode gerenciá-los como parte do seu script.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="9ab3e-109">Isso ajudará no desempenho em determinados cenários.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="9ab3e-110">O script de exemplo define o modo de cálculo como manual.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="9ab3e-111">Isso significa que a workbook só recalculará fórmulas quando o script diz a ela (ou você [calcula manualmente](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)por meio da interface do usuário ).</span><span class="sxs-lookup"><span data-stu-id="9ab3e-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="9ab3e-112">Em seguida, o script exibe o modo de cálculo atual e recalcula totalmente toda a workbook.</span><span class="sxs-lookup"><span data-stu-id="9ab3e-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="9ab3e-113">Código de exemplo: Modo de cálculo de controle</span><span class="sxs-lookup"><span data-stu-id="9ab3e-113">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="9ab3e-114">Vídeo de treinamento: Gerenciar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="9ab3e-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="9ab3e-115">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="9ab3e-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
