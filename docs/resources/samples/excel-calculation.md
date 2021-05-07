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
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="45d79-103">Gerenciar o modo de cálculo Excel</span><span class="sxs-lookup"><span data-stu-id="45d79-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="45d79-104">Este exemplo mostra como [](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="45d79-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="45d79-105">Você pode experimentar o script em qualquer arquivo Excel arquivo.</span><span class="sxs-lookup"><span data-stu-id="45d79-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="45d79-106">Cenário</span><span class="sxs-lookup"><span data-stu-id="45d79-106">Scenario</span></span>

<span data-ttu-id="45d79-107">No Excel na Web, o modo de cálculo de um arquivo pode ser controlado programaticamente usando APIs.</span><span class="sxs-lookup"><span data-stu-id="45d79-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="45d79-108">As ações a seguir são possíveis usando Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="45d79-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="45d79-109">Obter o modo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="45d79-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="45d79-110">De definir o modo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="45d79-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="45d79-111">Calcule Excel fórmulas para arquivos que estão definidos para o modo manual (também chamado de recálcula).</span><span class="sxs-lookup"><span data-stu-id="45d79-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="45d79-112">Código de exemplo: Modo de cálculo de controle</span><span class="sxs-lookup"><span data-stu-id="45d79-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="45d79-113">Vídeo de treinamento: Gerenciar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="45d79-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="45d79-114">[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="45d79-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
