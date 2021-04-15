---
title: Usar arquivos de macro em fluxos do Power Automate
description: Saiba como usar arquivos de macro ou arquivos xlsm em fluxos do Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: a7929fc485ae2118d30a4f2783538d0e04deca2a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755011"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="6a3b2-103">Como usar arquivos de macro em fluxos do Power Automate</span><span class="sxs-lookup"><span data-stu-id="6a3b2-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="6a3b2-104">[Fluxos do Power Automate](https://flow.microsoft.com/) fornecem conectores do Excel para ajudar a conectar arquivos do Excel com o restante de seus dados e aplicativos organizacionais, como Teams, Outlook e SharePoint. [](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)</span><span class="sxs-lookup"><span data-stu-id="6a3b2-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="6a3b2-105">No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).</span><span class="sxs-lookup"><span data-stu-id="6a3b2-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="A ação de script Power Automate Run mostrando nenhum arquivo de macro selecionado. O erro mostrado é 'Arquivo' é necessário.":::

<span data-ttu-id="6a3b2-107">Uma maneira de se livrar desse problema é incluir a ação "Obter Metadados de Arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a3b2-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="A ação de script Executar do Power Automate mostrando o arquivo de macro selecionado e nenhum erro de script executar.":::

> [!NOTE]
> <span data-ttu-id="6a3b2-109">Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector online do Excel.</span><span class="sxs-lookup"><span data-stu-id="6a3b2-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="6a3b2-110">Certifique-se de testar antes de implantar sua solução.</span><span class="sxs-lookup"><span data-stu-id="6a3b2-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="6a3b2-111">[![Assista a um vídeo sobre como usar XLSM na ação Executar Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre como usar XLSM na ação Executar Script")</span><span class="sxs-lookup"><span data-stu-id="6a3b2-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>
