---
title: Usar arquivos de macro em fluxos do Power Automate
description: Saiba como usar arquivos de macro ou arquivos xlsm em fluxos do Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571100"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="14919-103">Como usar arquivos de macro em fluxos do Power Automate</span><span class="sxs-lookup"><span data-stu-id="14919-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="14919-104">[Fluxos do Power Automate](https://flow.microsoft.com/) fornecem conectores do Excel para ajudar a conectar arquivos do Excel com o restante de seus dados e aplicativos organizacionais, como Teams, Outlook e SharePoint. [](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)</span><span class="sxs-lookup"><span data-stu-id="14919-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="14919-105">No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).</span><span class="sxs-lookup"><span data-stu-id="14919-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

![Nenhum xlsm na ação Executar Script](../images/no-xlsm.png)

<span data-ttu-id="14919-107">Uma maneira de se livrar desse problema é incluir a ação "Obter Metadados de Arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="14919-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

![xlsm na ação Executar Script](../images/xlsm-in-pa.png)

> [!NOTE]
> <span data-ttu-id="14919-109">Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector online do Excel.</span><span class="sxs-lookup"><span data-stu-id="14919-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="14919-110">Certifique-se de testar antes de implantar sua solução.</span><span class="sxs-lookup"><span data-stu-id="14919-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="14919-111">[![Assista a um vídeo sobre como usar XLSM na ação Executar Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre como usar XLSM na ação Executar Script")</span><span class="sxs-lookup"><span data-stu-id="14919-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>
