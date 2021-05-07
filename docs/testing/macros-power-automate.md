---
title: Usar arquivos de macro em Power Automate fluxos
description: Saiba como usar arquivos de macro ou arquivos xlsm Power Automate fluxos.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232652"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="8894c-103">Como usar arquivos de macro em fluxos Power Automate fluxos</span><span class="sxs-lookup"><span data-stu-id="8894c-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="8894c-104">[Power Automate fluxos](https://flow.microsoft.com/) fornecem conectores Excel para ajudar Excel a conectar arquivos Excel com o restante de seus dados organizacionais e [aplicativos,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) como Teams, Outlook e SharePoint.</span><span class="sxs-lookup"><span data-stu-id="8894c-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="8894c-105">No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).</span><span class="sxs-lookup"><span data-stu-id="8894c-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="A Power Automate executar script mostrando nenhum arquivo de macro selecionado. O erro mostrado é 'Arquivo' é necessário":::

<span data-ttu-id="8894c-107">Uma maneira de se livrar desse problema é incluir a ação "Obter metadados de arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8894c-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="A Power Automate executar script mostrando o arquivo de macro selecionado e nenhum erro executar script":::

> [!NOTE]
> <span data-ttu-id="8894c-109">Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector Excel online.</span><span class="sxs-lookup"><span data-stu-id="8894c-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="8894c-110">Certifique-se de testar antes de implantar sua solução.</span><span class="sxs-lookup"><span data-stu-id="8894c-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="8894c-111">Outros recursos</span><span class="sxs-lookup"><span data-stu-id="8894c-111">Other resources</span></span>

<span data-ttu-id="8894c-112">Assista ao vídeo do [YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre como usar um arquivo .xlsm em uma ação Executar Script.</span><span class="sxs-lookup"><span data-stu-id="8894c-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
