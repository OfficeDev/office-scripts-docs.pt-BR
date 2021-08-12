---
title: Usar arquivos de macro em Power Automate fluxos
description: Saiba como usar arquivos de macro ou arquivos xlsm Power Automate fluxos.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 67686ca5d677a2d04c47d6312a37fa6375bed4a2bef9ae7b6ee61bba2302bfb4
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847216"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Como usar arquivos de macro em fluxos Power Automate fluxos

[Power Automate fluxos](https://flow.microsoft.com/) fornecem conectores Excel para ajudar Excel a conectar arquivos Excel com o restante de seus dados organizacionais e [aplicativos,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) como Teams, Outlook e SharePoint.

No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).

:::image type="content" source="../images/no-xlsm.png" alt-text="A Power Automate executar script mostrando nenhum arquivo de macro selecionado. O erro mostrado é 'Arquivo' é necessário.":::

Uma maneira de se livrar desse problema é incluir a ação "Obter metadados de arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="A Power Automate executar script mostrando o arquivo de macro selecionado e nenhum erro de script executar.":::

> [!NOTE]
> Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector Excel online. Certifique-se de testar antes de implantar sua solução.

## <a name="other-resources"></a>Outros recursos

Assista ao vídeo do [YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre como usar um arquivo .xlsm em uma ação Executar Script.
