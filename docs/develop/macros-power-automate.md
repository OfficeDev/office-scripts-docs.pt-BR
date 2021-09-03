---
title: Usar arquivos de macro em Power Automate fluxos
description: Saiba como usar arquivos de macro ou arquivos xlsm Power Automate fluxos.
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866532"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Como usar arquivos de macro em fluxos Power Automate fluxos

O [conector Excel Online (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) no Power Automate normalmente funciona apenas com arquivos no formato Microsoft Excel Open XML Spreadsheet (.xlsx). [](https://flow.microsoft.com/) O navegador de arquivos limita sua seleção .xlsx arquivos dentro do conector. No entanto, os arquivos de macro são compatíveis com a ação de script Executar do **conector** se os metadados do arquivo são usados.

Em seu fluxo, use a **ação Obter metadados** de arquivo dos conectores [OneDrive for Business](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) ou [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) de arquivo. A **ação Executar script** aceita os metadados como um arquivo válido. Use o *conteúdo dinâmico de ID* retornado da **ação Obter metadados de** arquivo como o argumento "File" ao executar o script. A captura de tela a seguir mostra um fluxo fornecendo os metadados para um arquivo chamado "Test Macro File.xlsm" para uma ação **de script Executar.**

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Um fluxo com uma ação Obter metadados de arquivo passando os metadados de um arquivo de macro para uma ação executar script.":::

> [!WARNING]
> Alguns arquivos .xlsm, especialmente aqueles com controles ActiveX ou Formulário, podem não funcionar no conector Excel online. Certifique-se de testar antes de implantar sua solução.

## <a name="other-resources"></a>Outros recursos

Assista ao vídeo do [YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre como usar um arquivo .xlsm em uma ação Executar Script.
