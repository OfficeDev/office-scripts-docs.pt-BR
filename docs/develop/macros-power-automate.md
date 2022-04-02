---
title: Usar arquivos habilitados para macro em Power Automate fluxos
description: Saiba como usar arquivos habilitados para macro ou arquivos .xlsm em Power Automate fluxos.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585741"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>Como usar arquivos habilitados para macro em Power Automate fluxos

Você pode integrar seus arquivos .xlsm a um Power Automate fluxo. Isso permite que você comece a converter suas soluções de automação existentes em formatos baseados na Web. Observe que as macros contidas nos arquivos .xslm não podem ser Power Automate. Somente Office scripts estão habilitados lá.

O [conector Excel Online (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) no Power Automate normalmente é [](https://flow.microsoft.com/) limitado a arquivos no formato Microsoft Excel Planilha Open XML (.xlsx). Seu navegador de arquivos só permite selecionar .xlsx arquivos. No entanto, os arquivos habilitados para macro são compatíveis com a ação executar **script do conector** se os metadados do arquivo são usados.

Em seu fluxo, use a **ação Obter Metadados** de Arquivo [OneDrive for Business ou](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) conectores. A **ação Executar script** aceita os metadados como um arquivo válido. Use o *conteúdo dinâmico de ID* retornado da **ação Obter metadados de** arquivo como o argumento "File" ao executar o script. A captura de tela a seguir mostra um fluxo fornecendo os metadados para um arquivo chamado "Test Macro File.xlsm" para uma ação **de script Executar** .

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Um fluxo com uma ação Obter metadados de arquivo passando os metadados de um arquivo de macro para uma ação executar script.":::

> [!WARNING]
> Alguns arquivos .xlsm, especialmente aqueles com controles ActiveX ou Formulário, podem não funcionar no conector Excel online. Certifique-se de testar antes de implantar sua solução.

## <a name="other-resources"></a>Outros recursos

[Assista ao vídeo do YouTube de Sudhi Ramamurthy sobre como usar um arquivo .xlsm em uma ação Executar Script](https://youtu.be/o-H9BbywJQQ).
