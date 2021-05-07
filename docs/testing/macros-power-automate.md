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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Como usar arquivos de macro em fluxos Power Automate fluxos

[Power Automate fluxos](https://flow.microsoft.com/) fornecem conectores Excel para ajudar Excel a conectar arquivos Excel com o restante de seus dados organizacionais e [aplicativos,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) como Teams, Outlook e SharePoint.

No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).

:::image type="content" source="../images/no-xlsm.png" alt-text="A Power Automate executar script mostrando nenhum arquivo de macro selecionado. O erro mostrado é 'Arquivo' é necessário":::

Uma maneira de se livrar desse problema é incluir a ação "Obter metadados de arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="A Power Automate executar script mostrando o arquivo de macro selecionado e nenhum erro executar script":::

> [!NOTE]
> Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector Excel online. Certifique-se de testar antes de implantar sua solução.

## <a name="other-resources"></a>Outros recursos

Assista ao vídeo do [YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre como usar um arquivo .xlsm em uma ação Executar Script.
