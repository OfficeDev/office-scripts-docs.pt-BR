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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Como usar arquivos de macro em fluxos do Power Automate

[Fluxos do Power Automate](https://flow.microsoft.com/) fornecem conectores do Excel para ajudar a conectar arquivos do Excel com o restante de seus dados e aplicativos organizacionais, como Teams, Outlook e SharePoint. [](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)

No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).

:::image type="content" source="../images/no-xlsm.png" alt-text="A ação de script Power Automate Run mostrando nenhum arquivo de macro selecionado. O erro mostrado é 'Arquivo' é necessário.":::

Uma maneira de se livrar desse problema é incluir a ação "Obter Metadados de Arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="A ação de script Executar do Power Automate mostrando o arquivo de macro selecionado e nenhum erro de script executar.":::

> [!NOTE]
> Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector online do Excel. Certifique-se de testar antes de implantar sua solução.

[![Assista a um vídeo sobre como usar XLSM na ação Executar Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre como usar XLSM na ação Executar Script")
