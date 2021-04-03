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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Como usar arquivos de macro em fluxos do Power Automate

[Fluxos do Power Automate](https://flow.microsoft.com/) fornecem conectores do Excel para ajudar a conectar arquivos do Excel com o restante de seus dados e aplicativos organizacionais, como Teams, Outlook e SharePoint. [](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)

No entanto, os arquivos de macro não podem ser selecionados no menu suspenso do arquivo (consulte um exemplo na captura de tela a seguir).

![Nenhum xlsm na ação Executar Script](../images/no-xlsm.png)

Uma maneira de se livrar desse problema é incluir a ação "Obter Metadados de Arquivo" (OneDrive ou SharePoint) e usar a propriedade ID na ação "Executar Script", conforme mostrado na captura de tela a seguir.

![xlsm na ação Executar Script](../images/xlsm-in-pa.png)

> [!NOTE]
> Alguns XLSM (especialmente aqueles com controles ActiveX/Formulário) podem não funcionar no conector online do Excel. Certifique-se de testar antes de implantar sua solução.

[![Assista a um vídeo sobre como usar XLSM na ação Executar Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre como usar XLSM na ação Executar Script")
