---
title: Armazenamento e propriedade de arquivos de scripts do Office
description: Informações sobre como os scripts do Office são armazenados no Microsoft OneDrive e transferidos entre proprietários.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346858"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Armazenamento e propriedade de arquivos de scripts do Office

Os scripts do Office são armazenados como arquivos **. OSTs** no Microsoft onedrive. Isso permite que seus scripts existam fora de qualquer pasta de trabalho específica. Suas configurações do OneDrive controlam o acesso compartilhado e as permissões para todos os arquivos script **. OSTs** ; independente de qualquer configuração do Excel.

## <a name="file-storage"></a>Armazenamento de arquivos

Os scripts do Office são armazenados em seu OneDrive. Os arquivos **. OSTs** são encontrados na pasta **scripts//Documents/Office** . Quaisquer edições feitas nesses arquivos **. OSTs** , como renomear ou excluir arquivos, serão refletidas no editor de código e na Galeria de scripts.

Scripts que são compartilhados com uma de suas pastas de trabalho permanecem no OneDrive do criador do script. Eles não são copiados para nenhuma de suas pastas locais ou do OneDrive quando você executa o script compartilhado no Excel. O botão **fazer uma cópia** do editor de código salva uma cópia separada do script em seu onedrive. As alterações na cópia não afetam o script original.

### <a name="script-folders"></a>Pastas de script

A adição de pastas ao OneDrive ajuda a manter seus scripts organizados. Quaisquer pastas em **scripts do/Documents/Office/** são exibidas na seção **meus scripts** do editor de código. Observe que essas pastas não podem ser criadas ou excluídas usando o editor de código. Da mesma forma, os scripts não podem ser colocados em pastas ou movidos entre pastas usando o editor de código.

![Alguns scripts contidos em pastas, conforme exibido no painel de tarefas do editor de código](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a>Propriedade e retenção de arquivo

Os scripts do Office são armazenados no OneDrive de um usuário. Eles seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive. Para saber como lidar com os scripts criados e compartilhados por um usuário que está sendo removido da sua organização, confira [retenção e exclusão do OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Também consulte

- [Compartilhando scripts do Office no Excel para a Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Desfazer os efeitos de um script do Office](../testing/undo.md)
