---
title: Office Propriedade e armazenamento de arquivos de scripts
description: Informações sobre como Office scripts são armazenados em Microsoft OneDrive e transferidos entre proprietários.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232526"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Propriedade e armazenamento de arquivos de scripts

Office Scripts são armazenados como **arquivos .osts** em seu Microsoft OneDrive. Isso permite que seus scripts existam fora de qualquer workbook específico. Suas OneDrive configurações controlam o acesso compartilhado e as permissões para todos os arquivos **.osts de** script; independentemente de qualquer Excel configurações.

## <a name="file-storage"></a>Armazenamento de arquivos

Você Office scripts são armazenados em seu OneDrive. Os **arquivos .osts** são encontrados na **pasta /Documents/Office Scripts/.** Todas as edições feitas nesses **arquivos .osts,** como renomeação ou exclusão de arquivos, serão refletidas no Editor de Código e na Galeria de Scripts.

Os scripts compartilhados com uma de suas guias de trabalho permanecem no OneDrive. Eles não são copiados para nenhuma pasta local ou OneDrive quando você executar o script compartilhado em Excel. O **botão Fazer uma Cópia** do Editor de Código salva uma cópia separada do script em seu OneDrive. As alterações na cópia não afetam o script original.

### <a name="script-folders"></a>Pastas de script

Adicionar pastas ao seu OneDrive ajuda a manter seus scripts organizados. Todas as pastas em **/Documents/Office Scripts/** são exibidas na seção **Meus Scripts** do Editor de Código. Observe que essas pastas não podem ser criadas ou excluídas usando o Editor de Código. Da mesma forma, os scripts não podem ser colocados em pastas ou movidos entre pastas usando o Editor de Código.

:::image type="content" source="../images/script-folders.png" alt-text="A caixa de diálogo Novo Script no Editor de Código mostrando scripts contidos em pastas, conforme exibido no painel de tarefas":::

## <a name="file-ownership-and-retention"></a>Propriedade e retenção de arquivos

Office Os scripts são armazenados no OneDrive. Eles seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive. Para saber como lidar com os scripts criados e compartilhados por um usuário que está sendo removido da sua organização, confira [retenção e exclusão do OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Confira também

- [Compartilhando scripts do Office no Excel para a Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Desfazer os efeitos de um script do Office](../testing/undo.md)
