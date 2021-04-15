---
title: Armazenamento e propriedade de arquivos do Office Scripts
description: Informações sobre como os Scripts do Office são armazenados no Microsoft OneDrive e transferidos entre proprietários.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: bd868c1dbfd0b33d3cd9fc4ee774c654d86f9b07
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755102"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Armazenamento e propriedade de arquivos do Office Scripts

Os Scripts do Office são armazenados como **arquivos .osts** no Microsoft OneDrive. Isso permite que seus scripts existam fora de qualquer workbook específico. As configurações do OneDrive controlam o acesso compartilhado e as permissões de todos os arquivos **.osts de** script; independente de qualquer configuração do Excel.

## <a name="file-storage"></a>Armazenamento de arquivos

Os Scripts do Office são armazenados no OneDrive. Os **arquivos .osts** são encontrados na **pasta /Documents/Office Scripts/.** Todas as edições feitas nesses **arquivos .osts,** como renomeação ou exclusão de arquivos, serão refletidas no Editor de Código e na Galeria de Scripts.

Os scripts compartilhados com uma de suas guias de trabalho permanecem no OneDrive do criador do script. Eles não são copiados para nenhuma pasta local ou do OneDrive quando você executar o script compartilhado no Excel. O **botão Fazer uma Cópia** do Editor de Código salva uma cópia separada do script no OneDrive. As alterações na cópia não afetam o script original.

### <a name="script-folders"></a>Pastas de script

Adicionar pastas ao OneDrive ajuda a manter os scripts organizados. Todas as pastas em **/Documents/Office Scripts/** são exibidas na seção **Meus Scripts** do Editor de Código. Observe que essas pastas não podem ser criadas ou excluídas usando o Editor de Código. Da mesma forma, os scripts não podem ser colocados em pastas ou movidos entre pastas usando o Editor de Código.

:::image type="content" source="../images/script-folders.png" alt-text="A caixa de diálogo Novo Script no Editor de Código mostrando scripts contidos em pastas, conforme exibido no painel de tarefas.":::

## <a name="file-ownership-and-retention"></a>Propriedade e retenção de arquivos

Os Scripts do Office são armazenados no OneDrive de um usuário. Eles seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive. Para saber como lidar com os scripts criados e compartilhados por um usuário que está sendo removido da sua organização, confira [retenção e exclusão do OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Confira também

- [Compartilhando scripts do Office no Excel para a Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Desfazer os efeitos de um script do Office](../testing/undo.md)
