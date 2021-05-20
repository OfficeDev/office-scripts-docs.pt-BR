---
title: Office Armazenamento e propriedade de arquivos scripts
description: Informações sobre como Office Scripts são armazenadas em Microsoft OneDrive e transferidas entre proprietários.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545798"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Armazenamento e propriedade de arquivos scripts

Office Os scripts são armazenados como **arquivos .osts** em sua Microsoft OneDrive. Eles são armazenados separadamente de uma pasta de trabalho. Para dar acesso aos outros, [compartilhe o roteiro com uma Excel livro de trabalho](excel.md#sharing-scripts). Isso significa que você está ligando o script com o arquivo, não anexando-o. Quem tiver acesso ao arquivo Excel também poderá visualizar, executar ou fazer uma cópia do script.

A menos que compartilhe seus roteiros, ninguém mais pode acessá-los. Suas configurações de OneDrive controlam o acesso compartilhado e as permissões para todos os arquivos script **.osts,** independente de qualquer configuração Excel. Os scripts não podem ser vinculados a partir de um disco local ou locais de nuvem personalizados. Office Scripts só reconhece e executa um script se estiver em sua pasta OneDrive ou compartilhado com a pasta de trabalho.

## <a name="file-storage"></a>Armazenamento de arquivos

Você Office scripts estão armazenados em seu OneDrive. Os arquivos **.osts** são encontrados na pasta **/Documentos/Office Scripts/pasta.** Quaisquer edições feitas a esses arquivos **.osts,** como renomeação ou exclusão de arquivos, serão refletidas no Editor de Código e na Galeria de Scripts.

Scripts que são compartilhados com uma de suas pastas de trabalho permanecem no OneDrive do criador do roteiro. Eles não são copiados para nenhuma de suas pastas locais ou OneDrive quando você executa o script compartilhado em Excel. O botão **Fazer uma cópia** do Editor de Código salva uma cópia separada do script em sua OneDrive. Alterações na cópia não afetam o script original.

## <a name="file-ownership-and-retention"></a>Propriedade e retenção de arquivos

Office Os scripts são armazenados no OneDrive do usuário. Eles seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive. Para saber como lidar com os scripts criados e compartilhados por um usuário que está sendo removido da sua organização, confira [retenção e exclusão do OneDrive](/onedrive/retention-and-deletion).

Durante a edição, os arquivos são armazenados temporariamente no navegador. Você deve salvar o script antes de fechar a janela Excel para salvá-lo no local OneDrive. Não se esqueça de salvar o arquivo após as edições, ou então essas edições estarão apenas na versão do arquivo do navegador.

## <a name="see-also"></a>Confira também

- [Compartilhando scripts do Office no Excel para a Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Desfazer os efeitos do Scripts do Office](../testing/undo.md)
