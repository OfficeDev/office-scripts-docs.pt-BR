---
title: Office Propriedade e armazenamento de arquivos de scripts
description: Informações sobre como Office scripts são armazenados em Microsoft OneDrive e transferidos entre proprietários.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545798"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Propriedade e armazenamento de arquivos de scripts

Office Scripts são armazenados como **arquivos .osts** em seu Microsoft OneDrive. Eles são armazenados separadamente de uma workbook. Para dar acesso a outras pessoas, [compartilhe o script com uma Excel de trabalho](excel.md#sharing-scripts). Isso significa que você está vinculando o script com o arquivo, não anexando-o. Quem tiver acesso ao arquivo Excel também poderá exibir, executar ou fazer uma cópia do script.

A menos que você compartilhe seus scripts, ninguém mais poderá acessá-los. Suas OneDrive controlam o acesso compartilhado e as permissões para todos os arquivos **.osts** de script, independentemente de qualquer configuração Excel script. Os scripts não podem ser vinculados a partir de um disco local ou locais de nuvem personalizados. Office Os scripts só reconhecem e executam um script se ele estiver em sua pasta OneDrive ou compartilhado com a pasta de trabalho.

## <a name="file-storage"></a>Armazenamento de arquivos

Você Office scripts são armazenados em seu OneDrive. Os **arquivos .osts** são encontrados na **pasta /Documents/Office Scripts/.** Todas as edições feitas nesses **arquivos .osts,** como renomeação ou exclusão de arquivos, serão refletidas no Editor de Código e na Galeria de Scripts.

Os scripts compartilhados com uma de suas guias de trabalho permanecem no OneDrive. Eles não são copiados para nenhuma pasta local ou OneDrive quando você executar o script compartilhado em Excel. O **botão Fazer uma Cópia** do Editor de Código salva uma cópia separada do script em seu OneDrive. As alterações na cópia não afetam o script original.

## <a name="file-ownership-and-retention"></a>Propriedade e retenção de arquivos

Office Os scripts são armazenados no OneDrive. Eles seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive. Para saber como lidar com os scripts criados e compartilhados por um usuário que está sendo removido da sua organização, confira [retenção e exclusão do OneDrive](/onedrive/retention-and-deletion).

Durante a edição, os arquivos são temporariamente armazenados no navegador. Você deve salvar o script antes de fechar a janela Excel para salvá-lo no OneDrive local. Não se esqueça de salvar o arquivo após edições, ou essas edições só estarão na versão do arquivo do navegador.

## <a name="see-also"></a>Confira também

- [Compartilhando scripts do Office no Excel para a Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Desfazer os efeitos do Scripts do Office](../testing/undo.md)
