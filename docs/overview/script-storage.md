---
title: Office armazenamento e propriedade de arquivos de scripts
description: Informações sobre como Office scripts são armazenados em Microsoft OneDrive e transferidos entre proprietários.
ms.date: 05/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e2bc89db54ee5520c3b911ebd0f182777a78e2b
ms.sourcegitcommit: 8ae932e8b4e521fec8576ab16126eb9fe22a8dd7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2022
ms.locfileid: "65310754"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office armazenamento e propriedade de arquivos de scripts

Office scripts são armazenados como **arquivos .osts** em seu Microsoft OneDrive. Eles são armazenados separadamente de uma pasta de trabalho. Para dar acesso a outras pessoas, [compartilhe o script com uma Excel de trabalho](excel.md#share-office-scripts). Isso significa que você está vinculando o script ao arquivo, não anexando-o. Quem tiver acesso ao arquivo Excel também poderá exibir, executar ou fazer uma cópia do script.

A menos que você compartilhe seus scripts, ninguém mais poderá accessá-los. Suas OneDrive configurações controlam o acesso compartilhado e as permissões para todos os arquivos **.osts** de script, independentemente de Excel configurações. Os scripts não podem ser vinculados de um disco local ou locais de nuvem personalizados. Office scripts só reconhecerão e executarão um script se ele estiver em sua pasta OneDrive ou compartilhado com a pasta de trabalho.

## <a name="file-storage"></a>Armazenamento de arquivos

Você Office scripts são armazenados em seu OneDrive. Os **arquivos .osts** são encontrados na **pasta /Documents/Office Scripts/**. Todas as edições feitas nesses arquivos **.osts** , como renomear ou excluir arquivos, serão refletidas no Editor de Códigos e na Galeria de Scripts.

Os scripts que são compartilhados com uma de suas pastas de trabalho permanecem na conta do criador do script OneDrive. Eles não são copiados para nenhuma das pastas locais ou OneDrive quando você executa o script compartilhado em Excel. O **botão Fazer uma Cópia** do Editor de Códigos salva uma cópia separada do script em seu OneDrive. As alterações na cópia não afetam o script original.

### <a name="restore-deleted-scripts"></a>Restaurar scripts excluídos

Quando você exclui um script no Excel, ele vai para sua OneDrive lixeira. Para restaurar um script excluído, siga as etapas listadas em Restaurar arquivos [ou pastas excluídos no OneDrive](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f). Restaurar um **arquivo .osts** o retorna para a **lista Todos os scripts** .

Um script excluído é descompartilhado com a pasta de trabalho. Quando você restaura um script, ele **não retém** o acesso ao script. Você precisará compartilhar o script novamente.

Os scripts restaurados ainda funcionam conforme o esperado com Power Automate fluxos. Você não precisa recriar o conector de fluxo.

## <a name="file-ownership-and-retention"></a>Propriedade e retenção do arquivo

Office scripts são armazenados no banco de dados OneDrive. Eles seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive. Para saber como lidar com os scripts criados e compartilhados por um usuário que está sendo removido da sua organização, confira [retenção e exclusão do OneDrive](/onedrive/retention-and-deletion).

Durante a edição, os arquivos são armazenados temporariamente no navegador. Você deve salvar o script antes de fechar Excel janela para salvá-lo no OneDrive local. Não se esqueça de salvar o arquivo após as edições, caso contrário, essas edições só estarão na versão do navegador do arquivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar Office uso de scripts no nível de administrador

Descubra quais locatários estão usando Office Scripts com o log de auditoria no centro de conformidade. Para saber como usar essa ferramenta, visite [Pesquisar o log de auditoria no Centro de Conformidade & Segurança](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Para localizar quem está usando Office scripts com a ferramenta de pesquisa, `.osts` adicione o campo Arquivo **, pasta ou site**. Isso pesquisa todos os arquivos com a extensão Office arquivo Scripts. Se alguém em sua organização tiver usado o Office scripts, a atividade do usuário será exibida nos resultados da pesquisa de log de auditoria.

## <a name="see-also"></a>Confira também

- [Compartilhamento de Scripts do Office no Excel para a web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Desfazer os efeitos do Scripts do Office](../testing/undo.md)
