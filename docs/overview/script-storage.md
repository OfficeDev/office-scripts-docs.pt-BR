---
title: Office armazenamento e propriedade de arquivos de scripts
description: Informações sobre como Office scripts são armazenados em Microsoft OneDrive e transferidos entre proprietários.
ms.date: 06/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 17603660bcafa41f898b15b1226d11fa0d51b0a5
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128206"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office armazenamento e propriedade de arquivos de scripts

> [!IMPORTANT]
> SharePoint suporte para Office scripts está sendo distribuído e não está disponível para todos. É liberado lentamente para um número maior de usuários para garantir que está funcionando conforme o esperado. Esse recurso está sujeito a alterações com base em seus comentários.

Office scripts são armazenados como **arquivos .osts** no Microsoft OneDrive ou em uma SharePoint. Eles são armazenados separadamente de uma pasta de trabalho. Para dar aos usuários que estão fora do SharePoint acesso ao script, compartilhe o script com uma Excel [de trabalho](excel.md#share-office-scripts). Isso significa que você está vinculando o script ao arquivo, não anexando-o. Quem tiver acesso ao arquivo Excel também poderá exibir, executar ou fazer uma cópia do script.

Excel reconhece e executa um script somente se ele estiver em sua pasta OneDrive, uma pasta do Sharepoint ou compartilhado com a pasta de trabalho.

## <a name="onedrive"></a>OneDrive

O comportamento padrão é que Office scripts são armazenados em seu OneDrive. Os **arquivos .osts** são encontrados na **pasta /Documents/Office Scripts/**. Todas as edições feitas nesses arquivos **.osts** , como renomear ou excluir arquivos, serão refletidas no Editor de Códigos e na Galeria de Scripts.

Os scripts que são compartilhados com uma de suas pastas de trabalho permanecem na conta do criador do script OneDrive. Eles não são copiados para nenhuma das pastas locais ou OneDrive quando você executa o script compartilhado em Excel. O **botão Fazer uma Cópia** do Editor de Códigos salva uma cópia separada do script em seu OneDrive. As alterações na cópia não afetam o script original.

A menos que você compartilhe seus scripts pessoais, ninguém mais poderá accessá-los. Suas OneDrive configurações controlam o acesso compartilhado e as permissões para todos os arquivos **.osts** de script, independentemente de Excel configurações. Os scripts não podem ser vinculados de um disco local ou locais de nuvem personalizados.

## <a name="sharepoint"></a>Microsoft Office SharePoint Online

Office scripts salvos em um site SharePoint são de propriedade de sua equipe. Você e os membros da sua organização com o acesso apropriado podem executar e editar scripts SharePoint. Você também verá esses scripts aparecerem na Galeria **de** Scripts da guia Automatizar.

Para carregar um script de SharePoint, vá para Todos os **scripts** e selecione Exibir mais **scripts** na parte inferior da lista. Isso abre um seletor de arquivos em que você pode escolher arquivos **.osts** de qualquer SharePoint site ao qual você tenha acesso. Observe que os scripts SharePoint que você já abriu serão exibidos na lista de scripts recentes.

Para salvar um script SharePoint, vá para o menu **Mais opções (...)** e selecione **Salvar como**. Isso abre um seletor de arquivos no qual você pode selecionar pastas em seu SharePoint site. Salvar em um novo local cria uma cópia do script nesse local. A versão original ainda está em seu OneDrive ou em outro SharePoint local.

> [!IMPORTANT]
> Scripts com [chamadas externas](../develop/external-calls.md) não podem ser executados SharePoint. Você receberá um erro informando "Não há suporte para chamadas de acesso à rede no momento para scripts salvos em um SharePoint site".

> [!IMPORTANT]
> Power Automate **não dá** suporte a scripts armazenados SharePoint no momento.

## <a name="restore-deleted-scripts"></a>Restaurar scripts excluídos

Quando você exclui um script no Excel, ele vai para sua OneDrive ou SharePoint lixeira. Para restaurar um script excluído, siga as etapas listadas em Como recuperar itens ausentes, excluídos ou corrompidos no SharePoint e OneDrive para trabalho [ou escola](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87). Restaurar um **arquivo .osts** o retorna para a **lista Todos os scripts** .

Um script excluído é descompartilhado com a pasta de trabalho. Quando você restaura um script, ele **não retém** o acesso ao script. Você precisará compartilhar o script novamente.

Os scripts restaurados ainda funcionam conforme o esperado com Power Automate fluxos. Você não precisa recriar o conector de fluxo.

## <a name="file-ownership-and-retention"></a>Propriedade e retenção do arquivo

Office scripts seguem as políticas de retenção e exclusão especificadas por Microsoft OneDrive e Microsoft SharePoint. Para saber como lidar com scripts que foram criados e compartilhados por um usuário que está sendo removido da sua organização, consulte Saiba mais sobre retenção para SharePoint e [OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true).

Durante a edição, os arquivos são armazenados temporariamente no navegador. Você deve salvar o script antes de fechar Excel janela para salvá-lo no OneDrive local. Não se esqueça de salvar o arquivo após as edições, caso contrário, essas edições só estarão na versão do navegador do arquivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar Office uso de scripts no nível de administrador

Descubra quais locatários estão usando Office Scripts com o log de auditoria no centro de conformidade. Para saber como usar essa ferramenta, visite [Pesquisar o log de auditoria no Centro de Conformidade & Segurança](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Para localizar quem está usando Office scripts com a ferramenta de pesquisa, `.osts` adicione o campo Arquivo **, pasta ou site**. Isso pesquisa todos os arquivos com a extensão Office arquivo Scripts. Se alguém em sua organização tiver usado o Office scripts, a atividade do usuário será exibida nos resultados da pesquisa de log de auditoria.

## <a name="see-also"></a>Confira também

- [Compartilhamento de Scripts do Office no Excel para a web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Desfazer os efeitos do Scripts do Office](../testing/undo.md)
