---
title: Propriedade e armazenamento de arquivos de Scripts do Office
description: Informações sobre como os Scripts do Office são armazenados no Microsoft OneDrive e transferidos entre proprietários.
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 573f65f299c29b4f481c9a2e23ebe7e36181706b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572504"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Propriedade e armazenamento de arquivos de Scripts do Office

Os Scripts do Office são armazenados como **arquivos .osts** no Microsoft OneDrive ou em uma pasta do SharePoint. Eles são armazenados separadamente de uma pasta de trabalho. Para dar aos usuários que estão fora do site do SharePoint acesso ao script, [compartilhe o script com uma pasta de trabalho do Excel](excel.md#share-office-scripts). Isso significa que você está vinculando o script ao arquivo, não anexando-o. Quem tiver acesso ao arquivo do Excel também poderá exibir, executar ou fazer uma cópia do script.

O Excel só reconhecerá e executará um script se ele estiver em sua pasta do OneDrive, em uma pasta do Sharepoint ou compartilhado com a pasta de trabalho.

## <a name="onedrive"></a>OneDrive

O comportamento padrão é que os Scripts do Office são armazenados em seu OneDrive. Os **arquivos .osts** são encontrados na **pasta /Documents/Office Scripts/** . Todas as edições feitas nesses arquivos **.osts** , como renomear ou excluir arquivos, serão refletidas no Editor de Códigos e na Galeria de Scripts.

Os scripts compartilhados com uma de suas pastas de trabalho permanecem no OneDrive do criador do script. Eles não são copiados para nenhuma das pastas locais ou do OneDrive quando você executa o script compartilhado no Excel. O **botão Fazer uma Cópia** do Editor de Códigos salva uma cópia separada do script no OneDrive. As alterações na cópia não afetam o script original.

A menos que você compartilhe seus scripts pessoais, ninguém mais poderá accessá-los. Suas configurações do OneDrive controlam o acesso compartilhado e as permissões para todos os arquivos **.osts de** script, independentemente de qualquer configuração do Excel. Os scripts não podem ser vinculados de um disco local ou locais de nuvem personalizados.

## <a name="sharepoint"></a>SharePoint

Os Scripts do Office salvos em um site do SharePoint pertencem à sua equipe. Você e os membros da sua organização com o acesso apropriado podem executar e editar scripts do SharePoint. Você também verá esses scripts aparecerem na Galeria **de** Scripts da guia Automatizar.

Para carregar um script do SharePoint, acesse **Todos os scripts** e selecione Exibir mais **scripts** na parte inferior da lista. Isso abre um seletor de arquivos no qual você pode escolher arquivos **.osts** de qualquer site do SharePoint ao qual você tenha acesso. Observe que os scripts do SharePoint que você já abriu serão exibidos na lista de scripts recentes.

Para salvar um script no SharePoint, vá para o menu **Mais opções (...)** e selecione **Salvar como**. Isso abre um seletor de arquivos no qual você pode selecionar pastas em seu site do SharePoint. Salvar em um novo local cria uma cópia do script nesse local. A versão original ainda está em seu OneDrive ou em outro local do SharePoint.

> [!IMPORTANT]
> Scripts com [chamadas externas](../develop/external-calls.md) não podem ser executados do SharePoint. Você receberá um erro informando "Não há suporte para chamadas de acesso à rede no momento para scripts salvos em um site do SharePoint".

> [!IMPORTANT]
> O Power Automate **não dá** suporte a scripts armazenados no SharePoint no momento.

## <a name="restore-deleted-scripts"></a>Restaurar scripts excluídos

Quando você exclui um script no Excel, ele vai para a lixeira do OneDrive ou do SharePoint. Para restaurar um script excluído, siga as etapas listadas em Como recuperar itens ausentes, excluídos ou corrompidos no [SharePoint e no OneDrive para trabalho ou escola](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87). Restaurar um **arquivo .osts** o retorna para a **lista Todos os scripts** .

Um script excluído é descompartilhado com a pasta de trabalho. Quando você restaura um script, ele **não retém** o acesso ao script. Você precisará compartilhar o script novamente.

Os scripts restaurados ainda funcionam conforme o esperado com fluxos do Power Automate. Você não precisa recriar o conector de fluxo.

## <a name="file-ownership-and-retention"></a>Propriedade e retenção do arquivo

Os Scripts do Office seguem as políticas de retenção e exclusão especificadas pelo Microsoft OneDrive e pelo Microsoft SharePoint. Para saber como lidar com scripts que foram criados e compartilhados por um usuário que está sendo removido da sua organização, consulte Saiba mais sobre a retenção para [o SharePoint e o OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true).

Durante a edição, os arquivos são armazenados temporariamente no navegador. Você deve salvar o script antes de fechar a janela do Excel para salvá-la no local do OneDrive. Não se esqueça de salvar o arquivo após as edições, caso contrário, essas edições só estarão na versão do navegador do arquivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar o uso de Scripts do Office no nível de administrador

Descubra quem está usando Scripts do Office em sua organização com o log de auditoria do centro de conformidade. Detalhes sobre o log de auditoria são encontrados em [Pesquisar o log de auditoria no Centro de Conformidade do & Segurança](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Para auditar especificamente a atividade relacionada aos Scripts do Office como administrador, execute as etapas a seguir.

1. Em uma janela do navegador InPrivate (ou Incognito ou outro modo de acompanhamento limitado específico do navegador), abra e faça logon no [Centro de conformidade](https://compliance.microsoft.com/).
1. Vá para a **página Auditoria** .
1. *(Somente uma vez)* Na guia **Pesquisar** , selecione **Iniciar gravação de atividade de usuário e administrador**.

    > [!IMPORTANT]
    > Pode levar uma ou duas horas depois de ativar a gravação antes que todas as atividades no locatário sejam gravadas.

1. Defina as opções de pesquisa desejadas e pressione **Pesquisar**. **Filtre atividades** para **o script Executado na pasta de** trabalho para ver sempre que um script foi executado. Você também pode filtrar **o campo Arquivo, pasta ou site** para `.osts`. Isso revela quem em sua organização está criando ou modificando scripts.

    :::image type="content" source="../images/audit-log-example.png" alt-text="Algumas linhas de resultados da pesquisa de log de auditoria, incluindo a ação 'Executar script na pasta de trabalho' e o upload e a modificação de um arquivo .osts.":::

## <a name="see-also"></a>Confira também

- [Compartilhamento de Scripts do Office no Excel para a web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Desfazer os efeitos do Scripts do Office](../testing/undo.md)
