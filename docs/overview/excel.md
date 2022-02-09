---
title: Scripts do Office no Excel na Web
description: Uma breve introdução ao Gravador de ação e ao Editor de códigos de scripts do Office.
ms.topic: overview
ms.date: 02/04/2022
ms.localizationpriority: high
ms.openlocfilehash: b4222d4f3d99ccd26dde534659681a5919088450
ms.sourcegitcommit: b84d4c8dd31335e4e39b0da6ad25fd528cb9d8f3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62462528"
---
# <a name="office-scripts-in-excel-on-the-web"></a>Scripts do Office no Excel na Web

Os scripts do Office no Excel na Web permitem automatizar suas tarefas diárias. Você pode gravar suas ações do Excel com o Gravador de Ações, que cria um script de linguagem TypeScript. Você também pode criar e editar scripts com o Editor de códigos. Os scripts podem ser compartilhados com toda a organização, para que seus colegas possam automatizar os fluxos de trabalho.

Esta série de documentos ensina como usar essas ferramentas. Você será apresentado ao Gravador de ações e verá como gravar suas ações frequentes do Excel. Você também aprenderá a criar ou atualizar seus próprios scripts com o Editor de códigos.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Requisitos

Para utilizar os Scripts do Office, você precisará do seguinte.

1. [Excel na web](https://www.office.com/launch/excel) (outras plataformas, como para desktop, não são suportadas).
1. OneDrive for Business.
1. Qualquer licença comercial ou educacional do Microsoft 365 com acesso aos aplicativos para desktop do Microsoft Office 365, como:

    - Office 365 Business
    - Office 365 Business Premium
    - Office 365 ProPlus
    - Office 365 ProPlus para dispositivos
    - Office 365 Enterprise E3
    - Office 365 Enterprise E5
    - Office 365 A3
    - Office 365 A5

> [!NOTE]
> Se você atender aos requisitos e ainda não estiver vendo a guia **Automatizar**, é possível que o seu administrador tenha desabilitado o recurso ou que haja outro problema em seu ambiente. Siga as etapas em [Guia Automatizar não aparecem ou Scripts do Office não disponíveis](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable) para começar a usar os Scripts do Office.

## <a name="when-to-use-office-scripts"></a>Quando usar Scripts do Office

Os scripts permitem gravar e reproduzir suas ações do Excel em diferentes pastas de trabalho e planilhas. Se você perceber que vive fazendo as mesmas coisas o tempo inteiro, experimente transformar todo esse trabalho em um Script do Office fácil de executar. Execute seu script com um apertar de botão no Excel ou combine-o com o Power Automate para simplificar todo o fluxo de trabalho.

Como exemplo, digamos que você comece seu dia de trabalho abrindo um arquivo .csv em um site de contabilidade no Excel. Então você gasta alguns minutos excluindo colunas desnecessárias, formatando uma tabela, adicionando fórmulas e criando uma tabela dinâmica em uma nova planilha. As ações repetidas diariamente podem ser gravadas uma vez com o Gravador de ações. A partir daí, a execução do script cuidará da sua conversão .csv. Além de remover o risco de esquecer as etapas, você poderá compartilhar seu processo com outras pessoas sem precisar ensinar nada a elas. Os Scripts do Office permitem que você automatize suas tarefas comuns para que você e seu local de trabalho possam ser mais eficientes e produtivos.

## <a name="action-recorder"></a>Gravador de ações

:::image type="content" source="../images/action-recorder-intro.png" alt-text="Uma lista de ações gravada pelo Gravador de Ações.":::

O Gravador de Ações registra as ações que você executa no Excel e as salva na forma de um script. Com o Gravador de ações em execução, você pode capturar as ações do Excel enquanto edita células, altera a formatação e cria tabelas. O script resultante pode ser executado em outras planilhas e pastas de trabalho para recriar suas ações originais.

## <a name="code-editor"></a>Editor de códigos

:::image type="content" source="../images/code-editor-intro.png" alt-text="O Editor de Código mostrando o código de script usado neste tutorial.":::

Todos os scripts gravados com o Gravador de ações podem ser editados através do Editor de códigos. Isso permite que você ajuste e personalize o script para melhor atender às suas necessidades. Você também pode adicionar lógica e funcionalidade que não são acessíveis de forma direta pela interface do usuário do Excel, como instruções condicionais (se/senão) e loops.

> [!TIP]
> O Gravador de Ações tem um botão **Copiar como código** para registrar as ações no código do script sem salvar o script inteiro.
>
> :::image type="content" source="../images/action-recorder-copy-code.png" alt-text="O painel de tarefas do Gravador de Ações com o botão 'Copiar como código' destacado.":::

Nossos [tutoriais](../tutorials/excel-tutorial.md) fornecem uma maneira orientada e estruturada de aprender as funcionalidades dos Scripts do Office. Depois de concluir os tutoriais, leia [Fundamentos de script para os Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md) para saber mais sobre o Editor de Código e como escrever e editar seus próprios scripts. Para informações adicionais sobre o Editor de Código e como seu código de script é interpretado, leia [Ambiente do Editor de Código de Scripts do Office](code-editor-environment.md).

## <a name="share-scripts"></a>Compartilhar scripts

:::image type="content" source="../images/script-sharing.png" alt-text="A página de detalhes do script mostrando a opção &quot;Compartilhar com outras pessoas nesta pasta de trabalho&quot;.":::

Os scripts do Office podem ser compartilhados com outros usuários de uma pasta de trabalho do Excel. Quando você compartilha um script em uma pasta de trabalho compartilhada, todos com acesso à pasta de trabalho também podem visualizar e executar seu script.

Para obter mais detalhes sobre como compartilhar e cancelar o compartilhamento de scripts, confira [Compartilhando Scripts do Office no Excel para a Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b).

> [!NOTE]
> Saiba mais sobre como os scripts são armazenados no seu OneDrive em [Armazenamento de arquivos e propriedade de Scripts do Office](script-storage.md).

### <a name="create-script-buttons"></a>Criar botões de script

[!INCLUDE [Rollout information](../includes/rollout-note.md)]

:::image type="content" source="../images/run-from-button.png" alt-text="Um botão na planilha que executa um script quando clicado.":::

Ajude seus colegas a encontrar e executar seus scripts adicionando botões de script a uma pasta de trabalho. Com qualquer script, vá para o menu Mais opções **(...)** na página Detalhes do script ou no painel de tarefas do Editor de Código e selecione **Adicionar botão**. Isso cria um botão na guia de trabalho que executa o script associado quando selecionado. Ele também compartilha o script com a pasta de trabalho, para que todos com permissões de gravação para a pasta de trabalho possam usar sua automação útil.

A captura de tela a seguir mostra a página Detalhes do script de um script intitulado **Criar Relatório** e tem a opção **Adicionar botão** no menu **Mais opções** (...) realçada.

:::image type="content" source="../images/add-button.png" alt-text="A opção 'Adicionar botão' no menu da página Detalhes do script.":::

Esses botões de script também funcionam no Windows. Crie o botão no Excel na Web e os usuários no Windows podem executar o script pressionando um botão. Observe que só há suporte para a execução do botão no Windows. Se você precisar editar o script, precisará fazer isso por meio do aplicativo Web.

> [!NOTE]
> Para usar esse recurso no Windows, você precisa de uma cópia do Excel conectada a uma Microsoft 365 assinatura. Plataformas sem suporte ou versões mais antigas do Office exibem a forma usada para o botão de script, mas não é clicável.

Para parar de compartilhar um script por meio de um botão, vá para o menu Mais opções **(...)** na página Detalhes do script e selecione **Parar de Compartilhar**. Isso remove todos os botões que executem o script. A exclusão de um único botão remove o script desse botão, mesmo que a operação seja desfeita ou o botão seja cortado e passado.

## <a name="connect-office-scripts-to-power-automate"></a>Conectar Scripts do Office com o Power Automate

[O Power Automate](https://flow.microsoft.com/) é um serviço que ajuda você a criar fluxos de trabalho automatizados entre vários aplicativos e serviços. Os scripts do Office podem ser usados nesses fluxos de trabalho, permitindo que você controle seus scripts fora da pasta de trabalho. Você pode executar seus scripts em um cronograma, dispará-los em resposta a emails e muito mais. Visite o [Executar scripts do Office no Excel na Web com o tutorial do Power Automate](../tutorials/excel-power-automate-manual.md) para aprender como se conectar a esses serviços de automação.

## <a name="next-steps"></a>Próximas etapas

Conclua o [tutorial de Scripts do Office no Excel na web](../tutorials/excel-tutorial.md) para saber como criar seu primeiro script.

## <a name="see-also"></a>Confira também

- [Fundamentos de script para scripts do Office no Excel na Web](../develop/scripting-fundamentals.md)
- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Configurações dos scripts do Office no M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Introdução aos Scripts do Office no Excel](https://support.microsoft.com/office/9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Compartilhamento de Scripts do Office no Excel para a web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Centro de Desenvolvimento de Scripts do Office](https://developer.microsoft.com/office-scripts)
