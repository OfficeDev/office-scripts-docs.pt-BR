---
title: Visual Studio Code para Scripts do Office (versão prévia)
description: Como configurar o Editor de Código de Scripts do Office para se conectar com o VS Code para a Web.
ms.date: 11/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: fd9dd417610c8ad64fbd3fc50048ce56afdb4e28
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892027"
---
# <a name="visual-studio-code-for-office-scripts-preview"></a>Visual Studio Code para Scripts do Office (versão prévia)

[Visual Studio Code para a Web](https://vscode.dev/) permite que os usuários editem qualquer coisa de qualquer lugar. Conecte sua experiência de Scripts do Office a este editor de código popular para iniciar o script fora da pasta de trabalho.

:::image type="content" source="../images/vscode-script-editor.png" alt-text="Uma janela Excel na Web com o Editor de Código aberta ao lado de um VS Code na janela da Web com um script aberto.":::

Visual Studio Code tem algumas vantagens sobre o Editor de Código interno.

- Edição em tela inteira! Seu script não precisa mais compartilhar espaço na tela com a pasta de trabalho.
- Edite vários scripts ao mesmo tempo! Alternar rapidamente entre scripts para compartilhar código de suas outras automações.
- Extensões! Use suas extensões de VS Code favoritas para verificação ortográfica, formatação e qualquer outra coisa que ajude você a fazer o trabalho.

> [!NOTE]
> Este recurso está em versão prévia. Ele está sujeito a alterações com base nos comentários. Se você encontrar algum problema, denuncie-os por meio do botão **Comentários** no Excel. Veja a seguir problemas conhecidos com a versão atual do recurso.
>
> - Visual Studio Code só pode ser conectado aos Scripts do Office por meio de Excel na Web.
> - Essa conexão de Scripts do Office só está disponível com clientes do Excel em inglês.

## <a name="connect-visual-studio-code-to-office-scripts"></a>Conectar Visual Studio Code a scripts do Office

Siga estas etapas pontuais para conectar Visual Studio Code e Excel na Web.

1. Abra o **Editor de Código** de Scripts do Office.
2. No menu **Mais opções (...),** selecione **Configurações do Editor**.
3. Selecione **(versão prévia) Visual Studio Code conexão**.

:::image type="content" source="../images/vscode-enable-option.png" alt-text="O painel de tarefas configurações do editor mostrando uma caixa de seleção rotulada Visual Studio Code conexão.":::

Agora você pode editar e executar seus scripts de Visual Studio Code. Em qualquer script, acesse o menu **Mais opções (...)** e selecione **Abrir no VS Code**.

:::image type="content" source="../images/vscode-open-option.png" alt-text="A opção Abrir no VS Code que está sendo selecionada em uma lista ao lado de um script aberto.":::

## <a name="see-also"></a>Confira também

- [Ambiente do Editor de Código de Scripts do Office](../overview/code-editor-environment.md)
- [Visual Studio Code para a Web (documentação)](https://code.visualstudio.com/docs/editor/vscode-web)
