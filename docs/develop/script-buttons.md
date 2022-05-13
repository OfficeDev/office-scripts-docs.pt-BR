---
title: Executar Office scripts em Excel com botões
description: Adicione botões a pastas de trabalho que controlam Office scripts no Excel.
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: fde34d62f9abe897a8b93195ab37a75cfc73f619
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393681"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>Executar Office scripts em Excel com botões

Ajude seus colegas a encontrar e executar seus scripts adicionando botões de script a uma pasta de trabalho.

:::image type="content" source="../images/run-from-button.png" alt-text="Um botão na planilha que executa um script quando clicado.":::

## <a name="create-script-buttons"></a>Criar botões de script

Com qualquer script, vá para o menu Mais opções **(...)** na página de detalhes do script ou no painel de tarefas do Editor de Códigos e selecione **o botão Adicionar**. Isso cria um botão na guia de trabalho que executa o script associado quando selecionado. Ele também compartilha o script com a pasta de trabalho, para que todos com permissões de gravação para a pasta de trabalho possam usar sua automação útil.

A captura de tela a seguir mostra a página de detalhes do script de  um script intitulado Criar Tabela Dinâmica  e tem a opção de botão Adicionar no menu Mais opções **(...)** realçada.

:::image type="content" source="../images/add-button.png" alt-text="A opção 'Adicionar botão' no menu da página de detalhes do script.":::

## <a name="remove-script-buttons"></a>Remover botões de script

Para interromper o compartilhamento de um script por meio de um botão, vá para o menu Mais opções **(...)** na página de detalhes do script e selecione **Parar de compartilhar**. Isso remove todos os botões que executem o script. A exclusão de um único botão remove o script desse botão, mesmo que a operação seja desfeita ou o botão seja cortado e passado.

## <a name="script-buttons-with-excel-on-windows"></a>Botões de script com Excel no Windows

Esses botões de script também funcionam no Windows. Crie o botão no Excel na Web e os usuários Windows podem executar o script com o clique de um botão. Observe que você não pode editar scripts em Excel no Windows. Você só pode editar scripts em Excel na Web.

Algumas Office APIs de Scripts podem não ter suporte Excel em Windows, especialmente builds mais antigos. Elas incluem APIs e APIs mais recentes para recursos somente da Web. Se um script contiver APIs sem suporte, o script não será executado e, em vez disso, o painel de tarefas Status de Execução de **Script** exibirá uma mensagem de aviso informando: "Este script deve ser executado no momento Excel para a Web. Abra a pasta de trabalho no navegador e tente novamente ou entre em contato com o proprietário do script para obter ajuda."  

> [!IMPORTANT]
> Os botões de script [exigem que o WebView2](/deployoffice/webview2-install) funcione com Excel no Windows. Isso é instalado por padrão com as versões mais recentes do Excel na Área de Trabalho, mas se você não conseguir clicar em botões de scripts, visite Baixar o [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/#download-section) e baixe o mecanismo do navegador.
