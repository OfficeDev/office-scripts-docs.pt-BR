---
title: Solução de problemas dos scripts do Office
description: Dicas e técnicas de depuração para scripts do Office, bem como recursos da ajuda.
ms.date: 10/30/2020
localization_priority: Normal
ms.openlocfilehash: b45957bd336edce527397253cacec8cb09df715a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342875"
---
# <a name="troubleshooting-office-scripts"></a>Solução de problemas dos scripts do Office

Ao desenvolver scripts do Office, você pode cometer erros. Não há problema. Temos ferramentas que ajudam a encontrar os problemas e a fazer com que seus scripts funcionem perfeitamente.

## <a name="console-logs"></a>Logs do console

Às vezes, durante a solução de problemas, convém imprimir mensagens na tela. Eles podem mostrar o valor atual de variáveis ou quais caminhos de código estão sendo disparados. Para fazer isso, faça o log do texto no console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

As cadeias de caracteres passadas para `console.log` serão exibidas no console de registro em log do editor de código. Para ativar o console, pressione o botão **reticências** e selecione **logs...**

Os logs não afetam a pasta de trabalho.

## <a name="error-messages"></a>Mensagens de erro

Quando o script do Excel encontra um problema em execução, ele produz um erro. Você verá um pop-up de solicitação perguntando se deseja **exibir os logs**. Pressione esse botão para abrir o console e exibir quaisquer erros.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Guia automatizar não aparecendo ou os scripts do Office não estão disponíveis

As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **automatizada** não aparecendo no Excel na Web.

1. [Verifique se a licença do Microsoft 365 inclui scripts do Office](../overview/excel.md#requirements).
1. [Verifique se há suporte para o seu navegador](platform-limits.md#browser-support).
1. [Verifique se os cookies de terceiros estão habilitados](platform-limits.md#third-party-cookies).
1. [Certifique-se de que o administrador não desabilitou scripts do Office no centro de administração do Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>Recursos de ajuda

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores que desejam ajudar com problemas de codificação. Muitas vezes, você poderá encontrar a solução para o problema por meio de uma pesquisa rápida de estouro de pilha. Caso contrário, faça a pergunta e marque-a com a marca "Office-scripts". Não deixe de mencionar que você está criando um *script* do Office, não um *suplemento* do Office.

Se você encontrar um problema com a API JavaScript do Office, crie um problema no repositório do GitHub do [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) . Os membros da equipe do produto responderão a problemas e fornecerão mais assistência. A criação de um problema no repositório **OfficeDev/Office-js** indica que você encontrou uma falha na biblioteca da API JavaScript do Office para a qual a equipe de produto deve tratar.

Se houver um problema com o gravador de ação ou editor, envie comentários através do botão **ajuda > comentários** no Excel.

## <a name="see-also"></a>Também consulte

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Conceitos básicos de script para scripts do Office no Excel na Web](../develop/scripting-fundamentals.md)
- [Limites de plataforma com scripts do Office](platform-limits.md)
- [Melhorar o desempenho dos scripts do Office](../develop/web-client-performance.md)
- [Desfazer os efeitos de um script do Office](undo.md)
