---
title: Office Ambiente do Editor de Código de Scripts
description: Os pré-requisitos e informações de ambiente para Office scripts em Excel na Web.
ms.date: 05/27/2021
localization_priority: Normal
ms.openlocfilehash: 5b2f7afa193dc71e13a3d6763c9e8ff8344ee3e8be18e7e996f8431e03510509
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847090"
---
# <a name="office-scripts-code-editor-environment"></a>Office Ambiente do Editor de Código de Scripts

Office Os scripts são escritos em TypeScript ou JavaScript e usam as APIs JavaScript Office scripts para interagir com uma Excel de trabalho. O Editor de Código baseia-se Visual Studio Code, portanto, se você já usou esse ambiente antes, se sentirá em casa.

## <a name="scripting-language-typescript-or-javascript"></a>Idioma de script: TypeScript ou JavaScript

Os Scripts do Office são escritos em [TypeScript](https://www.typescriptlang.org/docs/home.html), que é um superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). O Gravador de Ações gera código em TypeScript e a documentação Office Scripts usa TypeScript. Como TypeScript é um superconjunto de JavaScript, qualquer código de script que você escrever em JavaScript funcionará muito bem.

Office Scripts são partes de código amplamente autoconstrutivas. Apenas uma pequena parte da funcionalidade do TypeScript é usada. Portanto, você pode editar scripts sem ter que aprender as complexidades de TypeScript. O Editor de Código também lida com a instalação, a compilação e a execução do código, para que você não precise se preocupar com nada além do próprio script. É possível aprender o idioma e criar scripts sem conhecimento de programação anterior. No entanto, se você for novo na programação, recomendamos aprender alguns conceitos básicos antes de prosseguir com Office Scripts:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Office Os scripts usam uma versão especializada das APIs javaScript Office para [Office de Office.](/office/dev/add-ins/overview/index) Embora haja semelhanças nas duas APIs, você não deve supor que o código possa ser portado entre as duas plataformas. As diferenças entre as duas plataformas são descritas no artigo Diferenças entre Office [scripts e Office de complementos.](../resources/add-ins-differences.md#apis) Você pode exibir todas as APIs disponíveis para seu script na documentação de referência da API Office [Scripts.](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>Suporte à biblioteca externa

Office Scripts não suportam o uso de bibliotecas JavaScript externas de terceiros. Atualmente, você não pode chamar qualquer biblioteca que não seja Office scripts de um script. Você ainda tem acesso a qualquer [objeto JavaScript interno,](../develop/javascript-objects.md)como [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense é um conjunto de recursos do Editor de Código que ajudam você a escrever código. Ele fornece realce automático, realce de erro de sintaxe e documentação de API em linha.

IntelliSense sugestões conforme você digita, semelhante ao texto sugerido em Excel. Pressionar a tecla Tab ou Enter insere o membro sugerido. Acionar IntelliSense local atual do cursor pressionando as teclas Ctrl+Space. Essas sugestões são especialmente úteis ao concluir um método. A assinatura do método exibida pelo IntelliSense contém uma lista de argumentos necessários, o tipo de cada argumento, se um determinado argumento é obrigatório ou opcional e o tipo de retorno do método.

Passe o cursor sobre um método, classe ou outro objeto de código para ver mais informações. Passe o mouse sobre um erro de sintaxe ou sugestão de código, representado por uma linha vermelha ou amarela, para ver sugestões sobre como corrigir o problema. Geralmente, IntelliSense fornece uma opção "Correção Rápida" para alterar automaticamente o código.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Uma mensagem de erro no texto de foco do Editor de Código com um botão &quot;Correção Rápida&quot;.":::

O Office de Código de Scripts usa o mesmo mecanismo IntelliSense que Visual Studio Code. Para saber mais sobre o recurso, visite os recursos Visual Studio Code [do IntelliSense.](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)

## <a name="keyboard-shortcuts"></a>Atalhos de teclado

A maioria dos atalhos de teclado para Visual Studio Code também funcionam no Editor de Código Office Scripts. Use os PDFs a seguir para saber mais sobre as opções disponíveis e aproveitar ao máximo o Editor de Código:

- [Atalhos de teclado para macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Atalhos de teclado para Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Confira também

- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Usar objetos internos do JavaScript nos scripts do Office](../develop/javascript-objects.md)
