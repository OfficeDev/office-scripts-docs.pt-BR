---
title: Ambiente do Editor de Código de Scripts do Office
description: Os pré-requisitos e as informações de ambiente para scripts do Office em Excel na Web.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891250"
---
# <a name="office-scripts-code-editor-environment"></a>Ambiente do Editor de Código de Scripts do Office

Os Scripts do Office são escritos em TypeScript ou JavaScript e usam as APIs JavaScript de Scripts do Office para interagir com uma pasta de trabalho do Excel. O Editor de Código é baseado em Visual Studio Code, portanto, se você já usou esse ambiente antes, você se sentirá em casa.

> [!TIP]
> Se você estiver familiarizado com Visual Studio Code, agora poderá usá-lo para escrever scripts. Visite [Visual Studio Code para Scripts do Office (versão prévia)](../develop/vscode-for-scripts.md) para experimentar esse recurso.

## <a name="scripting-language-typescript-or-javascript"></a>Linguagem de script: TypeScript ou JavaScript

Os Scripts do Office são escritos em [TypeScript](https://www.typescriptlang.org/docs/home.html), que é um superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). O Gravador de Ações gera código no TypeScript e a documentação do Office Scripts usa TypeScript. Como TypeScript é um superconjunto de JavaScript, qualquer código de script que você gravar no JavaScript funcionará muito bem.

Os Scripts do Office são em grande parte peças de código independentes. Apenas uma pequena parte da funcionalidade do TypeScript é usada. Portanto, você pode editar scripts sem precisar aprender os meandros do TypeScript. O Editor de Código também manipula a instalação, a compilação e a execução do código, portanto, você não precisa se preocupar com nada além do script em si. É possível aprender o idioma e criar scripts sem conhecimento de programação anterior. No entanto, se você for novo na programação, recomendamos aprender alguns fundamentos antes de prosseguir com os Scripts do Office:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Os Scripts do Office usam uma versão especializada das APIs JavaScript do Office para [Suplementos do Office](/office/dev/add-ins/overview/index). Embora haja semelhanças nas duas APIs, você não deve supor que o código possa ser portado entre as duas plataformas. As diferenças entre as duas plataformas são descritas no artigo [Diferenças entre Scripts do Office e Suplementos do Office](../resources/add-ins-differences.md#apis) . Você pode exibir todas as APIs disponíveis para o script na [documentação de referência da API de Scripts do Office](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Suporte à biblioteca externa

Os Scripts do Office não dão suporte ao uso de bibliotecas JavaScript externas de terceiros. Atualmente, você não pode chamar nenhuma biblioteca diferente das APIs de Scripts do Office de um script. Você ainda tem acesso a qualquer [objeto JavaScript interno](../develop/javascript-objects.md), como [Matemática](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>Intellisense

O IntelliSense é um conjunto de recursos do Editor de Código que ajudam você a escrever código. Ele fornece realce de erro de sintaxe e preenchimento automático e documentação de API embutida.

O IntelliSense dá sugestões conforme você digita, semelhante ao texto sugerido no Excel. Pressionar a tecla Tab ou Enter insere o membro sugerido. Acione o IntelliSense no local atual do cursor pressionando as teclas Ctrl+Space. Essas sugestões são especialmente úteis ao concluir um método. A assinatura do método exibida pelo IntelliSense contém uma lista de argumentos necessários, o tipo de cada argumento, se um determinado argumento é necessário ou opcional e o tipo de retorno do método.

Passe o cursor sobre um método, classe ou outro objeto de código para ver mais informações. Passe o mouse sobre um erro de sintaxe ou uma sugestão de código, representada por uma linha vermelha ou amarela, para ver sugestões sobre como corrigir o problema. Muitas vezes, o IntelliSense fornece uma opção "Correção Rápida" para alterar automaticamente o código.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Uma mensagem de erro no texto de mouse do Editor de Código com um botão 'Correção Rápida'.":::

O Editor de Código de Scripts do Office usa o mesmo mecanismo IntelliSense que Visual Studio Code. Para saber mais sobre o recurso, visite [os recursos do IntelliSense do Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Atalhos de teclado

A maioria dos atalhos de teclado para Visual Studio Code também funcionam no Editor de Código de Scripts do Office. Use os seguintes PDFs para saber mais sobre as opções disponíveis e aproveitar ao máximo o Editor de Código:

- [Atalhos de teclado para macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Atalhos de teclado para Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Confira também

- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Usar objetos internos do JavaScript nos scripts do Office](../develop/javascript-objects.md)
- [Visual Studio Code para Scripts do Office (versão prévia)](../develop/vscode-for-scripts.md)
