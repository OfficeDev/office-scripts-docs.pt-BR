---
title: Office Ambiente do Editor de Código de Scripts
description: Os pré-requisitos e informações ambientais para Office Scripts em Excel na Web.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545819"
---
# <a name="office-scripts-code-editor-environment"></a>Office Ambiente do Editor de Código de Scripts

Office Os scripts são gravados em TypeScript ou JavaScript e usam as APIs JavaScript Office Scripts para interagir com uma Excel livro de trabalho. O Editor de Código é baseado em Visual Studio Code, então se você já usou esse ambiente antes, você vai se sentir em casa.

## <a name="scripting-language-typescript-or-javascript"></a>Linguagem de script: TypeScript ou JavaScript

Office Os scripts são escritos no [TypeScript](https://www.typescriptlang.org/docs/home.html), que é um superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). O Gravador de Ação gera código no TypeScript e a documentação Office Scripts usa TypeScript. Uma vez que o TypeScript é um superconjunto de JavaScript, qualquer código de script que você escrever no JavaScript funcionará muito bem.

Office Scripts são em grande parte peças de código independentes. Apenas uma pequena parte da funcionalidade do TypeScript é usada. Portanto, você pode editar scripts sem ter que aprender os meandros do TypeScript. O Editor de Códigos também lida com a instalação, compilação e execução de código, para que você não precise se preocupar com nada além do script em si. É possível aprender a língua e criar roteiros sem conhecimento prévio de programação. No entanto, se você é novo na programação, recomendamos aprender alguns fundamentos antes de prosseguir com Office Scripts:

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Office Os scripts usam uma versão especializada das APIs javascript Office para [Office Add-ins](/office/dev/add-ins/overview/index). Embora existam semelhanças nas duas APIs, você não deve assumir que o código pode ser portado entre as duas plataformas. As diferenças entre as duas plataformas estão descritas nas Diferenças entre Office Scripts e Office artigo [Add-ins.](../resources/add-ins-differences.md#apis) Você pode visualizar todas as APIs disponíveis para o seu script na [documentação de referência da API scripts Office](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Suporte externo à biblioteca

Office Scripts não suportam o uso de bibliotecas JavaScript externas de terceiros. Atualmente, você não pode chamar nenhuma biblioteca além das APIs de scripts Office de um script. Você ainda tem acesso a qualquer [objeto JavaScript embutido,](../develop/javascript-objects.md)como [Matemática](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense é um recurso do Editor de Código que ajuda a evitar erros de digitação e sintaxe à medida que você edita seu script. Ele exibe possíveis nomes de objeto e campo à medida que você digita, bem como documentação inline para cada API.

O Excel Code Editor usa o mesmo motor IntelliSense que Visual Studio Code. Para saber mais sobre o recurso, visite [os recursos de IntelliSense da Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Atalhos de teclado

A maioria dos atalhos de teclado para Visual Studio Code também funcionam no Office Scripts Code Editor. Use os seguintes PDFs para saber sobre as opções disponíveis e aproveitar ao máximo o Editor de Código:

- [Atalhos de teclado para macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Atalhos de teclado para Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Confira também

- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Usar objetos internos do JavaScript nos scripts do Office](../develop/javascript-objects.md)
