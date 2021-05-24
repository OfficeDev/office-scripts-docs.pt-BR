---
title: Office Ambiente do Editor de Código de Scripts
description: Os pré-requisitos e informações de ambiente para Office scripts em Excel na Web.
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

Office Os scripts são escritos em TypeScript ou JavaScript e usam as APIs JavaScript Office scripts para interagir com uma Excel de trabalho. O Editor de Código baseia-se Visual Studio Code, portanto, se você já usou esse ambiente antes, se sentirá em casa.

## <a name="scripting-language-typescript-or-javascript"></a>Idioma de script: TypeScript ou JavaScript

Os Scripts do Office são escritos em [TypeScript](https://www.typescriptlang.org/docs/home.html), que é um superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). O Gravador de Ações gera código em TypeScript e a documentação Office Scripts usa TypeScript. Como TypeScript é um superconjunto de JavaScript, qualquer código de script que você escrever em JavaScript funcionará muito bem.

Office Scripts são partes de código amplamente autoconstrutivas. Apenas uma pequena parte da funcionalidade do TypeScript é usada. Portanto, você pode editar scripts sem ter que aprender as complexidades de TypeScript. O Editor de Código também lida com a instalação, a compilação e a execução do código, para que você não precise se preocupar com nada além do próprio script. É possível aprender o idioma e criar scripts sem conhecimento de programação anterior. No entanto, se você for novo na programação, recomendamos aprender alguns conceitos básicos antes de prosseguir com Office Scripts:

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Office Os scripts usam uma versão especializada das APIs javaScript Office para [Office de Office.](/office/dev/add-ins/overview/index) Embora haja semelhanças nas duas APIs, você não deve supor que o código possa ser portado entre as duas plataformas. As diferenças entre as duas plataformas são descritas no artigo Diferenças entre Office [scripts e Office de complementos.](../resources/add-ins-differences.md#apis) Você pode exibir todas as APIs disponíveis para seu script na documentação de referência da API Office [Scripts.](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>Suporte à biblioteca externa

Office Scripts não suportam o uso de bibliotecas JavaScript externas de terceiros. Atualmente, você não pode chamar qualquer biblioteca que não seja Office scripts de um script. Você ainda tem acesso a qualquer [objeto JavaScript interno,](../develop/javascript-objects.md)como [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense é um recurso editor de código que ajuda a evitar erros de digitação e sintaxe à medida que você edita o script. Ele exibe possíveis nomes de objetos e campos conforme você digita, bem como a documentação em linha para cada API.

O Excel de código usa o mesmo mecanismo IntelliSense que Visual Studio Code. Para saber mais sobre o recurso, [visite Visual Studio Code de IntelliSense Recursos.](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)

## <a name="keyboard-shortcuts"></a>Atalhos de teclado

A maioria dos atalhos de teclado para Visual Studio Code também funcionam no Editor de Código Office Scripts. Use os PDFs a seguir para saber mais sobre as opções disponíveis e aproveitar ao máximo o Editor de Código:

- [Atalhos de teclado para macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Atalhos de teclado para Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Confira também

- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Usar objetos internos do JavaScript nos scripts do Office](../develop/javascript-objects.md)
