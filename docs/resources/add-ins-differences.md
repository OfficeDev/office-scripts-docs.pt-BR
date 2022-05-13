---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: O comportamento e as diferenças de API entre Office scripts e Office suplementos.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd483f928e3e153b8a08537f6b333c3ea8d724dd
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393618"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre os scripts do Office e os suplementos do Office

Entenda as diferenças entre Office scripts e Office suplementos para saber quando usar cada um. Office scripts são projetados para serem feitos rapidamente por qualquer pessoa que deseja melhorar seu fluxo de trabalho. Office suplementos se integram à interface do usuário do Office para uma experiência mais interativa por meio de botões da faixa de opções e painéis de tarefas. Office suplementos também podem expandir funções Excel internas fornecendo funções personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes Office de extensibilidade. Scripts Office e suplementos web Office estão focados na Web e na colaboração, mas os scripts do Office atendem aos usuários finais (enquanto os suplementos da Web Office são destinados a desenvolvedores profissionais).":::

Office scripts são executados até a conclusão com um pressionamento de botão manual ou como uma etapa no [Power Automate](https://flow.microsoft.com/), enquanto os suplementos do Office continuam em execução dependendo de como eles estão configurados. Por exemplo, você pode configurar um Office suplemento para continuar em execução mesmo quando o painel de tarefas estiver fechado. Isso significa que Office suplementos mantêm o estado durante uma sessão, enquanto Office scripts não mantêm um estado interno entre execuções. Se a solução que você está criando exigir um estado mantido, visite a documentação de [suplementos do Office](/office/dev/add-ins) para saber mais sobre Office suplementos.

O restante deste artigo descreve as principais diferenças entre Office suplementos e Office scripts.

## <a name="platform-support"></a>Suporte à plataforma

Office suplementos são multiplataforma. Eles funcionam em Windows desktop, Mac, iOS e plataformas Web e fornecem a mesma experiência em cada uma. Qualquer exceção a isso é notada na documentação da API individual.

Office scripts atualmente só têm suporte para Excel na Web. Todo o gerenciamento de gravação, edição e script é feito na plataforma Web.

### <a name="script-support-for-excel-on-windows"></a>Suporte a script para Excel no Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>APIs

Embora as APIs Office JavaScript para suplementos Office e as APIs de scripts do Office compartilhem algumas funcionalidades, elas são plataformas diferentes. As APIs Office scripts são um subconjunto otimizado e síncrono do modelo Excel API JavaScript. A principal diferença é o uso do paradigma `load`/`sync` com suplementos. Além disso, os suplementos oferecem APIs para eventos e um conjunto mais amplo de funcionalidades fora Excel, conhecidas como APIs comuns.

### <a name="events"></a>Eventos

Office scripts não dão suporte a eventos no nível da pasta de [trabalho](/office/dev/add-ins/excel/excel-add-ins-events). Os scripts são disparados por usuários selecionando o **botão** Executar para um script ou por meio Power Automate. Cada script executa o código em um único `main` método e, em seguida, termina.

### <a name="common-apis"></a>APIs comuns

Office scripts não podem usar [APIs comuns](/javascript/api/office). Se você precisar de autenticação, janelas de diálogo ou outros recursos compatíveis apenas com APIs comuns, provavelmente precisará criar um suplemento do Office em vez de um script Office.

## <a name="see-also"></a>Confira também

- [Office scripts no Excel](../overview/excel.md)
- [Diferenças entre Office scripts e macros VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
