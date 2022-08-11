---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: O comportamento e as diferenças de API entre scripts do Office e suplementos do Office.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3df4daf04f963598d2cb31f82dd2c1c9923fdc8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281907"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre os scripts do Office e os suplementos do Office

Entenda as diferenças entre scripts do Office e suplementos do Office para saber quando usar cada um deles. Os Scripts do Office foram projetados para serem feitos rapidamente por qualquer pessoa que deseja melhorar seu fluxo de trabalho. Os Suplementos do Office integram-se à interface do usuário do Office para obter uma experiência mais interativa por meio de botões de faixa de opções e painéis de tarefas. Os Suplementos do Office também podem expandir funções internas do Excel fornecendo funções personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade do Office. Os Scripts do Office e os Suplementos da Web do Office se concentram na Web e na colaboração, mas os Scripts do Office atendem aos usuários finais (enquanto os Suplementos da Web do Office são destinados a desenvolvedores profissionais).":::

Os Scripts do Office são executados até a conclusão com um pressionamento de botão manual ou como uma etapa no [Power Automate](https://flow.microsoft.com/), enquanto os Suplementos do Office continuam em execução dependendo de como eles estão configurados. Por exemplo, você pode configurar um Suplemento do Office para continuar em execução mesmo quando o painel de tarefas estiver fechado. Isso significa que os Suplementos do Office mantêm o estado durante uma sessão, enquanto os Scripts do Office não mantêm um estado interno entre execuções. Se a solução que você está criando exigir um estado mantido, visite a documentação de [Suplementos do Office](/office/dev/add-ins) para saber mais sobre os Suplementos do Office.

O restante deste artigo descreve as principais diferenças entre suplementos do Office e scripts do Office.

## <a name="platform-support"></a>Suporte à plataforma

Os Suplementos do Office são multiplataforma. Eles funcionam em plataformas da Área de Trabalho do Windows, Mac, iOS e Web e fornecem a mesma experiência em cada uma delas. Qualquer exceção a isso é notada na documentação da API individual.

Atualmente, os Scripts do Office só têm suporte para Excel na Web. Todo o gerenciamento de gravação, edição e script é feito na plataforma Web.

### <a name="script-support-for-excel-on-windows"></a>Suporte a scripts para Excel no Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>APIs

Embora as APIs JavaScript do Office para Suplementos do Office e as APIs de Scripts do Office compartilhem algumas funcionalidades, elas são plataformas diferentes. As APIs de Scripts do Office são um subconjunto otimizado e síncrono do modelo de API JavaScript do Excel. A principal diferença é o uso do paradigma `load`/`sync` com suplementos. Além disso, os suplementos oferecem APIs para eventos e um conjunto mais amplo de funcionalidades fora do Excel, conhecidas como APIs comuns.

### <a name="events"></a>Events

Os Scripts do Office não dão suporte a eventos no nível da pasta de [trabalho](/office/dev/add-ins/excel/excel-add-ins-events). Os scripts são disparados por usuários selecionando o **botão** Executar para um script ou por meio do Power Automate. Cada script executa o código em uma única `main` função e, em seguida, termina.

### <a name="common-apis"></a>APIs comuns

Os Scripts do Office não podem [usar APIs comuns](/javascript/api/office). Se você precisar de autenticação, janelas de diálogo ou outros recursos compatíveis apenas com APIs comuns, provavelmente precisará criar um Suplemento do Office em vez de um Script do Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel](../overview/excel.md)
- [Diferenças entre scripts do Office e macros do VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
