---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: O comportamento e as diferenças de API entre Office scripts e Office de complementos.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7b199e8f3acdbe753fcaa2d1f4b6b5f11998b52b
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59328097"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre os scripts do Office e os suplementos do Office

Entenda as diferenças entre Office scripts e Office de Office para saber quando usar cada um deles. Office Os scripts foram projetados para serem feitos rapidamente por qualquer pessoa que procura melhorar seu fluxo de trabalho. Office Os complementos se integram à interface do usuário Office para uma experiência mais interativa por meio de botões de faixa de opções e painéis de tarefas. Office Os complementos também podem expandir funções de Excel por meio do fornecimento de funções personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções Office extensibilidade. Os Office Scripts e os Office Web Add-ins estão focados na Web e na colaboração, mas os scripts do Office atendem aos usuários finais (enquanto os Office Web São destinados a desenvolvedores profissionais).":::

Office Os scripts são executados para conclusão com uma pressionamento de botão manual ou como uma etapa em [Power Automate](https://flow.microsoft.com/), enquanto os Office de complementos continuam sendo executados dependendo de como eles são configurados. Por exemplo, você pode configurar um Office para continuar a ser executado mesmo quando o painel de tarefas estiver fechado. Isso significa que Office os complementos mantêm o estado durante uma sessão, enquanto Office scripts não mantêm um estado interno entre as executações. Se a solução que você está criando exigir um estado mantido, você deve visitar a documentação de Office de Office de [complementos](/office/dev/add-ins) para saber mais sobre os Office Desem.

O restante deste artigo descreve as principais diferenças entre os Office e Office Scripts.

## <a name="platform-support"></a>Suporte à plataforma

Office Os complementos são entre plataformas. Eles trabalham em Windows desktop, Mac, iOS e plataformas Web e fornecem a mesma experiência em cada uma delas. Qualquer exceção a isso é notada na documentação da API individual.

Office Atualmente, os scripts só têm suporte para Excel na Web. Toda a gravação, edição e execução é feita na plataforma Web.

## <a name="apis"></a>APIs

Embora as APIs Office JavaScript para Office e as APIs Office scripts compartilhem algumas funcionalidades, elas são plataformas diferentes. As OFFICE scripts são um subconjunto otimizado e síncrono do modelo de API JavaScript Excel JavaScript. A principal diferença é o uso do `load` / `sync` paradigma com os complementos. Além disso, os complementos oferecem APIs para eventos e um conjunto mais amplo de funcionalidades fora da Excel, conhecidas como APIs Comuns.

### <a name="events"></a>Events

Office Os scripts não suportam eventos de nível de [trabalho.](/office/dev/add-ins/excel/excel-add-ins-events) Os scripts são disparados por usuários selecionando o botão **Executar** para um script ou por meio Power Automate. Cada script executa o código em um único `main` método e termina.

### <a name="common-apis"></a>Common APIs

Office Os scripts não podem usar [APIs comuns.](/javascript/api/office) Se você precisar de autenticação, janelas de diálogo ou outros recursos com suporte apenas para APIs comuns, provavelmente precisará criar um Office Add-in em vez de um script Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre Office scripts e macros do VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
