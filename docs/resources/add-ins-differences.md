---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: O comportamento e as diferenças de API entre Office scripts e Office de complementos.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 018d210208bc78da894678d21e368864522cb83e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585602"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre os scripts do Office e os suplementos do Office

Entenda as diferenças entre Office scripts e Office de Office para saber quando usar cada um deles. Office scripts são projetados para serem feitos rapidamente por qualquer pessoa que procura melhorar seu fluxo de trabalho. Office os Complementos se integram à interface do usuário Office para uma experiência mais interativa por meio de botões de faixa de opções e painéis de tarefas. Office os complementos também podem expandir funções de Excel integrados fornecendo funções personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para soluções Office extensibilidade diferentes. Os Office scripts e os Office web add-ins estão focados na Web e na colaboração, mas os scripts do Office atendem aos usuários finais (enquanto os Office Web Add-ins são destinados a desenvolvedores profissionais).":::

Office Scripts são executados para conclusão com uma pressionamento de botão manual ou como uma etapa no [Power Automate, enquanto](https://flow.microsoft.com/) os Office de complementos continuam sendo executados dependendo de como eles são configurados. Por exemplo, você pode configurar um Office para continuar a ser executado mesmo quando o painel de tarefas estiver fechado. Isso significa que Office os complementos mantêm o estado durante uma sessão, enquanto Office scripts não mantêm um estado interno entre as executações. Se a solução que você está criando exigir um estado mantido, você deverá visitar a documentação de Office de Office de [Complementos](/office/dev/add-ins) para saber mais sobre os Office Desem.

O restante deste artigo descreve as principais diferenças entre os Office e Office Scripts.

## <a name="platform-support"></a>Suporte à plataforma

Office Os complementos são entre plataformas. Eles trabalham em Windows desktop, Mac, iOS e plataformas Web e fornecem a mesma experiência em cada uma delas. Qualquer exceção a isso é notada na documentação da API individual.

Office scripts atualmente são suportados apenas por Excel na Web. Todo o gerenciamento de gravação, edição e script é feito na plataforma Web.

### <a name="script-support-for-excel-on-windows"></a>Suporte a scripts para Excel no Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>APIs

Embora as APIs Office JavaScript para Office e as APIs Office Scripts do Office compartilhem algumas funcionalidades, elas são plataformas diferentes. As OFFICE scripts são um subconjunto otimizado e síncrono do modelo de API JavaScript Excel JavaScript. A principal diferença é o uso do paradigma `load`/`sync` com os complementos. Além disso, os complementos oferecem APIs para eventos e um conjunto mais amplo de funcionalidades fora da Excel, conhecidas como APIs Comuns.

### <a name="events"></a>Eventos

Office scripts não suportam eventos de nível de [trabalho.](/office/dev/add-ins/excel/excel-add-ins-events) Os scripts são disparados por usuários selecionando o botão **Executar** para um script ou por meio Power Automate. Cada script executa o código em um único `main` método e termina.

### <a name="common-apis"></a>APIs comuns

Office scripts não podem usar [APIs comuns](/javascript/api/office). Se você precisar de autenticação, janelas de diálogo ou outros recursos que só têm suporte para APIs comuns, provavelmente precisará criar um Office Add-in em vez de um script Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre Office scripts e macros do VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
