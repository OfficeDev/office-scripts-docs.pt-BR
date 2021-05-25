---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: O comportamento e as diferenças de API entre Office scripts e Office de complementos.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 5c30406867da05952dedda684f765df5e7a7e53f
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631675"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre os scripts do Office e os suplementos do Office

Office Os complementos e Office scripts têm muito em comum. Ambos oferecem controle automatizado de uma Excel uma API JavaScript. No entanto, as APIs Office scripts são uma versão especializada e síncrona da API Office JavaScript.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções Office extensibilidade. Os Office scripts e os Office web add-ins estão focados na Web e na colaboração, mas os scripts do Office atendem aos usuários finais (enquanto os Office Web Add-ins são destinados a desenvolvedores profissionais)":::

Office Os scripts são executados para conclusão com uma pressão de botão manual ou como uma etapa em [Power Automate](https://flow.microsoft.com/), enquanto os Office de complementos persistem enquanto seus painéis de tarefas estão abertos. Isso significa que os complementos podem manter o estado durante uma sessão, enquanto Office scripts não mantêm um estado interno entre as executações. Se você descobrir que Excel extensão do Excel precisa exceder os recursos da plataforma de scripts, visite Office documentação de Office de [Complementos](/office/dev/add-ins) para saber mais sobre Office Desempla.

O restante deste artigo descreve as principais diferenças entre os Office e Office Scripts.

## <a name="platform-support"></a>Suporte à plataforma

Office Os complementos são entre plataformas. Eles trabalham em Windows desktop, Mac, iOS e plataformas Web e fornecem a mesma experiência em cada uma delas. Qualquer exceção a isso é notada na documentação da API individual.

Office Atualmente, os scripts só têm suporte para Excel na Web. Toda a gravação, edição e execução é feita na plataforma Web.

## <a name="apis"></a>APIs

Embora as APIs Office JavaScript para Office e as APIs Office scripts compartilhem algumas funcionalidades, elas são plataformas diferentes. As OFFICE scripts são uma versão otimizada e síncrona do modelo Excel API JavaScript. A principal diferença é o uso do `load` / `sync` paradigma com os complementos. Além disso, os complementos oferecem APIs para eventos e um conjunto mais amplo de funcionalidades fora da Excel, conhecidas como APIs Comuns.

### <a name="events"></a>Eventos

Office Scripts não suportam [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada script executa o código em um único `main` método e termina. Ele não é reativado quando os eventos são disparados e, portanto, não pode registrar eventos.

### <a name="common-apis"></a>Common APIs

Office Os scripts não podem usar [APIs comuns.](/javascript/api/office) Se você precisar de autenticação, janelas de diálogo ou outros recursos com suporte apenas para APIs comuns, provavelmente precisará criar um Office Add-in em vez de um script Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre Office scripts e macros do VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
