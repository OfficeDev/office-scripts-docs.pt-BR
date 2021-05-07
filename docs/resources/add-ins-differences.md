---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: O comportamento e as diferenças de API entre Office scripts e Office de complementos.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232631"
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

Não há nenhuma versão síncrona das APIs Office JavaScript para Office de usuário. As APIs Office scripts padrão são exclusivas da plataforma e têm várias otimizações e alterações para evitar o uso do `load` / `sync` paradigma.

Algumas das [APIs Excel JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true) são compatíveis com as [APIs Office Scripts Async.](../develop/excel-async-model.md) Alguns exemplos e blocos de código de complemento podem ser portados para blocos `Excel.run` com conversão mínima. Embora as duas plataformas compartilhem a funcionalidade, há lacunas. Os dois principais conjuntos de API que Office os Complementos têm, mas Office scripts não são eventos e apIs comuns.

### <a name="events"></a>Eventos

Office Scripts não suportam [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada script executa o código em um único `main` método e termina. Ele não é reativado quando os eventos são disparados e, portanto, não pode registrar eventos.

### <a name="common-apis"></a>APIs comuns

Office Os scripts não podem usar [APIs comuns.](/javascript/api/office) Se você precisar de autenticação, janelas de diálogo ou outros recursos com suporte apenas para APIs comuns, provavelmente precisará criar um Office Add-in em vez de um script Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre Office scripts e macros do VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
