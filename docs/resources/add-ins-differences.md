---
title: Diferenças entre os scripts do Office e os suplementos do Office
description: As diferenças de comportamento e API entre scripts do Office e os complementos do Office.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755095"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre os scripts do Office e os suplementos do Office

Os Complementos do Office e scripts do Office têm muito em comum. Ambos oferecem controle automatizado de uma planilha do Excel uma API JavaScript. No entanto, as APIs de Scripts do Office são uma versão especializada e síncrona da API JavaScript do Office.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade do Office. Tanto os Scripts do Office quanto os Complementos da Web do Office estão focados na Web e na colaboração, mas os Scripts do Office atendem aos usuários finais (enquanto os Complementos da Web do Office são destinados a desenvolvedores profissionais).":::

Os Scripts do Office são executados para conclusão com uma pressão de botão manual ou como uma etapa no [Power Automate](https://flow.microsoft.com/), enquanto os Complementos do Office persistem enquanto os painéis de tarefas estão abertos. Isso significa que os complementos podem manter o estado durante uma sessão, enquanto os Scripts do Office não mantêm um estado interno entre as executações. Se você descobrir que sua extensão do Excel precisa exceder os recursos da plataforma de scripts, visite a documentação de [Complementos](/office/dev/add-ins) do Office para saber mais sobre os Complementos do Office.

O restante deste artigo descreve as principais diferenças entre os Complementos do Office e scripts do Office.

## <a name="platform-support"></a>Suporte à plataforma

Os Complementos do Office são entre plataformas. Eles funcionam em plataformas da Web, mac, iOS e desktop do Windows e fornecem a mesma experiência em cada uma delas. Qualquer exceção a isso é notada na documentação da API individual.

Atualmente, os Scripts do Office só têm suporte para o Excel na Web. Toda a gravação, edição e execução é feita na plataforma Web.

## <a name="apis"></a>APIs

Não há nenhuma versão síncrona das APIs JavaScript do Office para Os Complementos do Office. As APIs padrão do Office Scripts são exclusivas da plataforma e têm várias otimizações e alterações para evitar o uso do `load` / `sync` paradigma.

Algumas das [APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true) são compatíveis com as [APIs async de Scripts do Office.](../develop/excel-async-model.md) Alguns exemplos e blocos de código de complemento podem ser portados para blocos `Excel.run` com conversão mínima. Embora as duas plataformas compartilhem a funcionalidade, há lacunas. Os dois principais conjuntos de API que os Complementos do Office têm, mas os Scripts do Office não são eventos e as APIs comuns.

### <a name="events"></a>Eventos

Scripts do Office não suportam [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada script executa o código em um único `main` método e termina. Ele não é reativado quando os eventos são disparados e, portanto, não pode registrar eventos.

### <a name="common-apis"></a>APIs comuns

Scripts do Office não podem usar [APIs comuns.](/javascript/api/office) Se você precisar de autenticação, janelas de diálogo ou outros recursos com suporte apenas para APIs comuns, provavelmente precisará criar um Add-in do Office em vez de um Script do Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre scripts do Office e macros do VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
