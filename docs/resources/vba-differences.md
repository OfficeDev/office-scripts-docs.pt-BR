---
title: Diferenças entre scripts Office e macros VBA
description: As diferenças de comportamento e API entre Office Scripts e Excel macros VBA.
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 612a5f21d935fd262a6e9fd12a3431956105636a
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545585"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferenças entre scripts Office e macros VBA

Office Scripts e macros VBA têm muito em comum. Ambos permitem que os usuários automatizem soluções através de um gravador de ação fácil de usar e permitam edições dessas gravações. Ambas as estruturas são projetadas para capacitar pessoas que podem não se considerar programadoras para criar pequenos programas em Excel.
A diferença fundamental é que as macros VBA são desenvolvidas para soluções de desktop e Office Scripts são projetados para soluções seguras baseadas em nuvem. Atualmente, Office Scripts só são suportados em Excel na Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade Office. Tanto Office scripts quanto macros VBA foram projetadas para ajudar os usuários finais a criar soluções, mas Office Scripts são construídos para a web e colaboração (enquanto o VBA é para a área de trabalho)":::

Este artigo descreve as principais diferenças entre as macros VBA (bem como vba em geral) e Office Scripts. Como Office Scripts só estão disponíveis para Excel, esse é o único anfitrião que está sendo discutido aqui.

## <a name="platform-and-ecosystem"></a>Plataforma e ecossistema

O VBA foi projetado para a área de trabalho e Office Scripts são projetados para a web. O VBA pode interagir com a área de trabalho do usuário para se conectar com tecnologias semelhantes, como COM e OLE. No entanto, a VBA não tem uma maneira conveniente de chamar a internet.

Office Os scripts usam um tempo de execução universal para JavaScript. Isso dá comportamento consistente e acessibilidade, independentemente da máquina ser usada para executar o script. Eles também podem fazer chamadas para outros serviços web.

## <a name="security"></a>Segurança

As macros VBA têm a mesma autorização de segurança que Excel. Isso lhes dá acesso total à sua área de trabalho. Office Os scripts só têm acesso à pasta de trabalho, não à máquina que hospeda a pasta de trabalho. Além disso, nenhum token de autenticação JavaScript pode ser compartilhado com scripts. Isso significa que o script não tem nem os tokens do usuário de login, nem existem recursos de API para fazer login em um serviço externo, então eles são incapazes de usar tokens existentes para fazer chamadas externas em nome do usuário.

Os administradores têm três opções para macros VBA: permitir todas as macros do inquilino, não permitir macros no inquilino ou permitir apenas macros com certificados assinados. Essa falta de granularidade torna difícil isolar um único ator ruim. Atualmente, Office scripts estão ligados ou desligados para um inquilino. No entanto, estamos trabalhando para dar aos administradores mais controle sobre scripts individuais e criadores de scripts.

## <a name="coverage"></a>cobertura

Atualmente, o VBA oferece uma cobertura mais completa dos recursos Excel, especialmente aqueles disponíveis no cliente desktop. Office Os roteiros cobrem quase todos os cenários para Excel na Web. Além disso, à medida que os novos recursos estreiam na web, Office Scripts os suportarão tanto para as APIs do Action Recorder quanto do JavaScript.

Office Scripts não suportam [eventos](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)de nível Excel . Os scripts só são executados quando um usuário os inicia manualmente ou quando um fluxo de Power Automate chama o script.

## <a name="power-automate"></a>Power Automate

Office Os scripts podem ser executados através de Power Automate. Sua pasta de trabalho pode ser atualizada através de fluxos programados ou orientados a eventos, permitindo que você automatize fluxos de trabalho sem sequer abrir Excel. Isso significa que, desde que sua pasta de trabalho seja armazenada em OneDrive (e acessível a Power Automate), um fluxo pode executar seus scripts independentemente de você e sua organização usarem a área de trabalho, Mac ou cliente da Web da Excel.

A VBA não tem um conector Power Automate. Todos os cenários VBA suportados envolvem um usuário que atende à execução da macro.

Experimente os [scripts de chamada a partir de um](../tutorials/excel-power-automate-manual.md) tutorial de fluxo de Power Automate manual para começar a aprender sobre Power Automate. Você também pode verificar a amostra [de lembretes de tarefas automatizadas](scenarios/task-reminders.md) para ver Office Scripts conectados a Teams através de Power Automate em um cenário do mundo real.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Execute Office scripts com Power Automate](../develop/power-automate-integration.md)
- [Diferenças entre os scripts do Office e os suplementos do Office](add-ins-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Referência do VBA do Excel](/office/vba/api/overview/excel)
