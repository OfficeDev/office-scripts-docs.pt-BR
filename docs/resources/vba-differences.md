---
title: Diferenças entre Office scripts e macros VBA
description: O comportamento e as diferenças de API entre Office Scripts e Excel macros VBA.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60e4fba6e63967302066f544b76fb20a8c8630a6
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393611"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferenças entre Office scripts e macros VBA

Office scripts e macros VBA têm muito em comum. Ambos permitem que os usuários automatizem soluções por meio de um gravador de ações fácil de usar e permitem edições dessas gravações. Ambas as estruturas foram projetadas para capacitar pessoas que podem não se considerar programadores para criar pequenos programas Excel.

A diferença fundamental é que as macros VBA são desenvolvidas para soluções de área de trabalho e Office scripts são projetados para soluções seguras baseadas em nuvem. Atualmente, Office scripts só têm suporte em Excel na Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções Office extensibilidade. Scripts Office e macros VBA foram projetados para ajudar os usuários finais a criar soluções, mas scripts Office são criados para a Web e colaboração (enquanto o VBA é para a área de trabalho).":::

Este artigo descreve as principais diferenças entre as macros do VBA (bem como o VBA em geral) e Office Scripts. Como Office scripts estão disponíveis apenas para Excel, esse é o único host que está sendo discutido aqui.

## <a name="platform-and-ecosystem"></a>Plataforma e ecossistema

O VBA é compatível com Excel no Windows Mac. Office scripts tem suporte Excel na Web.

As duas soluções foram projetadas para suas respectivas plataformas. O VBA pode interagir com a área de trabalho de um usuário para se conectar com tecnologias semelhantes, como COM e OLE. No entanto, o VBA não tem uma maneira conveniente de chamar a Internet. Office scripts usam um runtime universal para JavaScript. Isso fornece comportamento e acessibilidade consistentes, independentemente do computador que está sendo usado para executar o script. Eles também podem fazer chamadas para outros serviços Web.

### <a name="script-support-for-excel-on-windows"></a>Suporte a script para Excel no Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>Segurança

As macros VBA têm a mesma permissão de segurança que Excel. Isso dá a eles acesso completo à sua área de trabalho. Office scripts têm acesso apenas à pasta de trabalho, não ao computador que hospeda a pasta de trabalho. Além disso, nenhum token de autenticação JavaScript pode ser compartilhado com scripts. Isso significa que o script não tem os tokens do usuário conectado nem há recursos de API para entrar em um serviço externo, portanto, não é possível usar tokens existentes para fazer chamadas externas em nome do usuário.

Os administradores têm três opções para macros VBA: permitir todas as macros no locatário, não permitir macros no locatário ou permitir somente macros com certificados assinados. Essa falta de granularidade torna difícil isolar um único ator ruim. Atualmente, Office scripts podem estar desativados para um locatário inteiro, para um locatário inteiro ou para um grupo de usuários em um locatário. Os administradores também têm controle sobre quem pode compartilhar scripts com outras pessoas e quem pode usar scripts em Power Automate.

## <a name="coverage"></a>Cobertura

Atualmente, o VBA oferece uma cobertura mais completa de Excel recursos, especialmente aqueles disponíveis no cliente da área de trabalho. Office scripts abrangem quase todos os cenários para Excel na Web. Além disso, à medida que novos recursos estrearem na Web, Office scripts darão suporte a eles para o Gravador de Ações e APIs JavaScript.

Office scripts não dão suporte Excel eventos de nível [de servidor](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects). Os scripts só são executados quando um usuário os inicia manualmente ou quando um fluxo Power Automate chama o script.

## <a name="power-automate"></a>Power Automate

Office scripts podem ser executados por meio Power Automate. Sua pasta de trabalho pode ser atualizada por meio de fluxos agendados ou controlados por eventos, permitindo automatizar fluxos de trabalho sem sequer abrir Excel. Isso significa que, desde que sua pasta de trabalho esteja armazenada no OneDrive (e acessível ao Power Automate), um fluxo poderá executar seus scripts, independentemente de você e sua organização usarem a área de trabalho, o Mac ou o cliente Web do Excel.

O VBA não tem um Power Automate conector. Todos os cenários de VBA com suporte envolvem um usuário que está participando da execução da macro.

Experimente os [scripts de chamada em um tutorial de fluxo de Power Automate](../tutorials/excel-power-automate-manual.md) manual para começar a aprender sobre Power Automate. Você também pode conferir o exemplo de [lembretes de tarefas automatizadas](scenarios/task-reminders.md) para ver Office scripts conectados ao Teams por meio Power Automate em um cenário do mundo real.

## <a name="see-also"></a>Confira também

- [Office scripts no Excel](../overview/excel.md)
- [Executar Office scripts com Power Automate](../develop/power-automate-integration.md)
- [Diferenças entre os scripts do Office e os suplementos do Office](add-ins-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Referência do VBA do Excel](/office/vba/api/overview/excel)
