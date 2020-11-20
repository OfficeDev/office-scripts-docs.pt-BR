---
title: Diferenças entre scripts do Office e macros VBA
description: As diferenças de comportamento e API entre scripts do Office e macros VBA do Excel.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 7b9186d03489a43836c6e9da7bd28e0abc135f63
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342882"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferenças entre scripts do Office e macros VBA

Scripts do Office e macros do VBA têm muito em comum. Ambos permitem que os usuários automatizem soluções por meio de um gravador de ação fácil de usar e permitir edições dessas gravações. Ambas as estruturas foram projetadas para capacitar as pessoas que podem não considerar os programadores para criar pequenos programas no Excel.
A diferença fundamental é que as macros do VBA são desenvolvidas para soluções de área de trabalho e scripts do Office são projetadas com suporte e segurança entre plataformas como os princípios de orientação. Atualmente, os scripts do Office só têm suporte no Excel na Web.

![Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade do Office. Tanto os scripts do Office quanto as macros do VBA foram projetados para ajudar os usuários finais a criar soluções, mas os scripts do Office são criados para a Web e colaboração (enquanto o VBA é para a área de trabalho).)](../images/office-programmability-diagram.png)

Este artigo descreve as principais diferenças entre as macros VBA (bem como VBA em geral) e scripts do Office. Como os scripts do Office estão disponíveis apenas para o Excel, esse é o único host abordado aqui.

## <a name="platform-and-ecosystem"></a>Plataforma e ecossistema

O VBA foi projetado para a área de trabalho e os scripts do Office foram projetados para a Web. O VBA pode interagir com a área de trabalho de um usuário para se conectar com tecnologias semelhantes, como COM e OLE. No entanto, o VBA não tem uma maneira conveniente de fazer chamadas para a Internet.

Os scripts do Office usam um tempo de execução universal para JavaScript. Isso fornece comportamento e acessibilidade consistentes, independentemente da máquina que está sendo usada para executar o script. Eles também podem fazer chamadas para outros serviços Web.

## <a name="security"></a>Segurança

As macros do VBA têm o mesmo espaço livre de segurança que o Excel. Isso lhes dá acesso total à sua área de trabalho. Os scripts do Office só têm acesso à pasta de trabalho, não à máquina que hospeda a pasta de trabalho. Além disso, nenhum token de autenticação JavaScript pode ser compartilhado com scripts. Isso significa que o script não tem os tokens do usuário conectado ou não há recursos de API para entrar em um serviço externo, portanto, não é possível usar tokens existentes para fazer chamadas externas em nome do usuário.

Os administradores têm três opções para macros VBA: permitir todas as macros no locatário, não permitir macros no locatário ou permitir somente macros com certificados assinados. Essa falta de granularidade dificulta a isolação de um único ator ruim. Atualmente, os scripts do Office estão ativados ou desativados para um locatário. No entanto, estamos trabalhando para dar aos administradores mais controle sobre scripts individuais e criadores de scripts.

## <a name="coverage"></a>Funcionamento

Atualmente, o VBA oferece uma cobertura mais completa dos recursos do Excel, particularmente aqueles disponíveis no cliente de desktop. Os scripts do Office cobrem quase todos os cenários do Excel na Web. Além disso, como novos recursos na Web, os scripts do Office oferecerão suporte para o gravador de ação e as APIs JavaScript.

## <a name="power-automate"></a>Power Automate

Os scripts do Office podem ser executados através da automatização de energia. Sua pasta de trabalho pode ser atualizada por meio de fluxos agendados ou orientados por eventos, permitindo que você automatize fluxos de trabalho sem precisar abrir o Excel. Isso significa que, desde que a pasta de trabalho seja armazenada no OneDrive (e acessível para automatização de energia), um fluxo pode executar seus scripts independentemente de você e sua organização usar a área de trabalho, Mac ou cliente Web do Excel.

O VBA não tem um conector automatizado de energia. Todos os cenários de VBA suportados envolveram um usuário que participa da execução da macro.

## <a name="see-also"></a>Também consulte

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre os scripts do Office e os suplementos do Office](add-ins-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Referência do VBA do Excel](/office/vba/api/overview/excel)
