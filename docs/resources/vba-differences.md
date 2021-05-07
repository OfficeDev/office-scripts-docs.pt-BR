---
title: Diferenças entre Office scripts e macros do VBA
description: O comportamento e as diferenças de API entre Office Scripts e Excel VBA.
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: ca571e2adad81a87b99696a652a3c49209b870ab
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232841"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferenças entre Office scripts e macros do VBA

Office Scripts e macros VBA têm muito em comum. Ambos permitem que os usuários automatizem soluções por meio de um gravador de ações fácil de usar e permitem edições dessas gravações. Ambas as estruturas foram projetadas para capacitar pessoas que podem não se considerar programadores para criar pequenos programas em Excel.
A diferença fundamental é que as macros do VBA são desenvolvidas para soluções de área de trabalho e Office scripts são projetados com suporte entre plataformas e segurança como os princípios orientadores. Atualmente, Office scripts são suportados apenas em Excel na Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para soluções Office extensibilidade diferentes. Tanto Office scripts quanto macros VBA foram projetados para ajudar os usuários finais a criar soluções, mas Office scripts são criados para a Web e colaboração (enquanto o VBA é para a área de trabalho)":::

Este artigo descreve as principais diferenças entre as macros do VBA (bem como o VBA em geral) e Office Scripts. Como Office scripts estão disponíveis apenas para Excel, esse é o único host que está sendo discutido aqui.

## <a name="platform-and-ecosystem"></a>Plataforma e ecossistema

O VBA foi projetado para a área de trabalho e Office scripts foram projetados para a Web. O VBA pode interagir com a área de trabalho de um usuário para se conectar com tecnologias semelhantes, como COM e OLE. No entanto, o VBA não tem uma maneira conveniente de chamar a Internet.

Office Scripts usam um tempo de execução universal para JavaScript. Isso oferece comportamento e acessibilidade consistentes, independentemente do computador que está sendo usado para executar o script. Eles também podem fazer chamadas para outros serviços Web.

## <a name="security"></a>Segurança

As macros VBA têm a mesma autorização de segurança que Excel. Isso dá a eles acesso total à sua área de trabalho. Office Os scripts só têm acesso à workbook, não ao computador que hospeda a workbook. Além disso, nenhum token de autenticação JavaScript pode ser compartilhado com scripts. Isso significa que o script não tem os tokens do usuário interno nem há recursos de API para entrar em um serviço externo, portanto, eles não podem usar tokens existentes para fazer chamadas externas em nome do usuário.

Os administradores têm três opções para macros VBA: permitir todas as macros no locatário, não permitir macros no locatário ou permitir somente macros com certificados assinados. Essa falta de granularidade dificulta a isolação de um único ator ruim. Atualmente, Office scripts estão ativas ou desligadas para um locatário. No entanto, estamos trabalhando para dar aos administradores mais controle sobre scripts individuais e criadores de scripts.

## <a name="coverage"></a>Cobertura

Atualmente, o VBA oferece uma cobertura mais completa de recursos Excel, especialmente aqueles disponíveis no cliente da área de trabalho. Office Os scripts abrangem quase todos os cenários para Excel na Web. Além disso, à medida que novos recursos estream na Web, os scripts Office os darão suporte para as APIs Action Recorder e JavaScript.

Office Scripts não suportam eventos Excel [nível](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects). Os scripts só são executados quando um usuário os inicia manualmente ou quando um fluxo Power Automate chama o script.

## <a name="power-automate"></a>Power Automate

Office Os scripts podem ser executados Power Automate. Sua workbook pode ser atualizada por meio de fluxos agendados ou orientados por eventos, o que permite automatizar fluxos de trabalho sem nem mesmo abrir Excel. Isso significa que, desde que sua OneDrive seja armazenada no OneDrive (e acessível ao Power Automate), um fluxo pode executar seus scripts independentemente de você e sua organização usarem a área de trabalho, Mac ou cliente Web do Excel.

O VBA não tem um Power Automate conector. Todos os cenários do VBA com suporte envolviam um usuário que participava da execução da macro.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre os scripts do Office e os suplementos do Office](add-ins-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Referência do VBA do Excel](/office/vba/api/overview/excel)
