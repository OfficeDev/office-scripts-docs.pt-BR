---
title: Diferenças entre scripts do Office e macros do VBA
description: As diferenças de comportamento e API entre scripts do Office e macros do Excel VBA.
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: a56409a5de3eb07876faa88bfbfe78eeca59f70f
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755018"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferenças entre scripts do Office e macros do VBA

Scripts do Office e macros VBA têm muito em comum. Ambos permitem que os usuários automatizem soluções por meio de um gravador de ações fácil de usar e permitem edições dessas gravações. Ambas as estruturas foram projetadas para capacitar pessoas que podem não se considerar programadores para criar pequenos programas no Excel.
A diferença fundamental é que as macros do VBA são desenvolvidas para soluções de área de trabalho e os Scripts do Office são projetados com suporte entre plataformas e segurança como princípios orientadores. Atualmente, os Scripts do Office só têm suporte no Excel na Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade do Office. Tanto os Scripts do Office quanto as macros do VBA foram projetados para ajudar os usuários finais a criar soluções, mas os Scripts do Office são criados para a Web e a colaboração (enquanto o VBA é para a área de trabalho).":::

Este artigo descreve as principais diferenças entre as macros do VBA (bem como o VBA em geral) e scripts do Office. Como os Scripts do Office estão disponíveis apenas para o Excel, esse é o único host que está sendo discutido aqui.

## <a name="platform-and-ecosystem"></a>Plataforma e ecossistema

O VBA foi projetado para a área de trabalho e os Scripts do Office foram projetados para a Web. O VBA pode interagir com a área de trabalho de um usuário para se conectar com tecnologias semelhantes, como COM e OLE. No entanto, o VBA não tem uma maneira conveniente de chamar a Internet.

Os Scripts do Office usam um tempo de execução universal para JavaScript. Isso oferece comportamento e acessibilidade consistentes, independentemente do computador que está sendo usado para executar o script. Eles também podem fazer chamadas para outros serviços Web.

## <a name="security"></a>Segurança

As macros do VBA têm a mesma autorização de segurança do Excel. Isso dá a eles acesso total à sua área de trabalho. Os Scripts do Office só têm acesso à caixa de trabalho, não ao computador que hospeda a workbook. Além disso, nenhum token de autenticação JavaScript pode ser compartilhado com scripts. Isso significa que o script não tem os tokens do usuário interno nem há recursos de API para entrar em um serviço externo, portanto, eles não podem usar tokens existentes para fazer chamadas externas em nome do usuário.

Os administradores têm três opções para macros VBA: permitir todas as macros no locatário, não permitir macros no locatário ou permitir somente macros com certificados assinados. Essa falta de granularidade dificulta a isolação de um único ator ruim. Atualmente, os Scripts do Office estão ativas ou desligadas para um locatário. No entanto, estamos trabalhando para dar aos administradores mais controle sobre scripts individuais e criadores de scripts.

## <a name="coverage"></a>Cobertura

Atualmente, o VBA oferece uma cobertura mais completa dos recursos do Excel, especialmente aqueles disponíveis no cliente da área de trabalho. Os Scripts do Office abrangem quase todos os cenários do Excel na Web. Além disso, à medida que novos recursos estream na Web, os Scripts do Office darão suporte a eles para as APIs Action Recorder e JavaScript.

Os Scripts do Office não suportam eventos no nível [do](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)Excel. Os scripts só são executados quando um usuário os inicia manualmente ou quando um fluxo do Power Automate chama o script.

## <a name="power-automate"></a>Power Automate

Os Scripts do Office podem ser executados por meio do Power Automate. Sua planilha pode ser atualizada por meio de fluxos agendados ou orientados por eventos, o que permite automatizar fluxos de trabalho sem sequer abrir o Excel. Isso significa que, desde que sua planilha seja armazenada no OneDrive (e acessível ao Power Automate), um fluxo pode executar seus scripts independentemente de você e sua organização usarem a área de trabalho, Mac ou cliente Web do Excel.

O VBA não tem um conector do Power Automate. Todos os cenários do VBA com suporte envolviam um usuário que participava da execução da macro.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Diferenças entre os scripts do Office e os suplementos do Office](add-ins-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Referência do VBA do Excel](/office/vba/api/overview/excel)
