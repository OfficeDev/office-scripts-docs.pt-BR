---
title: Limites e requisitos da plataforma com scripts Office
description: Limites de recursos e suporte ao navegador para scripts Office quando usados com Excel na Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545578"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites e requisitos da plataforma com scripts Office

Existem algumas limitações de plataforma das quais você deve estar ciente ao desenvolver Office Scripts. Este artigo detalha o suporte ao navegador e os limites de dados para scripts Office para Excel na Web.

## <a name="browser-support"></a>Suporte do navegador

Office Os scripts funcionam em qualquer navegador que [suporte Office para a web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). No entanto, alguns recursos JavaScript não são suportados no Internet Explorer 11 (IE 11). Quaisquer recursos introduzidos no [ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11. Se as pessoas na sua organização ainda usarem esse navegador, certifique-se de testar seus scripts nesse ambiente ao compartilhá-los.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceiros

Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **Automate** em Excel na Web. Verifique as configurações do seu navegador se a guia não está sendo exibida. Se você estiver usando uma sessão privada do navegador, talvez seja necessário ree habilitar essa configuração todas as vezes.

> [!NOTE]
> Alguns navegadores se referem a esta configuração como "todos os cookies", em vez de "cookies de terceiros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instruções para ajustar as configurações de cookies em navegadores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Borda](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites de dados

Há limites sobre quanto Excel dados podem ser transferidos de uma só vez e quantas transações individuais Power Automate podem ser realizadas.

### <a name="excel"></a>Excel

Excel para a web tem as seguintes limitações ao fazer chamadas para a pasta de trabalho através de um script:

- As solicitações e respostas são limitadas a **5MB**.
- Um alcance é limitado a **cinco milhões de células.**

Se você estiver encontrando erros ao lidar com grandes conjuntos de dados, tente usar várias faixas menores em vez de faixas maiores. Por exemplo, consulte a amostra [de conjunto de dados Write.](../resources/samples/write-large-dataset.md) Você também pode usar APIs como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para atingir células específicas em vez de grandes intervalos.

### <a name="power-automate"></a>Power Automate

Ao usar Office Scripts com Power Automate, cada usuário é limitado a **400 chamadas para a ação Executar script por dia**. Este limite é zerado às 12:00 AM UTC.

A plataforma Power Automate também possui limitações de uso, que podem ser encontradas nos seguintes artigos:

- [Limites e configuração em Power Automate](/power-automate/limits-and-config)
- [Problemas e limitações conhecidos para o conector online Excel (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Confira também

- [Solução de problemas Office Scripts](troubleshooting.md)
- [Desfazer os efeitos do Scripts do Office](undo.md)
- [Melhore o desempenho de seus scripts de Office](../develop/web-client-performance.md)
- [Roteirizando fundamentos para roteiros Office em Excel na Web](../develop/scripting-fundamentals.md)
