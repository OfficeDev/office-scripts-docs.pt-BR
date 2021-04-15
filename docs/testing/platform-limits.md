---
title: Limites e requisitos de plataforma com scripts do Office
description: Limites de recursos e suporte ao navegador para Scripts do Office quando usados com o Excel na Web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: ef733562fb3caa8261fbbd8382923927a46cb7d4
ms.sourcegitcommit: 5ca286615a11d282e3f80023d22d36a039800eed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2021
ms.locfileid: "51689763"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites e requisitos de plataforma com scripts do Office

Há algumas limitações de plataforma que você deve estar ciente ao desenvolver scripts do Office. Este artigo detalha o suporte ao navegador e os limites de dados para Scripts do Office para Excel na Web.

## <a name="browser-support"></a>Suporte do navegador

Os Scripts do Office funcionam em qualquer navegador [compatível com o Office para a Web.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) No entanto, alguns recursos JavaScript não são suportados no Internet Explorer 11 (IE 11). Quaisquer recursos introduzidos [no ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11. Se as pessoas em sua organização ainda usarem esse navegador, teste seus scripts nesse ambiente ao compartilhar.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceiros

Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **Automatizar** no Excel na Web. Verifique as configurações do navegador se a guia não está sendo exibida. Se você estiver usando uma sessão privada do navegador, talvez seja necessário reabilitar essa configuração sempre.

> [!NOTE]
> Alguns navegadores referem-se a essa configuração como "todos os cookies", em vez de "cookies de terceiros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instruções para ajustar configurações de cookie em navegadores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Borda](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites de dados

Há limites sobre a quantidade de dados do Excel que podem ser transferidos de uma só vez e quantas transações individuais do Power Automate podem ser conduzidas.

### <a name="excel"></a>Excel

O Excel para a Web tem as seguintes limitações ao fazer chamadas para a planilha por meio de um script:

- Solicitações e respostas são limitadas a **5 MB**.
- Um intervalo é limitado a **cinco milhões de células**.

Se você estiver encontrando erros ao lidar com conjuntos de dados grandes, tente usar vários intervalos menores em vez de intervalos maiores. Você também pode APIs como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para direcionar células específicas em vez de intervalos grandes.

### <a name="power-automate"></a>Power Automate

Ao usar scripts do Office com o Power Automate, cada usuário é limitado a **400** chamadas para a ação Executar Script por dia . Esse limite é redefinido às 00:00 UTC.

A plataforma Power Automate também tem limitações de uso, que podem ser encontradas nos seguintes artigos:

- [Limites e configuração no Power Automate](/power-automate/limits-and-config)
- [Problemas conhecidos e limitações para o conector do Excel Online (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Confira também

- [Solução de problemas dos scripts do Office](troubleshooting.md)
- [Desfazer os efeitos de um script do Office](undo.md)
- [Melhorar o desempenho dos scripts do Office](../develop/web-client-performance.md)
- [Scripts básicos para scripts do Office no Excel na Web](../develop/scripting-fundamentals.md)
