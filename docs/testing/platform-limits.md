---
title: Limites e requisitos da plataforma com scripts do Office
description: Limites de recursos e suporte ao navegador para Scripts do Office quando usados com Excel na Web.
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891242"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites e requisitos da plataforma com scripts do Office

Há algumas limitações de plataforma das quais você deve estar ciente ao desenvolver scripts do Office. Este artigo detalha o suporte do navegador e os limites de dados para scripts do Office para Excel na Web.

## <a name="browser-support"></a>Suporte do navegador

Os Scripts do Office funcionam em qualquer navegador compatível [com Office para a Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). No entanto, alguns recursos JavaScript não têm suporte no Internet Explorer 11 (IE 11). Todos os recursos introduzidos no [ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11. Se as pessoas em sua organização ainda usarem esse navegador, não deixe de testar seus scripts nesse ambiente ao compartilhá-los.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceiros

Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **Automatizar** no Excel na Web. Verifique as configurações do navegador se a guia não está sendo exibida. Se você estiver usando uma sessão privada do navegador, talvez seja necessário habilitar novamente essa configuração sempre.

> [!NOTE]
> Alguns navegadores se referem a essa configuração como "todos os cookies", em vez de "cookies de terceiros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instruções para ajustar configurações de cookie em navegadores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Borda](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites de dados

Há limites sobre a quantidade de dados do Excel que podem ser transferidos ao mesmo tempo e quantas transações individuais do Power Automate podem ser realizadas.

### <a name="excel"></a>Excel

Excel para a Web tem as seguintes limitações ao fazer chamadas para a pasta de trabalho por meio de um script:

- As solicitações e respostas são limitadas a **5MB**.
- Um intervalo é limitado a **cinco milhões de células**.

Se você estiver encontrando erros ao lidar com conjuntos de dados grandes, tente usar vários intervalos menores em vez de intervalos maiores. Para obter um exemplo, consulte o exemplo [Gravar um conjunto de dados grande](../resources/samples/write-large-dataset.md) . Você também pode usar APIs como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) para direcionar células específicas em vez de grandes intervalos.

Os limites do Excel que não são específicos para scripts do Office podem ser encontrados no artigo [Especificações e limites do Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

### <a name="power-automate"></a>Power Automate

Ao usar scripts do Office com o Power Automate, cada usuário é limitado a **1.600 chamadas para a ação Executar Script por dia**. Esse limite é redefinido às 12:00 UTC.

A plataforma Power Automate também tem limitações de uso, que podem ser encontradas nos artigos a seguir.

- [Limites e configuração no Power Automate](/power-automate/limits-and-config)
- [Problemas e limitações conhecidos para o conector do Excel Online (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Se você tiver um script de longa execução, esteja ciente do [tempo limite de 120 segundos para operações síncronas do Power Automate](/power-automate/limits-and-config#timeout). Você precisará [otimizar seu script](../develop/web-client-performance.md) ou dividir sua automação do Excel em vários scripts.

## <a name="see-also"></a>Confira também

- [Especificações e limites do Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Solucionar problemas de scripts do Office](troubleshooting.md)
- [Desfazer os efeitos do Scripts do Office](undo.md)
- [Melhorar o desempenho dos scripts do Office](../develop/web-client-performance.md)
