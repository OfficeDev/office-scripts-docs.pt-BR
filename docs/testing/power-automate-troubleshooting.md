---
title: Solução de problemas Office Scripts em execução em Power Automate
description: Dicas, informações da plataforma e problemas conhecidos com a integração entre Office Scripts e Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545564"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Solução de problemas Office Scripts em execução em Power Automate

Power Automate permite que você leve sua automação de script Office para o próximo nível. No entanto, como Power Automate executa scripts em seu nome em sessões independentes de Excel, há algumas coisas importantes a notar.

> [!TIP]
> Se você está apenas começando a usar Office Scripts com Power Automate, por favor, comece com [Run Office Scripts com Power Automate](../develop/power-automate-integration.md) para aprender sobre as plataformas.

## <a name="avoid-relative-references"></a>Evite referências relativas

Power Automate executa seu script na pasta de trabalho Excel escolhida em seu nome. A pasta de trabalho pode estar fechada quando isso acontecer. Qualquer API que dependa do estado atual do usuário, como, por mais `Workbook.getActiveWorksheet` que, não se comporte de forma diferente em Power Automate. Isso ocorre porque as APIs são baseadas em uma posição relativa da visão ou cursor do usuário e essa referência não existe em um fluxo Power Automate.

Algumas APIs de referência relativa jogam erros em Power Automate. Outros têm um comportamento padrão que implica o estado de um usuário. Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos. Isso torna o fluxo de Power Automate consistente, mesmo que as planilhas sejam reorganizadas.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Métodos de script que falham quando executados Power Automate flui

Os seguintes métodos jogarão um erro e falharão quando chamados de um script em um fluxo de Power Automate.

| Classe | Método |
|--|--|
| [Gráfico](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script com um comportamento padrão em fluxos de Power Automate

Os seguintes métodos utilizam um comportamento padrão, em vez do estado atual de qualquer usuário.

| Classe | Método | Power Automate comportamento |
|--|--|--|
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Retorna a primeira planilha na pasta de trabalho ou a planilha atualmente ativada pelo `Worksheet.activate` método. |
| [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca a planilha como a planilha ativa para fins de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Selecione pastas de trabalho com o controle do navegador de arquivos

Ao construir a etapa de **script executar** de um fluxo de Power Automate, você precisa selecionar qual pasta de trabalho faz parte do fluxo. Use o navegador de arquivos para selecionar sua pasta de trabalho, em vez de digitar manualmente o nome da pasta de trabalho.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="A ação de script Power Automate Run mostrando a opção do navegador de arquivos Show Picker":::

Para obter mais contexto sobre a limitação Power Automate e uma discussão sobre possíveis soluções alternativas para a seleção dinâmica de livros, consulte [este segmento no Power Automate Community da Microsoft](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Diferenças de fuso horário

Excel arquivos não têm uma localização ou fuso horário inerentes. Toda vez que um usuário abre a pasta de trabalho, sua sessão usa o fuso horário local do usuário para cálculos de data. Power Automate sempre usa UTC.

Se o seu script usa datas ou horários, pode haver diferenças comportamentais quando o script é testado localmente versus quando ele é executado através de Power Automate. Power Automate permite converter, formatar e ajustar os tempos. Consulte [Trabalhando com datas e horários dentro de seus fluxos](https://flow.microsoft.com/blog/working-with-dates-and-times/) para obter instruções sobre como usar essas funções em Power Automate e [ `main` Parâmetros: Passe dados para um script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) para aprender a fornecer informações de tempo para o script.

## <a name="see-also"></a>Confira também

- [Solução de problemas Office Scripts](troubleshooting.md)
- [Execute Office scripts com Power Automate](../develop/power-automate-integration.md)
- [Excel Documentação de referência do conector online (business)](/connectors/excelonlinebusiness/)
