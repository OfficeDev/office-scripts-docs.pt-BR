---
title: Solução de problemas de informações para Power Automate com Office Scripts
description: Dicas, informações da plataforma e problemas conhecidos com a integração entre Office Scripts e Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: bcfedb8db88d74f16e46c604121bceff3c7c7382
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232645"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Solução de problemas de informações para Power Automate com Office Scripts

Power Automate permite que você leve sua automação Office Script para o próximo nível. No entanto, como Power Automate executa scripts em seu nome em sessões Excel independentes, há algumas coisas importantes a observar.

> [!TIP]
> Se você estiver apenas começando a usar Office scripts com Power Automate, comece com [Executar scripts](../develop/power-automate-integration.md) Office com Power Automate para saber mais sobre as plataformas.

## <a name="avoid-using-relative-references"></a>Evite usar referências relativas

Power Automate executa seu script na Excel de trabalho escolhida em seu nome. A workbook pode ser fechada quando isso acontece. Qualquer API que se basei no estado atual do usuário, como , pode se comportar de forma `Workbook.getActiveWorksheet` diferente Power Automate. Isso porque as APIs se baseiam em uma posição relativa do cursor ou exibição do usuário e essa referência não existe em um fluxo Power Automate usuário.

Algumas APIs de referência relativa lançam erros Power Automate. Outras têm um comportamento padrão que implica no estado de um usuário. Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos. Isso torna o fluxo Power Automate consistente, mesmo que as planilhas sejam reorganizadas.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Métodos de script que falham ao executar Power Automate fluxos

Os métodos a seguir lançarão um erro e falharão quando chamados de um script em Power Automate fluxo.

| Classe | Método |
|--|--|
| [Gráfico](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script com um comportamento padrão em fluxos Power Automate fluxos

Os métodos a seguir usam um comportamento padrão, em vez do estado atual de qualquer usuário.

| Classe | Método | Power Automate comportamento |
|--|--|--|
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Retorna a primeira planilha da pasta de trabalho ou a planilha atualmente ativada pelo `Worksheet.activate` método. |
| [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca a planilha como a planilha ativa para fins de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Selecionar pasta de trabalho com o controle do navegador de arquivos

Ao criar a **etapa Executar script** de um fluxo Power Automate, você precisa selecionar qual workbook faz parte do fluxo. Use o navegador de arquivos para selecionar sua pasta de trabalho, em vez de digitar manualmente o nome da pasta de trabalho.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="A Power Automate executar script mostrando a opção Mostrar navegador de arquivos do Se picker":::

Para obter mais contexto sobre a limitação Power Automate e uma discussão sobre possíveis soluções alternativas para a seleção dinâmica de workbooks, consulte este thread no [Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Diferenças de fuso horário

Excel arquivos não têm um local ou zona de tempo inerente. Sempre que um usuário abre a workbook, sua sessão usa o período de tempo local desse usuário para cálculos de data. Power Automate sempre usa UTC.

Se o script usa datas ou horas, pode haver diferenças comportamentais quando o script é testado localmente em comparação com quando ele é executado por Power Automate. Power Automate permite converter, formatar e ajustar tempos. Consulte Trabalhando com Datas e [Horas](https://flow.microsoft.com/blog/working-with-dates-and-times/) dentro de seus fluxos para obter instruções sobre como usar essas funções no Power Automate e [ `main` Parâmetros:](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) Passar dados para um script para saber como fornecer essas informações de tempo para o script.

## <a name="see-also"></a>Confira também

- [Solução de problemas dos scripts do Office](troubleshooting.md)
- [Executar Office scripts com Power Automate](../develop/power-automate-integration.md)
- [Excel Documentação de referência do conector Online (Business)](/connectors/excelonlinebusiness/)
