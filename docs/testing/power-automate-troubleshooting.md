---
title: Informações de solução de problemas do Power Automate com scripts do Office
description: Dicas, informações de plataforma e problemas conhecidos com a integração entre scripts do Office e o Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b0f5b2f542216789f0d96f309cb7d799d201ba0f
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933263"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Informações de solução de problemas do Power Automate com scripts do Office

O Power Automate permite que você leve sua automação de Script do Office para o próximo nível. No entanto, como o Power Automate executa scripts em seu nome em sessões independentes do Excel, há algumas coisas importantes a observar.

> [!TIP]
> Se você estiver começando a usar scripts do Office com o Power Automate, comece com Executar scripts do Office com o [Power Automate](../develop/power-automate-integration.md) para saber mais sobre as plataformas.

## <a name="avoid-using-relative-references"></a>Evite usar referências relativas

O Power Automate executa seu script na planilha escolhida do Excel em seu nome. A workbook pode ser fechada quando isso acontece. Qualquer API que depende do estado atual do usuário, como, pode se comportar de `Workbook.getActiveWorksheet` maneira diferente no Power Automate. Isso porque as APIs se baseiam em uma posição relativa da exibição ou do cursor do usuário e essa referência não existe em um fluxo do Power Automate.

Algumas APIs de referência relativa lançam erros no Power Automate. Outros têm um comportamento padrão que implica no estado de um usuário. Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos. Isso torna o fluxo do Power Automate consistente, mesmo que as planilhas sejam reorganizadas.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Métodos de script que falham ao executar fluxos do Power Automate

Os métodos a seguir lançarão um erro e falharão quando chamados de um script em um fluxo do Power Automate.

| Classe | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script com um comportamento padrão em fluxos do Power Automate

Os métodos a seguir usam um comportamento padrão, em vez do estado atual de qualquer usuário.

| Classe | Method | Comportamento do Power Automate |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Retorna a primeira planilha na pasta de trabalho ou a planilha atualmente ativada pelo `Worksheet.activate` método. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca a planilha como a planilha ativa para fins `Workbook.getActiveWorksheet` de. |

## <a name="select-workbooks-with-the-file-browser-control"></a>Selecionar as pasta de trabalho com o controle do navegador de arquivos

Ao criar a **etapa Executar script** de um fluxo do Power Automate, você precisa selecionar qual a agenda faz parte do fluxo. Use o navegador de arquivos para selecionar sua pasta de trabalho, em vez de digitar manualmente o nome da pasta de trabalho.

![A opção de navegador de arquivo ao criar uma ação "Executar script" no Power Automate](../images/power-automate-file-browser.png)

Para obter mais contexto sobre a limitação do Power Automate e uma discussão sobre possíveis soluções alternativas para a seleção dinâmica de workbooks, consulte este thread na Comunidade do [Microsoft Power Automate.](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)

## <a name="time-zone-differences"></a>Diferenças de fuso horário

Os arquivos do Excel não têm um local inerente ou um zona de tempo. Sempre que um usuário abre a agenda, sua sessão usa o zona de tempo local desse usuário para cálculos de data. O Power Automate sempre usa UTC.

Se o script usa datas ou horas, pode haver diferenças comportamentais quando o script é testado localmente versus quando é executado por meio do Power Automate. O Power Automate permite converter, formatar e ajustar tempos. Consulte [](https://flow.microsoft.com/blog/working-with-dates-and-times/) Trabalhando com datas e horas dentro de seus fluxos para obter instruções sobre como usar essas funções no Power Automate e [ `main` parâmetros:](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) passando dados para um script para saber como fornecer essas informações de tempo para o script.

## <a name="see-also"></a>Confira também

- [Solução de problemas dos scripts do Office](troubleshooting.md)
- [Executar scripts do Office com o Power Automate](../develop/power-automate-integration.md)
- [Documentação de referência do conector do Excel Online (Business)](/connectors/excelonlinebusiness/)
