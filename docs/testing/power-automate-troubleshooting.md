---
title: Solucionar Office scripts em execução no Power Automate
description: Dicas, informações da plataforma e problemas conhecidos com a integração entre Office Scripts e Power Automate.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 028c34003a6f6b00c9afc67450b249b938d445fb
ms.sourcegitcommit: 634ad2061e683ae1032c1e0b55b00ac577adc34f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/03/2021
ms.locfileid: "60725622"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Solucionar Office scripts em execução no Power Automate

Power Automate permite que você leve sua automação Office Script para o próximo nível. No entanto, como Power Automate executa scripts em seu nome em sessões Excel independentes, há algumas coisas importantes a observar.

> [!TIP]
> Se você estiver apenas começando a usar Office scripts com Power Automate, comece com [Executar scripts](../develop/power-automate-integration.md) Office com Power Automate para saber mais sobre as plataformas.

## <a name="avoid-relative-references"></a>Evitar referências relativas

Power Automate executa seu script na Excel de trabalho escolhida em seu nome. A workbook pode ser fechada quando isso acontece. Qualquer API que se basei no estado atual do usuário, como , pode se comportar de forma `Workbook.getActiveWorksheet` diferente Power Automate. Isso porque as APIs se baseiam em uma posição relativa do cursor ou exibição do usuário e essa referência não existe em um fluxo Power Automate usuário.

Algumas APIs de referência relativa lançam erros Power Automate. Outras têm um comportamento padrão que implica no estado de um usuário. Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos. Isso torna o fluxo Power Automate consistente, mesmo que as planilhas sejam reorganizadas.

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>Métodos de script que falham quando executados em Power Automate fluxos

Os métodos a seguir lançam um erro e falham quando chamados de um script em Power Automate fluxo.

| Classe | Método |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script com um comportamento padrão em fluxos Power Automate fluxos

Os métodos a seguir usam um comportamento padrão, em vez do estado atual de qualquer usuário.

| Classe | Método | Power Automate comportamento |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Retorna a primeira planilha da pasta de trabalho ou a planilha atualmente ativada pelo `Worksheet.activate` método. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca a planilha como a planilha ativa para fins de `Workbook.getActiveWorksheet` . |

## <a name="data-refresh-not-supported-in-power-automate"></a>Atualização de dados não suportada em Power Automate

Office Os scripts não podem atualizar dados quando executados Power Automate. Métodos como `PivotTable.refresh` não fazer nada quando chamado em um fluxo. Além disso, Power Automate não dispara uma atualização de dados para fórmulas que usam links de workbook.

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>Métodos de script que não fazem nada quando executados Power Automate fluxos

Os métodos a seguir não fazem nada em um script quando chamados por Power Automate. Eles ainda retornam com êxito e não lançam erros.

| Classe | Método |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>Selecionar pasta de trabalho com o controle do navegador de arquivos

Ao criar a **etapa Executar script** de um fluxo Power Automate, você precisa selecionar qual workbook faz parte do fluxo. Use o navegador de arquivos para selecionar sua pasta de trabalho, em vez de digitar manualmente o nome da pasta de trabalho.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="A Power Automate executar script mostrando a opção Mostrar navegador de arquivo do Se picker.":::

Para obter mais contexto sobre a limitação Power Automate e uma discussão sobre possíveis soluções alternativas para a seleção dinâmica de workbooks, consulte este thread no [Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="pass-entire-arrays-as-script-parameters"></a>Passar matrizes inteiras como parâmetros de script

Power Automate permite que os usuários passem matrizes para conectores como uma variável ou como elementos individuais na matriz. O padrão é passar elementos individuais, que constrói a matriz no fluxo. Para scripts ou outros conectores que têm matrizes inteiras como argumentos, você precisa selecionar o **botão Alternar** para inserir todo o botão de matriz para passar a matriz como um objeto completo. Este botão está no canto superior direito de cada campo de entrada de parâmetro de matriz.

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="O botão a ser alternado para inserir uma matriz inteira em uma caixa de entrada do campo de controle.":::

## <a name="time-zone-differences"></a>Diferenças de fuso horário

Excel arquivos não têm um local ou zona de tempo inerente. Sempre que um usuário abre a workbook, sua sessão usa o período de tempo local desse usuário para cálculos de data. Power Automate sempre usa UTC.

Se o script usa datas ou horas, pode haver diferenças comportamentais quando o script é testado localmente em comparação com quando ele é executado por Power Automate. Power Automate permite converter, formatar e ajustar tempos. Consulte Trabalhando com Datas e [Horas](https://flow.microsoft.com/blog/working-with-dates-and-times/) dentro de seus fluxos para obter instruções sobre como usar essas funções no Power Automate [ `main` e Parâmetros:](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) Passar dados para um script para saber como fornecer essas informações de tempo para o script.

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>Campos de parâmetro de script ou saída retornada não aparecendo no Power Automate

Há dois motivos para que os parâmetros ou os dados retornados de um script não sejam refletidos com precisão no Power Automate de fluxo.

- A assinatura de script (os parâmetros ou o valor de retorno) foi alterada desde que o **conector Excel Business (Online)** foi adicionado.
- A assinatura de script usa tipos sem suporte. Verifique seus tipos em relação às [](../develop/power-automate-integration.md#return-data-from-a-script) listas sob os [parâmetros](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) e retorna seções de Executar Office [Scripts com Power Automate](../develop/power-automate-integration.md) artigo.

A assinatura de um script é armazenada com o **conector Excel Business (Online)** quando ele é criado. Remova o conector antigo e crie um novo para obter os parâmetros mais recentes e retornar valores para a **ação Executar script.**

## <a name="see-also"></a>Confira também

- [Solucionar Office scripts](troubleshooting.md)
- [Executar Office scripts com Power Automate](../develop/power-automate-integration.md)
- [Excel Documentação de referência do conector Online (Business)](/connectors/excelonlinebusiness/)
