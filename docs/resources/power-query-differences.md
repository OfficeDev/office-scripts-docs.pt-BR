---
title: Quando usar Power Query ou Scripts do Office
description: Os cenários mais adequados para as plataformas Power Query e Office Scripts.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585902"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Quando usar Power Query ou Scripts do Office

[Power Query](https://powerquery.microsoft.com) e Office scripts são soluções de automação poderosas para Excel. Ambas as soluções permitem Excel os usuários limpem e transformem dados em guias de trabalho. Um único Power Query ou Office Script pode ser atualizado e reprisado em novos dados para produzir resultados consistentes, o que economiza tempo e permite que você trabalhe com as informações resultantes mais rapidamente.

Este artigo fornece uma visão geral de quando você pode favorecer uma plataforma em relação à outra. Em geral, o Power Query é bom para puxar e transformar dados de grandes fontes de dados externas e scripts do Office são bons para soluções rápidas e centradas em Excel e integrações [Power Automate.](../develop/power-automate-integration.md)

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Grandes fontes de dados e recuperação de dados: Power Query

Recomendamos Power Query ao lidar com fontes de dados de plataformas com suporte.

Power Query tem [conexões de dados integrados](https://powerquery.microsoft.com/connectors/) a centenas de fontes. Power Query é especialmente projetado para tarefas de recuperação, transformação e combinação de dados. Quando você precisa de dados de uma dessas fontes, Power Query oferece uma maneira sem código de trazer esses dados para Excel na forma de que você precisa.

Essas Power Query são projetadas para conjuntos de dados grandes. Eles não têm os mesmos limites de [transferência que](../testing/platform-limits.md) Power Automate ou Excel para a Web.

Office scripts oferecem uma solução leve para fontes de dados menores ou fontes de dados não cobertas por Power Query conectores. Isso inclui o [uso ou `fetch` APIs REST](../develop/external-calls.md) ou obter informações de fontes de dados ad hoc, como um cartão [Teams adaptável](../resources/scenarios/task-reminders.md).

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Formatação, visualizações e controle programático: Office Scripts

Recomendamos Office scripts quando suas necessidades vão além da importação e transformação de dados.

Quase tudo o que você pode fazer manualmente Excel interface do usuário é possível com Office Scripts. Eles são ótimos para aplicar formatação consistente a guias de trabalho. Os scripts criam gráficos, tabelas dinâmicas, formas, imagens e outras visualizações de planilha. Os scripts também dão controle preciso sobre as posições, tamanhos, cores e outros atributos dessas visualizações.

A inclusão do código TypeScript oferece um alto grau de personalização. A lógica de controle programática, como instruções `if...else` , torna o script robusto. Isso permite que você faça coisas como ler dados condicionalmente sem depender de fórmulas Excel complexas ou examinar a workbook em busca de alterações inesperadas antes de alterar a workbook.

A formatação pode ser aplicada com Power Query por meio Excel [modelos](https://templates.office.com/power-query-tutorial-tm11414620). No entanto, os modelos são atualizados no nível individual ou da organização, enquanto Office scripts oferecem controle de acesso mais granular.

## <a name="power-automate-integrations"></a>Power Automate integrações

Office scripts oferecem mais opções para Power Automate integração. Os scripts são personalizados para suas soluções. Você define [a entrada e a saída do script](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), para que ele funcione com qualquer outro conector ou dados no fluxo. A captura de tela a seguir mostra um exemplo Power Automate fluxo que passa dados de um cartão adaptável Teams para um script Office.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Uma captura de tela que mostra o conector Excel Online (Business) no designer de fluxo. O conector está usando a ação Executar script para tomar a entrada de um cartão adaptável Teams e formi-lo para um script.":::

Power Query é usado no conector [SQL Server](https://powerquery.microsoft.com/flow/) Power Automate. A [ação Transformar dados Power Query](/connectors/sql/#transform-data-using-power-query) permite que você crie uma consulta Power Automate. Embora essa seja uma ferramenta poderosa para uso com SQL Server, ela limita Power Query à fonte de entrada, conforme mostrado na captura de tela de fluxo a seguir.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Uma captura de tela que mostra o conector SQL Server no designer de fluxo. O conector está usando a ação Transformar dados Power Query.":::

## <a name="platform-dependencies"></a>Dependências da plataforma

Office scripts atualmente está disponível apenas para Excel na Web. Power Query atualmente está disponível apenas para Excel desktop. Ambos podem ser usados Power Automate, o que permite que o fluxo funcione com Excel de trabalho armazenadas em OneDrive.

## <a name="see-also"></a>Confira também

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query com Excel](https://powerquery.microsoft.com/excel/)
- [Executar Office scripts com Power Automate](../develop/power-automate-integration.md)
