---
title: Office Exemplos de scripts
description: Exemplos Office scripts e cenários disponíveis.
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0d11e15a7e839f33a74ca8ad7f1d09dd7711347c
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59334931"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Exemplos e cenários de scripts

Esta seção contém Office [de automação baseadas](../../overview/excel.md) em Scripts que ajudam os usuários finais a alcançar a automação de tarefas diárias. Ele contém cenários realistas que os usuários de negócios enfrentam e fornece soluções detalhadas juntamente com links de vídeo instrucional passo a passo.

Para cada um dos projetos em [Noções Básicas](#basics) e Além das noções [básicas,](#beyond-the-basics)confira o código-fonte, os vídeos passo a passo do [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)e muito mais.

Em [Cenários,](#scenarios)incluímos alguns exemplos de cenários maiores que demonstram casos de uso no mundo real.

Também saudamos [as contribuições da comunidade.](#community-contributions-and-fun-samples)

## <a name="basics"></a>Noções básicas

| Projeto | Detalhes |
|---------|---------|
| [Noções básicas de scripts](../excel-samples.md) | Esses exemplos demonstram blocos de construção fundamentais para Office Scripts. |
| [Adicionar comentários em Excel](add-excel-comments.md) | Este exemplo adiciona comentários a uma célula, incluindo @mentioning um colega. |
| [Adicionar imagens a uma pasta de trabalho](add-image-to-workbook.md) | Este exemplo adiciona uma imagem a uma planilha e copia uma imagem entre folhas.|
| [Copiar várias Excel tabelas em uma única tabela](copy-tables-combine.md) | Este exemplo combina dados de várias Excel tabelas em uma única tabela que inclui todas as linhas. |

## <a name="beyond-the-basics"></a>Além do básico

Confira o seguinte projeto de ponta a ponta que automatiza cenários de exemplo, juntamente com scripts completos, exemplos Excel arquivos usados e [vídeos (hospedados no YouTube)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Projeto | Detalhes |
|---------|---------|
| [Combinar planilhas em uma única pasta de trabalho](combine-worksheets-into-single-workbook.md) | Este exemplo usa Office scripts e Power Automate para puxar dados de outras guias de trabalho para uma única workbook. |
| [Converter arquivos CSV em Excel pastas de trabalho](convert-csv.md) | Este exemplo usa Office scripts e Power Automate para criar .xlsx arquivos .csv arquivos. |
| [Workbooks de referência cruzada](excel-cross-reference.md) | Este exemplo usa Office scripts e Power Automate para fazer referência cruzada e validar informações em diferentes workbooks. |
| [Contar linhas em branco em uma planilha específica ou em todas as planilhas](count-blank-rows.md) | Este exemplo detecta se há linhas em branco em planilhas nas quais você antecipa a presença dos dados e relata a contagem de linhas em branco para uso em um fluxo de Power Automate. |
| [Imagens de gráfico de email e tabela](email-images-chart-table.md) | Este exemplo usa Office scripts e ações Power Automate para criar um gráfico e enviar esse gráfico como uma imagem por email. |
| [Chamadas de busca externas](external-fetch-calls.md) | Este exemplo usa `fetch` para obter informações do GitHub para o script. |
| [Filtrar Excel tabela e obter intervalo visível](filter-table-get-visible-range.md) | Este exemplo filtra uma Excel e retorna o intervalo visível como um objeto JSON. Esse JSON poderia ser fornecido para um fluxo Power Automate como parte de uma solução maior. |
| [Gerenciar o modo de cálculo Excel](excel-calculation.md) | Este exemplo mostra como usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts. |
| [Mover linhas entre tabelas](move-rows-across-tables.md) | Este exemplo mostra como mover linhas entre tabelas salvando filtros e, em seguida, processamento e reaplicação dos filtros. |
| [Saída Excel dados como JSON](get-table-data.md) | Esta solução mostra como Excel dados de tabela como JSON a ser usado Power Automate. |
| [Remover hiperlinks de cada célula em uma Excel de trabalho](remove-hyperlinks-from-cells.md) | Este exemplo limpa todos os hiperlinks da planilha atual. |
| [Executar um script em todos os arquivos do Excel em uma pasta](automate-tasks-on-all-excel-files-in-folder.md) | Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta OneDrive for Business (também pode ser usado para uma pasta SharePoint de dados). Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions um colega. |
| [Escrever um grande conjuntos de dados](write-large-dataset.md) | Este exemplo mostra como enviar um intervalo grande como sub-intervalos menores. |

## <a name="scenarios"></a>Cenários

Office Os scripts podem automatizar partes da sua rotina diária. Essas tarefas diárias geralmente existem em ecossistemas exclusivos, com Excel de trabalho que são configuradas de maneiras específicas. Esses exemplos de cenário maiores demonstram esses casos de uso no mundo real. Eles incluem o Office scripts e as guias de trabalho, para que você possa ver o cenário de ponta a ponta.

| Cenário | Detalhes |
|---------|---------|
| [Analisar downloads da Web](../scenarios/analyze-web-downloads.md) | Esse cenário apresenta um script que analisado registros de tráfego da Web para determinar o país de origem de um usuário. Ele mostra as habilidades de análise de texto, usando subfunções em scripts, aplicando formatação condicional e trabalhando com tabelas. |
| [Buscar e representar graficamente os dados do nível de água do NOAA](../scenarios/noaa-data-fetch.md) | Esse cenário usa um script de Office para obter dados de uma fonte externa (o banco de dados [do NoAA Tides e Currents)](https://tidesandcurrents.noaa.gov/)e fazer um gráfico das informações resultantes. Ele realça as habilidades de uso `fetch` para obter dados e usar gráficos. |
| [Calculadora de notas](../scenarios/grade-calculator.md) | Esse cenário apresenta um script que valida o registro de um instrutor para as notas da classe. Ele mostra as habilidades de verificação de erros, formatação de células e expressões regulares. |
| [Agendar entrevistas no Teams](../scenarios/schedule-interviews-in-teams.md) | Este cenário mostra como usar uma planilha Excel para gerenciar os horários de reunião de entrevista e fazer um fluxo para agendar reuniões em Teams. |
| [Lembretes de tarefas](../scenarios/task-reminders.md) | Esse cenário usa um script Office em um fluxo Power Automate para enviar lembretes aos colegas de trabalho para atualizar o status de um projeto. Ele realça as habilidades de Power Automate integração e transferência de dados de e para scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contribuições e exemplos divertidos

Saudamos [as contribuições](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nossa Office de Scripts! Sinta-se à vontade para criar uma solicitação pull para revisão.

| Project | Detalhes |
|---------|---------|
| [Jogo da Vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | O blog "Ready Player Zero" de Yutao Huang no Excel Tech Community inclui um script para modelar O [*Jogo da Vida de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)John Conway. |
| [Animação de saudações de estações](community-seasons-greetings.md) | Este script foi contribuído [por Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) no espírito da estação de feriado! É um script divertido que mostra uma árvore de Natal cantoria Excel na Web usando Office Scripts. |

## <a name="try-it-out"></a>Experimente

Esses exemplos são de código aberto. Experimente você mesmo. Você precisará de uma conta de trabalho ou de estudante da Microsoft do trabalho ou da escola com uma licença para Microsoft 365 assinatura (E3 ou superior). Basta ir para https://office.com entrar em sua conta e começar.

## <a name="leave-a-comment"></a>Deixar um comentário

Sinta-se à vontade para deixar um comentário, fazer uma sugestão ou registrar um problema usando a seção **Comentários** na parte inferior da página de documentação do exemplo específico.
