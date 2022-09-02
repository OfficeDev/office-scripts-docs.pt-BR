---
title: Exemplos de Scripts do Office
description: Exemplos e cenários de Scripts do Office disponíveis.
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5798da37bd4166d18b41c005c4d8cc8a4b6c401d
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572483"
---
# <a name="office-scripts-samples-and-scenarios"></a>Exemplos e cenários de Scripts do Office

Esta seção contém [soluções de automação baseadas em Scripts do Office](../../overview/excel.md) que ajudam os usuários finais a alcançar a automação de tarefas diárias. Ele contém cenários realistas que os usuários empresariais enfrentam e fornece soluções detalhadas junto com links de vídeo instrucionais passo a passo.

Para cada um dos projetos em [Noções Básicas](#basics) e Além dos [conceitos básicos](#beyond-the-basics), confira o código-fonte, os vídeos passo a passo do [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) e muito mais.

Em [Cenários](#scenarios), incluímos alguns exemplos de cenários maiores que demonstram casos de uso do mundo real.

Também damos [boas-vindas às contribuições da comunidade](#community-contributions-and-fun-samples). Esses exemplos são código aberto.

> [!IMPORTANT]
> Certifique-se de atender aos pré-requisitos dos Scripts do Office antes de experimentar os exemplos. Os requisitos para sua assinatura e conta do Microsoft 365 são [encontrados na seção "Requisitos" da visão geral dos Scripts do Office para Excel](../../overview/excel.md#requirements).

## <a name="basics"></a>Noções básicas

| Project | Detalhes |
|---------|---------|
| [Noções básicas de scripts](excel-samples.md) | Esses exemplos demonstram blocos de construção fundamentais para Scripts do Office. |
| [Adicionar comentários no Excel](add-excel-comments.md) | Este exemplo adiciona comentários a uma célula, incluindo @mentioning um colega. |
| [Adicionar imagens a uma pasta de trabalho](add-image-to-workbook.md) | Este exemplo adiciona uma imagem a uma pasta de trabalho e copia uma imagem entre planilhas.|
| [Copiar várias tabelas do Excel em uma única tabela](copy-tables-combine.md) | Este exemplo combina dados de várias tabelas do Excel em uma única tabela que inclui todas as linhas. |
| [Criar um sumário da pasta de trabalho](table-of-contents.md) | Este exemplo cria um sumário com links para cada planilha. |
| [Remover filtros de coluna de tabela](clear-table-filter-for-active-cell.md) | Este exemplo limpa todos os filtros de uma coluna de tabela. |
| [Registrar alterações diárias no Excel e reportá-las com um fluxo do Power Automate](report-day-to-day-changes.md) | Este exemplo usa um fluxo agendado do Power Automate para registrar leituras diárias e relatar as alterações. |

## <a name="beyond-the-basics"></a>Além do básico

Confira o projeto de ponta a ponta a seguir que automatiza cenários de exemplo junto com scripts completos, arquivos do Excel de exemplo usados e vídeos (hospedados [no YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Detalhes |
|---------|---------|
| [Combinar planilhas em uma única pasta de trabalho](combine-worksheets-into-single-workbook.md) | Este exemplo usa Scripts do Office e o Power Automate para efetuar pull de dados de outras pastas de trabalho em uma única pasta de trabalho. |
| [Converter arquivos CSV em pastas de trabalho do Excel](convert-csv.md) | Este exemplo usa Scripts do Office e o Power Automate para criar .xlsx arquivos .csv arquivos. |
| [Pastas de trabalho de referência cruzada](excel-cross-reference.md) | Este exemplo usa Scripts do Office e o Power Automate para fazer referência cruzada e validar informações em pastas de trabalho diferentes. |
| [Contar linhas em branco em uma planilha específica ou em todas as planilhas](count-blank-rows.md) | Este exemplo detecta se há linhas em branco em planilhas em que você prevê que os dados estejam presentes e, em seguida, relata a contagem de linhas em branco para uso em um fluxo do Power Automate. |
| [Email imagens de tabela e gráfico](email-images-chart-table.md) | Este exemplo usa scripts do Office e ações do Power Automate para criar um gráfico e enviar esse gráfico como uma imagem por email. |
| [Chamadas de busca externas](external-fetch-calls.md) | Este exemplo usa `fetch` para obter informações do GitHub para o script. |
| [Gerenciar o modo de cálculo no Excel](excel-calculation.md) | Este exemplo mostra como usar o modo de cálculo e calcular métodos em Excel na Web usando scripts do Office. |
| [Mover linhas entre tabelas](move-rows-across-tables.md) | Este exemplo mostra como mover linhas entre tabelas salvando filtros e, em seguida, processando e reaplicação dos filtros. |
| [Saída de dados do Excel como JSON](get-table-data.md) | Esta solução mostra como gerar dados de tabela do Excel como JSON para usar no Power Automate. |
| [Remover hiperlinks de cada célula em uma planilha do Excel](remove-hyperlinks-from-cells.md) | Este exemplo limpa todos os hiperlinks da planilha atual. |
| [Executar um script em todos os arquivos do Excel em uma pasta](automate-tasks-on-all-excel-files-in-folder.md) | Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta no OneDrive for Business (também pode ser usado para uma pasta do SharePoint). Ele executa cálculos nos arquivos do Excel, adiciona formatação e insere um comentário que @mentions um colega. |
| [Escrever um grande conjuntos de dados](write-large-dataset.md) | Este exemplo mostra como enviar um intervalo grande como subintervalos menores. |

## <a name="scenarios"></a>Cenários

Os Scripts do Office podem automatizar partes da sua rotina diária. Essas tarefas diárias geralmente existem em ecossistemas exclusivos, com pastas de trabalho do Excel configuradas de maneiras específicas. Esses exemplos de cenários maiores demonstram esses casos de uso do mundo real. Eles incluem os Scripts do Office e as pastas de trabalho, para que você possa ver o cenário de ponta a ponta.

| Cenário | Detalhes |
|---------|---------|
| [Analisar downloads da Web](../scenarios/analyze-web-downloads.md) | Esse cenário apresenta um script que analisa registros de tráfego da Web para determinar o país de origem de um usuário. Ele demonstra as habilidades de análise de texto, usando subfunções em scripts, aplicando formatação condicional e trabalhando com tabelas. |
| [Buscar e representar graficamente os dados do nível de água do NOAA](../scenarios/noaa-data-fetch.md) | Esse cenário usa um Script do Office para efetuar pull de dados de uma fonte externa (o banco de dados [Tides e Currents do NOAA](https://tidesandcurrents.noaa.gov/)) e grafar as informações resultantes. Ele destaca as habilidades de uso para `fetch` obter dados e usar gráficos. |
| [Calculadora de notas](../scenarios/grade-calculator.md) | Esse cenário apresenta um script que valida o registro de um instrutor para as notas da classe. Ele demonstra as habilidades de verificação de erros, formatação de célula e expressões regulares. |
| [Agendar entrevistas no Teams](../scenarios/schedule-interviews-in-teams.md) | Este cenário mostra como usar uma planilha do Excel para gerenciar horários de reunião de entrevista e fazer um fluxo para agendar reuniões no Teams. |
| [Lembretes de tarefas](../scenarios/task-reminders.md) | Este cenário usa um Script do Office em um fluxo do Power Automate para enviar lembretes a colegas de trabalho para atualizar o status de um projeto. Ele destaca as habilidades de integração e transferência de dados do Power Automate de e para scripts. |

## <a name="community-contributions-and-fun-samples"></a>Contribuições da comunidade e amostras divertidas

Damos [as boas-vindas](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) às contribuições da nossa comunidade de Scripts do Office! Fique à vontade para criar uma solicitação de pull para revisão.

| Project | Detalhes |
|---------|---------|
| [Jogo da Vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | O blog "Ready Player Zero", de Yutao Huang, no Excel Tech Community, inclui um script para o modelo The [*Game of Life de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) John Tutorial. |
| [Botão Relógio do Sistema](../scenarios/punch-clock.md) | Este script foi contribuido por [Brian Gonzalez](https://github.com/b-gonzalez). O cenário apresenta um script e um botão de script que registra a hora atual. |
| [Animação de saudações de estações](community-seasons-greetings.md) | Este roteiro foi contribuido [por Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) no espírito da temporada de feriados! É um script divertido que mostra uma árvore de Natal cantando no Excel na Web usando Scripts do Office. |

## <a name="leave-a-comment"></a>Deixe um comentário

Fique à vontade para deixar um comentário, fazer uma sugestão ou registrar um problema usando a  seção Comentários na parte inferior da página de documentação do exemplo específico.
