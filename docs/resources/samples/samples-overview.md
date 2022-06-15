---
title: Office exemplos de scripts
description: Exemplos Office scripts e cenários disponíveis.
ms.date: 06/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 059a21711f36273207f1e65ba844ba23c502ffe0
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088047"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office exemplos e cenários de scripts

Esta seção contém [soluções Office automação baseadas em Scripts](../../overview/excel.md) que ajudam os usuários finais a alcançar a automação de tarefas diárias. Ele contém cenários realistas que os usuários empresariais enfrentam e fornece soluções detalhadas junto com links de vídeo instrucionais passo a passo.

Para cada um dos projetos em [Noções Básicas](#basics) e Além dos [conceitos básicos](#beyond-the-basics), confira o código-fonte, os vídeos passo a passo do [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) e muito mais.

Em [Cenários](#scenarios), incluímos alguns exemplos de cenários maiores que demonstram casos de uso do mundo real.

Também damos [boas-vindas às contribuições da comunidade](#community-contributions-and-fun-samples). Esses exemplos são código aberto.

> [!IMPORTANT]
> Certifique-se de atender aos pré-requisitos para Office scripts antes de experimentar os exemplos. Os requisitos para sua assinatura Microsoft 365 conta são [encontrados na seção "Requisitos" Office scripts do Excel visão geral](../../overview/excel.md#requirements).

## <a name="basics"></a>Noções básicas

| Project | Detalhes |
|---------|---------|
| [Noções básicas de scripts](../excel-samples.md) | Esses exemplos demonstram blocos de construção fundamentais para Office Scripts. |
| [Adicionar comentários no Excel](add-excel-comments.md) | Este exemplo adiciona comentários a uma célula, incluindo @mentioning um colega. |
| [Adicionar imagens a uma pasta de trabalho](add-image-to-workbook.md) | Este exemplo adiciona uma imagem a uma pasta de trabalho e copia uma imagem entre planilhas.|
| [Copiar várias Excel tabelas em uma única tabela](copy-tables-combine.md) | Este exemplo combina dados de várias Excel tabelas em uma única tabela que inclui todas as linhas. |
| [Criar um sumário da pasta de trabalho](table-of-contents.md) | Este exemplo cria um sumário com links para cada planilha. |

## <a name="beyond-the-basics"></a>Além do básico

Confira o projeto de ponta a ponta a seguir, que automatiza cenários de exemplo, juntamente com scripts completos, exemplos Excel arquivos usados e vídeos (hospedados no [YouTube)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Project | Detalhes |
|---------|---------|
| [Combinar planilhas em uma única pasta de trabalho](combine-worksheets-into-single-workbook.md) | Este exemplo usa Office Scripts e Power Automate para efetuar pull de dados de outras pastas de trabalho em uma única pasta de trabalho. |
| [Converter arquivos CSV em Excel de trabalho](convert-csv.md) | Este exemplo usa Office scripts e Power Automate para criar .xlsx arquivos .csv arquivos. |
| [Pastas de trabalho de referência cruzada](excel-cross-reference.md) | Este exemplo usa Office scripts e Power Automate para fazer referência cruzada e validar informações em pastas de trabalho diferentes. |
| [Contar linhas em branco em uma planilha específica ou em todas as planilhas](count-blank-rows.md) | Este exemplo detecta se há linhas em branco em planilhas em que você prevê que os dados estejam presentes e, em seguida, relata a contagem de linhas em branco para uso em um fluxo de Power Automate. |
| [Imagens de tabela e gráfico de email](email-images-chart-table.md) | Este exemplo usa Office scripts e Power Automate ações para criar um gráfico e enviar esse gráfico como uma imagem por email. |
| [Chamadas de busca externas](external-fetch-calls.md) | Este exemplo usa `fetch` para obter informações de GitHub para o script. |
| [Filtrar Excel tabela e obter intervalo visível](filter-table-get-visible-range.md) | Este exemplo filtra uma Excel e retorna o intervalo visível como um objeto JSON. Esse JSON pode ser fornecido a um fluxo Power Automate como parte de uma solução maior. |
| [Gerenciar o modo de cálculo no Excel](excel-calculation.md) | Este exemplo mostra como usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts. |
| [Mover linhas entre tabelas](move-rows-across-tables.md) | Este exemplo mostra como mover linhas entre tabelas salvando filtros e, em seguida, processando e reaplicação dos filtros. |
| [Saída Excel dados como JSON](get-table-data.md) | Esta solução mostra como gerar Excel dados da tabela como JSON a serem Power Automate. |
| [Remover hiperlinks de cada célula em uma Excel de trabalho](remove-hyperlinks-from-cells.md) | Este exemplo limpa todos os hiperlinks da planilha atual. |
| [Executar um script em todos os arquivos do Excel em uma pasta](automate-tasks-on-all-excel-files-in-folder.md) | Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta no OneDrive for Business (também pode ser usado para uma pasta SharePoint dados). Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions um colega. |
| [Escrever um grande conjuntos de dados](write-large-dataset.md) | Este exemplo mostra como enviar um intervalo grande como subintervalos menores. |

## <a name="scenarios"></a>Cenários

Office scripts podem automatizar partes da sua rotina diária. Essas tarefas diárias geralmente existem em ecossistemas exclusivos, com Excel de trabalho que são configuradas de maneiras específicas. Esses exemplos de cenários maiores demonstram esses casos de uso do mundo real. Eles incluem os scripts Office e as pastas de trabalho, para que você possa ver o cenário de ponta a ponta.

| Cenário | Detalhes |
|---------|---------|
| [Analisar downloads da Web](../scenarios/analyze-web-downloads.md) | Esse cenário apresenta um script que analisa registros de tráfego da Web para determinar o país de origem de um usuário. Ele demonstra as habilidades de análise de texto, usando subfunções em scripts, aplicando formatação condicional e trabalhando com tabelas. |
| [Buscar e representar graficamente os dados do nível de água do NOAA](../scenarios/noaa-data-fetch.md) | Esse cenário usa um script Office para efetuar pull de dados de uma fonte externa (o banco de dados [Tides e Currents do NOAA](https://tidesandcurrents.noaa.gov/)) e grafar as informações resultantes. Ele destaca as habilidades de uso para `fetch` obter dados e usar gráficos. |
| [Calculadora de notas](../scenarios/grade-calculator.md) | Esse cenário apresenta um script que valida o registro de um instrutor para as notas da classe. Ele demonstra as habilidades de verificação de erros, formatação de célula e expressões regulares. |
| [Agendar entrevistas no Teams](../scenarios/schedule-interviews-in-teams.md) | Este cenário mostra como usar uma planilha Excel para gerenciar horários de reunião de entrevista e fazer um fluxo para agendar reuniões em Teams. |
| [Lembretes de tarefas](../scenarios/task-reminders.md) | Esse cenário usa um Office script em um fluxo Power Automate para enviar lembretes aos colegas de trabalho para atualizar o status de um projeto. Ele destaca as habilidades de Power Automate integração e transferência de dados de e para scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contribuições e exemplos divertidos

Damos [as boas-vindas](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) às contribuições da nossa Office scripts! Fique à vontade para criar uma solicitação de pull para revisão.

| Project | Detalhes |
|---------|---------|
| [Jogo da Vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | O blog "Ready Player Zero", de Yutao Huang no Excel Tech Community, inclui um script para o modelo De Vida de John [*Classes.*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) |
| [Botão Relógio do Sistema](../scenarios/punch-clock.md) | Este script foi contribuido por [Brian Gonzalez](https://github.com/b-gonzalez). O cenário apresenta um script e um botão de script que registra a hora atual. |
| [Animação de saudações de estações](community-seasons-greetings.md) | Este roteiro foi contribuido [por Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) no espírito da temporada de feriados! É um script divertido que mostra uma árvore de Natal cantando em Excel na Web usando Office Scripts. |

## <a name="leave-a-comment"></a>Deixe um comentário

Fique à vontade para deixar um comentário, fazer uma sugestão ou registrar um problema usando a  seção Comentários na parte inferior da página de documentação do exemplo específico.
