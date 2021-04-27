---
title: Office Exemplos de scripts
description: Exemplos Office scripts e cenários disponíveis.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: dc09db00cb63e6873b255360aff17ad2a56fa89e
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026825"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Exemplos e cenários de scripts

Esta seção contém Office [de automação baseadas](../../overview/excel.md) em Scripts que ajudam os usuários finais a alcançar a automação de tarefas diárias. Ele contém cenários realistas que os usuários de negócios enfrentam e fornece soluções detalhadas juntamente com links de vídeo instrucional passo a passo.

Para cada um dos projetos em [Noções Básicas](#basics) e Além das noções [básicas,](#beyond-the-basics)confira o código-fonte, os vídeos passo a passo do [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)e muito mais.

Em [Cenários,](#scenarios)incluímos alguns exemplos de cenários maiores que demonstram casos de uso no mundo real.

Também saudamos [as contribuições da comunidade.](#community-contributions)

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Noções básicas

| Project | Detalhes |
|---------|---------|
| [Noções básicas de scripts](../excel-samples.md) | Esses exemplos demonstram blocos de construção fundamentais para Office Scripts. |
| [Saiba noções básicas sobre como usar o objeto Range Office Scripts](range-basics.md) | Este artigo mostra as noções básicas de uso do objeto Range e suas APIs. Este é um tópico base que será usado em todos os outros projetos. |

## <a name="beyond-the-basics"></a>Além do básico

Confira o seguinte projeto de ponta a ponta que automatiza cenários de exemplo juntamente com scripts completos, exemplos Excel arquivos usados e [vídeos](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Project | Detalhes |
|---------|---------|
| [Adicionar comentários em Excel](add-excel-comments.md) | Este exemplo mostra como adicionar comentários a uma célula, incluindo @mentioning um colega. |
| [Contar linhas em branco em uma planilha específica ou em todas as planilhas](count-blank-rows.md) | Este exemplo detecta se há linhas em branco em planilhas nas quais você antecipa a presença dos dados e relata a contagem de linhas em branco para uso em um fluxo de Power Automate. |
| [Fazer referência cruzada e formatar um Excel arquivo](excel-cross-reference.md) | Esta solução mostra como dois arquivos Excel podem ser referenciados e formatados usando Office Scripts e Power Automate. |
| [Imagens de gráfico de email e tabela](email-images-chart-table.md) | Este exemplo usa Office scripts e ações Power Automate para criar um gráfico e enviar esse gráfico como uma imagem por email. |
| [Chamadas de busca externas](external-fetch-calls.md) | Este exemplo usa `fetch` para obter informações do GitHub para o script. |
| [Filtrar Excel tabela e obter intervalo visível](filter-table-get-visible-range.md) | Este exemplo filtra uma Excel e retorna o intervalo visível como um objeto JSON. Esse JSON poderia ser fornecido para um fluxo Power Automate como parte de uma solução maior. |
| [Gerar um identificador exclusivo em uma workbook](document-number-generator.md) | Esse cenário ajuda um usuário a gerar um número de documento exclusivo com um formato específico e a adicionar uma entrada a um intervalo ou tabela. |
| [Gerenciar o modo de cálculo Excel](excel-calculation.md) | Este exemplo mostra como usar o modo de cálculo e calcular métodos em Excel na Web usando Office Scripts. |
| [Mesclar várias Excel tabelas em uma única tabela](copy-tables-combine.md) | Este exemplo combina dados de várias Excel tabelas em uma única tabela que inclui todas as linhas. |
| [Mover linhas entre tabelas](move-rows-across-tables.md) | Este exemplo mostra como mover linhas entre tabelas salvando filtros e, em seguida, processamento e reaplicação dos filtros. |
| [Saída Excel dados como JSON](get-table-data.md) | Esta solução mostra como Excel dados de tabela como JSON a ser usado Power Automate. |
| [Remover hiperlinks de cada célula em uma Excel de trabalho](remove-hyperlinks-from-cells.md) | Este exemplo limpa todos os hiperlinks da planilha atual. |
| [Executar um script em todos os arquivos do Excel em uma pasta](automate-tasks-on-all-excel-files-in-folder.md) | Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta OneDrive for Business (também pode ser usado para uma pasta SharePoint de dados). Ele executa cálculos nos arquivos Excel, adiciona formatação e insere um comentário que @mentions um colega. |
| [Enviar uma Teams de dados Excel dados](send-teams-invite-from-excel-data.md) | Esta solução mostra como usar Office scripts e ações Power Automate para selecionar linhas do arquivo Excel e usá-lo para enviar um convite de reunião Teams e atualizar Excel. |

## <a name="scenarios"></a>Cenários

Office Os scripts podem automatizar partes da sua rotina diária. Essas tarefas diárias geralmente existem em ecossistemas exclusivos, com Excel de trabalho que são configuradas de maneiras específicas. Esses exemplos de cenário maiores demonstram esses casos de uso no mundo real. Eles incluem o Office scripts e as guias de trabalho, para que você possa ver o cenário de ponta a ponta.

| Cenário | Detalhes |
|---------|---------|
| [Analisar downloads da Web](../scenarios/analyze-web-downloads.md) | Esse cenário apresenta um script que analisado registros de tráfego da Web para determinar o país de origem de um usuário. Ele mostra as habilidades de análise de texto, usando subfunções em scripts, aplicando formatação condicional e trabalhando com tabelas. |
| [Buscar e representar graficamente os dados do nível de água do NOAA](../scenarios/noaa-data-fetch.md) | Esse cenário usa um script de Office para obter dados de uma fonte externa (o banco de dados [do NoAA Tides e Currents)](https://tidesandcurrents.noaa.gov/)e fazer um gráfico das informações resultantes. Ele realça as habilidades de uso `fetch` para obter dados e usar gráficos. |
| [Calculadora de notas](../scenarios/grade-calculator.md) | Esse cenário apresenta um script que valida o registro de um instrutor para as notas da classe. Ele mostra as habilidades de verificação de erros, formatação de células e expressões regulares. |
| [Lembretes de tarefas](../scenarios/task-reminders.md) | Esse cenário usa um script Office em um fluxo Power Automate para enviar lembretes aos colegas de trabalho para atualizar o status de um projeto. Ele realça as habilidades de Power Automate integração e transferência de dados de e para scripts. |

## <a name="community-contributions"></a>Contribuições da comunidade

Saudamos [as contribuições](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nossa Office de Scripts! Sinta-se à vontade para criar uma solicitação pull para revisão.

| Project | Detalhes |
|---------|---------|
| [Animação de saudações de estações](community-seasons-greetings.md) | Este script foi contribuído [por Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) no espírito da estação de feriado! É um script divertido que mostra uma árvore de Natal cantoria Excel na Web usando Office Scripts. |

## <a name="try-it-out"></a>Experimente

Esses exemplos são de código aberto. Experimente você mesmo. Você precisará de uma conta de trabalho ou de estudante da Microsoft do trabalho ou da escola com uma licença para Microsoft 365 assinatura (E3 ou superior). Basta ir para https://office.com entrar em sua conta e começar.

## <a name="leave-a-comment"></a>Deixar um comentário

Sinta-se à vontade para deixar um comentário, fazer uma sugestão ou registrar um problema usando a seção **Comentários** na parte inferior da página de documentação do exemplo específico.
