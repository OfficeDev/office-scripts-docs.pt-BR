---
title: Exemplos de Scripts do Office
description: Exemplos e cenários disponíveis do Office Scripts.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de0e99cbac7fcdeb1a3d3c43dd72ce53ed5847dd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571083"
---
# <a name="office-scripts-samples-and-scenarios"></a>Exemplos e cenários de Scripts do Office

Esta seção contém [soluções de automação baseadas](../../overview/excel.md) em Scripts do Office que ajudam os usuários finais a alcançar a automação de tarefas diárias. Ele contém cenários realistas que os usuários de negócios enfrentam e fornece soluções detalhadas juntamente com links de vídeo instrucional passo a passo.

Para cada um dos projetos em [Noções Básicas](#basics) e Além das noções [básicas,](#beyond-the-basics)confira o código-fonte, os vídeos passo a passo do [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)e muito mais.

Em [Cenários,](#scenarios)incluímos alguns exemplos de cenários maiores que demonstram casos de uso no mundo real.

Também saudamos [as contribuições da comunidade.](#community-contributions)

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Noções básicas

| Projeto | Detalhes |
|---------|---------|
| [Noções básicas de script](../excel-samples.md) | Esses exemplos demonstram blocos de construção fundamentais para scripts do Office. |
| [Saiba noções básicas sobre como usar o objeto Range em Scripts do Office](range-basics.md) | Este artigo mostra as noções básicas de uso do objeto Range e suas APIs. Este é um tópico base que será usado em todos os outros projetos. |

## <a name="beyond-the-basics"></a>Além das noções básicas

Confira o seguinte projeto de ponta a ponta que automatiza cenários de exemplo juntamente com scripts completos, arquivos do Excel de exemplo usados e [vídeos](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Projeto | Detalhes |
|---------|---------|
| [Adicionar comentários no Excel](add-excel-comments.md) | Este exemplo mostra como adicionar comentários a uma célula, incluindo @mentioning um colega. |
| [Contar linhas em branco em uma planilha específica ou em todas as planilhas](count-blank-rows.md) | Este exemplo detecta se há linhas em branco em planilhas nas quais você antecipa a presença dos dados e relata a contagem de linhas em branco para uso em um fluxo do Power Automate. |
| [Referência cruzada e formatar um arquivo do Excel](excel-cross-reference.md) | Esta solução mostra como dois arquivos do Excel podem ser cruzados e formatados usando Scripts do Office e Power Automate. |
| [Imagens de gráfico de email e tabela](email-images-chart-table.md) | Este exemplo usa scripts do Office e ações do Power Automate para criar um gráfico e enviar esse gráfico como uma imagem por email. |
| [Filtrar tabela do Excel e obter intervalo visível](filter-table-get-visible-range.md) | Este exemplo filtra uma tabela do Excel e retorna o intervalo visível como um objeto JSON. Esse JSON pode ser fornecido a um fluxo do Power Automate como parte de uma solução maior. |
| [Gerar um identificador exclusivo em uma workbook](document-number-generator.md) | Esse cenário ajuda um usuário a gerar um número de documento exclusivo com um formato específico e a adicionar uma entrada a um intervalo ou tabela. |
| [Gerenciar o modo de cálculo no Excel](excel-calculation.md) | Este exemplo mostra como usar o modo de cálculo e calcular métodos no Excel na Web usando Scripts do Office. |
| [Mesclar várias tabelas do Excel em uma única tabela](copy-tables-combine.md) | Este exemplo combina dados de várias tabelas do Excel em uma única tabela que inclui todas as linhas. |
| [Mover linhas entre tabelas](move-rows-across-tables.md) | Este exemplo mostra como mover linhas entre tabelas salvando filtros e, em seguida, processamento e reaplicação dos filtros. |
| [Dados de saída do Excel como JSON](get-table-data.md) | Esta solução mostra como saída de dados de tabela do Excel como JSON a ser usado no Power Automate. |
| [Remover hiperlinks de cada célula em uma planilha do Excel](remove-hyperlinks-from-cells.md) | Este exemplo limpa todos os hiperlinks da planilha atual. |
| [Executar um script em todos os arquivos do Excel em uma pasta](automate-tasks-on-all-excel-files-in-folder.md) | Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta no OneDrive for Business (também pode ser usado para uma pasta do SharePoint). Ele executa cálculos nos arquivos do Excel, adiciona formatação e insere um comentário que @mentions um colega. |
| [Enviar uma reunião do Teams a partir de dados do Excel](send-teams-invite-from-excel-data.md) | Esta solução mostra como usar scripts do Office e ações do Power Automate para selecionar linhas do arquivo do Excel e usá-la para enviar um convite de reunião do Teams e atualizar o Excel. |

## <a name="scenarios"></a>Cenários

Os Scripts do Office podem automatizar partes da sua rotina diária. Essas tarefas diárias geralmente existem em ecossistemas exclusivos, com as planilhas do Excel que são configuradas de maneiras específicas. Esses exemplos de cenário maiores demonstram esses casos de uso no mundo real. Eles incluem os Scripts do Office e as workbooks, para que você possa ver o cenário de ponta a ponta.

| Cenário | Detalhes |
|---------|---------|
| [Analisar downloads da Web](../scenarios/analyze-web-downloads.md) | Esse cenário apresenta um script que analisado registros de tráfego da Web para determinar o país de origem de um usuário. Ele mostra as habilidades de análise de texto, usando subfunções em scripts, aplicando formatação condicional e trabalhando com tabelas. |
| [Buscar e representar graficamente os dados do nível de água do NOAA](../scenarios/noaa-data-fetch.md) | Esse cenário usa um Script do Office para obter dados de uma fonte externa (o banco de dados [DesaA E Currents](https://tidesandcurrents.noaa.gov/)) e fazer um gráfico das informações resultantes. Ele realça as habilidades de uso `fetch` para obter dados e usar gráficos. |
| [Calculadora de notas](../scenarios/grade-calculator.md) | Esse cenário apresenta um script que valida o registro de um instrutor para as notas da classe. Ele mostra as habilidades de verificação de erros, formatação de células e expressões regulares. |
| [Lembretes de tarefas](../scenarios/task-reminders.md) | Esse cenário usa um Script do Office em um fluxo do Power Automate para enviar lembretes aos colegas de trabalho para atualizar o status de um projeto. Ele destaca as habilidades da integração do Power Automate e a transferência de dados de scripts de e para. |

## <a name="community-contributions"></a>Contribuições da comunidade

Saudamos [as contribuições](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nossa comunidade de Scripts do Office! Sinta-se à vontade para criar uma solicitação pull para revisão.

| Projeto | Detalhes |
|---------|---------|
| [Animação de saudações de estações](community-seasons-greetings.md) | Este script foi contribuído [por Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) no espírito da estação de feriado! É um script divertido que mostra uma árvore de Natal cantoria no Excel na Web usando Scripts do Office. |

## <a name="try-it-out"></a>Experimente

Esses exemplos são de código aberto. Experimente você mesmo. Você precisará de uma conta de estudante ou de trabalho da Microsoft do trabalho ou da escola com uma licença para a assinatura do Microsoft 365 (E3 ou superior). Basta ir para https://office.com entrar em sua conta e começar.

## <a name="leave-a-comment"></a>Deixar um comentário

Sinta-se à vontade para deixar um comentário, fazer uma sugestão ou registrar um problema usando a seção **Comentários** na parte inferior da página de documentação do exemplo específico.
