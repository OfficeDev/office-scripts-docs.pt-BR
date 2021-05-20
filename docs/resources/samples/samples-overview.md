---
title: Office Amostras de scripts
description: Disponível Office scripts amostras e cenários.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0ea9a8a8986681fca0e45784e2923c1d3b34576d
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545706"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Scripts amostras e cenários

Esta seção contém soluções de automação baseadas em [scripts Office](../../overview/excel.md) que ajudam os usuários finais a alcançar a automação de tarefas diárias. Ele contém cenários realistas que os usuários de negócios enfrentam e fornece soluções detalhadas, juntamente com links de vídeo instrutivos passo a passo.

Para cada um dos projetos em [Noções Básicas](#basics) e [Além do básico,](#beyond-the-basics)confira o código-fonte, vídeos passo a passo [**do YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)e muito mais.

Em [Cenários,](#scenarios)incluímos algumas amostras de cenários maiores que demonstram casos de uso no mundo real.

Também recebemos [contribuições da comunidade.](#community-contributions-and-fun-samples)

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Noções básicas

| Project | Detalhes |
|---------|---------|
| [Noções básicas de scripts](../excel-samples.md) | Essas amostras demonstram blocos fundamentais de construção para Office Scripts. |
| [Adicione comentários em Excel](add-excel-comments.md) | Esta amostra adiciona comentários a uma célula, incluindo @mentioning um colega. |
| [Adicionar imagens a uma pasta de trabalho](add-image-to-workbook.md) | Esta amostra adiciona uma imagem a uma pasta de trabalho e copia uma imagem através de folhas.|
| [Copie várias tabelas de Excel em uma única tabela](copy-tables-combine.md) | Esta amostra combina dados de várias tabelas Excel em uma única tabela que inclui todas as linhas. |

## <a name="beyond-the-basics"></a>Além do básico

Confira o projeto de ponta a ponta que automatiza cenários de amostra, juntamente com scripts completos, Excel de amostra usados e [vídeos (hospedados no YouTube)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Project | Detalhes |
|---------|---------|
| [Conte linhas em branco em uma folha específica ou em todas as folhas](count-blank-rows.md) | Esta amostra detecta se há alguma linha em branco nas folhas onde você antecipa que os dados estejam presentes e, em seguida, relatar a contagem de linhas em branco para uso em um fluxo de Power Automate. |
| [Gráfico de e-mail e imagens de tabela](email-images-chart-table.md) | Esta amostra usa Office Scripts e Power Automate ações para criar um gráfico e enviar esse gráfico como uma imagem por e-mail. |
| [Chamadas de busca externa](external-fetch-calls.md) | Esta amostra usa `fetch` para obter informações de GitHub para o script. |
| [Filtrar Excel tabela e obter alcance visível](filter-table-get-visible-range.md) | Esta amostra filtra uma Excel tabela e retorna o alcance visível como um objeto JSON. Este JSON poderia ser fornecido a um fluxo de Power Automate como parte de uma solução maior. |
| [Gerenciar o modo de cálculo em Excel](excel-calculation.md) | Esta amostra mostra como usar o modo de cálculo e calcular métodos em Excel na Web usando scripts Office. |
| [Mova linhas através das mesas](move-rows-across-tables.md) | Esta amostra mostra como mover linhas entre as tabelas salvando filtros e, em seguida, processando e reaplicando os filtros. |
| [Dados de Excel de saída como JSON](get-table-data.md) | Esta solução mostra como produzir dados de tabela Excel como JSON para usar em Power Automate. |
| [Remova hiperlinks de cada célula em uma planilha Excel](remove-hyperlinks-from-cells.md) | Esta amostra limpa todos os hiperlinks da planilha atual. |
| [Executar um script em todos os arquivos do Excel em uma pasta](automate-tasks-on-all-excel-files-in-folder.md) | Este projeto executa um conjunto de tarefas de automação em todos os arquivos situados em uma pasta em OneDrive for Business (também pode ser usado para uma pasta SharePoint). Ele realiza cálculos sobre os arquivos Excel, adiciona formatação e insere um comentário que @mentions um colega. |
| [Escrever um grande conjuntos de dados](write-large-dataset.md) | Esta amostra mostra como enviar uma grande gama como subranges menores. |

## <a name="scenarios"></a>Cenários

Office Scripts podem automatizar partes de sua rotina diária. Essas tarefas cotidianas geralmente existem em ecossistemas únicos, com Excel livros de trabalho que são configuradas de maneiras particulares. Essas amostras de cenários maiores demonstram tais casos de uso no mundo real. Eles incluem tanto o Office Scripts quanto os livros de trabalho, para que você possa ver o cenário de ponta a ponta.

| Cenário | Detalhes |
|---------|---------|
| [Analisar downloads da Web](../scenarios/analyze-web-downloads.md) | Este cenário apresenta um script que analisa registros de tráfego da Web para determinar o país de origem de um usuário. Ele mostra as habilidades de análise de texto, usando subfunções em scripts, aplicando formatação condicional e trabalhando com tabelas. |
| [Buscar e representar graficamente os dados do nível de água do NOAA](../scenarios/noaa-data-fetch.md) | Esse cenário usa um Script Office para extrair dados de uma fonte externa (o [banco de dados NOAA Tides and Currents)](https://tidesandcurrents.noaa.gov/)e gráfico das informações resultantes. Ele destaca as habilidades de usar `fetch` para obter dados e usar gráficos. |
| [Calculadora de notas](../scenarios/grade-calculator.md) | Este cenário apresenta um roteiro que valida o registro de um instrutor para as notas de suas aulas. Ele mostra as habilidades de verificação de erros, formatação celular e expressões regulares. |
| [Lembretes de tarefas](../scenarios/task-reminders.md) | Este cenário usa um Script Office em um fluxo de Power Automate para enviar lembretes aos colegas de trabalho para atualizar o status de um projeto. Destaca as habilidades de integração Power Automate e transferência de dados de e para scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contribuições e amostras divertidas

Damos boas-vindas [às contribuições](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) da nossa comunidade de scripts Office! Sinta-se livre para criar um pedido de tração para revisão.

| Project | Detalhes |
|---------|---------|
| [Jogo da Vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | O blog "Ready Player Zero", de Yutao Huang, no Excel Tech Community inclui um roteiro para o modelo [*The Game of Life, de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)John Conway. |
| [Animação de saudações de estações](community-seasons-greetings.md) | Este roteiro foi contribuído por [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) no espírito da temporada de férias! É um roteiro divertido que mostra uma árvore de Natal cantando em Excel na Web usando Office Scripts. |

## <a name="try-it-out"></a>Experimente

Estas amostras são de código aberto. Experimente você mesmo. Você precisará de uma conta de trabalho ou escola da Microsoft do trabalho ou da escola com uma licença para Microsoft 365 assinatura (E3 ou superior). Basta ir até https://office.com entrar na sua conta e começar.

## <a name="leave-a-comment"></a>Deixe um comentário

Sinta-se à vontade para deixar um comentário, fazer uma sugestão ou registrar um problema usando a seção **Feedback** na parte inferior da página de documentação da amostra específica.
