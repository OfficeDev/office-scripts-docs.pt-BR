---
title: Usar chamadas de busca externa em Scripts do Office
description: Aprenda a fazer chamadas de API externas em scripts Office.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545749"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar chamadas de busca externa em Scripts do Office

Este script obtém informações básicas sobre os repositórios de GitHub do usuário. Mostra como usar `fetch` em um cenário simples. Para obter mais informações sobre o uso `fetch` ou outras chamadas [externas, leia o suporte de chamadas de API externas em scripts Office](../../develop/external-calls.md)

Você pode saber mais sobre as APIs do GItHub que estão sendo usadas na [referência GitHub API](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user). Você também pode ver a saída bruta de chamada de API visitando `https://api.github.com/users/{USERNAME}/repos` em um navegador da Web (certifique-se de substituir o espaço reservado {USERNAME} pelo seu ID GitHub).

![Obtenha o exemplo de informações de repositórios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de amostra: Obtenha informações básicas sobre os repositórios de GitHub do usuário

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License 
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de treinamento: Como fazer chamadas externas de API

[Assista Sudhi Ramamurthy andar através desta amostra no YouTube](https://youtu.be/fulP29J418E).
