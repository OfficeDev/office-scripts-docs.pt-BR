---
title: Usar chamadas de busca externa em Scripts do Office
description: Saiba como fazer chamadas de API externas Office Scripts.
ms.date: 05/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: d957e0536e8574681f2ec752f23f9e6ba07f5fd2
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59335745"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar chamadas de busca externa em Scripts do Office

Este script obtém informações básicas sobre os repositórios de GitHub do usuário. Ele mostra como usar `fetch` em um cenário simples. Para obter mais informações sobre como `fetch` usar ou outras chamadas [externas,](../../develop/external-calls.md) leia Suporte a chamadas de API externa em Office Scripts

Você pode saber mais sobre as APIs GItHub que estão sendo usadas na [referência GitHub API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Você também pode ver a saída de chamada de API bruta visitando um navegador da Web (certifique-se de substituir o espaço reservado {USERNAME} pelo seu `https://api.github.com/users/{USERNAME}/repos` ID GitHub de usuário).

![Obter exemplo de informações de repositórios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de exemplo: obter informações básicas sobre os repositórios GitHub do usuário

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

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de treinamento: como fazer chamadas de API externas

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/fulP29J418E).
