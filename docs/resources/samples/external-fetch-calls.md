---
title: Usar chamadas de busca externa em Scripts do Office
description: Saiba como fazer chamadas de API externas Office Scripts.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 721bfa39eea1e9973efc7fd13efa5bac734b76dd
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232519"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar chamadas de busca externa em Scripts do Office

Este script obtém informações básicas sobre os repositórios de GitHub do usuário. Ele mostra como usar `fetch` em um cenário simples.

Você pode saber mais sobre as APIs GItHub que estão sendo usadas na [referência GitHub API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Você também pode ver a saída de chamada de API bruta visitando um navegador da Web (certifique-se de substituir o espaço reservado `https://api.github.com/users/{USERNAME}/repos` {USERNAME} pela ID do Github).

![Obter exemplo de informações de repositórios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de exemplo: obter informações básicas sobre os repositórios GitHub do usuário

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de treinamento: como fazer chamadas de API externas

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/fulP29J418E).
