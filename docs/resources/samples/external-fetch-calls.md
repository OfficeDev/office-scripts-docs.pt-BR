---
title: Usar chamadas de busca externa em Scripts do Office
description: Saiba como fazer chamadas de API externas Office Scripts.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 569d74f1ca8996cd8fe8a4ba3163445d57676d27
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088089"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar chamadas de busca externa em Scripts do Office

Esse script obtém informações básicas sobre repositórios de GitHub do usuário. Ele mostra como usar em `fetch` um cenário simples. Para obter mais informações sobre como usar `fetch` ou outras chamadas externas, leia o suporte à chamada à [API externa Office Scripts](../../develop/external-calls.md). Para obter informações sobre como trabalhar com objetos [JSON](https://www.w3schools.com/whatis/whatis_json.asp) como o que é retornado pelas APIs do GitHub, leia Usar JSON para passar dados de e para [Office Scripts](../../develop/use-json.md).

Saiba mais sobre as APIs do GItHub que estão sendo usadas na referência [GitHub API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Você também pode ver a saída de chamada de API `https://api.github.com/users/{USERNAME}/repos` bruta visitando um navegador da Web (substitua o espaço reservado {USERNAME} pela sua ID GitHub configuração).

![Obter exemplo de informações de repositórios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de exemplo: obter informações básicas sobre repositórios GitHub usuário

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();

  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
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

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de treinamento: Como fazer chamadas à API externas

[Veja Sudhi Ramamurthy percorrer este exemplo no YouTube](https://youtu.be/fulP29J418E).
