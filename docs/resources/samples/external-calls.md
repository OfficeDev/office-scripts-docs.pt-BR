---
title: Fazer chamadas de API externas em Scripts do Office
description: Saiba como fazer chamadas de API externas em Scripts do Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: d0abfa0bb1adedc7535059ed359b8053d9f1c84d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571098"
---
# <a name="external-api-calls-from-office-scripts"></a>Chamadas de API externas de Scripts do Office

Scripts do Office permitem [suporte limitado a chamada de API externa](../../develop/external-calls.md).

> [!IMPORTANT]
>
> * Não há nenhuma maneira de entrar ou usar o tipo OAuth2 de fluxos de autenticação. Todas as chaves e credenciais devem ser codificadas (ou leitura de outra fonte).
> * Não há infraestrutura para armazenar credenciais e chaves da API. Isso terá que ser gerenciado pelo usuário.
> * Chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejáveis ou dados externos a serem trazidos para as guias de trabalho internas. O administrador pode estabelecer proteção de firewall contra essas chamadas. Verifique as políticas locais antes de confiar em chamadas externas.
> * Se um script usar uma chamada de API, ele não funcionará em um cenário do Power Automate. Você terá que usar a ação HTTP do Power Automate ou ações equivalentes para puxar dados ou pressioná-los para um serviço externo.
> * Uma chamada de API externa envolve uma sintaxe assíncrona da API e requer um conhecimento ligeiramente avançado da maneira como a comunicação assíncrona funciona.
> * Verifique a quantidade de transferência de dados antes de assumir uma dependência. Por exemplo, retirar todo o conjuntos de dados externos pode não ser a melhor opção e, em vez disso, a paginação deve ser usada para obter dados em partes.

## <a name="useful-knowledge-and-resources"></a>Conhecimento e recursos úteis

* [API REST](https://en.wikipedia.org/wiki/Representational_state_transfer): A maneira mais provável é usar a chamada da API.
* [ `async` : Entenda como isso funciona. `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Entenda como isso funciona.

## <a name="steps"></a>Etapas

1. Marque sua `main` função como uma função assíncrona adicionando `async` prefixo. Por exemplo, `async function main(workbook: ExcelScript.Workbook)`.
1. Qual tipo de chamada de API você está fazendo? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? Consulte o material da API REST para obter detalhes.
1. Obtenha o ponto de extremidade da API de serviço, requisitos de autenticação, headers, etc.
1. Defina a entrada ou saída `interface` para ajudar na conclusão do código e na verificação do tempo de desenvolvimento. Consulte [vídeo](#training-video-how-to-make-external-api-calls) para obter detalhes.
1. Código, teste, otimizar. Você pode criar uma função para sua rotina de chamada de API para torná-la reutilizável de outras partes do seu script ou para reutilização em um script diferente (copiar colar se torna muito mais fácil dessa maneira).

## <a name="scenario"></a>Cenário

Este script obtém informações básicas sobre repositórios do GitHub do usuário.

![Obter exemplo de informações de repositórios](../../images/git.png)

## <a name="resources-used-in-the-sample"></a>Recursos usados no exemplo

1. [Obter referência da API Github de repositórios.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. Saída de chamada da API: vá para um navegador da Web ou qualquer interface HTTP e digite , substituindo o espaço reservado `https://api.github.com/users/{USERNAME}/repos` {USERNAME} pela ID do Github.
1. Informações buscadas: repo.name, repo.size, repo.owner.id, repo.license?. name

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de exemplo: obter informações básicas sobre repositórios do GitHub do usuário

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

[![Assista a um vídeo sobre como fazer chamadas de API externas](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vídeo sobre como fazer chamadas de API externas")
