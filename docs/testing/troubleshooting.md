---
title: Solução de problemas Office Scripts
description: Depuração de dicas e técnicas para Office Scripts, bem como recursos de ajuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545550"
---
# <a name="troubleshoot-office-scripts"></a>Solução de problemas Office Scripts

À medida que você desenvolve Office Scripts, você pode cometer erros. Está tudo bem, está tudo bem. Você tem as ferramentas para ajudar a encontrar os problemas e fazer seus roteiros funcionarem perfeitamente.

## <a name="types-of-errors"></a>Tipos de erros

Office Os erros de scripts caem em uma das duas categorias:

* Compilar erros ou avisos em tempo de compilação
* Erros de tempo de execução

### <a name="compile-time-errors"></a>Erros de tempo de compilação

Erros e avisos de tempo de compilação são mostrados inicialmente no Editor de Código. Estes são mostrados pelos sublinhados vermelhos ondulados no editor. Eles também são exibidos na guia **Problemas** na parte inferior do painel de tarefas do Editor de Código. A seleção do erro dará mais detalhes sobre o problema e sugerirá soluções. Os erros de tempo de compilação devem ser resolvidos antes de executar o script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Um erro do compilador mostrado no texto do hover do Editor de código":::

Você também pode ver sublinhas de aviso laranja e mensagens informacionais cinzas. Isso indica sugestões de desempenho ou outras possibilidades onde o script pode ter efeitos não intencionais. Tais avisos devem ser examinados de perto antes de rejeií-los.

### <a name="runtime-errors"></a>Erros de tempo de execução

Erros de tempo de execução acontecem por causa de problemas lógicos no script. Isso pode ser porque um objeto usado no script não está na pasta de trabalho, uma tabela é formatada de forma diferente do previsto, ou alguma outra pequena discrepância entre os requisitos do script e a pasta de trabalho atual. O script a seguir gera um erro quando uma planilha chamada "TestSheet" não está presente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Mensagens de console

Ambos os erros de tempo de compilação e tempo de execução exibem mensagens de erro no console quando um script é executado. Eles dão um número de linha onde o problema foi encontrado. Tenha em mente que a causa raiz de qualquer problema pode ser uma linha de código diferente do indicado no console.

A imagem a seguir mostra a saída do console para o erro [explícito `any` ](../develop/typescript-restrictions.md) do compilador. Observe o texto `[5, 16]` no início da sequência de erros. Isso indica que o erro está na linha 5, começando pelo caractere 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O console Code Editor exibindo uma mensagem de erro explícita 'qualquer'":::

A imagem a seguir mostra a saída do console para um erro de tempo de execução. Aqui, o script tenta adicionar uma planilha com o nome de uma planilha existente. Novamente, observe a "Linha 2" que precede o erro para mostrar qual linha investigar.
:::image type="content" source="../images/runtime-error-console.png" alt-text="O console Code Editor exibindo um erro da chamada 'addWorksheet'":::

## <a name="console-logs"></a>Logs de console

Imprima mensagens na tela com a `console.log` instrução. Esses registros podem mostrar o valor atual das variáveis ou quais caminhos de código estão sendo acionados. Para fazer isso, chame `console.log` qualquer objeto como parâmetro. Normalmente, `string` um é o tipo mais fácil de ler no console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

As strings passadas `console.log` são exibidas no console de registro do Editor de Código, na parte inferior do painel de tarefas. Os registros são encontrados na guia **Saída,** embora a guia ganhe automaticamente o foco quando um registro é gravado.

Os registros não afetam a pasta de trabalho.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Automatize a guia não aparecendo ou Office Scripts indisponíveis

As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **Automate** que não aparece em Excel na Web.

1. [Certifique-se de que sua licença de Microsoft 365 inclui scripts Office](../overview/excel.md#requirements).
1. [Verifique se seu navegador está suportado](platform-limits.md#browser-support).
1. [Certifique-se de que cookies de terceiros estão ativados](platform-limits.md#third-party-cookies).
1. [Certifique-se de que o administrador não desabilitou Office Scripts no centro administrativo Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Solução de problemas de scripts em Power Automate

Para obter informações específicas para executar scripts através de Power Automate, consulte [Troubleshoot Office Scripts em execução em Power Automate](power-automate-troubleshooting.md).

## <a name="help-resources"></a>Recursos de ajuda

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores dispostos a ajudar com problemas de codificação. Muitas vezes, você será capaz de encontrar a solução para o seu problema através de uma pesquisa rápida stack overflow. Se não, faça sua pergunta e marque-a com a tag "office-scripts". Não deixe de mencionar que você está criando um *script* Office , não um *complemento* Office .

Se você encontrar um problema com a API javascript Office, crie um problema no repositório [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub. Os membros da equipe de produtos responderão às questões e prestarão assistência adicional. Criar um problema no repositório **OfficeDev/office-js** indica que você encontrou uma falha na biblioteca de API JavaScript Office que a equipe do produto deve abordar.

Se houver algum problema com o Gravador de Ação ou Editor, envie feedback através do botão **Ajuda > Feedback** em Excel.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas no Scripts do Office](../develop/best-practices.md)
- [Limites de plataforma com scripts Office](platform-limits.md)
- [Melhore o desempenho de seus scripts de Office](../develop/web-client-performance.md)
- [Solução de problemas Office Scripts em execução no PowerAutomate](power-automate-troubleshooting.md)
- [Desfazer os efeitos do Scripts do Office](undo.md)
