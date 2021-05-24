---
title: Solucionar Office scripts
description: Dicas e técnicas de depuração para Office scripts, bem como recursos de ajuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545550"
---
# <a name="troubleshoot-office-scripts"></a>Solucionar Office scripts

À medida que você Office scripts, você pode cometer erros. Não há problema. Você tem as ferramentas para ajudar a encontrar os problemas e fazer seus scripts funcionarem perfeitamente.

## <a name="types-of-errors"></a>Tipos de erros

Office Os erros de scripts se enquadram em uma das duas categorias:

* Erros ou avisos em tempo de compilação
* Erros de tempo de execução

### <a name="compile-time-errors"></a>Erros em tempo de compilação

Erros e avisos de tempo de compilação são mostrados inicialmente no Editor de Código. Eles são mostrados pelos sublinhados vermelho ondulados no editor. Eles também são exibidos na guia **Problemas** na parte inferior do painel de tarefas Editor de Código. Selecionar o erro dará mais detalhes sobre o problema e sugerirá soluções. Erros em tempo de compilação devem ser resolvidos antes de executar o script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Um erro de compilador mostrado no texto de foco do Editor de Código":::

Você também pode ver sublinhados de aviso laranja e mensagens informativas cinzas. Elas indicam sugestões de desempenho ou outras possibilidades em que o script pode ter efeitos não intencional. Esses avisos devem ser examinados de perto antes de descartá-los.

### <a name="runtime-errors"></a>Erros de tempo de execução

Erros de tempo de execução ocorrem devido a problemas de lógica no script. Isso pode ser porque um objeto usado no script não está na guia de trabalho, uma tabela é formatada de forma diferente do previsto ou alguma outra pequena discrepância entre os requisitos do script e a atual. O script a seguir gera um erro quando uma planilha chamada "TestSheet" não está presente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Mensagens de console

Erros de tempo de compilação e tempo de execução exibem mensagens de erro no console quando um script é executado. Eles dão um número de linha onde o problema foi encontrado. Lembre-se de que a causa raiz de qualquer problema pode ser uma linha de código diferente da indicada no console.

A imagem a seguir mostra a saída do console para [o erro explícito `any` ](../develop/typescript-restrictions.md) do compilador. Observe o texto `[5, 16]` no início da cadeia de caracteres de erro. Isso indica que o erro está na linha 5, começando no caractere 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O console do Editor de Código exibindo uma mensagem de erro explícita &quot;qualquer&quot;":::

A imagem a seguir mostra a saída do console para um erro de tempo de execução. Aqui, o script tenta adicionar uma planilha com o nome de uma planilha existente. Novamente, observe a "Linha 2" anterior ao erro para mostrar qual linha investigar.
:::image type="content" source="../images/runtime-error-console.png" alt-text="O console do Editor de Código exibindo um erro da chamada 'addWorksheet'":::

## <a name="console-logs"></a>Logs de console

Imprimir mensagens na tela com a `console.log` instrução. Esses logs podem mostrar o valor atual das variáveis ou quais caminhos de código estão sendo disparados. Para fazer isso, chame `console.log` qualquer objeto como parâmetro. Normalmente, um `string` é o tipo mais fácil de ler no console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

As cadeias de caracteres passadas para são exibidas no console de registro em log do Editor de Código, na `console.log` parte inferior do painel de tarefas. Os logs são encontrados na guia **Saída,** embora a guia automaticamente obtém o foco quando um log é gravado.

Os logs não afetam a agenda de trabalho.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Guia Automatizar não aparecendo ou Office Scripts indisponíveis

As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **Automatizar** que não aparecem no Excel na Web.

1. [Certifique-se de Microsoft 365 sua licença de Office Scripts](../overview/excel.md#requirements).
1. [Verifique se o navegador tem suporte](platform-limits.md#browser-support).
1. [Verifique se os cookies de terceiros estão habilitados.](platform-limits.md#third-party-cookies)
1. [Verifique se o administrador não desabilitou Office scripts no Microsoft 365 de administração.](/microsoft-365/admin/manage/manage-office-scripts-settings)

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Solucionar problemas de scripts em Power Automate

Para obter informações específicas sobre como executar scripts Power Automate, consulte [Troubleshoot Office Scripts em execução em Power Automate](power-automate-troubleshooting.md).

## <a name="help-resources"></a>Recursos de ajuda

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores dispostos a ajudar com problemas de codificação. Muitas vezes, você poderá encontrar a solução para seu problema por meio de uma pesquisa rápida de Estouro de Pilha. Se não, faça sua pergunta e marque-a com a marca "office-scripts". Certifique-se de mencionar que você está criando um *script* de Office , não um Office *Desem.*

Se você encontrar um problema com a API JavaScript Office, crie um problema no repositório [officeDev/office-js GitHub](https://github.com/OfficeDev/office-js) repositório. Os membros da equipe do produto responderão a problemas e fornecerão mais assistência. A criação de um problema no repositório **officeDev/office-js** indica que você encontrou uma falha na biblioteca de API JavaScript Office que a equipe do produto deve resolver.

Se houver um problema com o Gravador de Ações ou Editor, envie comentários por meio do botão **Ajuda > Comentários** no Excel.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas nos Scripts do Office ](../develop/best-practices.md)
- [Limites da plataforma com Office Scripts](platform-limits.md)
- [Melhorar o desempenho de seus Office Scripts](../develop/web-client-performance.md)
- [Solucionar Office scripts em execução no PowerAutomate](power-automate-troubleshooting.md)
- [Desfazer os efeitos do Scripts do Office](undo.md)
