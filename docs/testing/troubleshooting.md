---
title: Solucionar problemas de scripts do Office
description: Dicas e técnicas de depuração para Scripts do Office, bem como recursos de ajuda.
ms.date: 10/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4fe4a9b17d51d078403d1a46abed774d38eeaa80
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495464"
---
# <a name="troubleshoot-office-scripts"></a>Solucionar problemas de scripts do Office

À medida que desenvolve scripts do Office, você pode cometer erros. Não faz mal. Você tem as ferramentas para ajudar a encontrar os problemas e fazer com que seus scripts funcionem perfeitamente.

> [!NOTE]
> Para obter conselhos de solução de problemas específicos para Scripts do Office com o Power Automate, consulte Solucionar problemas de [scripts do Office em execução no Power Automate](power-automate-troubleshooting.md).

## <a name="types-of-errors"></a>Tipos de erros

Os erros de Scripts do Office se enquadram em uma das duas categorias:

* Erros ou avisos de tempo de compilação
* Erros de runtime

### <a name="compile-time-errors"></a>Erros de tempo de compilação

Erros e avisos em tempo de compilação são inicialmente mostrados no Editor de Códigos. Eles são mostrados pelos sublinhados vermelhos ondulados no editor. Eles também são exibidos na guia **Problemas** na parte inferior do painel de tarefas do Editor de Códigos. Selecionar o erro fornecerá mais detalhes sobre o problema e sugerirá soluções. Erros de tempo de compilação devem ser resolvidos antes de executar o script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Um erro do compilador mostrado no texto de foco do Editor de Códigos.":::

Você também pode ver sublinhados de aviso laranja e mensagens informativas cinza. Isso indica sugestões de desempenho ou outras possibilidades em que o script pode ter efeitos não intencionais. Esses avisos devem ser examinados de perto antes de descartá-los.

### <a name="runtime-errors"></a>Erros de runtime

Erros de runtime ocorrem devido a problemas lógicos no script. Isso pode ocorrer porque um objeto usado no script não está na pasta de trabalho, uma tabela é formatada de forma diferente do previsto ou alguma outra pequena discrepância entre os requisitos do script e a pasta de trabalho atual. O script a seguir gera um erro quando uma planilha chamada "TestSheet" não está presente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Mensagens do console

Os erros de tempo de compilação e de runtime exibem mensagens de erro no console quando um script é executado. Eles fornecem um número de linha onde o problema foi encontrado. Tenha em mente que a causa raiz de qualquer problema pode ser uma linha de código diferente da indicada no console.

A imagem a seguir mostra a saída do console para [o erro explícito `any`](../develop/typescript-restrictions.md) do compilador. Observe o texto `[5, 16]` no início da cadeia de caracteres de erro. Isso indica que o erro está na linha 5, começando no caractere 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="O console do Editor de Códigos exibindo uma mensagem de erro explícita 'any'.":::

A imagem a seguir mostra a saída do console para um erro de runtime. Aqui, o script tenta adicionar uma planilha com o nome de uma planilha existente. Novamente, observe a "Linha 2" que precede o erro para mostrar qual linha investigar.
:::image type="content" source="../images/runtime-error-console.png" alt-text="O console do Editor de Códigos exibindo um erro da chamada 'addWorksheet'.":::

## <a name="console-logs"></a>Logs do console

Imprima mensagens na tela com a instrução `console.log` . Esses logs podem mostrar o valor atual das variáveis ou quais caminhos de código estão sendo disparados. Para fazer isso, chame `console.log` qualquer objeto como um parâmetro. Normalmente, um `string` é o tipo mais fácil de ler no console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

As cadeias de caracteres `console.log` passadas são exibidas no console de log do Editor de Códigos, na parte inferior do painel de tarefas. Os logs são encontrados na **guia** Saída, embora a guia ganhe automaticamente o foco quando um log é gravado.

Os logs não afetam a pasta de trabalho.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Guia Automatizar não aparecendo ou Scripts do Office indisponíveis

As etapas a seguir devem ajudar a solucionar problemas relacionados à guia **Automatizar** que não aparecem Excel na Web.

1. [Verifique se sua licença do Microsoft 365 inclui Scripts do Office](../overview/excel.md#requirements).
1. [Verifique se o navegador tem suporte](platform-limits.md#browser-support).
1. [Verifique se os cookies de terceiros estão habilitados](platform-limits.md#third-party-cookies).
1. [Verifique se o administrador não desabilitou os Scripts do Office no Centro de administração do Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. Verifique se você não está conectado como um usuário externo ou convidado ao seu locatário.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

> [!NOTE]
> Há um problema conhecido que impede que scripts armazenados no SharePoint sempre apareçam na lista usada recentemente. Isso ocorre quando o administrador desativa os Serviços Web do Exchange (EWS). Seus scripts baseados no SharePoint ainda são acessíveis e utilizáveis por meio da caixa de diálogo de arquivo.

## <a name="help-resources"></a>Recursos de ajuda

[O Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores dispostos a ajudar com problemas de codificação. Muitas vezes, você poderá encontrar a solução para o problema por meio de uma pesquisa rápida do Stack Overflow. Caso contrário, faça sua pergunta e marque-a com a marca "office-scripts". Lembre-se de mencionar que você está criando um *Script* do Office, não um *Suplemento do* Office.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas nos Scripts do Office ](../develop/best-practices.md)
- [Limites de plataforma com scripts do Office](platform-limits.md)
- [Melhorar o desempenho dos scripts do Office](../develop/web-client-performance.md)
- [Solucionar problemas de scripts do Office em execução no PowerAutomate](power-automate-troubleshooting.md)
- [Desfazer os efeitos do Scripts do Office](undo.md)
