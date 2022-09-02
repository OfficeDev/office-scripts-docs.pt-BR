---
title: Registrar alterações diárias no Excel e reportá-las com um fluxo do Power Automate
description: Saiba como usar scripts do Office e o Power Automate para controlar alterações de valor em uma pasta de trabalho
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572634"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>Registrar alterações diárias no Excel e reportá-las com um fluxo do Power Automate

Os Scripts do Power Automate e do Office são combinados para lidar com tarefas repetitivas para você. Neste exemplo, você tem a tarefa de gravar uma única leitura numérica em uma pasta de trabalho todos os dias e relatar a alteração desde ontem. Você criará um fluxo para obter essa leitura, registrá-la na pasta de trabalho e relatar a alteração por meio de um email.

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe [daily-readings.xlsx](daily-readings.xlsx) para uma pasta de trabalho pronta para uso. Adicione o script a seguir para experimentar o exemplo por conta própria!

## <a name="sample-code-record-and-report-daily-readings"></a>Código de exemplo: registrar e relatar leituras diárias

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>Fluxo de exemplo: relatar alterações diárias

Siga estas etapas para criar um [fluxo do Power Automate](https://powerautomate.microsoft.com/) para o exemplo.

1. Crie um fluxo **de nuvem agendado**.
1. Agende o fluxo para repetir a **cada 1 dia**.

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="A etapa de criação de fluxo mostrando que ela será repetida todos os dias.":::
1. Selecionar **Criar**.
1. Em um fluxo real, você adicionará uma etapa que obtém seus dados. Os dados podem vir de outra pasta de trabalho, um cartão adaptável do Teams ou qualquer outra fonte. Para testar o exemplo, faça um número de teste. Adicione uma nova etapa com a **ação Inicializar variável** . Forneça os valores a seguir.
    1. **Nome**: Entrada
    1. **Tipo**: Inteiro
    1. **Valor**: 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="A ação Inicializar variável com os valores fornecidos.":::
1. Adicione uma nova etapa com o conector **do Excel Online (Business)** com a **ação Executar script** . Use os valores a seguir para a ação.
    1. **Localização**: OneDrive for Business
    1. **Biblioteca de Documentos**: OneDrive
    1. **Arquivo**: daily-readings.xlsx *(escolhido por meio do navegador de arquivos)*
    1. **Script**: seu nome de script
    1. **newData**: Entrada *(conteúdo dinâmico)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="A ação Executar script com os valores fornecidos.":::
1. O script retorna a diferença de leitura diária como conteúdo dinâmico chamado "resultado". Para o exemplo, você pode enviar por email as informações para si mesmo. Crie uma nova etapa que use o **conector do Outlook** com a ação Enviar **um email (V2) (** ou qualquer cliente de email que você preferir). Use os valores a seguir para concluir a ação.
    1. **Para**: seu endereço de email
    1. **Assunto**: Alteração de leitura diária
    1. **Corpo**: resultado "Diferença de ontem" *(conteúdo dinâmico do Excel)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="O conector completo do Outlook no Power Automate.":::
1. Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo. Certifique-se de permitir o acesso quando solicitado.
