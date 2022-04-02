---
title: Combinar as guias de trabalho em uma única workbook
description: Saiba como usar Office scripts e Power Automate para criar planilhas de mesclagem de outras pasta de trabalho em uma única pasta de trabalho.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: f90980f2e2d1f125f4ca2ffb80822f13ecdeed0e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585881"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>Combinar planilhas em uma única pasta de trabalho

Este exemplo mostra como puxar dados de várias guias de trabalho para uma única e centralizada. Ele usa dois scripts: um para recuperar informações de uma pasta de trabalho e outro para criar novas planilhas com essas informações. Ele combina os scripts em um fluxo Power Automate que atua em uma pasta OneDrive inteira.

> [!IMPORTANT]
> Este exemplo copia apenas os valores das outras guias de trabalho. Ele não preserva formatação, gráficos, tabelas ou outros objetos.

## <a name="scenario"></a>Cenário

1. Crie um novo Excel em seu OneDrive e adicione dois scripts deste exemplo a ele.
1. Crie uma pasta em seu OneDrive e adicione uma ou mais pastas de trabalho com dados a ela.
1. Crie um fluxo para obter todos os arquivos dessa pasta.
1. Use o **script retornar dados da planilha** para obter os dados de cada planilha em cada uma das planilhas.
1. Use o **script Adicionar planilhas** para criar uma nova planilha em uma única pasta de trabalho para cada planilha em todos os outros arquivos.

## <a name="sample-code-return-worksheet-data"></a>Código de exemplo: Retornar dados da planilha

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>Código de exemplo: Adicionar planilhas

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automate fluxo: Combinar planilhas em uma única pasta de trabalho

1. Entre [Power Automate e](https://flow.microsoft.com) crie um novo fluxo **de nuvem instantâneo**.
1. Escolha **Disparar manualmente um fluxo e** selecione **Criar**.
1. Obter todos os arquivos na pasta. Neste exemplo, vamos usar uma pasta chamada "output". Adicione uma **nova etapa** que usa o **conector OneDrive for Business** e os **arquivos list na ação de** pasta. Forneça o caminho da pasta que contém os .csv arquivos.
    * **Pasta**: /output

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="O conector OneDrive for Business no Power Automate.":::
1. Execute o **script retornar dados da planilha** para obter todos os dados de cada uma das planilhas. Adicione o **conector Excel Online (Business)** com a **ação Executar script**. Use os seguintes valores para a ação. Observe que, quando você adicionar a *ID* do arquivo, Power Automate envolverá a ação em um **Aplicar** a cada controle, para que a ação seja executada em todos os arquivos.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: *ID* (conteúdo dinâmico de **arquivos de lista na pasta**)
    * **Script**: Retornar dados da planilha
1. Execute o **script Adicionar planilhas** no novo arquivo Excel que você criou. Isso adicionará os dados de todas as outras guias de trabalho. Após a ação **executar script** anterior e dentro da **ação Aplicar** a cada controle, adicione um conector Excel **Online (Business)** com a **ação Executar script**. Use os seguintes valores para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: seu arquivo
    * **Script**: Adicionar planilhas
    * **workbookName**: *Name* (conteúdo dinâmico de **Arquivos de lista na pasta**)
    * **worksheetInformation** (depois de selecionar o botão **Alternar** para inserir toda a matriz, consulte a observação a seguir à próxima imagem): *resultado* (conteúdo dinâmico do **script Executar**)

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="As duas ações de script Executar dentro do controle Aplicar a cada controle.":::
    > [!NOTE]
    > Selecione o **botão Alternar para inserir toda a matriz** para adicionar o objeto array diretamente, em vez de itens individuais para a matriz.
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="O botão a ser alternado para inserir uma matriz inteira em uma caixa de entrada do campo de controle.":::
1. Salve o fluxo. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da **guia Meus fluxos** . Certifique-se de permitir o acesso quando solicitado.
1. Seu Excel arquivo de Excel agora deve ter novas planilhas.
