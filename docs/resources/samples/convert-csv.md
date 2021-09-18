---
title: Converter arquivos CSV em Excel pastas de trabalho
description: Saiba como usar Office scripts e Power Automate para criar .xlsx arquivos .csv arquivos.
ms.date: 07/19/2021
ms.localizationpriority: medium
ms.openlocfilehash: 213c6caab1d1b20d566aa0e79630c1a9b50554f7
ms.sourcegitcommit: 5ec904cbb1f2cc00a301a5ba7ccb8ae303341267
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/18/2021
ms.locfileid: "59447475"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Converter arquivos CSV em Excel pastas de trabalho

Muitos serviços exportam dados como arquivos CSV (valores separados por vírgula). Essa solução automatiza o processo de conversão desses arquivos CSV em Excel pastas de trabalho no formato .xlsx arquivo. Ele usa um [fluxo Power Automate](https://flow.microsoft.com) para encontrar arquivos com a extensão .csv em uma pasta OneDrive e um Script Office para copiar os dados do arquivo .csv para uma nova pasta de trabalho Excel.

## <a name="solution"></a>Solução

1. Armazene os .csv e um arquivo "Template" em .xlsx em uma pasta OneDrive em branco.
1. Crie um Office Script para analisar os dados CSV em um intervalo.
1. Crie um Power Automate fluxo para ler os arquivos .csv e passar seu conteúdo para o script.

## <a name="sample-files"></a>Exemplo de arquivos

Baixe <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> para obter o arquivo Template.xlsx e dois arquivos .csv exemplo. Extraia os arquivos em uma pasta em sua OneDrive. Este exemplo supõe que a pasta é chamada de "saída".

Adicione o script a seguir e crie um fluxo usando as etapas fornecidas para experimentar o exemplo você mesmo!

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Código de exemplo: inserir valores separados por vírgulas em uma workbook

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate fluxo: criar novos .xlsx arquivos

1. Entre [Power Automate](https://flow.microsoft.com) e crie um novo **fluxo de nuvem agendado.**
1. De definir o fluxo como **Repetir a cada** "1" "Dia" e selecione **Criar**.
1. Obter o arquivo Excel modelo. Essa é a base para todos os arquivos .csv convertidos. Adicione uma **nova etapa que** usa o conector **OneDrive for Business** e a ação Obter **conteúdo de** arquivo. Forneça o caminho do arquivo para o arquivo "Template.xlsx".
    * **Arquivo**: /output/Template.xlsx
1. Renomeie a etapa **Obter** conteúdo de arquivo indo para o menu Obter conteúdo de **arquivo(...)** dessa etapa (no canto superior direito do conector) e selecionando a opção **Renomear.** Altere o nome da etapa para "Obter Excel modelo".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="O conector OneDrive for Business no Power Automate, renomeado para Get Excel template.":::
1. Obter todos os arquivos na pasta "saída". Adicione uma **nova etapa que** usa o conector **OneDrive for Business** e os arquivos list na **ação de** pasta. Forneça o caminho da pasta que contém os .csv arquivos.
    * **Pasta**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="O conector OneDrive for Business no Power Automate.":::
1. Adicione uma condição para que o fluxo só opere .csv arquivos. Adicione uma **nova etapa** que é o **controle Condição.** Use os seguintes valores para a **condição**.
    * **Escolha um valor:** *Nome* (conteúdo dinâmico de **Arquivos de lista na pasta**). Observe que esse conteúdo dinâmico tem vários resultados, portanto, um Aplicar a **cada** controle *de* valor envolve a **Condição**.
    * **termina com** (na lista de menus suspensos)
    * **Escolha um valor**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="O controle Condição concluída com o controle Aplicar a cada controle ao seu redor.":::
1. O restante do fluxo está na seção **Se sim,** já que só queremos agir .csv arquivos. Obter um arquivo .csv individual adicionando uma **nova** etapa que usa o conector OneDrive for Business **e** a **ação Obter conteúdo de** arquivo. Use a **ID do** conteúdo dinâmico de **Arquivos de lista na pasta**.
    * **Arquivo**: *ID* (conteúdo dinâmico dos arquivos **de lista na etapa de** pasta)
1. Renomeie a nova **etapa Obter conteúdo de** arquivo como "Obter .csv arquivo". Isso ajuda a distinguir esse arquivo do Excel modelo.
1. Faça o novo arquivo .xlsx, usando o modelo Excel como o conteúdo base. Adicione uma **nova etapa** que usa o **conector OneDrive for Business** e a ação **Criar arquivo.** Use os seguintes valores.
    * **Caminho da pasta**: /output
    * **Nome do** arquivo : *nome* sem .xlsx (escolha o conteúdo dinâmico Nome sem extensão dos arquivos **de** lista na pasta e digite manualmente ".xlsx" depois dele)
    * **Conteúdo do arquivo**: *conteúdo do arquivo* (conteúdo dinâmico de Obter Excel **modelo**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="O arquivo Get .csv e Criar etapas de arquivo do Power Automate fluxo.":::
1. Execute o script para copiar dados para a nova workbook. Adicione o **conector Excel Online (Business)** com a **ação Executar script.** Use os seguintes valores para a ação.
    * **Localização**: OneDrive for Business
    * **Biblioteca de Documentos**: OneDrive
    * **Arquivo**: *ID* (conteúdo dinâmico de **Criar arquivo**)
    * **Script**: Converter CSV
    * **csv**: *Conteúdo de arquivo* (conteúdo dinâmico de Obter .csv **arquivo**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="O conector Excel online (Business) concluído Power Automate.":::
1. Salve o fluxo. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.
1. Você deve encontrar novos arquivos .xlsx na pasta "output", juntamente com os arquivos .csv originais. As novas pastas de trabalho contêm os mesmos dados dos arquivos CSV.
