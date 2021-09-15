---
title: Adicionar imagens a uma pasta de trabalho
description: Saiba como usar Office Scripts para adicionar uma imagem a uma planilha e copiá-la entre planilhas.
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0c4b3446df8de280b6cb557e291504ceed5ee7f7
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326853"
---
# <a name="add-images-to-a-workbook"></a>Adicionar imagens a uma pasta de trabalho

Este exemplo mostra como trabalhar com imagens usando um script Office no Excel.

## <a name="scenario"></a>Cenário

As imagens ajudam com identidade visual, identidade visual e modelos. Eles ajudam a tornar uma workbook mais do que apenas uma tabela enorme.

O primeiro exemplo copia uma imagem de uma planilha para outra. Isso pode ser usado para colocar o logotipo da sua empresa na mesma posição em cada planilha.

O segundo exemplo copia uma imagem de uma URL. Isso pode ser usado para copiar fotos que um colega armazenou em uma pasta compartilhada para uma pasta de trabalho relacionada.

## <a name="sample-excel-file"></a>Exemplo Excel arquivo

Baixe <a href="add-images.xlsx">add-images.xlsx</a> para uma workbook pronta para uso. Adicione os scripts a seguir e experimente o exemplo você mesmo!

## <a name="sample-code-copy-an-image-across-worksheets"></a>Código de exemplo: copiar uma imagem entre planilhas

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Código de exemplo: adicionar uma imagem de uma URL a uma workbook

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
