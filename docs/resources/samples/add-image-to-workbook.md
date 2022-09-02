---
title: Adicionar imagens a uma pasta de trabalho
description: Saiba como usar scripts do Office para adicionar uma imagem a uma pasta de trabalho e copiá-la entre planilhas.
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78c7779cf4d524ed62bf8d419135863228b23d33
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572602"
---
# <a name="add-images-to-a-workbook"></a>Adicionar imagens a uma pasta de trabalho

Este exemplo mostra como trabalhar com imagens usando um Script do Office no Excel.

## <a name="scenario"></a>Cenário

As imagens ajudam com identidade visual, identidade visual e modelos. Eles ajudam a fazer uma pasta de trabalho mais do que apenas uma mesa gigante.

O primeiro exemplo copia uma imagem de uma planilha para outra. Isso pode ser usado para colocar o logotipo da sua empresa na mesma posição em cada planilha.

O segundo exemplo copia uma imagem de uma URL. Isso pode ser usado para copiar fotos que um colega armazenou em uma pasta compartilhada para uma pasta de trabalho relacionada.

## <a name="sample-excel-file"></a>Arquivo de exemplo do Excel

Baixe [add-images.xlsx](add-images.xlsx) para uma pasta de trabalho pronta para uso. Adicione os scripts a seguir e experimente o exemplo por conta própria!

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Código de exemplo: adicionar uma imagem de uma URL a uma pasta de trabalho

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
