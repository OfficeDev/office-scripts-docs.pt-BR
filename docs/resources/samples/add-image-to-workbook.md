---
title: Adicionar imagens a uma pasta de trabalho
description: Aprenda a usar Office Scripts para adicionar uma imagem a uma pasta de trabalho e copiá-la através de folhas.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 99c3cc2cacf6e535bdb882bb8414d23fd105be35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546022"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="c7f52-103">Adicionar imagens a uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="c7f52-103">Add images to a workbook</span></span>

<span data-ttu-id="c7f52-104">Esta amostra mostra como trabalhar com imagens usando um script Office em Excel.</span><span class="sxs-lookup"><span data-stu-id="c7f52-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="c7f52-105">Cenário</span><span class="sxs-lookup"><span data-stu-id="c7f52-105">Scenario</span></span>

<span data-ttu-id="c7f52-106">As imagens ajudam com marcas, identidade visual e modelos.</span><span class="sxs-lookup"><span data-stu-id="c7f52-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="c7f52-107">Eles ajudam a fazer um livro de trabalho mais do que apenas uma mesa gigante.</span><span class="sxs-lookup"><span data-stu-id="c7f52-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="c7f52-108">A primeira amostra copia uma imagem de uma planilha para outra.</span><span class="sxs-lookup"><span data-stu-id="c7f52-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="c7f52-109">Isso pode ser usado para colocar o logotipo da sua empresa na mesma posição em cada folha.</span><span class="sxs-lookup"><span data-stu-id="c7f52-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="c7f52-110">A segunda amostra copia uma imagem de uma URL.</span><span class="sxs-lookup"><span data-stu-id="c7f52-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="c7f52-111">Isso pode ser usado para copiar fotos que um colega armazenava em uma pasta compartilhada para uma pasta de trabalho relacionada.</span><span class="sxs-lookup"><span data-stu-id="c7f52-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="c7f52-112">Arquivo de Excel de amostra</span><span class="sxs-lookup"><span data-stu-id="c7f52-112">Sample Excel file</span></span>

<span data-ttu-id="c7f52-113">Baixe o arquivo <a href="add-images.xlsx">add-images.xlsx</a> usado nessas amostras e experimente você mesmo!</span><span class="sxs-lookup"><span data-stu-id="c7f52-113">Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="c7f52-114">Código de amostra: Copie uma imagem através de planilhas</span><span class="sxs-lookup"><span data-stu-id="c7f52-114">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="c7f52-115">Código de amostra: Adicione uma imagem de uma URL a uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="c7f52-115">Sample code: Add an image from a URL to a workbook</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```