---
title: Adicionar comentários no Excel
description: Saiba como usar scripts do Office para adicionar comentários em uma planilha.
ms.date: 03/29/2021
localization_priority: Normal
ms.openlocfilehash: aaaf26df6973bd081290b0fbb67edecad8627e53
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571158"
---
# <a name="add-comments-in-excel"></a>Adicionar comentários no Excel

Este exemplo mostra como adicionar comentários a uma célula, [incluindo @mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) um colega.

## <a name="example-scenario"></a>Cenário de exemplo

* O líder da equipe mantém o cronograma de turnos. O líder da equipe atribui uma ID de funcionário ao registro de turno.
* O líder da equipe deseja notificar o funcionário. Adicionando um comentário que @mentions o funcionário, o funcionário é enviado por email com uma mensagem personalizada da planilha.
* Posteriormente, o funcionário pode exibir a guia de trabalho e responder ao comentário por conveniência.

## <a name="solution"></a>Solução

1. O script extrai informações dos funcionários da planilha do funcionário.
1. Em seguida, o script adiciona um comentário (incluindo o email de funcionário relevante) à célula apropriada no registro de turno.
1. Os comentários existentes na célula são removidos antes de adicionar o novo comentário.

## <a name="sample-code-add-comments"></a>Código de exemplo: Adicionar comentários

Baixe o arquivo <a href="excel-comments.xlsx">excel-comments.xlsx</a> usado neste exemplo e experimente você mesmo!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
    console.log(employees); 

    const scheduleSheet = workbook.getWorksheet('Schedule');
    const table = scheduleSheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const scheduleData = range.getTexts();

    for (let i=0; i < scheduleData.length; i++) {
      let eId = scheduleData[i][3];

      let employeeInfo = employees.find(e => e[0] === eId);
      if (employeeInfo) {
        console.log("Found a match " + employeeInfo);
        let adminNotes = scheduleData[i][4];
        try { 
          let comment = workbook.getCommentByCell(range.getCell(i, 5));
          comment.delete();
        } catch {
            console.log("Ignore if there is no existing comment in the cell");
        }
        workbook.addComment(range.getCell(i,5), {
          mentions: [{
            email: employeeInfo[1],
            id: 0,
            name: employeeInfo[2]
          }],
          richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
        }, ExcelScript.ContentType.mention);        
        
      } else {
        console.log("No match for: " + eId);
      }
    }
    return;
}
```

## <a name="training-video-add-comments"></a>Vídeo de treinamento: Adicionar comentários

[![Assista a um vídeo passo a passo sobre como adicionar comentários em um arquivo do Excel](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Vídeo passo a passo sobre como adicionar comentários em um arquivo do Excel")
