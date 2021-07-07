---
title: Adicionar comentários em Excel
description: Saiba como usar Office scripts para adicionar comentários em uma planilha.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 77e308d020281c71751e2652f8dbaec00c263e44
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313908"
---
# <a name="add-comments-in-excel"></a>Adicionar comentários em Excel

Este exemplo mostra como adicionar comentários a uma célula, [incluindo @mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) um colega.

## <a name="example-scenario"></a>Cenário de exemplo

* O líder da equipe mantém o cronograma de turnos. O líder da equipe atribui uma ID de funcionário ao registro de turno.
* O líder da equipe deseja notificar o funcionário. Adicionando um comentário que @mentions o funcionário, o funcionário é enviado por email com uma mensagem personalizada da planilha.
* Posteriormente, o funcionário pode exibir a guia de trabalho e responder ao comentário por conveniência.

## <a name="solution"></a>Solução

1. O script extrai informações dos funcionários da planilha do funcionário.
1. Em seguida, o script adiciona um comentário (incluindo o email de funcionário relevante) à célula apropriada no registro de turno.
1. Os comentários existentes na célula são removidos antes de adicionar o novo comentário.

## <a name="sample-excel-file"></a>Exemplo Excel arquivo

Baixe <a href="excel-comments.xlsx">excel-comments.xlsx</a> para uma workbook pronta para uso. Adicione o seguinte script para experimentar o exemplo você mesmo!

## <a name="sample-code-add-comments"></a>Código de exemplo: Adicionar comentários

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a>Vídeo de treinamento: Adicionar comentários

[Assista a Sudhi Ramamurthy passar por este exemplo no YouTube](https://youtu.be/CpR78nkaOFw).
