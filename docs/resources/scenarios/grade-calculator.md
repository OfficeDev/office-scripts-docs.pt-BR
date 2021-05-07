---
title: 'Office Cenário de exemplo de scripts: calculadora de notas'
description: Um exemplo que determina a porcentagem e as notas de carta para uma classe de alunos.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: e2ef6e7522fc88219bf6ba40900a1ecceecb263b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232694"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office Cenário de exemplo de scripts: calculadora de notas

Nesse cenário, você é um instrutor que depende das notas de fim de semestre de cada aluno. Você está inserindo as pontuações para suas atribuições e testes à medida que vai. Agora, é hora de determinar os destinos dos alunos.

Você desenvolverá um script que totaliza as notas para cada categoria de ponto. Em seguida, ele atribuirá uma nota de carta a cada aluno com base no total. Para ajudar a garantir a precisão, você adicionará algumas verificações para ver se as pontuações individuais são muito baixas ou altas. Se a pontuação de um aluno for menor que zero ou mais do que o valor de ponto possível, o script sinaliza a célula com um preenchimento vermelho e não totaliza os pontos do aluno. Isso será uma indicação clara de quais registros você precisa verificar duas vezes. Você também adicionará algumas formatações básicas às notas para poder exibir rapidamente a parte superior e inferior da classe.

## <a name="scripting-skills-covered"></a>Habilidades de script abordadas

- Formatação de células
- Verificação de erros
- Expressões regulares
- Formatação condicional

## <a name="setup-instructions"></a>Instruções de instalação

1. Baixe <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> para seu OneDrive.

2. Abra a workbook com Excel para a Web.

3. Na guia **Automatizar,** abra **Todos os Scripts.**

4. No painel de tarefas Editor de **Código,** pressione **Novo Script** e colar o seguinte script no editor.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = (studentData[0][1] as string).match(/\d+/);
      const midtermMaxMatches = (studentData[0][2] as string).match(/\d+/);
      const finalMaxMatches = (studentData[0][3] as string).match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = (studentData[i][1] as number) + (studentData[i][2] as number) + (studentData[i][3] as number);
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
            break;
          case total < 70:
            grade = "D";
            break;
          case total < 80:
            grade = "C";
            break;
          case total < 90:
            grade = "B";
            break;
          default:
            grade = "A";
            break;
        }
    
        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#443300",
          "#FFEE22",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting: ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({ formula1, operator });
    }
    ```

5. Renomeie o script para **Calculadora de Notas** e salve-o.

## <a name="running-the-script"></a>Executando o script

Execute o **script Calculadora** de Notas na única planilha. O script totaliza as notas e atribui a cada aluno uma nota de carta. Se qualquer nota individual tiver mais pontos do que a atribuição ou teste valerá, a nota ofensiva será marcada em vermelho e o total não será calculado. Além disso, todas as notas 'A' são realçadas em verde, enquanto as notas 'D' e 'F' são realçadas em amarelo.

### <a name="before-running-the-script"></a>Antes de executar o script

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Uma planilha que mostra linhas de pontuações para alunos":::

### <a name="after-running-the-script"></a>Depois de executar o script

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Uma planilha que mostra os dados de pontuação do aluno com células inválidas em totais vermelhos para linhas de alunos válidas":::
