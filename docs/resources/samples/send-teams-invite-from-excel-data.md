---
title: Envie uma reunião de Teams a partir de dados Excel
description: Aprenda a usar Office Scripts para enviar uma reunião Teams a partir de dados Excel.
ms.date: 05/06/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: 85b39d7e3d1008dee01e7fe9c690116be1d7e5d8
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545627"
---
# <a name="send-teams-meeting-from-excel-data"></a>Envie Teams reunião a partir de dados Excel

Esta solução mostra como usar Office Scripts e ações Power Automate para selecionar linhas de Excel arquivo e usá-lo para enviar um convite de reunião Teams e atualizar Excel.

## <a name="example-scenario"></a>Cenário de exemplo

* Um recrutador de RH gerencia o cronograma de entrevistas dos candidatos em um arquivo Excel.
* O recrutador precisa enviar o convite Teams reunião para o candidato e entrevistadores. As regras de negócios são selecionar:

    a Convida apenas para aqueles para quem o convite ainda não foi enviado como registrado na coluna de arquivos.

    b Datas de entrevista no futuro (sem datas passadas).

* O recrutador precisa atualizar o arquivo Excel com a confirmação de que todas as reuniões Teams foram enviadas para os registros elegíveis.

A solução tem 3 partes:

1. Office Script para extrair dados de uma tabela com base nas condições e retorna uma matriz de objetos como dados JSON.
1. Os dados são então enviados ao Teams **Criar uma ação de reunião Teams** para enviar convites. Envie uma Teams reunião por instância na matriz JSON.
1. Envie os mesmos dados JSON para outro script Office para atualizar o status do convite.

## <a name="sample-excel-file"></a>Arquivo de Excel de amostra

Baixe o arquivo <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> usado nesta solução e experimente você mesmo!

## <a name="sample-code-select-filtered-rows-from-table-as-json"></a>Código de amostra: Selecione linhas filtradas da tabela como JSON

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  console.log("Current date time: " + new Date().toUTCString());
  const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue() as number;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet('Interviews');
  const table = sheet.getTables()[0];
  const dataRows: string[][] = table.getRangeBetweenHeaderAndTotal().getTexts();

  // Convert the table rows into InterviewInvite objects for the flow.
  const recordDetails: RecordDetail[] = returnObjectFromValues(dataRows);
  const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
  console.log(JSON.stringify(inviteRecords));
  return inviteRecords;
}

/**
 * Converts table values into a RecordDetail array.
 */
function returnObjectFromValues(values: string[][]): RecordDetail[] {
  let objectArray: BasicObj[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }
    objectArray.push(object);
  }
  return objectArray as RecordDetail[];
}

/**
 * Generate interview records by selecting required columns.
 * @param records Input records from the table of interviews.
 * @param mins Number of minutes to add to the start date-time.
 */
function generateInterviewRecords(records: RecordDetail[], mins: number): InterviewInvite[] {
  const interviewInvites: InterviewInvite[] = [];

  records.forEach((record) => {
    // Interviewer 1
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time1'])) > new Date()) {
      console.log("selected " + new Date(record['Start time1']).toUTCString());
      let startTime = new Date(record['Start time1']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer1,
        InterviewerEmail: record['Interviewer1 email'],
        StartTime: startTime,
        FinishTime: finishTime
      });
    } else {
      console.log("Rejected " + (new Date(record['Start time1']).toUTCString()));
    }
    // Interviewer 2 
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time2'])) > new Date()) {
      console.log("selected " + new Date(record['Start time2']).toUTCString());


      let startTime = new Date(record['Start time2']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer2,
        InterviewerEmail: record['Interviewer2 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time2']).toUTCString()))

    }
  })
  return interviewInvites;
}

/**
 * Add minutes to start date-time.
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date-time
 */
function addMins(startDateTime: Date, mins: number) {
  return new Date(startDateTime.getTime() + mins * 60 * 1000);
}

// Basic key-value pair object.
interface BasicObj {
  [key: string]: string | number | boolean
}

// Input record that matches the table data.
interface RecordDetail extends BasicObj {
  ID: string
  'Invite to interview': string
  Candidate: string
  'Candidate email': string
  'Candidate contact': string
  Interviewer1: string
  'Interviewer1 email': string
  Interviewer2: string
  'Interviewer2 email': string
  'Start time1': string
  'Start time2': string
}

// Output record.
interface InterviewInvite extends BasicObj {
  ID: string
  Candidate: string
  CandidateEmail: string
  CandidateContact: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
}
```

## <a name="sample-code-mark-as-invited"></a>Código amostral: Mark como convidado

```TypeScript
function main(workbook: ExcelScript.Workbook, completedInvitesString: string) {
    completedInvitesString = `[
      {
        "ID": "10",
        "Candidate": "Adele ",
        "CandidateEmail": "AdeleV@M365x904181.OnMicrosoft.com",
        "CandidateContact": "1234567899",
        "Interviewer": "Megan",
        "InterviewerEmail": "MeganB@M365x904181.OnMicrosoft.com",
        "StartTime": "2020-11-03T18:30:00Z",
        "FinishTime": "2020-11-03T22:45:00Z"
      },
      {
        "ID": "30",
        "Candidate": "Allan ",
        "CandidateEmail": "AllanD@M365x904181.OnMicrosoft.com",
        "CandidateContact": "1234567978",
        "Interviewer": "Raul",
        "InterviewerEmail": "RaulR@M365x904181.OnMicrosoft.com",
        "StartTime": "2020-11-03T23:00:00Z",
        "FinishTime": "2020-11-03T23:45:00Z"
      }
    ]`;
    let completedInvites = JSON.parse(completedInvitesString) as InterviewInvite[];
    const sheet = workbook.getWorksheet('Interviews');
    const range = sheet.getTables()[0].getRange();
    const dataRows = range.getValues();
    for (let i=0; i < dataRows.length; i++) {
        for (let invite of completedInvites) {
            if (String(dataRows[i][0]) === invite.ID) {
                range.getCell(i,1).setValue(true);
            }
        }
    }
    return;
}


// Invite record.
interface InterviewInvite  {
    ID: string
    Candidate: string
    CandidateEmail: string
    CandidateContact: string
    Interviewer: string
    InterviewerEmail: string
    StartTime: string
    FinishTime: string
}
```

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Vídeo de treinamento: Envie uma reunião de Teams a partir de dados Excel

[Assista Sudhi Ramamurthy andar através desta amostra no YouTube](https://youtu.be/HyBdx52NOE8).
