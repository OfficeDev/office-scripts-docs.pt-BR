---
title: Agendar entrevistas no Teams
description: Saiba como usar Office scripts para enviar uma reunião Teams de Excel dados.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cb24da12637add805d86da4d07ce878509c6a5f6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313726"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office Cenário de exemplo de scripts: Agendar entrevistas em Teams

Nesse cenário, você é um recrutador de RH agendando reuniões de entrevista com candidatos em Teams. Você gerencia o agendamento de entrevista de candidatos em um arquivo Excel arquivo. Você precisará enviar o convite Teams reunião para o candidato e os entrevistadores. Em seguida, você precisa atualizar o arquivo Excel com a confirmação de que Teams reuniões foram enviadas.

A solução tem três etapas combinadas em um único Power Automate fluxo.

1. Um script extrai dados de uma tabela e retorna uma matriz de objetos como dados JSON.
1. Os dados são enviados para o Teams **Criar uma ação de Teams de** reunião para enviar convites.
1. Os mesmos dados JSON são enviados para outro script para atualizar o status do convite.

## <a name="scripting-skills-covered"></a>Habilidades de script abordadas

* Power Automate fluxos
* Teams integração
* Análise de tabela

## <a name="sample-excel-file"></a>Exemplo Excel arquivo

Baixe o arquivo <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> usado nesta solução e experimente você mesmo! Certifique-se de alterar pelo menos um dos endereços de email para que você receba um convite.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Código de exemplo: extrair dados de tabela para agendar convites

Adicione esse script à sua coleção de scripts. **Nomeá-lo Agendar** Entrevistas para o fluxo.

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a>Código de exemplo: Marcar linhas como convidados

Adicione esse script à sua coleção de scripts. **Nomeá-lo Gravar Convites Enviados** para o fluxo.

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Fluxo de exemplo: execute os scripts de agendamento de entrevista e envie as Teams reuniões

1. Criar um novo **fluxo de nuvem instantânea.**
1. Escolha **Disparar manualmente um fluxo e** selecione **Criar**.
1. Adicione uma **nova etapa que** usa o conector Excel Online **(Business)** e a **ação Executar script.** Conclua o conector com os seguintes valores.
    1. **Localização**: OneDrive for Business
    1. **Biblioteca de Documentos**: OneDrive
    1. **Arquivo**: hr-interviews.xlsx *(Escolhido por meio do navegador de arquivos)*
    1. **Script**: Agendar Entrevistas Captura de tela do conector Excel :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="online (Business)"::: concluído para obter dados de entrevista da Power Automate.
1. Adicione uma **nova etapa** que usa a ação Criar uma **Teams de** reunião. À medida que você seleciona o conteúdo dinâmico Excel conector de Excel, um **Apply a cada** bloco será gerado para seu fluxo. Conclua o conector com os seguintes valores.
    1. **ID do calendário**: Calendário
    1. **Assunto**: Entrevista contoso
    1. **Mensagem**: **Mensagem** (o Excel valor)
    1. **Fuso horário**: Hora Padrão do Pacífico
    1. **Hora de início**: **StartTime** (o Excel valor)
    1. **Hora de** término : **FinishTime** (o Excel valor)
    1. **Participantes obrigatórios**: **CandidateEmail** ; **InterviewerEmail** (os Excel) Captura de tela do conector de Teams concluído para agendar reuniões :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="no Power Automate.":::
1. Na mesma opção **Aplicar a cada** bloco, adicione outro conector Excel Online **(Business)** com a **ação Executar script.** Use os seguintes valores.
    1. **Localização**: OneDrive for Business
    1. **Biblioteca de Documentos**: OneDrive
    1. **Arquivo**: hr-interviews.xlsx *(Escolhido por meio do navegador de arquivos)*
    1. **Script**: Gravar Convites Enviados
    1. **invites**: **result** (the Excel value) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Screenshot of the completed Excel Online (Business) connector"::: to record that invites have been sent in Power Automate.
1. Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Vídeo de treinamento: enviar uma reunião Teams de dados Excel dados

[Assista a Sudhi Ramamurthy](https://youtu.be/HyBdx52NOE8)passar por uma versão deste exemplo no YouTube . Sua versão usa um script mais robusto que lida com a alteração de colunas e os horários de reunião obsoletos.
