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
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="5406f-103">Office Cenário de exemplo de scripts: Agendar entrevistas em Teams</span><span class="sxs-lookup"><span data-stu-id="5406f-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="5406f-104">Nesse cenário, você é um recrutador de RH agendando reuniões de entrevista com candidatos em Teams.</span><span class="sxs-lookup"><span data-stu-id="5406f-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="5406f-105">Você gerencia o agendamento de entrevista de candidatos em um arquivo Excel arquivo.</span><span class="sxs-lookup"><span data-stu-id="5406f-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="5406f-106">Você precisará enviar o convite Teams reunião para o candidato e os entrevistadores.</span><span class="sxs-lookup"><span data-stu-id="5406f-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="5406f-107">Em seguida, você precisa atualizar o arquivo Excel com a confirmação de que Teams reuniões foram enviadas.</span><span class="sxs-lookup"><span data-stu-id="5406f-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="5406f-108">A solução tem três etapas combinadas em um único Power Automate fluxo.</span><span class="sxs-lookup"><span data-stu-id="5406f-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="5406f-109">Um script extrai dados de uma tabela e retorna uma matriz de objetos como dados JSON.</span><span class="sxs-lookup"><span data-stu-id="5406f-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="5406f-110">Os dados são enviados para o Teams **Criar uma ação de Teams de** reunião para enviar convites.</span><span class="sxs-lookup"><span data-stu-id="5406f-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="5406f-111">Os mesmos dados JSON são enviados para outro script para atualizar o status do convite.</span><span class="sxs-lookup"><span data-stu-id="5406f-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="5406f-112">Habilidades de script abordadas</span><span class="sxs-lookup"><span data-stu-id="5406f-112">Scripting skills covered</span></span>

* <span data-ttu-id="5406f-113">Power Automate fluxos</span><span class="sxs-lookup"><span data-stu-id="5406f-113">Power Automate flows</span></span>
* <span data-ttu-id="5406f-114">Teams integração</span><span class="sxs-lookup"><span data-stu-id="5406f-114">Teams integration</span></span>
* <span data-ttu-id="5406f-115">Análise de tabela</span><span class="sxs-lookup"><span data-stu-id="5406f-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="5406f-116">Exemplo Excel arquivo</span><span class="sxs-lookup"><span data-stu-id="5406f-116">Sample Excel file</span></span>

<span data-ttu-id="5406f-117">Baixe o arquivo <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> usado nesta solução e experimente você mesmo!</span><span class="sxs-lookup"><span data-stu-id="5406f-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="5406f-118">Certifique-se de alterar pelo menos um dos endereços de email para que você receba um convite.</span><span class="sxs-lookup"><span data-stu-id="5406f-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="5406f-119">Código de exemplo: extrair dados de tabela para agendar convites</span><span class="sxs-lookup"><span data-stu-id="5406f-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="5406f-120">Adicione esse script à sua coleção de scripts.</span><span class="sxs-lookup"><span data-stu-id="5406f-120">Add this script to your script collection.</span></span> <span data-ttu-id="5406f-121">**Nomeá-lo Agendar** Entrevistas para o fluxo.</span><span class="sxs-lookup"><span data-stu-id="5406f-121">Name it **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="5406f-122">Código de exemplo: Marcar linhas como convidados</span><span class="sxs-lookup"><span data-stu-id="5406f-122">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="5406f-123">Adicione esse script à sua coleção de scripts.</span><span class="sxs-lookup"><span data-stu-id="5406f-123">Add this script to your script collection.</span></span> <span data-ttu-id="5406f-124">**Nomeá-lo Gravar Convites Enviados** para o fluxo.</span><span class="sxs-lookup"><span data-stu-id="5406f-124">Name it **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="5406f-125">Fluxo de exemplo: execute os scripts de agendamento de entrevista e envie as Teams reuniões</span><span class="sxs-lookup"><span data-stu-id="5406f-125">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="5406f-126">Criar um novo **fluxo de nuvem instantânea.**</span><span class="sxs-lookup"><span data-stu-id="5406f-126">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="5406f-127">Escolha **Disparar manualmente um fluxo e** selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="5406f-127">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="5406f-128">Adicione uma **nova etapa que** usa o conector Excel Online **(Business)** e a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="5406f-128">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="5406f-129">Conclua o conector com os seguintes valores.</span><span class="sxs-lookup"><span data-stu-id="5406f-129">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="5406f-130">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="5406f-130">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="5406f-131">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5406f-131">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="5406f-132">**Arquivo**: hr-interviews.xlsx *(Escolhido por meio do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="5406f-132">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Script**: Agendar Entrevistas Captura de tela do conector Excel :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="online (Business)"::: concluído para obter dados de entrevista da Power Automate.
1. <span data-ttu-id="5406f-134">Adicione uma **nova etapa** que usa a ação Criar uma **Teams de** reunião.</span><span class="sxs-lookup"><span data-stu-id="5406f-134">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="5406f-135">À medida que você seleciona o conteúdo dinâmico Excel conector de Excel, um **Apply a cada** bloco será gerado para seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="5406f-135">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="5406f-136">Conclua o conector com os seguintes valores.</span><span class="sxs-lookup"><span data-stu-id="5406f-136">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="5406f-137">**ID do calendário**: Calendário</span><span class="sxs-lookup"><span data-stu-id="5406f-137">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="5406f-138">**Assunto**: Entrevista contoso</span><span class="sxs-lookup"><span data-stu-id="5406f-138">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="5406f-139">**Mensagem**: **Mensagem** (o Excel valor)</span><span class="sxs-lookup"><span data-stu-id="5406f-139">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="5406f-140">**Fuso horário**: Hora Padrão do Pacífico</span><span class="sxs-lookup"><span data-stu-id="5406f-140">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="5406f-141">**Hora de início**: **StartTime** (o Excel valor)</span><span class="sxs-lookup"><span data-stu-id="5406f-141">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="5406f-142">**Hora de** término : **FinishTime** (o Excel valor)</span><span class="sxs-lookup"><span data-stu-id="5406f-142">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Participantes obrigatórios**: **CandidateEmail** ; **InterviewerEmail** (os Excel) Captura de tela do conector de Teams concluído para agendar reuniões :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="no Power Automate.":::
1. <span data-ttu-id="5406f-144">Na mesma opção **Aplicar a cada** bloco, adicione outro conector Excel Online **(Business)** com a **ação Executar script.**</span><span class="sxs-lookup"><span data-stu-id="5406f-144">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="5406f-145">Use os seguintes valores.</span><span class="sxs-lookup"><span data-stu-id="5406f-145">Use the following values.</span></span>
    1. <span data-ttu-id="5406f-146">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="5406f-146">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="5406f-147">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5406f-147">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="5406f-148">**Arquivo**: hr-interviews.xlsx *(Escolhido por meio do navegador de arquivos)*</span><span class="sxs-lookup"><span data-stu-id="5406f-148">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="5406f-149">**Script**: Gravar Convites Enviados</span><span class="sxs-lookup"><span data-stu-id="5406f-149">**Script**: Record Sent Invites</span></span>
    1. **invites**: **result** (the Excel value) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Screenshot of the completed Excel Online (Business) connector"::: to record that invites have been sent in Power Automate.
1. <span data-ttu-id="5406f-151">Salve o fluxo e experimente-o. Use o **botão Testar** na página do editor de fluxo ou execute o fluxo através da guia **Meus fluxos.** Certifique-se de permitir o acesso quando solicitado.</span><span class="sxs-lookup"><span data-stu-id="5406f-151">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="5406f-152">Vídeo de treinamento: enviar uma reunião Teams de dados Excel dados</span><span class="sxs-lookup"><span data-stu-id="5406f-152">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="5406f-153">[Assista a Sudhi Ramamurthy](https://youtu.be/HyBdx52NOE8)passar por uma versão deste exemplo no YouTube .</span><span class="sxs-lookup"><span data-stu-id="5406f-153">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="5406f-154">Sua versão usa um script mais robusto que lida com a alteração de colunas e os horários de reunião obsoletos.</span><span class="sxs-lookup"><span data-stu-id="5406f-154">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
