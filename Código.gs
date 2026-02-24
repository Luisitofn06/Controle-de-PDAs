const SHEET_ID = '1xcm_9ViPSx7gkaN8AsTzO-OKPlZ0bo4UhN3q7VnQ-YA'; // <--- TOME CUIDADO AQUI

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Shopee PDA Control')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no');
}

function getAppDados() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 1. Busca Saldo
  const abaSaldos = ss.getSheetByName('Saldos');
  let dadosSaldos = abaSaldos.getDataRange().getValues();
  let meuSaldoRow = dadosSaldos.find(r => r[0] == email);
  
  if (!meuSaldoRow) {
    abaSaldos.appendRow([email, 0]);
    meuSaldoRow = [email, 0];
  }

  // 2. Busca Usuários para Troca (Exceto eu e Estoque)
  const listaColegas = dadosSaldos
    .filter(r => r[0] != email && r[0] != 'ESTOQUE' && r[0] != 'Email')
    .map(r => r[0]);

  // 3. Busca Notificações (Trocas onde sou Destino e está PENDENTE)
  const abaTrans = ss.getSheetByName('Transacoes');
  const dadosTrans = abaTrans.getDataRange().getValues();
  
  // Recebimentos Pendentes (Preciso aceitar)
  const entradasPendentes = dadosTrans
    .filter(r => r[3] == email && r[6] == 'PENDENTE' && r[5] == 'TROCA')
    .map(r => ({ id: r[0], de: r[2], qtd: r[4], data: new Date(r[1]).toLocaleTimeString() }));

  // Envios que fiz e ainda não aceitaram (Só para visualizar)
  const saidasPendentes = dadosTrans
    .filter(r => r[2] == email && r[6] == 'PENDENTE' && r[5] == 'TROCA')
    .map(r => ({ para: r[3], qtd: r[4] }));

  return {
    email: email,
    saldo: meuSaldoRow[1],
    colegas: listaColegas,
    inbox: entradasPendentes,
    outbox: saidasPendentes
  };
}

// --- AÇÕES ---

function novaTroca(destino, qtd) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const abaTrans = ss.getSheetByName('Transacoes');
  const abaSaldos = ss.getSheetByName('Saldos');
  const eu = Session.getActiveUser().getEmail();
  
  // Valida Saldo Antes
  const dadosSaldos = abaSaldos.getDataRange().getValues();
  const meuSaldo = dadosSaldos.find(r => r[0] == eu)[1];
  
  // Soma o que já está pendente de saída para não deixar saldo negativo virtual
  // (Lógica avançada: Se tenho 10, envio 5 pra A e 6 pra B, o sistema deve travar o segundo)
  // Simplificação: Verifica saldo bruto agora.
  if (meuSaldo < qtd) return { success: false, msg: "Saldo insuficiente!" };

  // Cria ID único
  const id = new Date().getTime();
  
  abaTrans.appendRow([id, new Date(), eu, destino, qtd, 'TROCA', 'PENDENTE']);
  
  return { success: true, msg: "Solicitação enviada! Aguardando aceite do colega." };
}

function responderTroca(idTransacao, aceitar) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const abaTrans = ss.getSheetByName('Transacoes');
  const abaSaldos = ss.getSheetByName('Saldos');
  const eu = Session.getActiveUser().getEmail(); // Eu sou o DESTINO (quem aceita)

  const dadosTrans = abaTrans.getDataRange().getValues();
  const indexTrans = dadosTrans.findIndex(r => r[0] == idTransacao);
  
  if (indexTrans == -1) return { success: false, msg: "Transação não encontrada." };
  
  const transacao = dadosTrans[indexTrans]; // [id, data, DE, PARA, qtd, tipo, status]
  const qtd = Number(transacao[4]);
  const remetente = transacao[2];
  
  if (aceitar) {
    // 1. Movimenta Saldos
    const dadosSaldos = abaSaldos.getDataRange().getValues();
    const idxRemetente = dadosSaldos.findIndex(r => r[0] == remetente);
    const idxEu = dadosSaldos.findIndex(r => r[0] == eu);
    
    // Valida se o remetente ainda tem saldo (caso ele tenha tentado burlar)
    if (dadosSaldos[idxRemetente][1] < qtd) {
       abaTrans.getRange(indexTrans + 1, 7).setValue('CANCELADO_SALDO');
       return { success: false, msg: "O remetente não tem mais saldo suficiente." };
    }
    
    // Subtrai do Remetente
    abaSaldos.getRange(idxRemetente + 1, 2).setValue(dadosSaldos[idxRemetente][1] - qtd);
    // Adiciona para Mim
    abaSaldos.getRange(idxEu + 1, 2).setValue(dadosSaldos[idxEu][1] + qtd);
    
    // 2. Atualiza Status
    abaTrans.getRange(indexTrans + 1, 7).setValue('CONCLUIDO');
    return { success: true, msg: "PDAs recebidos com sucesso!" };
    
  } else {
    // Rejeitar
    abaTrans.getRange(indexTrans + 1, 7).setValue('REJEITADO');
    return { success: true, msg: "Transferência recusada." };
  }
}
