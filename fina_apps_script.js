// ── Configuração — preencha aqui ─────────────────────────
const BOT_TOKEN          = 'SEU_TOKEN_AQUI';
const SHEET_ID           = 'SEU_SHEET_ID_AQUI';
const AUTHORIZED_USER_ID = 'SEU_ID_TELEGRAM_AQUI'; // deixe '' para liberar pra todos

// ── Recebe mensagens do Telegram ──────────────────────────
function doPost(e) {
  const data    = JSON.parse(e.postData.contents);
  const message = data.message;
  if (!message) return;

  const chatId = message.chat.id;
  const userId = message.from.id.toString();
  const text   = message.text || '';

  if (AUTHORIZED_USER_ID && userId !== AUTHORIZED_USER_ID) {
    sendMessage(chatId, 'Acesso não autorizado.');
    return;
  }

  handleMessage(chatId, text);
}

// ── Roteador de mensagens ─────────────────────────────────
function handleMessage(chatId, text) {
  const lower   = text.toLowerCase().trim();
  const now     = new Date();
  const dateStr = Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM/yyyy');
  const month   = Utilities.formatDate(now, 'America/Sao_Paulo', 'MM/yyyy');

  if (lower.startsWith('/start')) {
    sendMessage(chatId,
      '👋 Olá! Sou a *FINA*, sua assistente financeira.\n\n' +
      '*Como me usar:*\n\n' +
      '💳 `paguei C6 209,76`\n' +
      '💸 `gastei 80 padaria`\n' +
      '💚 `recebi 5000 salário`\n' +
      '📌 `nova conta escola 650 dia 5`\n\n' +
      '📊 /saldo\n' +
      '📋 /contas\n' +
      '📈 /resumo'
    );

  } else if (lower.startsWith('/saldo')) {
    handleSaldo(chatId, month);

  } else if (lower.startsWith('/contas')) {
    handleContas(chatId);

  } else if (lower.startsWith('/resumo')) {
    handleResumo(chatId, month);

  } else if (lower.startsWith('paguei') || lower.startsWith('pag ')) {
    const clean  = lower.replace(/^(paguei|pag)\s*/, '').trim();
    const amount = extractAmount(clean);
    const desc   = removeAmount(clean) || 'Pagamento';
    if (!amount) { sendMessage(chatId, 'Não entendi. Ex: `paguei C6 209,76`'); return; }
    appendRow('Pagamentos', ['Data', 'Mês', 'Descrição', 'Valor'], [dateStr, month, desc, amount]);
    sendMessage(chatId, `✅ *Pagamento registrado!*\n📋 ${desc}\n💰 R$ ${fmt(amount)}\n📅 ${dateStr}`);

  } else if (lower.startsWith('gastei') || lower.startsWith('gasto ')) {
    const clean  = lower.replace(/^(gastei|gasto)\s*/, '').trim();
    const amount = extractAmount(clean);
    const desc   = removeAmount(clean) || 'Gasto';
    if (!amount) { sendMessage(chatId, 'Não entendi. Ex: `gastei 50 mercado`'); return; }
    appendRow('Despesas', ['Data', 'Mês', 'Descrição', 'Valor'], [dateStr, month, desc, amount]);
    sendMessage(chatId, `💸 *Gasto registrado!*\n📋 ${desc}\n💰 R$ ${fmt(amount)}\n📅 ${dateStr}`);

  } else if (lower.startsWith('recebi') || lower.startsWith('entrada ')) {
    const clean  = lower.replace(/^(recebi|entrada)\s*/, '').trim();
    const amount = extractAmount(clean);
    const desc   = removeAmount(clean) || 'Entrada';
    if (!amount) { sendMessage(chatId, 'Não entendi. Ex: `recebi 5000 salário`'); return; }
    appendRow('Receitas', ['Data', 'Mês', 'Descrição', 'Valor'], [dateStr, month, desc, amount]);
    sendMessage(chatId, `💚 *Entrada registrada!*\n📋 ${desc}\n💰 R$ ${fmt(amount)}\n📅 ${dateStr}`);

  } else if (lower.startsWith('nova conta') || lower.startsWith('nova ')) {
    const clean  = lower.replace(/^(nova conta|nova)\s*/, '').trim();
    const amount = extractAmount(clean);
    const day    = extractDay(clean);
    const desc   = removeAmount(clean).replace(/dia\s+\d{1,2}/i, '').trim() || 'Conta';
    if (!amount) { sendMessage(chatId, 'Não entendi. Ex: `nova conta escola 650 dia 5`'); return; }
    const venc = day ? `Dia ${day}` : '—';
    appendRow('Contas Fixas', ['Descrição', 'Valor', 'Vencimento', 'Adicionado em'], [desc, amount, venc, dateStr]);
    sendMessage(chatId, `📌 *Conta fixa adicionada!*\n📋 ${desc}\n💰 R$ ${fmt(amount)}\n📅 Vence: ${venc}`);

  } else {
    sendMessage(chatId,
      'Não entendi. Tente:\n\n' +
      '💳 `paguei C6 209,76`\n' +
      '💸 `gastei 80 padaria`\n' +
      '💚 `recebi 5000 salário`\n' +
      '📌 `nova conta escola 650 dia 5`\n\n' +
      '📊 /saldo   📋 /contas   📈 /resumo'
    );
  }
}

// ── Saldo do mês ──────────────────────────────────────────
function handleSaldo(chatId, month) {
  const r = sumSheet('Receitas', month);
  const d = sumSheet('Despesas', month);
  const p = sumSheet('Pagamentos', month);
  const saldo = r - d - p;
  sendMessage(chatId,
    `📊 *Saldo — ${month}*\n\n` +
    `💚 Entradas:    R$ ${fmt(r)}\n` +
    `💸 Gastos:      R$ ${fmt(d)}\n` +
    `✅ Pagamentos:  R$ ${fmt(p)}\n` +
    `──────────────────────\n` +
    `${saldo >= 0 ? '✅' : '🔴'} *Saldo: R$ ${fmt(saldo)}*`
  );
}

// ── Lista de contas fixas ─────────────────────────────────
function handleContas(chatId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName('Contas Fixas');
  if (!ws || ws.getLastRow() < 2) {
    sendMessage(chatId, 'Nenhuma conta fixa ainda.\n\nUse: `nova conta escola 650 dia 5`');
    return;
  }
  const rows = ws.getRange(2, 1, ws.getLastRow() - 1, 3).getValues();
  let msg = '📋 *Contas Fixas*\n\n';
  let total = 0;
  rows.forEach(r => {
    if (!r[0]) return;
    msg += `• ${r[0]} — R$ ${fmt(r[1])} (${r[2]})\n`;
    total += parseFloat(r[1]) || 0;
  });
  msg += `\n💰 *Total mensal: R$ ${fmt(total)}*`;
  sendMessage(chatId, msg);
}

// ── Resumo do mês ─────────────────────────────────────────
function handleResumo(chatId, month) {
  const ss        = SpreadsheetApp.openById(SHEET_ID);
  const receitas  = getMonthRows(ss, 'Receitas',   month);
  const despesas  = getMonthRows(ss, 'Despesas',   month);
  const pagamentos= getMonthRows(ss, 'Pagamentos', month);

  const totalR = receitas.reduce  ((s, r) => s + (parseFloat(r[3]) || 0), 0);
  const totalD = despesas.reduce  ((s, r) => s + (parseFloat(r[3]) || 0), 0);
  const totalP = pagamentos.reduce((s, r) => s + (parseFloat(r[3]) || 0), 0);
  const saldo  = totalR - totalD - totalP;

  let msg = `📈 *Resumo — ${month}*\n\n`;

  if (receitas.length) {
    msg += '💚 *Entradas:*\n';
    receitas.forEach(r => msg += `  ${r[0]} · ${r[2]} — R$ ${fmt(r[3])}\n`);
    msg += `  *Total: R$ ${fmt(totalR)}*\n\n`;
  }
  if (pagamentos.length) {
    msg += '✅ *Pagamentos:*\n';
    pagamentos.forEach(r => msg += `  ${r[0]} · ${r[2]} — R$ ${fmt(r[3])}\n`);
    msg += `  *Total: R$ ${fmt(totalP)}*\n\n`;
  }
  if (despesas.length) {
    msg += '💸 *Gastos:*\n';
    despesas.forEach(r => msg += `  ${r[0]} · ${r[2]} — R$ ${fmt(r[3])}\n`);
    msg += `  *Total: R$ ${fmt(totalD)}*\n\n`;
  }

  msg += `──────────────────────\n${saldo >= 0 ? '✅' : '🔴'} *Saldo: R$ ${fmt(saldo)}*`;
  sendMessage(chatId, msg);
}

// ── Utilitários ───────────────────────────────────────────
function extractAmount(text) {
  const m = text.match(/(\d{1,6}[.,]\d{2}|\d{1,6})/);
  return m ? parseFloat(m[1].replace(',', '.')) : null;
}

function extractDay(text) {
  const m = text.match(/dia\s+(\d{1,2})/i);
  return m ? parseInt(m[1]) : null;
}

function removeAmount(text) {
  return text.replace(/\d{1,6}[.,]\d{2}|\d{1,6}/, '').trim();
}

function fmt(v) {
  return parseFloat(v || 0).toFixed(2).replace('.', ',');
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let ws   = ss.getSheetByName(name);
  if (!ws) { ws = ss.insertSheet(name); ws.appendRow(headers); }
  else if (ws.getLastRow() === 0) { ws.appendRow(headers); }
  return ws;
}

function appendRow(sheetName, headers, row) {
  getOrCreateSheet(sheetName, headers).appendRow(row);
}

function sumSheet(sheetName, month) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(sheetName);
  if (!ws || ws.getLastRow() < 2) return 0;
  return ws.getRange(2, 1, ws.getLastRow() - 1, 4).getValues()
    .reduce((sum, r) => r[1] === month ? sum + (parseFloat(r[3]) || 0) : sum, 0);
}

function getMonthRows(ss, sheetName, month) {
  const ws = ss.getSheetByName(sheetName);
  if (!ws || ws.getLastRow() < 2) return [];
  return ws.getRange(2, 1, ws.getLastRow() - 1, 4).getValues()
    .filter(r => r[1] === month && r[0]);
}

function sendMessage(chatId, text) {
  UrlFetchApp.fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'Markdown' })
  });
}

// ── Rodar UMA VEZ para ativar o webhook ───────────────────
function ativarWebhook() {
  const webhookUrl = ScriptApp.getService().getUrl();
  const resp = UrlFetchApp.fetch(
    `https://api.telegram.org/bot${BOT_TOKEN}/setWebhook?url=${webhookUrl}`
  );
  Logger.log(resp.getContentText());
}
