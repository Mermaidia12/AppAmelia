// ════════════════════════════════════════════════════════════
// AHSD — Google Apps Script (com suporte a CORS/JSONP)
// Cole este código em: Extensões → Apps Script na sua planilha
// Depois: Implantar → Nova implantação → App da Web
//   - Executar como: Eu mesmo
//   - Quem tem acesso: Qualquer pessoa
// Copie a URL e cole nas Configurações do admin.html
// ════════════════════════════════════════════════════════════

const SHEET_RESPOSTAS = 'Respostas';
const SHEET_LINKS     = 'Links Gerados';
const SCRIPT_VERSION  = '2026-04-27-delete-items';

// ── POST: salvar dados ──────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const tipo = data.tipo || '';

    if (tipo === 'ping' || tipo === '__ping') {
      return resposta({ ok: true, method: 'POST', version: SCRIPT_VERSION });
    } else if (tipo === 'link') {
      salvarLink(ss, data);
    } else if (tipo === 'resposta') {
      salvarResposta(ss, data);
    } else if (tipo === 'limparTudo') {
      limparTudoGerado(ss);
    } else if (tipo === 'excluirLink') {
      excluirLinkGerado(ss, data.linkId || '');
    } else if (tipo === 'excluirTeste') {
      excluirTesteFeito(ss, data.linkId || '', data.instrumento || '');
    }

    return resposta({ ok: true });
  } catch(err) {
    return resposta({ ok: false, error: err.toString() });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function limparTudoGerado(ss) {
  limparOuCriarSheet(ss, SHEET_LINKS, ['Data', 'Rótulo', 'Questionários', 'Link ID', 'URL', 'Respostas']);
  limparOuCriarSheet(ss, SHEET_RESPOSTAS, ['Data','Link ID','Aluno','Escola','Série','Responsável','Instrumento','Seção','Pergunta','Resposta']);
}

function limparOuCriarSheet(ss, nome, cabecalho) {
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.appendRow(cabecalho);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  sheet.getRange(1,1,1,cabecalho.length).setValues([cabecalho]);
  sheet.getRange(1,1,1,cabecalho.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
}

function excluirLinkGerado(ss, linkId) {
  if (!linkId) return;
  const sheet = ss.getSheetByName(SHEET_LINKS);
  if (!sheet || sheet.getLastRow() < 2) return;
  const ids = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(linkId)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function excluirTesteFeito(ss, linkId, instrumento) {
  if (!linkId || !instrumento) return;
  const sheet = ss.getSheetByName(SHEET_RESPOSTAS);
  if (!sheet || sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  let removido = false;
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (String(row[1]) === String(linkId) && String(row[6]) === String(instrumento)) {
      sheet.deleteRow(i + 2);
      removido = true;
    }
  }
  if (removido) ajustarContadorLink(ss, linkId, -1);
}

function ajustarContadorLink(ss, linkId, delta) {
  const sheet = ss.getSheetByName(SHEET_LINKS);
  if (!sheet || !linkId || sheet.getLastRow() < 2) return;
  const ids = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(linkId)) {
      const cell = sheet.getRange(i + 2, 6);
      const atual = Number(cell.getValue() || 0);
      cell.setValue(Math.max(0, atual + delta));
      break;
    }
  }
}

function salvarLink(ss, data) {
  let sheet = ss.getSheetByName(SHEET_LINKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LINKS);
    const hdr = ['Data', 'Rótulo', 'Questionários', 'Link ID', 'URL', 'Respostas'];
    sheet.appendRow(hdr);
    sheet.getRange(1,1,1,hdr.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
  }
  sheet.appendRow([
    new Date().toLocaleString('pt-BR'),
    data.label || '',
    (data.questionarios || []).join(', '),
    data.linkId || '',
    data.url    || '',
    0
  ]);
}

function salvarResposta(ss, data) {
  let sheet = ss.getSheetByName(SHEET_RESPOSTAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RESPOSTAS);
    const hdr = ['Data','Link ID','Aluno','Escola','Série','Responsável','Instrumento','Seção','Pergunta','Resposta'];
    sheet.appendRow(hdr);
    sheet.getRange(1,1,1,hdr.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
  }
  const now = new Date().toLocaleString('pt-BR');
  const rows = [];
  if (data.respostas && typeof data.respostas === 'object') {
    Object.entries(data.respostas).forEach(([secao, perguntas]) => {
      if (typeof perguntas === 'object') {
        Object.entries(perguntas).forEach(([pergunta, resp]) => {
          rows.push([now, data.linkId||'', data.aluno||'', data.escola||'',
                     data.serie||'', data.responsavel||'', data.instrumento||'',
                     secao, pergunta, resp||'']);
        });
      }
    });
  }
  if (rows.length) {
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 10).setValues(rows);
  }
  // Incrementa contador de respostas no sheet de Links
  const ls = ss.getSheetByName(SHEET_LINKS);
  if (ls && data.linkId) {
    const last = ls.getLastRow();
    if (last >= 2) {
      const ids = ls.getRange(2, 4, last-1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === data.linkId) {
          const cell = ls.getRange(i+2, 6);
          cell.setValue((cell.getValue()||0) + 1);
          break;
        }
      }
    }
  }
}

// ── GET: buscar dados (suporta JSONP para evitar CORS) ──────
function doGet(e) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const tipo     = e.parameter.tipo     || 'links';
  const callback = e.parameter.callback || '';   // JSONP callback

  let result;

  if (tipo === 'ping') {
    result = { ok: true, method: 'GET', version: SCRIPT_VERSION };
  } else if (tipo === 'links') {
    const sheet = ss.getSheetByName(SHEET_LINKS);
    if (!sheet || sheet.getLastRow() < 2) {
      result = [];
    } else {
      const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
      result = rows.map(r => ({
        data: r[0], label: r[1], questionarios: r[2],
        linkId: r[3], url: r[4], respostas: r[5]
      }));
    }
  } else if (tipo === 'respostas') {
    const sheet = ss.getSheetByName(SHEET_RESPOSTAS);
    if (!sheet || sheet.getLastRow() < 2) {
      result = [];
    } else {
      const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 10).getValues();
      result = rows.map(r => ({
        data: r[0], linkId: r[1], aluno: r[2], escola: r[3],
        serie: r[4], responsavel: r[5], instrumento: r[6],
        secao: r[7], pergunta: r[8], resposta: r[9]
      }));
    }
  } else {
    result = { error: 'tipo inválido' };
  }

  const json = JSON.stringify(result);

  // Se vier com ?callback=xxx → retorna JSONP (resolve CORS)
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
