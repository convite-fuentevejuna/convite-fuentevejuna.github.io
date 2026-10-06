// Backend das confirmações do convite Fuenteovejuna.
// Script vinculado à planilha "Fuenteovejuna — Confirmações".
// Publicado como Web App: Executar como "Eu", acesso "Qualquer pessoa".

const DATAS = {
  '16/10': 'Sex 16/10 · 19h',
  '17/10': 'Sáb 17/10 · 18h',
  '20/10': 'Ter 20/10 · 19h',
  '24/10': 'Sáb 24/10 · 19h',
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    let nome = String(d.nome || '').trim().slice(0, 80);
    const pessoas = Math.max(1, Math.min(10, parseInt(d.pessoas, 10) || 1));
    const sessao = DATAS[d.data];
    if (nome.length < 2 || !sessao) return json({ ok: false, error: 'dados inválidos' });
    if (/^[=+\-@]/.test(nome)) nome = "'" + nome;
    sheets().conf.appendRow([new Date(), nome, pessoas, sessao]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true });
}

function sheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let conf = ss.getSheetByName('Confirmações');
  if (!conf) {
    conf = ss.getSheets()[0];
    conf.setName('Confirmações');
    conf.getRange('A1:D1').setValues([['Quando', 'Nome', 'Pessoas', 'Sessão']]).setFontWeight('bold');
    conf.setFrozenRows(1);
    conf.setColumnWidth(1, 150);
    conf.setColumnWidth(2, 220);
    conf.setColumnWidth(4, 140);
  }
  let resumo = ss.getSheetByName('Resumo');
  if (!resumo) {
    resumo = ss.insertSheet('Resumo');
    resumo.getRange('A1:C1').setValues([['Sessão', 'Lugares', 'Confirmações']]).setFontWeight('bold');
    const rows = Object.values(DATAS).map((s, i) => {
      const r = i + 2;
      return [s, `=SUMIF('Confirmações'!D:D,A${r},'Confirmações'!C:C)`, `=COUNTIF('Confirmações'!D:D,A${r})`];
    });
    rows.push(['Total', '=SUM(B2:B5)', '=SUM(C2:C5)']);
    resumo.getRange(2, 1, rows.length, 3).setValues(rows);
    resumo.getRange('A6:C6').setFontWeight('bold');
    resumo.setColumnWidth(1, 160);
  }
  return { conf, resumo };
}

// Rode uma vez no editor para criar as abas e autorizar o script.
function setup() {
  sheets();
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
