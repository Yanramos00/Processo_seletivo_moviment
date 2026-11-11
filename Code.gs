/**
 * Clínica Moviment — processo seletivo
 * Backend (Google Apps Script) ligado à planilha de resultados.
 *
 * Como publicar:
 * 1. Abra a planilha de resultados > Extensões > Apps Script.
 * 2. Cole este arquivo como Code.gs e crie um arquivo HTML chamado "admin" com o conteúdo de admin.html.
 * 3. Implantar > Nova implantação > Tipo: Aplicativo da Web.
 *    - Executar como: Eu
 *    - Quem tem acesso: Qualquer pessoa   (necessário para os candidatos enviarem a prova)
 *    O painel só abre para a sua conta: doGet e getRows verificam o dono do script.
 * 4. Copie a URL e cole em ENDPOINT no arquivo da prova.
 */

const SHEET_NAME = 'Respostas';
const TZ = 'America/Sao_Paulo';
const HEADERS = [
  'ID', 'Recebido em', 'Nome', 'WhatsApp', 'E-mail', 'Início', 'Fim',
  'Motivo de encerramento', 'Nota (de 10)', 'Respostas', 'Acertos'
];

// Gabarito: índice correto de cada questão (0 = A, 1 = B, 2 = C, 3 = D), na ordem da prova.
const GABARITO = [1, 2, 0, 3, 1, 2, 0, 3, 1, 0];

function doGet(e) {
  if (!isOwner_()) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Acesso negado.</p>');
  }
  return HtmlService.createHtmlOutputFromFile('admin')
    .setTitle('Painel do processo seletivo · Clínica Moviment')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.id || !d.name) return json_({ ok: false, error: 'dados incompletos' });

    const sheet = getSheet_();
    // Evita duplicar a mesma tentativa (ex.: envio no fechamento da página + no clique em finalizar)
    const found = sheet.createTextFinder(String(d.id)).matchEntireCell(true).findNext();
    if (found) return json_({ ok: true, duplicate: true });

    const answers = Array.isArray(d.answers) ? d.answers : [];
    let score = 0;
    const marks = GABARITO.map((correct, i) => {
      const a = answers[i];
      if (a === correct) score++;
      return a === correct ? 'V' : 'X';
    });
    const letters = GABARITO.map((_, i) => {
      const a = answers[i];
      return (a === null || a === undefined) ? '-' : 'ABCD'[a];
    });

    sheet.appendRow([
      String(d.id),
      new Date(),
      String(d.name).slice(0, 120),
      String(d.phone || '').slice(0, 20),
      String(d.email || '').slice(0, 120),
      new Date(d.started),
      new Date(d.finished),
      String(d.reason || ''),
      score,
      letters.join(' '),
      marks.join(' ')
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Chamado pelo painel (google.script.run). Só o dono do script recebe os dados.
function getRows() {
  if (!isOwner_()) throw new Error('Acesso negado.');
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return values.map(r => ({
    id: String(r[0]),
    recebido: fmtDate_(r[1]),
    nome: String(r[2]),
    whatsapp: String(r[3]),
    email: String(r[4]),
    inicio: fmtDate_(r[5]),
    fim: fmtDate_(r[6]),
    motivo: String(r[7]),
    nota: Number(r[8]),
    respostas: String(r[9]).split(' '),
    acertos: String(r[10]).split(' ')
  }));
}

function isOwner_() {
  const active = Session.getActiveUser().getEmail();
  return !!active && active === Session.getEffectiveUser().getEmail();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function fmtDate_(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'dd/MM/yyyy HH:mm') : String(v || '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
