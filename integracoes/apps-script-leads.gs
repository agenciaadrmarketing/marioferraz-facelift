// Cole este código em Extensões > Apps Script da planilha "Leads - LP Facelift | Dr. Mário Ferraz".
// Cabeçalho da aba Leads (linha 1):
// Data/hora | Nome | E-mail | WhatsApp | Formato | Unidade | utm_source | utm_medium | utm_campaign | utm_term | gclid | Página | Status | Consulta agendada? | Observações
const SHEET_NAME = 'Leads';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    const txt = v => String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, 500);
    sh.appendRow([
      new Date(),
      txt(p.nome),
      txt(p.email),
      "'" + String(p.whatsapp || '').replace(/\D/g, ''),
      txt(p.formato),
      txt(p.unidade),
      txt(p.utm_source),
      txt(p.utm_medium),
      txt(p.utm_campaign),
      txt(p.utm_term),
      txt(p.gclid),
      txt(p.pagina)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
