// Leads da LP Facelift (facelift.marioferraz.com.br) -> planilha + aviso por e-mail.
// Cole em Extensões > Apps Script da planilha, salve e publique como App da Web
// (Executar como: Eu · Quem pode acessar: Qualquer pessoa). A URL /exec vai no index.html.

var EMAIL_AVISO = 'rastajohny091@gmail.com,caioanderle30@gmail.com';
var PLANILHA_ID = '1WHLXae9VAv9EEz1RY4qBAcrRPhX9odDBnD19HjrFljo';
var FUSO = 'America/Sao_Paulo';

var COLUNAS = ['Data','Nome','Email','WhatsApp','Formato','Unidade',
               'Origem do form','utm_source','utm_medium','utm_campaign',
               'utm_term','utm_content','gclid','fbclid','Pagina',
               'Status','Consulta agendada?','Observacoes','Notificacao'];

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (p.nome || p.email || p.whatsapp) return salvarLead(p);
  return ContentService.createTextOutput('LP Facelift Dr. Mario Ferraz - endpoint ativo');
}

function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); }
  catch (err) { d = (e && e.parameter) ? e.parameter : {}; }
  return salvarLead(d);
}

// Evita que um valor comece com = + - @ e vire fórmula na planilha.
function txt(v) {
  return String(v == null ? '' : v).trim().replace(/^[=+\-@]/, "'$&").slice(0, 500);
}

function formatoLegivel(f) {
  if (f === 'online') return 'Consulta digital';
  if (f === 'presencial') return 'Presencial';
  return f || '';
}

function salvarLead(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.openById(PLANILHA_ID).getSheets()[0];

    if (sh.getLastRow() === 0) {
      sh.appendRow(COLUNAS);
      sh.getRange(1, 1, 1, COLUNAS.length)
        .setFontWeight('bold').setBackground('#2b2620').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }

    var nome  = txt(d.nome);
    var email = txt(d.email);
    var whats = String(d.whatsapp || '').replace(/\D/g, '');
    var formato = formatoLegivel(d.formato);
    var unidade = d.formato === 'online' ? '' : txt(d.unidade);

    // Ignora o mesmo WhatsApp enviado de novo em menos de 2 minutos.
    var last = sh.getLastRow();
    if (last > 1 && whats) {
      var check = sh.getRange(Math.max(2, last - 4), 1, Math.min(5, last - 1), 4).getValues();
      for (var i = 0; i < check.length; i++) {
        var quando = check[i][0], fone = String(check[i][3] || '').replace(/\D/g, '');
        if (fone === whats && quando instanceof Date && (new Date() - quando) < 120000) {
          return ContentService.createTextOutput('duplicado');
        }
      }
    }

    var agora = new Date();
    sh.appendRow([agora, nome, email, whats ? "'" + whats : '', formato, unidade,
      txt(d.origem_form), txt(d.utm_source), txt(d.utm_medium), txt(d.utm_campaign),
      txt(d.utm_term), txt(d.utm_content), txt(d.gclid), txt(d.fbclid),
      txt(d.pagina), '', '', '', '']);

    var linha = sh.getLastRow(), resultado = '';
    try { resultado = notificar(nome, email, whats, formato, unidade, d, agora); }
    catch (errMail) { resultado = 'ERRO: ' + errMail; }
    sh.getRange(linha, COLUNAS.length).setValue(resultado);

    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('erro: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function notificar(nome, email, whats, formato, unidade, d, agora) {
  var lista = EMAIL_AVISO.split(',').map(function (x) { return x.trim(); })
    .filter(function (x) { return x.indexOf('@') > 0; });
  if (!lista.length) return 'sem destinatarios';

  var link = whats ? 'https://wa.me/' + (whats.length <= 11 ? '55' + whats : whats) : '';
  var quando = Utilities.formatDate(agora, FUSO, "dd/MM/yyyy 'as' HH:mm");
  var origem = [d.utm_source, d.utm_medium, d.utm_campaign].filter(String).join(' / ');
  var planilha = 'https://docs.google.com/spreadsheets/d/' + PLANILHA_ID + '/edit';
  var primeiro = (nome || '').split(' ')[0];
  var atendimento = formato + (unidade ? ' · ' + unidade : '');
  var form = d.origem_form === 'popup' ? 'Popup' : (d.origem_form === 'formulario' ? 'Formulario da pagina' : d.origem_form);

  // Paleta da LP: grafite, creme e dourado.
  var FUNDO = '#f4f1ec', CARTAO = '#ffffff', BORDA = '#ebe6de', ESCURO = '#2b2620', DOURADO = '#b8995a', CINZA = '#7a7268';

  function esc(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function linhaTab(rotulo, valor, href, ultima) {
    if (!valor) return '';
    var borda = ultima ? '' : 'border-bottom:1px solid ' + BORDA + ';';
    var conteudo = href
      ? '<a href="' + esc(href) + '" style="color:' + ESCURO + ';text-decoration:underline;word-break:break-all;">' + esc(valor) + '</a>'
      : '<span style="color:' + ESCURO + ';font-weight:bold;word-break:break-word;">' + esc(valor) + '</span>';
    return '<tr>' +
      '<td bgcolor="' + CARTAO + '" style="padding:14px 24px;' + borda + 'font:14px Arial,sans-serif;color:' + CINZA + ';vertical-align:top;width:36%;">' + esc(rotulo) + '</td>' +
      '<td bgcolor="' + CARTAO + '" style="padding:14px 24px;' + borda + 'font:14px Arial,sans-serif;color:' + ESCURO + ';vertical-align:top;">' + conteudo + '</td>' +
      '</tr>';
  }

  var linhas =
    linhaTab('Nome', nome) +
    linhaTab('E-mail', email, 'mailto:' + email) +
    linhaTab('WhatsApp', whats, link || null) +
    linhaTab('Atendimento', atendimento) +
    linhaTab('Formulario', form) +
    linhaTab('Origem', origem || 'acesso direto') +
    linhaTab('gclid', d.gclid) +
    linhaTab('Pagina', d.pagina, d.pagina) +
    linhaTab('Recebido em', quando, null, true);

  var botao = link
    ? '<tr><td align="center" bgcolor="' + CARTAO + '" style="padding:28px 24px 10px;text-align:center;">' +
        '<a href="' + esc(link) + '" style="display:inline-block;background:' + ESCURO + ';padding:16px 34px;font:bold 13px Arial,sans-serif;letter-spacing:2px;text-transform:uppercase;color:#ffffff;text-decoration:none;">' +
        'Chamar ' + esc(primeiro || 'o paciente') + ' no WhatsApp</a>' +
        '</td></tr>'
    : '';

  var html =
  '<!DOCTYPE html><html><head><meta charset="utf-8">' +
  '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body bgcolor="' + FUNDO + '" style="margin:0;padding:0;background:' + FUNDO + ';">' +
  '<div style="display:none;max-height:0;overflow:hidden;">Novo lead Facelift: ' + esc(nome) + ' - ' + esc(atendimento) + '</div>' +
  '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="' + FUNDO + '">' +
  '<tr><td align="center" style="padding:24px 12px;">' +
  '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="' + CARTAO + '" style="max-width:520px;background:' + CARTAO + ';border-top:3px solid ' + DOURADO + ';">' +

  '<tr><td bgcolor="' + ESCURO + '" style="padding:26px 24px;background:' + ESCURO + ';">' +
  '<div style="font:11px Arial,sans-serif;letter-spacing:3px;color:' + DOURADO + ';text-transform:uppercase;">Instituto Mario Ferraz · Facelift</div>' +
  '<div style="font:28px Georgia,serif;color:#ffffff;padding-top:8px;">Novo lead</div>' +
  '<div style="font:14px Arial,sans-serif;color:#d9d3c9;padding-top:8px;line-height:1.5;">Um paciente preencheu o formulario da LP de Facelift.</div>' +
  '</td></tr>' +

  '<tr><td style="padding:0;">' +
  '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
  linhas + botao +
  '<tr><td align="center" style="padding:16px 24px 28px;text-align:center;">' +
  '<a href="' + esc(planilha) + '" style="font:13px Arial,sans-serif;color:' + CINZA + ';text-decoration:underline;">Ver todos os leads na planilha</a>' +
  '</td></tr>' +
  '</table></td></tr>' +

  '</table>' +
  '<div style="font:11px Arial,sans-serif;color:' + CINZA + ';padding-top:16px;text-align:center;">Aviso automatico - LP Facelift Dr. Mario Ferraz · ADR Marketing</div>' +
  '</td></tr></table></body></html>';

  var texto = 'NOVO LEAD - LP FACELIFT\n\nNome: ' + (nome || '-') +
    '\nEmail: ' + (email || '-') +
    '\nWhatsApp: ' + (whats || '-') +
    '\nAtendimento: ' + (atendimento || '-') +
    '\nFormulario: ' + (form || '-') +
    '\nOrigem: ' + (origem || '-') +
    '\ngclid: ' + (d.gclid || '-') +
    '\nPagina: ' + (d.pagina || '-') +
    '\nRecebido em: ' + quando +
    (link ? '\n\nAbrir conversa: ' + link : '');

  var assunto = 'NOVO LEAD Facelift - ' + (nome || 'sem nome') + (atendimento ? ' (' + atendimento + ')' : '');
  var res = [];
  for (var i = 0; i < lista.length; i++) {
    try {
      MailApp.sendEmail({ to: lista[i], subject: assunto, body: texto, htmlBody: html, name: 'LP Facelift - Dr. Mario Ferraz' });
      res.push('OK ' + lista[i]);
    } catch (e1) {
      try {
        GmailApp.sendEmail(lista[i], assunto, texto, { htmlBody: html, name: 'LP Facelift - Dr. Mario Ferraz' });
        res.push('OK(gmail) ' + lista[i]);
      } catch (e2) { res.push('FALHOU ' + lista[i] + ' (' + e2 + ')'); }
    }
  }
  return res.join(' | ');
}

// Rode esta função no editor (botão Executar) para testar a planilha e o e-mail.
function testar() {
  Logger.log(salvarLead({
    nome: 'Teste Facelift', email: 'teste@teste.com', whatsapp: '(19) 99999-9999',
    formato: 'presencial', unidade: 'Campinas', origem_form: 'popup',
    utm_source: 'google', utm_medium: 'cpc', utm_campaign: 'facelift', gclid: 'TESTE123',
    pagina: 'https://facelift.marioferraz.com.br/'
  }).getContent());
}
