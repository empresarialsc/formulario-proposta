/**
 * Formulário de Informações para Proposta — Empresarial Assessoria Contábil
 * Backend em Google Apps Script (script vinculado à planilha).
 *
 * O que faz a cada envio do formulário:
 *   1. Valida os dados recebidos
 *   2. Grava nas abas "Respostas" (1 linha por envio) e "Empresas" (1 linha por CNPJ)
 *   3. Gera um PDF com todas as informações e salva no Google Drive
 *   4. Envia o PDF por e-mail para os destinatários configurados abaixo
 */

/* =================== CONFIGURAÇÃO (edite apenas este bloco) =================== */
const CONFIG = {
  // E-mails que receberão o PDF a cada envio (um ou mais)
  DESTINATARIOS: ['comercial@seudominio.com.br'],

  // true = o responsável pelo preenchimento também recebe uma cópia do PDF
  COPIA_AO_RESPONSAVEL: false,

  // Pasta do Drive onde os PDFs serão guardados. Deixe PASTA_ID vazio para o
  // script criar/usar automaticamente uma pasta com o nome abaixo.
  PASTA_ID: '',
  NOME_PASTA_PDFS: 'Formulários de Precificação (PDFs)',

  // (opcional) URL pública do logo, ex.: https://SEU-USUARIO.github.io/SEU-REPO/logo.png
  LOGO_URL: '',

  REMETENTE_NOME: 'Empresarial Assessoria Contábil — Formulário de Precificação',
  FUSO: 'America/Sao_Paulo'
};
/* ============================================================================== */

const ABA_RESPOSTAS = 'Respostas';
const ABA_EMPRESAS = 'Empresas';
const CAB_RESPOSTAS = ['Protocolo', 'Data/Hora', 'Responsável', 'E-mail', 'Telefone', 'Qtd. CNPJs',
  'CNPJs', 'Razões sociais', 'Observações', 'Link do PDF', 'Status do e-mail', 'Status da precificação'];
const CAB_EMPRESAS = ['Protocolo', 'Data/Hora', 'Nº', 'CNPJ', 'Tipo', 'Razão social', 'Atividade', 'Município/UF',
  'Regime tributário', 'Faturamento médio mensal (R$)', 'Notas fiscais/mês', 'Sócios',
  'Funcionários ativos', 'Recibos de autônomo/mês', 'Particularidades'];

const MAX_EMPRESAS = 30;

/* ------------------------------ ENDPOINTS ------------------------------ */

// Abrir a URL do Web App no navegador deve mostrar "OK" (útil para testar a implantação)
function doGet() {
  return ContentService.createTextOutput('OK — formulário de precificação ativo.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let travou = false;
  try {
    lock.waitLock(30000);
    travou = true;

    const bruto = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Campo-isca (anti-robô): se vier preenchido, finge sucesso e descarta
    if (bruto.website) return json_({ ok: true });

    const d = validar_(bruto);
    const ss = SpreadsheetApp.getActive();
    const abaR = aba_(ss, ABA_RESPOSTAS, CAB_RESPOSTAS);
    const abaE = aba_(ss, ABA_EMPRESAS, CAB_EMPRESAS);

    // Idempotência: reenvio do mesmo protocolo não duplica registros
    if (protocoloExiste_(abaR, d.id)) return json_({ ok: true, duplicado: true, id: d.id });

    const dataHora = Utilities.formatDate(new Date(), CONFIG.FUSO, 'dd/MM/yyyy HH:mm:ss');

    // PDF + Drive (falha aqui não impede o registro dos dados)
    let pdf = null, linkPdf = '', statusEmail = '';
    try {
      pdf = gerarPdf_(d, dataHora);
      const arquivo = pasta_().createFile(pdf);
      linkPdf = arquivo.getUrl();
    } catch (errPdf) {
      console.error('PDF/Drive: ' + errPdf);
      statusEmail = 'PDF não gerado: ' + errPdf.message;
    }

    // Grava dados
    abaR.appendRow([
      d.id, dataHora, d.responsavel.nome, d.responsavel.email, d.responsavel.telefone,
      d.empresas.length,
      d.empresas.map(x => x.cnpj).join(' | '),
      d.empresas.map(x => x.razaoSocial).join(' | '),
      d.observacoes, linkPdf, statusEmail || 'Pendente', 'Pendente'
    ].map(seguro_));
    const linhaResposta = abaR.getLastRow();

    d.empresas.forEach((x, i) => {
      abaE.appendRow([
        d.id, dataHora, i + 1, x.cnpj, x.tipo, x.razaoSocial, x.atividade, x.municipioUf,
        x.regime, x.faturamentoMensal, x.notasMes, x.socios, x.funcionarios, x.autonomos,
        x.complexidades.join('; ')
      ].map(seguro_));
    });
    SpreadsheetApp.flush();

    // E-mail (mesmo sem PDF, avisa a equipe — nesse caso sem anexo)
    const colStatus = CAB_RESPOSTAS.indexOf('Status do e-mail') + 1;
    try {
      enviarEmail_(d, pdf, linkPdf, dataHora, ss.getUrl());
      abaR.getRange(linhaResposta, colStatus).setValue(
        (pdf ? 'Enviado em ' : 'Enviado SEM PDF (falha ao gerar) em ') + dataHora);
    } catch (errMail) {
      console.error('E-mail: ' + errMail);
      abaR.getRange(linhaResposta, colStatus).setValue('FALHA no envio: ' + errMail.message);
    }

    return json_({ ok: true, id: d.id });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (travou) { try { lock.releaseLock(); } catch (_) {} }
  }
}

/* ------------------------------ VALIDAÇÃO ------------------------------ */

function validar_(b) {
  const txt = (v, max) => String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  const int = (v, min, max) => Math.min(max, Math.max(min, parseInt(v, 10) || 0));

  const id = txt(b.id, 40);
  if (!/^PF-\d{8}-[A-Z0-9]{6}$/.test(id)) throw new Error('Protocolo inválido.');

  const r = b.responsavel || {};
  const resp = { nome: txt(r.nome, 120), email: txt(r.email, 120), telefone: txt(r.telefone, 20) };
  if (resp.nome.length < 3) throw new Error('Nome do responsável inválido.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(resp.email)) throw new Error('E-mail do responsável inválido.');
  if (resp.telefone.replace(/\D/g, '').length < 10) throw new Error('Telefone do responsável inválido.');

  if (!Array.isArray(b.empresas) || b.empresas.length < 1) throw new Error('Informe ao menos um CNPJ.');
  if (b.empresas.length > MAX_EMPRESAS) throw new Error('Máximo de ' + MAX_EMPRESAS + ' CNPJs por envio.');

  const vistos = {};
  const empresas = b.empresas.map(function (x, i) {
    const cnpjNum = String(x.cnpj || '').replace(/\D/g, '');
    if (!cnpjValido_(cnpjNum)) throw new Error('CNPJ inválido no bloco ' + (i + 1) + '.');
    if (vistos[cnpjNum]) throw new Error('CNPJ repetido no bloco ' + (i + 1) + '.');
    vistos[cnpjNum] = true;
    const razao = txt(x.razaoSocial, 200);
    if (razao.length < 2) throw new Error('Razão social obrigatória no bloco ' + (i + 1) + '.');
    return {
      cnpj: cnpjNum.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5'),
      tipo: txt(x.tipo, 40),
      razaoSocial: razao,
      atividade: txt(x.atividade, 250),
      municipioUf: txt(x.municipioUf, 120),
      regime: txt(x.regime, 60),
      faturamentoMensal: int(x.faturamentoMensal, 0, 1e12),
      notasMes: int(x.notasMes, 0, 1e7),
      socios: int(x.socios, 1, 10000),
      funcionarios: int(x.funcionarios, 0, 1e6),
      autonomos: int(x.autonomos, 0, 1e6),
      complexidades: (Array.isArray(x.complexidades) ? x.complexidades : []).slice(0, 20).map(c => txt(c, 120))
    };
  });

  return { id: id, responsavel: resp, empresas: empresas, observacoes: txt(b.observacoes, 3000) };
}

function cnpjValido_(c) {
  if (c.length !== 14 || /^(\d)\1+$/.test(c)) return false;
  function dv(base) {
    const pesos = base.length === 12 ? [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2] : [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
    let s = 0;
    for (let i = 0; i < base.length; i++) s += Number(base.charAt(i)) * pesos[i];
    const r = s % 11;
    return r < 2 ? 0 : 11 - r;
  }
  return dv(c.slice(0, 12)) === Number(c.charAt(12)) && dv(c.slice(0, 13)) === Number(c.charAt(13));
}

/* ------------------------------ PLANILHA / DRIVE ------------------------------ */

function aba_(ss, nome, cabecalho) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
      .setFontWeight('bold').setBackground('#27306c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function protocoloExiste_(sh, id) {
  if (sh.getLastRow() < 2) return false;
  const achou = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return !!achou;
}

// Evita injeção de fórmula na planilha (valores que começam com = + - @)
function seguro_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function pasta_() {
  if (CONFIG.PASTA_ID) return DriveApp.getFolderById(CONFIG.PASTA_ID);
  const it = DriveApp.getFoldersByName(CONFIG.NOME_PASTA_PDFS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.NOME_PASTA_PDFS);
}

/* ------------------------------ PDF ------------------------------ */

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function brl_(n) {
  return 'R$ ' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',00';
}

function logoBase64_() {
  if (!CONFIG.LOGO_URL) return '';
  try {
    const blob = UrlFetchApp.fetch(CONFIG.LOGO_URL, { muteHttpExceptions: false }).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    console.error('Logo: ' + e);
    return '';
  }
}

function htmlPdf_(d, dataHora) {
  const logo = logoBase64_();
  const th = 'background:#27306c;color:#ffffff;text-align:left;padding:6px 8px;font-size:11px;';
  const tdL = 'padding:6px 8px;border-bottom:1px solid #dfe4ee;font-size:11px;color:#6b7280;';
  const tdV = 'padding:6px 8px;border-bottom:1px solid #dfe4ee;font-size:11px;font-weight:bold;color:#1e2233;';

  function linha(rot, val) {
    return '<tr><td width="38%" style="' + tdL + '">' + esc_(rot) + '</td><td style="' + tdV + '">' + esc_((val === 0 || (val !== '' && val != null)) ? val : '—') + '</td></tr>';
  }

  let h = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,Helvetica,sans-serif;color:#1e2233;">';
  h += '<table width="100%" style="width:100%;border-bottom:3px solid #1e7ebe;margin-bottom:14px;"><tr>' +
    '<td style="width:35%;padding-bottom:8px;">' + (logo ? '<img src="' + logo + '" width="80" height="42">' :
      '<b style="font-size:16px;color:#27306c;">Empresarial Assessoria Contábil</b>') + '</td>' +
    '<td style="text-align:right;padding-bottom:8px;"><div style="font-size:9px;color:#6b7280;letter-spacing:1px;">FORMULÁRIO DE INFORMAÇÕES PARA PROPOSTA</div>' +
    '<div style="font-size:15px;font-weight:bold;color:#1e7ebe;">Protocolo ' + esc_(d.id) + '</div>' +
    '<div style="font-size:10px;color:#6b7280;">Recebido em ' + esc_(dataHora) + '</div></td></tr></table>';

  h += '<table width="100%" style="width:100%;border-collapse:collapse;margin-bottom:14px;"><tr><th colspan="2" style="' + th + '">RESPONSÁVEL PELO PREENCHIMENTO</th></tr>' +
    linha('Nome', d.responsavel.nome) + linha('E-mail', d.responsavel.email) + linha('Telefone / WhatsApp', d.responsavel.telefone) + '</table>';

  d.empresas.forEach(function (x, i) {
    h += '<table width="100%" style="width:100%;border-collapse:collapse;margin-bottom:14px;page-break-inside:avoid;">' +
      '<tr><th colspan="2" style="' + th + '">CNPJ ' + (i + 1) + ' DE ' + d.empresas.length + ' — ' + esc_(x.tipo.toUpperCase()) + '</th></tr>' +
      linha('CNPJ', x.cnpj) + linha('Razão social', x.razaoSocial) + linha('Atividade principal', x.atividade) +
      linha('Município/UF', x.municipioUf) + linha('Regime tributário', x.regime) +
      linha('Faturamento médio mensal', brl_(x.faturamentoMensal)) +
      linha('Notas fiscais por mês', x.notasMes) + linha('Nº de sócios', x.socios) +
      linha('Funcionários ativos', x.funcionarios) + linha('Recibos de autônomo por mês', x.autonomos) +
      linha('Particularidades', x.complexidades.length ? x.complexidades.join('; ') : 'Nenhuma informada') + '</table>';
  });

  if (d.observacoes) {
    h += '<table width="100%" style="width:100%;border-collapse:collapse;"><tr><th style="' + th + '">OBSERVAÇÕES</th></tr>' +
      '<tr><td style="padding:8px;font-size:11px;">' + esc_(d.observacoes).replace(/\n/g, '<br>') + '</td></tr></table>';
  }
  h += '</body></html>';
  return h;
}

function gerarPdf_(d, dataHora) {
  const nomeArq = 'Formulario_' + d.id + '_' + d.empresas[0].razaoSocial.replace(/[^A-Za-z0-9]+/g, '_').slice(0, 40) + '.pdf';
  return Utilities.newBlob(htmlPdf_(d, dataHora), 'text/html', 'form.html')
    .getAs('application/pdf').setName(nomeArq);
}

/* ------------------------------ E-MAIL ------------------------------ */

function enviarEmail_(d, pdf, linkPdf, dataHora, urlPlanilha) {
  const lista = d.empresas.map(function (x) {
    return '<li>' + esc_(x.cnpj) + ' — ' + esc_(x.razaoSocial) + ' (' + esc_(x.tipo) + ', ' + esc_(x.regime) + ')</li>';
  }).join('');
  const corpo =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1e2233;">' +
    '<p>Novo formulário de informações para proposta recebido.</p>' +
    '<p><b>Protocolo:</b> ' + esc_(d.id) + '<br><b>Recebido em:</b> ' + esc_(dataHora) + '</p>' +
    '<p><b>Responsável:</b> ' + esc_(d.responsavel.nome) + '<br><b>E-mail:</b> ' + esc_(d.responsavel.email) +
    '<br><b>Telefone:</b> ' + esc_(d.responsavel.telefone) + '</p>' +
    '<p><b>CNPJs informados (' + d.empresas.length + '):</b></p><ul>' + lista + '</ul>' +
    '<p>' + (pdf ? 'O PDF completo segue em anexo.' : '<b>Atenção: não foi possível gerar o PDF; consulte os dados na planilha.</b>') + (linkPdf ? ' <a href="' + esc_(linkPdf) + '">Abrir no Drive</a>.' : '') +
    ' <a href="' + esc_(urlPlanilha) + '">Abrir planilha de respostas</a>.</p></div>';

  const para = CONFIG.DESTINATARIOS.slice();
  if (CONFIG.COPIA_AO_RESPONSAVEL) para.push(d.responsavel.email);

  MailApp.sendEmail({
    to: para.join(','),
    subject: 'Formulário de precificação — ' + d.empresas[0].razaoSocial +
      (d.empresas.length > 1 ? ' (+' + (d.empresas.length - 1) + ' CNPJ)' : '') + ' — ' + d.id,
    htmlBody: corpo,
    attachments: pdf ? [pdf] : [],
    name: CONFIG.REMETENTE_NOME,
    replyTo: d.responsavel.email
  });
}

/* ------------------------------ UTIL ------------------------------ */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Execute esta função UMA vez no editor (botão ▶ Executar) para conceder as permissões
 * (Planilhas, Drive, Gmail e acesso externo para baixar o logo) antes de publicar o Web App.
 */
function autorizar() {
  SpreadsheetApp.getActive().getName();
  DriveApp.getRootFolder().getName();
  MailApp.getRemainingDailyQuota();
  if (CONFIG.LOGO_URL) UrlFetchApp.fetch(CONFIG.LOGO_URL);
  Logger.log('Permissões concedidas. Agora publique em Implantar > Nova implantação > App da Web.');
}
