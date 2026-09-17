/**
 * Valora Med Club — recebe as respostas das duas páginas e grava na planilha.
 *
 * Implantado como Web App (executar como: eu / acesso: qualquer pessoa).
 * As páginas fazem POST com um JSON no corpo:
 *  - LP principal (Próximo Passo Valora):
 *      { fonte, passo, nome, especialidade, atuacao, faturamento, assunto }
 *  - LP com vídeo (Valora Med Club / mentoria):
 *      { fonte, nome, crm, whatsapp, email, instagram, interesse }
 */

var ABA = 'Inscrições';
var CABECALHO = [
  'Data/hora',
  'Fonte',
  'Próximo passo escolhido',
  'Nome',
  'Especialidade',
  'Modelo de atuação hoje',
  'Faturamento médio mensal',
  'Assunto de maior interesse',
  'CRM',
  'WhatsApp',
  'E-mail',
  'Instagram',
  'Interesse'
];
// Cabeçalho da versão anterior (sem Fonte): usado só para migrar a aba sem desalinhar as linhas.
var CABECALHO_ANTIGO = [
  'Data/hora', 'Próximo passo escolhido', 'Nome', 'Especialidade',
  'Modelo de atuação hoje', 'Faturamento médio mensal', 'Assunto de maior interesse'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var dados = lerCorpo_(e);
    var aba = abaInscricoes_();
    var linha = [
      new Date(),
      limpa_(dados.fonte) || 'LP principal',
      limpa_(dados.passo),
      limpa_(dados.nome),
      limpa_(dados.especialidade),
      limpa_(dados.atuacao),
      limpa_(dados.faturamento),
      limpa_(dados.assunto),
      limpa_(dados.crm),
      limpa_(dados.whatsapp),
      limpa_(dados.email),
      limpa_(dados.instagram),
      limpa_(dados.interesse)
    ];
    aba.appendRow(linha);
    var r = aba.getLastRow();
    aba.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    aba.getRange(r, 10).setNumberFormat('@'); // WhatsApp sempre como texto
    return resposta_({ ok: true, linha: r });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta_({ ok: true, servico: 'Valora — respostas das páginas (LP principal e LP com vídeo)' });
}

/* ---------- apoio ---------- */

function lerCorpo_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('Corpo vazio.');
  var txt = e.postData.contents;
  try {
    return JSON.parse(txt);
  } catch (_) {
    // fallback: application/x-www-form-urlencoded
    return e.parameter || {};
  }
}

/** Garante a aba com o cabeçalho atual: cria a aba, migra da versão anterior (insere a coluna Fonte) ou reescreve a linha 1. */
function abaInscricoes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA);
  if (!aba) {
    aba = ss.insertSheet(ABA, 0);
  }
  var atual = aba.getLastRow() === 0 ? [] : aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), 1)).getValues()[0].map(String);
  var igual = function (cab) { return atual.length >= cab.length && cab.every(function (c, i) { return atual[i] === c; }); };
  if (igual(CABECALHO)) return aba;

  if (igual(CABECALHO_ANTIGO)) {
    // migração: a coluna Fonte entra na posição B e as linhas existentes continuam alinhadas
    aba.insertColumnBefore(2);
    aba.getRange(2, 2, Math.max(aba.getLastRow() - 1, 1), 1).setValue('LP principal');
  } else if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHO);
  }
  aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), CABECALHO.length)).clearContent();
  aba.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);

  var cab = aba.getRange(1, 1, 1, CABECALHO.length);
  cab.setFontWeight('bold').setBackground('#211F1E').setFontColor('#FFE2B9');
  aba.setFrozenRows(1);
  aba.setColumnWidths(1, CABECALHO.length, 160);
  aba.setColumnWidth(3, 360);  // Próximo passo escolhido
  aba.setColumnWidth(4, 240);  // Nome
  aba.setColumnWidth(8, 300);  // Assunto de maior interesse
  aba.setColumnWidth(13, 320); // Interesse
  return aba;
}

function limpa_(v) {
  return v == null ? '' : String(v).trim();
}

function resposta_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Roda uma vez no editor para autorizar e ajustar a aba com o cabeçalho novo. */
function preparar() {
  abaInscricoes_();
}

/** Teste manual no editor: grava uma linha de cada página. */
function testeLocal() {
  var a = { postData: { contents: JSON.stringify({
    fonte: 'LP principal',
    passo: 'Quero entender se a Mentoria Valora faz sentido para mim',
    nome: 'Teste Automático',
    especialidade: 'Cardiologia',
    atuacao: 'Misto (convênio e particular)',
    faturamento: 'De R$ 40 mil a R$ 80 mil',
    assunto: 'Precificação e equipe'
  }) } };
  var b = { postData: { contents: JSON.stringify({
    fonte: 'LP com vídeo',
    nome: 'Teste Automático',
    crm: 'CRM/SP 123456',
    whatsapp: '(12) 99999-9999',
    email: 'teste@exemplo.com',
    instagram: '@teste',
    interesse: 'Gestão de consultório, Técnicas de vendas'
  }) } };
  Logger.log(doPost(a).getContent());
  Logger.log(doPost(b).getContent());
}
