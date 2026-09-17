/**
 * Valora Med Club — recebe as inscrições da página
 * "Inscrição para a Reunião Estratégica" e grava na planilha.
 *
 * Implantado como Web App (executar como: eu / acesso: qualquer pessoa).
 * A página faz POST com um JSON no corpo:
 * { nome, telefone, especialidade, faturamento, dia, opcao1, opcao2 }
 */

var ABA = 'Inscrições';
var CABECALHO = [
  'Data/hora',
  'Nome',
  'Telefone',
  'Especialidade',
  'Faturamento médio mensal',
  'Melhor dia',
  '1ª opção de horário',
  '2ª opção de horário'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var dados = lerCorpo_(e);
    var aba = abaInscricoes_();
    var linha = [
      new Date(),
      limpa_(dados.nome),
      limpa_(dados.telefone),
      limpa_(dados.especialidade),
      limpa_(dados.faturamento),
      limpa_(dados.dia),
      limpa_(dados.opcao1),
      limpa_(dados.opcao2)
    ];
    aba.appendRow(linha);
    var r = aba.getLastRow();
    aba.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    aba.getRange(r, 3).setNumberFormat('@'); // telefone sempre como texto
    return resposta_({ ok: true, linha: r });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta_({ ok: true, servico: 'Valora — inscrições da Reunião Estratégica' });
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

function abaInscricoes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA);
  if (!aba) {
    aba = ss.insertSheet(ABA, 0);
  }
  if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHO);
    var cab = aba.getRange(1, 1, 1, CABECALHO.length);
    cab.setFontWeight('bold').setBackground('#211F1E').setFontColor('#FFE2B9');
    aba.setFrozenRows(1);
    aba.setColumnWidths(1, CABECALHO.length, 180);
    aba.setColumnWidth(2, 260);
  }
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

/** Roda uma vez no editor para autorizar e criar a aba com cabeçalho. */
function preparar() {
  abaInscricoes_();
}

/** Teste manual no editor: grava uma linha de exemplo. */
function testeLocal() {
  var e = { postData: { contents: JSON.stringify({
    nome: 'Teste Automático',
    telefone: '(12) 9 9674-6674',
    especialidade: 'Cardiologia',
    faturamento: 'De R$ 40 mil a R$ 80 mil',
    dia: 'Segunda-feira',
    opcao1: 'Segunda-feira, 10h',
    opcao2: 'Terça-feira, 14h30'
  }) } };
  Logger.log(doPost(e).getContent());
}
