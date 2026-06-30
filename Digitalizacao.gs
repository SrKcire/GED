// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Digitalizacao.gs | Dossiês de alunos e controle de digitalização
// ═══════════════════════════════════════════════════════════════════════════

function buscarAluno(termo, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return Object.assign({ alunos: [] }, _respostaAcessoNegado_());
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Base.Alunos');
    if (!sheet || sheet.getLastRow() < 2) return { alunos: [] };

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    const t = String(termo).trim().toLowerCase();
    if (!t) return { alunos: [] };

    const resultado = [];
    dados.forEach((row, idx) => {
      const matricula = String(row[0]).trim();
      const dossie    = String(row[1]).trim();
      const nome      = String(row[2]).trim();
      const pai       = String(row[4]).trim();
      const mae       = String(row[5]).trim();
      if (!matricula && !nome) return;
      if (nome.toLowerCase().includes(t) || matricula.toLowerCase().includes(t)) {
        resultado.push({ linha: idx + 2, matricula, dossie, nome, pai, mae });
      }
    });

    return { alunos: resultado.slice(0, 10) };
  } catch(e) { return { alunos: [], erro: e.message }; }
}

function salvarDossie(linha, dossie, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    if (!linha || !dossie) return { ok: false, erro: 'Dados inválidos' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Base.Alunos');
    if (!sheet) return { ok: false, erro: 'Aba Base.Alunos não encontrada' };
    sheet.getRange(linha, 2).setValue(dossie);
    return { ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}

function registrarDigitalizacao(dados, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    if (!dados || !dados.matricula) return { ok: false, erro: 'Matrícula obrigatória' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DIGITALIZADOS');
    if (!sheet) return { ok: false, erro: 'Aba DIGITALIZADOS não encontrada' };

    const agora = new Date();
    const data  = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    sheet.appendRow([data, dados.matricula||'', dados.dossie||'', dados.nome||'', dados.mae||'', dados.pai||'']);
    return { ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}

function getDigitalizados(emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return Object.assign({ registros: [] }, _respostaAcessoNegado_());
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DIGITALIZADOS');
    if (!sheet || sheet.getLastRow() < 2) return { registros: [] };
    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    const registros = dados
      .filter(row => row[0] || row[1])
      .map(row => ({
        data:      String(row[0]).trim(),
        matricula: String(row[1]).trim(),
        dossie:    String(row[2]).trim(),
        nome:      String(row[3]).trim(),
        mae:       String(row[4]).trim(),
        pai:       String(row[5]).trim(),
      }))
      .reverse();
    return { registros };
  } catch(e) { return { registros: [], erro: e.message }; }
}

function getDashboardDigitalizacao(emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const ss       = SpreadsheetApp.openById(SHEET_ID);
    const sheetDig = ss.getSheetByName('DIGITALIZADOS');
    const sheetBase= ss.getSheetByName('Base.Alunos');
    const tz       = Session.getScriptTimeZone();
    const hoje     = new Date();
    hoje.setHours(0,0,0,0);

    let total = 0, esteMes = 0, hoje_ = 0;
    const porMes = {};

    if (sheetDig && sheetDig.getLastRow() > 1) {
      const dados = sheetDig.getRange(2, 1, sheetDig.getLastRow()-1, 1).getValues();
      dados.forEach(row => {
        if (!row[0]) return;
        total++;
        const d = new Date(row[0]);
        if (isNaN(d)) return;
        d.setHours(0,0,0,0);
        if (d.getTime() === hoje.getTime()) hoje_++;
        if (d.getMonth() === hoje.getMonth() && d.getFullYear() === hoje.getFullYear()) esteMes++;
        const key = Utilities.formatDate(d, tz, 'MM/yyyy');
        porMes[key] = (porMes[key] || 0) + 1;
      });
    }

    let dossiesSalvos = 0;
    if (sheetBase && sheetBase.getLastRow() > 1) {
      const dossies = sheetBase.getRange(2, 2, sheetBase.getLastRow()-1, 1).getValues();
      dossies.forEach(r => { if (String(r[0]).trim()) dossiesSalvos++; });
    }

    const meses = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
      const key   = Utilities.formatDate(d, tz, 'MM/yyyy');
      const label = Utilities.formatDate(d, tz, 'MMM yyyy');
      meses.push({ key, label, count: porMes[key] || 0 });
    }

    return { total, esteMes, hoje: hoje_, dossiesSalvos, meses, ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}
