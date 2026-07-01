// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Descarte.gs | Solicitação e aprovação de descarte de documentos
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Cria aba "Descartes" se não existir, com cabeçalho padronizado.
 * @private
 */
function _garantirAbaDescartes_(ss) {
  let sheet = ss.getSheetByName('Descartes');
  if (!sheet) {
    sheet = ss.insertSheet('Descartes');
    sheet.getRange(1, 1, 1, 9).setValues([COLUNAS_DESCARTES]);
    const header = sheet.getRange(1, 1, 1, 9);
    header.setBackground('#dc2626');
    header.setFontColor('#FFFFFF');
    header.setFontWeight('bold');
  }
  return sheet;
}

/**
 * Registra uma solicitação de descarte de documentos.
 * Disponível para perfil Usuário (departamento).
 *
 * @param {Object} dados
 *   dados.responsavel  string
 *   dados.email        string
 *   dados.departamento string
 *   dados.documentos   string  (descrição livre dos documentos a descartar)
 *   dados.justificativa string
 */
function solicitarDescarte(dados) {
  try {
    // SEGURANÇA: solicitação de descarte é sempre autoatribuída — deriva
    // responsavel/departamento do e-mail autenticado em vez de confiar
    // nos campos enviados pelo cliente (evita personificação).
    const infoU = _infoUsuario_(dados && dados.email);
    if (!infoU.encontrado || !infoU.ativo) {
      return { sucesso: false, erro: 'Usuário não identificado ou inativo.' };
    }
    dados.responsavel  = infoU.nome;
    dados.departamento = infoU.departamento;

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = _garantirAbaDescartes_(ss);

    const codigo       = gerarCodigoUnico('Descartes', 'D', dados.departamento || '');
    const proximaLinha = Math.max(sheet.getLastRow(), 1) + 1;
    const agora        = new Date();

    sheet.getRange(proximaLinha, 1, 1, 9).setValues([[
      codigo,
      agora,
      dados.responsavel   || '',
      dados.email         || '',
      dados.departamento  || '',
      dados.documentos    || '',
      dados.justificativa || '',
      'PENDENTE',
      ''
    ]]);

    return { sucesso: true, codigo };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  ADM — LISTAGEM E DECISÃO
// ───────────────────────────────────────────────────────────────────────────

/**
 * Lista solicitações de descarte.
 * @param {'pendentes'|'todos'} filtro
 */
function getDescartes(filtro, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return Object.assign({ descartes: [] }, _respostaAcessoNegado_());
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = _garantirAbaDescartes_(ss);
    if (sheet.getLastRow() < 2) return { descartes: [] };

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues();
    const tz    = Session.getScriptTimeZone();

    const descartes = [];
    dados.forEach((row, idx) => {
      if (!String(row[0]).trim()) return;
      if (filtro === 'pendentes' && String(row[7]).trim() !== 'PENDENTE') return;
      descartes.push({
        linha:         idx + 2,
        codigo:        String(row[0]).trim(),
        data:          row[1] ? Utilities.formatDate(new Date(row[1]), tz, 'dd/MM/yyyy HH:mm') : '',
        responsavel:   String(row[2]).trim(),
        email:         String(row[3]).trim(),
        departamento:  String(row[4]).trim(),
        documentos:    String(row[5]).trim(),
        justificativa: String(row[6]).trim(),
        status:        String(row[7]).trim(),
        observacaoAdm: String(row[8]).trim(),
      });
    });

    return { descartes };
  } catch(e) { return { descartes: [], erro: e.message }; }
}

/**
 * Aprova ou recusa uma solicitação de descarte.
 * @param {string} codigo
 * @param {'APROVADO'|'RECUSADO'} decisao
 * @param {string} observacao  motivo da recusa ou observação da aprovação
 */
function aprovarDescarte(codigo, decisao, observacao, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = _garantirAbaDescartes_(ss);
    if (sheet.getLastRow() < 2) return { ok: false, erro: 'Nenhuma solicitação encontrada.' };

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues();
    let encontrou = false;

    dados.forEach((row, idx) => {
      if (String(row[0]).trim() !== codigo) return;
      const linhaReal = idx + 2;
      sheet.getRange(linhaReal, 8).setValue(decisao);
      if (observacao) sheet.getRange(linhaReal, 9).setValue(observacao);
      encontrou = true;
    });

    if (!encontrou) return { ok: false, erro: 'Código não encontrado.' };
    return { ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}

/** Recusa uma solicitação de descarte (alias semântico). */
function recusarDescarte(codigo, motivo, emailSolicitante) {
  return aprovarDescarte(codigo, 'RECUSADO', motivo, emailSolicitante);
}
