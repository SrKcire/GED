// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Arquivamento.gs | Solicitação, aprovação e cancelamento de arquivamentos
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Salva os dados de arquivamento na aba "Arquivamento".
 *
 * Colunas:
 *   A CÓDIGO  B DATA  C RESPONSÁVEL  D EMAIL  E DEPARTAMENTO
 *   F CÓDIGO/ID  G DESCRIÇÃO/TIPO  H DATA DO DOCUMENTO
 *   I INFORMAÇÃO  J ÁREA  K TEMPORALIDADE  L STATUS
 *   M CAIXA  N MOTIVO DO CANCELAMENTO  O DATA DE VENCIMENTO
 */
function salvarDados(dados) {
  // SEGURANÇA: nunca confiar em responsavel/departamento enviados pelo
  // cliente para identificar quem está criando o registro. O e-mail é
  // conferido contra a aba "Usuários" e os campos de identidade são
  // derivados no servidor — evita que um usuário logado personifique
  // outra pessoa/departamento via chamada direta de google.script.run.
  const infoU = _infoUsuario_(dados && dados.email);
  if (!infoU.encontrado || !infoU.ativo) {
    return { sucesso: false, erro: 'Usuário não identificado ou inativo.' };
  }
  dados.responsavel = infoU.nome;
  // GED pode registrar arquivamento em nome de qualquer departamento
  // (não há formulário dedicado por departamento); demais perfis só
  // podem gravar no próprio departamento.
  if (infoU.perfil !== 'GED') dados.departamento = infoU.departamento;

  const planilha = SpreadsheetApp.openById(SHEET_ID);
  let sheet = planilha.getSheetByName("Arquivamento");
  if (!sheet) sheet = planilha.insertSheet("Arquivamento");

  const cabecalho = sheet.getRange(1, 1, 1, 15).getValues()[0];
  const vazio = cabecalho.every(c => c === '' || c === null);

  if (vazio) {
    sheet.getRange(1, 1, 1, 15).setValues([COLUNAS_ARQUIVAMENTO]);
    const header = sheet.getRange(1, 1, 1, 15);
    header.setBackground('#2563EB');
    header.setFontColor('#FFFFFF');
    header.setFontWeight('bold');
  }

  const prefArq      = (dados.tipo === '1') ? 'AM' : 'AG';
  const codigo       = gerarCodigoUnico('Arquivamento', prefArq, dados.departamento || '');
  const proximaLinha = Math.max(sheet.getLastRow(), 1) + 1;
  const agora        = new Date();
  let linhas         = [];

  if (dados.tipo === '1') {
    linhas = dados.documentos.map(doc => [
      codigo, agora, dados.responsavel || '', dados.email || '', dados.departamento || '',
      doc.id || '', doc.tipoSolicitacao || '', doc.dataDoc || '', doc.setor || '',
      doc.area || '', doc.temporalidade || '', '', '', '', '',
    ]);
  } else if (dados.tipo === '3') {
    linhas = dados.documentos.map(doc => [
      codigo, agora, dados.responsavel || '', dados.email || '', dados.departamento || '',
      '', doc.nome || '', doc.dataDoc || '', doc.projeto || '',
      doc.area || '', doc.temporalidade || '', '', '', '', doc.dataVencimento || '',
    ]);
  } else if (dados.tipo === '6') {
    linhas = dados.documentos.map(doc => [
      codigo, agora, dados.responsavel || '', dados.email || '', dados.departamento || '',
      '', doc.nome || '', doc.dataDoc || '', doc.projeto || '',
      doc.area || '', doc.temporalidade || '', '', '', '', '',
    ]);
  } else if (dados.tipo === '7') {
    linhas = dados.documentos.map(doc => [
      codigo, agora, dados.responsavel || '', dados.email || '', dados.departamento || '',
      '', doc.nome || '', doc.dataDoc || '', '',
      doc.area || '', doc.temporalidade || '', '', '', '', '',
    ]);
  } else {
    linhas = dados.documentos.map(doc => [
      codigo, agora, dados.responsavel || '', dados.email || '', dados.departamento || '',
      '', doc.nome || '', doc.dataDoc || '', doc.projeto || '',
      doc.area || '', doc.temporalidade || '', '', '', '', '',
    ]);
  }

  if (linhas.length > 0) {
    sheet.getRange(proximaLinha, 1, linhas.length, 15).setValues(linhas);
  }

  return { sucesso: true, codigo };
}


// ───────────────────────────────────────────────────────────────────────────
//  ADM — LISTAGEM E GERÊNCIA DE LOTES
// ───────────────────────────────────────────────────────────────────────────

function getArquivamentos(aba, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return Object.assign({ lotes: [] }, _respostaAcessoNegado_());
  try {
    aba = aba || 'pendentes';
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Arquivamento');
    if (!sheet || sheet.getLastRow() < 2) return { lotes: [] };

    if (sheet.getMaxColumns() < 13) sheet.insertColumnsAfter(sheet.getMaxColumns(), 13 - sheet.getMaxColumns());

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 13).getValues();
    const tz    = Session.getScriptTimeZone();

    const lotesMap = {};
    dados.forEach((row, idx) => {
      const cod = String(row[0]).trim();
      if (!cod) return;

      const status      = String(row[11]).trim() || 'PENDENTE';
      const isHistorico = (status === 'APROVADO' || status === 'CANCELADO');

      if (aba === 'pendentes' && isHistorico) return;
      if (aba === 'historico'  && !isHistorico) return;

      if (!lotesMap[cod]) {
        lotesMap[cod] = {
          codigo:       cod,
          tipo:         cod.startsWith('AM.') ? 'solicitacao' : 'geral',
          data:         row[1] ? Utilities.formatDate(new Date(row[1]), tz, 'dd/MM/yyyy HH:mm') : '',
          responsavel:  String(row[2]).trim(),
          email:        String(row[3]).trim(),
          departamento: String(row[4]).trim(),
          status,
          documentos:   [],
          linhas:       []
        };
      }

      const dataDocRaw      = row[7];
      const dataDocFormatada = dataDocRaw
        ? (dataDocRaw instanceof Date
            ? Utilities.formatDate(dataDocRaw, tz, 'dd/MM/yyyy')
            : String(dataDocRaw).trim())
        : '';

      const codigoIdRaw = row[5];
      const codigoIdStr = (codigoIdRaw !== '' && codigoIdRaw !== null && codigoIdRaw !== undefined)
        ? String(codigoIdRaw).trim()
        : '';

      lotesMap[cod].documentos.push({
        linha:         idx + 2,
        codigoId:      codigoIdStr,
        descricao:     String(row[6]).trim(),
        dataDoc:       dataDocFormatada,
        informacao:    String(row[8]).trim(),
        area:          String(row[9]).trim(),
        temporalidade: String(row[10]).trim(),
        caixa:         String(row[12]).trim(),
      });
      lotesMap[cod].linhas.push(idx + 2);
    });

    const lotes = Object.values(lotesMap).sort((a, b) => {
      const parse = s => {
        if (!s) return 0;
        const [d, h] = s.split(' ');
        const [dd, mm, yy] = d.split('/');
        return new Date(yy+'-'+mm+'-'+dd+'T'+(h||'00:00')+':00').getTime();
      };
      return parse(b.data) - parse(a.data);
    });

    return { lotes };
  } catch(e) { return { lotes: [], erro: e.message }; }
}

function excluirDocumentoArquivamento(linha, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Arquivamento');
    if (!sheet) return { ok: false, erro: 'Aba não encontrada' };
    sheet.deleteRow(linha);
    return { ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}

/**
 * Aprova um lote de arquivamento e grava nos índices Mock.
 *
 * Mapeamento Documentos Gerais (AG.*) → Mock.Infra:
 *   col G (descricao)  → col A (PROJETO)
 *   col I (informacao) → col B (DESCRIÇÃO)
 *   col H (dataDoc)    → col C (DATA DO DOCUMENTO)
 *   col J (area)       → col D (ÁREA)
 *   col K (temporalidade) → col E (TEMPORALIDADE)
 *   col M (caixa)      → col F (CAIXA)
 *
 * Mapeamento Solicitação de Material (AM.*) → Mock.Infra.S.M:
 *   col I (informacao) → col A (SETOR)
 *   col H (dataDoc)    → col B (DATA DO DOCUMENTO)
 *   col G (descricao)  → col C (TIPO DE SOLICITAÇÃO)
 *   col F (codigoId)   → col D (ID DA SOLICITAÇÃO)
 *   col K (temporalidade) → col E (TEMPORALIDADE)
 *   col M (caixa)      → col F (CAIXA)
 */
function aprovarArquivamento(codigo, documentosComCaixa, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const ss        = SpreadsheetApp.openById(SHEET_ID);
    const sheetArq  = ss.getSheetByName('Arquivamento');
    const sheetMock = ss.getSheetByName('Mock.Infra');
    const sheetSM   = ss.getSheetByName('Mock.Infra.S.M');

    if (!sheetArq || !sheetMock || !sheetSM)
      return { ok: false, erro: 'Aba não encontrada' };

    while (sheetArq.getMaxColumns() < 13) sheetArq.insertColumnAfter(sheetArq.getMaxColumns());

    if (sheetMock.getLastRow() === 0)
      sheetMock.appendRow(['PROJETO', 'DESCRIÇÃO', 'DATA DO DOCUMENTO', 'ÁREA', 'TEMPORALIDADE', 'CAIXA']);
    if (sheetSM.getLastRow() === 0)
      sheetSM.appendRow(['SETOR', 'DATA DO DOCUMENTO', 'TIPO DE SOLICITAÇÃO', 'ID DA SOLICITAÇÃO', 'TEMPORALIDADE', 'CAIXA']);

    const isSolicitacao = codigo.startsWith('AM.');

    for (const doc of documentosComCaixa) {
      if (!doc.caixa || !String(doc.caixa).trim())
        return { ok: false, erro: 'Preencha o número da caixa para todos os documentos antes de aprovar.' };
      if (!isSolicitacao && (!doc.area || !String(doc.area).trim()))
        return { ok: false, erro: 'Preencha a Área para todos os documentos antes de aprovar.' };
      if (!doc.temporalidade || !String(doc.temporalidade).trim())
        return { ok: false, erro: 'Preencha a Temporalidade para todos os documentos antes de aprovar.' };
    }

    const dadosArq = sheetArq.getRange(2, 1, sheetArq.getLastRow()-1, 13).getValues();

    dadosArq.forEach((row, idx) => {
      if (String(row[0]).trim() !== codigo) return;
      const linhaReal = idx + 2;
      const docMatch  = documentosComCaixa.find(d => d.linha === linhaReal);
      sheetArq.getRange(linhaReal, 12).setValue('APROVADO');
      if (docMatch) {
        sheetArq.getRange(linhaReal, 10).setValue(docMatch.area || '');
        sheetArq.getRange(linhaReal, 11).setValue(docMatch.temporalidade || '');
        sheetArq.getRange(linhaReal, 13).setValue(docMatch.caixa);
      }
    });

    documentosComCaixa.forEach(doc => {
      if (isSolicitacao) {
        sheetSM.appendRow([
          doc.informacao    || '',
          doc.dataDoc       || '',
          doc.descricao     || '',
          doc.codigoId      || '',
          doc.temporalidade || '',
          doc.caixa         || ''
        ]);
      } else {
        sheetMock.appendRow([
          doc.informacao    || '',
          doc.descricao     || '',
          doc.dataDoc       || '',
          doc.area          || '',
          doc.temporalidade || '',
          doc.caixa         || ''
        ]);
      }
    });

    const primeiraLinha = dadosArq.find(r => String(r[0]).trim() === codigo);
    const responsavel   = primeiraLinha ? String(primeiraLinha[2]).trim() : '';
    const emailDest     = primeiraLinha ? String(primeiraLinha[3]).trim() : '';
    const departamento  = primeiraLinha ? String(primeiraLinha[4]).trim() : '';

    if (emailDest) {
      enviarEmailAprovacaoArquivamento(codigo, {
        responsavel, email: emailDest, departamento,
        documentos: documentosComCaixa,
        isSolicitacao
      });
    }

    return { ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}

function cancelarArquivamento(codigo, motivoCancelamento, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Arquivamento');
    if (!sheet || sheet.getLastRow() < 2) return { ok: true };

    while (sheet.getMaxColumns() < 14) sheet.insertColumnAfter(sheet.getMaxColumns());

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 14).getValues();
    let responsavel = '', email = '', departamento = '';

    dados.forEach((row, idx) => {
      if (String(row[0]).trim() !== codigo) return;
      const linhaReal = idx + 2;
      sheet.getRange(linhaReal, 12).setValue('CANCELADO');
      if (motivoCancelamento) sheet.getRange(linhaReal, 14).setValue(motivoCancelamento);
      if (!responsavel) {
        responsavel  = String(row[2]).trim();
        email        = String(row[3]).trim();
        departamento = String(row[4]).trim();
      }
    });

    if (!email) {
      const prim = dados.find(r => String(r[0]).trim() === codigo);
      if (prim) {
        responsavel  = String(prim[2]).trim();
        email        = String(prim[3]).trim();
        departamento = String(prim[4]).trim();
      }
    }

    if (email) {
      enviarEmailCancelamentoArquivamento(codigo, { responsavel, email, departamento, motivoCancelamento });
    }

    return { ok: true };
  } catch(e) { return { ok: false, erro: e.message }; }
}
