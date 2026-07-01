// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Retirada.gs | Solicitação, liberação, adiamento e devolução de documentos
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Busca documentos disponíveis para retirada.
 * Para perfil não-GED, filtra pelo departamento do usuário logado.
 * Marca como "em retirada" os que já possuem solicitação aberta.
 */
function buscarDocumentosRetirada(departamento, incluirMock, emailSolicitante) {
  try {
    // SEGURANÇA: o departamento/flag de mock enviados pelo cliente só são
    // confiáveis para o admin GED. Para os demais perfis, ignora o valor
    // recebido e usa o departamento real do usuário (evita que alguém
    // chame esta função diretamente pedindo o acervo de outro departamento).
    const infoU = _infoUsuario_(emailSolicitante);
    if (!infoU.encontrado || !infoU.ativo) return { documentos: [], erro: 'Usuário não identificado ou inativo.' };
    if (infoU.perfil !== 'GED') {
      departamento = infoU.departamento;
      incluirMock  = infoU.departamento === 'Infraestrutura';
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);

    const abaRetiradas = ss.getSheetByName('Retiradas');
    const bloqueados   = new Set();

    if (abaRetiradas && abaRetiradas.getLastRow() > 1) {
      const linhasRet = abaRetiradas
        .getRange(2, 1, abaRetiradas.getLastRow() - 1, 14)
        .getValues();

      linhasRet.forEach(row => {
        const statusRetirada  = String(row[9]  || '').trim().toUpperCase();
        const statusDevolucao = String(row[11] || '').trim().toUpperCase();

        const devBloqueado = statusDevolucao === 'PENDENTE' || statusDevolucao === 'ADIADO';
        const retAberta    = (statusRetirada === 'PENDENTE' || statusRetirada === 'LIBERADO') && statusDevolucao === '';

        if (!devBloqueado && !retAberta) return;

        const projeto   = String(row[5] || '').trim().toUpperCase();
        const descricao = String(row[6] || '').trim().toUpperCase();
        const caixa     = String(row[7] || '').trim().toUpperCase();
        bloqueados.add(`${descricao}||${projeto}||${caixa}`);
      });
    }

    const isGED = !departamento || departamento === 'GED';
    let dadosAcervo = [];

    if (!isGED) {
      const nomeAba = DEPT_ABA_MAP[departamento];
      if (nomeAba) {
        const abaAcervo = ss.getSheetByName(nomeAba);
        if (abaAcervo && abaAcervo.getLastRow() > 1)
          dadosAcervo = abaAcervo.getRange(2, 1, abaAcervo.getLastRow() - 1, 5).getValues();
      }
    } else {
      Object.values(DEPT_ABA_MAP).forEach(nomeAba => {
        const aba = ss.getSheetByName(nomeAba);
        if (aba && aba.getLastRow() > 1)
          dadosAcervo = dadosAcervo.concat(aba.getRange(2, 1, aba.getLastRow() - 1, 5).getValues());
      });
    }

    let dadosMock = [];
    if (incluirMock) {
      const abaMock = ss.getSheetByName('Mock.Infra');
      if (abaMock && abaMock.getLastRow() > 1)
        dadosMock = abaMock.getRange(2, 1, abaMock.getLastRow() - 1, 5).getValues();
    }

    let dadosSM = [];
    if (incluirMock) {
      const abaSM = ss.getSheetByName('Mock.Infra.S.M');
      if (abaSM && abaSM.getLastRow() > 1) {
        const maxCols = abaSM.getLastColumn();
        const rawSM   = abaSM.getRange(2, 1, abaSM.getLastRow() - 1, maxCols).getValues();
        dadosSM = rawSM.map(row => {
          const setor     = String(row[0] || '').trim();
          const tipoSol   = String(row[2] || '').trim();
          const idSol     = String(row[3] || '').trim();
          const caixa     = maxCols >= 6 ? String(row[5] || '').trim() : '';
          const descricao = tipoSol + (idSol ? ' Nº ' + idSol : '');
          return [setor, descricao, caixa, '', caixa];
        });
      }
    }

    const vistos      = new Set();
    const todasLinhas = [...dadosAcervo, ...dadosMock, ...dadosSM].filter(row => {
      if (!row[0] && !row[1]) return false;
      const desc  = String(row[1] || '').trim().toUpperCase();
      const proj  = String(row[0] || '').trim().toUpperCase();
      const cx    = String(row[2] || row[4] || '').trim().toUpperCase();
      const chave = `${desc}||${proj}||${cx}`;
      if (vistos.has(chave)) return false;
      vistos.add(chave);
      return true;
    });

    const documentos = todasLinhas.map(row => {
      const projeto   = String(row[0] || '').trim();
      const descricao = String(row[1] || '').trim();
      const caixa     = String(row[2] || row[4] || '').trim();
      const chave     = `${descricao.toUpperCase()}||${projeto.toUpperCase()}||${caixa.toUpperCase()}`;
      return { projeto, descricao, caixa, emRetirada: bloqueados.has(chave) };
    });

    return { documentos };
  } catch (e) {
    return { erro: e.message };
  }
}

/**
 * Salva uma solicitação de retirada na aba "Retiradas".
 */
function salvarRetirada(dados) {
  try {
    // SEGURANÇA: solicitação de retirada é sempre autoatribuída — deriva
    // responsavel/departamento do e-mail autenticado em vez de confiar
    // nos campos enviados pelo cliente (evita personificação).
    const infoU = _infoUsuario_(dados && dados.email);
    if (!infoU.encontrado || !infoU.ativo) {
      return { sucesso: false, erro: 'Usuário não identificado ou inativo.' };
    }
    dados.responsavel  = infoU.nome;
    dados.departamento = infoU.departamento;

    const planilha = SpreadsheetApp.openById(SHEET_ID);
    let sheet = planilha.getSheetByName("Retiradas");
    if (!sheet) sheet = planilha.insertSheet("Retiradas");

    const cabecalho = sheet.getRange(1, 1, 1, 15).getValues()[0];
    const vazio = cabecalho.every(c => c === '' || c === null);

    if (vazio) {
      sheet.getRange(1, 1, 1, 15).setValues([COLUNAS_RETIRADAS]);
      const header = sheet.getRange(1, 1, 1, 15);
      header.setBackground('#7c3aed');
      header.setFontColor('#FFFFFF');
      header.setFontWeight('bold');
    } else if (cabecalho[13] === '' || cabecalho[13] === null || cabecalho[14] === '' || cabecalho[14] === null) {
      // Migração: aba já existia com apenas 13 colunas de cabeçalho, mas
      // atualizarStatusLote grava adiamento (14) e motivo/justificativa (15)
      // desde sempre — preenche os títulos que faltam sem tocar no resto.
      sheet.getRange(1, 14, 1, 2).setValues([[COLUNAS_RETIRADAS[13], COLUNAS_RETIRADAS[14]]]);
    }

    const codigo       = gerarCodigoUnico('Retiradas', 'S', dados.departamento || '');
    const proximaLinha = Math.max(sheet.getLastRow(), 1) + 1;
    const agora        = new Date();

    const linhas = dados.documentos.map(doc => [
      codigo, agora, dados.responsavel || '', dados.email || '', dados.departamento || '',
      doc.projeto || '', doc.descricao || '', doc.caixa || '', dados.motivo || '',
      'PENDENTE', '', '', '', '', ''
    ]);

    if (linhas.length > 0) {
      sheet.getRange(proximaLinha, 1, linhas.length, 15).setValues(linhas);
    }

    return { sucesso: true, codigo };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  ADM — LISTAGEM E AÇÕES EM LOTES DE RETIRADA
// ───────────────────────────────────────────────────────────────────────────

function getRetiradas(setor, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return Object.assign({ lotes: [] }, _respostaAcessoNegado_());
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Retiradas");
    if (!sheet || sheet.getLastRow() < 2) return { lotes: [] };
    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 15).getValues();

    const lotesMap = {};
    dados.forEach((row, idx) => {
      const cod  = String(row[0]).trim();
      const dept = String(row[4]).trim();
      if (!cod) return;
      if (setor && setor !== 'todos' && dept !== setor) return;
      if (!lotesMap[cod]) {
        lotesMap[cod] = {
          codigo:             cod,
          data:               row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
          responsavel:        String(row[2]).trim(),
          email:              String(row[3]).trim(),
          departamento:       dept,
          motivo:             String(row[8]).trim(),
          statusRetirada:     String(row[9]).trim(),
          dataRetirada:       row[10] ? Utilities.formatDate(new Date(row[10]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          statusDevolucao:    String(row[11]).trim(),
          dataDevolucao:      row[12] ? Utilities.formatDate(new Date(row[12]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          motivoCancelamento: String(row[14] || '').trim(),
          documentos: [],
          linhas: []
        };
      }
      lotesMap[cod].documentos.push({
        projeto:   String(row[5]).trim(),
        descricao: String(row[6]).trim(),
        caixa:     String(row[7]).trim()
      });
      lotesMap[cod].linhas.push(idx + 2);
    });

    return { lotes: Object.values(lotesMap) };
  } catch(e) { return { lotes: [], erro: e.message }; }
}

function atualizarStatusLote(codigo, acao, dadosExtra, emailSolicitante) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Retiradas");
    if (!sheet || sheet.getLastRow() < 2) return { sucesso: false };

    while (sheet.getMaxColumns() < 15) sheet.insertColumnAfter(sheet.getMaxColumns());

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 15).getValues();
    const agora = new Date();

    // GUARD DE PERFIL:
    // LIBERAR / ADIAR / DEVOLVER são exclusivas do ADM (GED).
    // CANCELAR é permitida ao ADM (qualquer lote) ou ao próprio usuário
    // dono do lote (cancelamento da própria solicitação pendente).
    const ehGed = _ehGedAtivo_(emailSolicitante);
    if (acao !== 'CANCELAR' && !ehGed) {
      return _respostaAcessoNegado_();
    }
    if (acao === 'CANCELAR' && !ehGed) {
      const loteDoUsuario = dados.find(r => String(r[0]).trim() === codigo);
      const emailDoLote   = loteDoUsuario ? String(loteDoUsuario[3]).trim().toLowerCase() : '';
      if (!emailSolicitante || emailDoLote !== String(emailSolicitante).trim().toLowerCase()) {
        return _respostaAcessoNegado_();
      }
    }

    // Guard contra ações duplicadas (duplo-clique / requisições repetidas)
    if (acao === 'CANCELAR') {
      const loteRow = dados.find(r => String(r[0]).trim() === codigo);
      if (!loteRow) return { sucesso: false, erro: 'Lote não encontrado.' };
      const statusRetAtual = String(loteRow[9]).trim();
      if (statusRetAtual === 'CANCELADO') {
        return { sucesso: false, erro: 'Este lote já foi cancelado.' };
      }
    }

    if (acao === 'LIBERAR') {
      const loteRow = dados.find(r => String(r[0]).trim() === codigo);
      if (!loteRow) return { sucesso: false, erro: 'Lote não encontrado.' };
      const statusRetAtual = String(loteRow[9]).trim();
      if (statusRetAtual === 'LIBERADO' || statusRetAtual === 'CANCELADO') {
        return { sucesso: false, erro: 'Este lote já foi processado (status atual: ' + statusRetAtual + ').' };
      }
    }

    if (acao === 'DEVOLVER') {
      const loteRow = dados.find(r => String(r[0]).trim() === codigo);
      if (!loteRow) return { sucesso: false, erro: 'Lote não encontrado.' };
      const statusDevAtual = String(loteRow[11]).trim();
      if (statusDevAtual === 'DEVOLVIDO') {
        return { sucesso: false, erro: 'Este lote já foi devolvido.' };
      }
    }

    if (acao === 'ADIAR') {
      const loteRow = dados.find(r => String(r[0]).trim() === codigo);
      if (!loteRow) return { sucesso: false, erro: 'Lote não encontrado.' };

      const statusDev = String(loteRow[11]).trim();
      if (statusDev === 'ADIADO') {
        return { sucesso: false, erro: 'Este lote já foi adiado. Não é possível adiar novamente.' };
      }

      const dataEscolhida = new Date(dadosExtra.dataFinal + 'T00:00:00');
      const limiteMax = new Date(agora);
      limiteMax.setDate(limiteMax.getDate() + 10);
      limiteMax.setHours(23,59,59,999);

      if (dataEscolhida > limiteMax) {
        return { sucesso: false, erro: 'A data de adiamento não pode ultrapassar 10 dias corridos a partir de hoje.' };
      }

      const dataMin = new Date(agora);
      dataMin.setHours(0,0,0,0);
      if (dataEscolhida < dataMin) {
        return { sucesso: false, erro: 'A data de adiamento não pode ser no passado.' };
      }
    }

    dados.forEach((row, idx) => {
      if (String(row[0]).trim() !== codigo) return;
      const linha = idx + 2;

      if (acao === 'LIBERAR') {
        sheet.getRange(linha, 10).setValue('LIBERADO');
        sheet.getRange(linha, 11).setValue(agora);
        sheet.getRange(linha, 11).setNumberFormat('dd/MM/yyyy HH:mm');
        sheet.getRange(linha, 12).setValue('PENDENTE');
        dadosExtra.codigo = codigo;
      } else if (acao === 'CANCELAR') {
        sheet.getRange(linha, 10).setValue('CANCELADO');
        sheet.getRange(linha, 12).setValue('CANCELADO');
        if (dadosExtra.motivoCancelamento) {
          sheet.getRange(linha, 15).setValue(dadosExtra.motivoCancelamento);
        }
        dadosExtra.codigo = codigo;
      } else if (acao === 'ADIAR') {
        const dataAdiamento = new Date(dadosExtra.dataFinal + 'T00:00:00');
        sheet.getRange(linha, 12).setValue('ADIADO');
        sheet.getRange(linha, 14).setValue(dataAdiamento);
        sheet.getRange(linha, 14).setNumberFormat('dd/MM/yyyy');
        if (dadosExtra.justificativa) {
          sheet.getRange(linha, 15).setValue(dadosExtra.justificativa);
        }
        dadosExtra.codigo = codigo;
      } else if (acao === 'DEVOLVER') {
        sheet.getRange(linha, 12).setValue('DEVOLVIDO');
        sheet.getRange(linha, 13).setValue(agora);
        sheet.getRange(linha, 13).setNumberFormat('dd/MM/yyyy HH:mm');
      }
    });

    if (acao === 'LIBERAR' || acao === 'CANCELAR' || acao === 'ADIAR') {
      if (!dadosExtra.email || !dadosExtra.responsavel) {
        const dadosPlan = dados.find(r => String(r[0]).trim() === codigo);
        if (dadosPlan) {
          dadosExtra.email        = dadosExtra.email        || String(dadosPlan[3]).trim();
          dadosExtra.responsavel  = dadosExtra.responsavel  || String(dadosPlan[2]).trim();
          dadosExtra.departamento = dadosExtra.departamento || String(dadosPlan[4]).trim();
        }
      }
      enviarEmailRetirada(codigo, acao, dadosExtra);
    }

    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

/**
 * Trigger diário: verifica vencimentos de retiradas e envia cobranças.
 * Instalar como time-driven trigger (diário, ~10h).
 */
function verificarVencimentosRetiradas() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Retiradas");
    if (!sheet || sheet.getLastRow() < 2) return;

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 14).getValues();
    const hoje  = new Date();
    hoje.setHours(0,0,0,0);

    const lotesMap = {};
    dados.forEach(row => {
      const codigo = String(row[0]).trim();
      if (!codigo) return;
      if (!lotesMap[codigo]) {
        lotesMap[codigo] = {
          codigo,
          email:       String(row[3]).trim(),
          responsavel: String(row[2]).trim(),
          statusRet:   String(row[9]).trim(),
          statusDev:   String(row[11]).trim(),
          dataRet:     row[10],
          dataAdiam:   row[13],
          documentos:  []
        };
      }
      lotesMap[codigo].documentos.push({
        projeto:   String(row[5]).trim(),
        descricao: String(row[6]).trim(),
        caixa:     String(row[7]).trim()
      });
    });

    Object.values(lotesMap).forEach(lote => {
      const { codigo, email, responsavel, statusRet, statusDev, dataRet, dataAdiam } = lote;
      if (!codigo || !email) return;

      if (statusRet === 'LIBERADO' && statusDev === 'PENDENTE' && dataRet) {
        const diasUteis = calcularDiasUteis(new Date(dataRet), hoje);
        if (diasUteis >= 5) {
          enviarEmailCobranca5Dias(email, responsavel, codigo, dataRet, diasUteis, lote.documentos);
        }
      }

      if (statusRet === 'LIBERADO' && statusDev === 'ADIADO' && dataAdiam) {
        const dataVenc = new Date(dataAdiam);
        dataVenc.setHours(0,0,0,0);
        if (hoje >= dataVenc) {
          enviarEmailCobrancaAdiamentoVencido(email, responsavel, codigo, dataAdiam, lote.documentos);
        }
      }
    });

    Logger.log('verificarVencimentosRetiradas concluído — ' + new Date());
  } catch(e) { Logger.log('Erro verificarVencimentosRetiradas: ' + e.message); }
}
