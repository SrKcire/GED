// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Consulta.gs | Leitura e busca do acervo (sem escrita de estado)
// ═══════════════════════════════════════════════════════════════════════════

// ───────────────────────────────────────────────────────────────────────────
//  VERIFICAÇÃO DE DUPLICATAS
// ───────────────────────────────────────────────────────────────────────────

/**
 * Verifica se os documentos informados já existem nas abas Mock.
 * Tipo '1' → compara por ID + tipo de solicitação (Mock.Infra.S.M)
 * Demais   → compara por nome + projeto (Mock.Infra)
 */
function verificarDuplicatasMock(tipo, documentos) {
  try {
    const ss         = SpreadsheetApp.openById(SHEET_ID);
    const duplicatas = [];

    if (tipo === '1') {
      const sheet = ss.getSheetByName('Mock.Infra.S.M');
      if (!sheet || sheet.getLastRow() < 2) return { duplicatas: [] };

      const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

      documentos.forEach((doc, idx) => {
        const idDoc   = String(doc.id || '').trim();
        const tipoSol = String(doc.tipoSolicitacao || '').trim().toUpperCase();
        if (!idDoc) return;

        const jaExiste = dados.some(row => {
          const idMock   = String(row[3] || '').trim();
          const tipoMock = String(row[2] || '').trim().toUpperCase();
          return idMock === idDoc && tipoMock === tipoSol;
        });

        if (jaExiste) {
          duplicatas.push({
            linha: idx + 1,
            mensagem: `Linha ${idx + 1}: ID "${idDoc}" com tipo "${tipoSol}" já existe no acervo (Mock.Infra.S.M).`
          });
        }
      });

    } else {
      const sheet = ss.getSheetByName('Mock.Infra');
      if (!sheet || sheet.getLastRow() < 2) return { duplicatas: [] };

      const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

      documentos.forEach((doc, idx) => {
        const nome    = String(doc.nome    || '').trim().toUpperCase();
        const projeto = String(doc.projeto || '').trim().toUpperCase();
        if (!nome) return;

        const jaExiste = dados.some(row => {
          const descMock = String(row[1] || '').trim().toUpperCase();
          const projMock = String(row[0] || '').trim().toUpperCase();
          const projetoOk = !projeto || projMock === projeto;
          return descMock === nome && projetoOk;
        });

        if (jaExiste) {
          duplicatas.push({
            linha: idx + 1,
            mensagem: `Linha ${idx + 1}: Documento "${nome}"${projeto ? ' / Projeto "' + projeto + '"' : ''} já existe no acervo (Mock.Infra).`
          });
        }
      });
    }

    return { duplicatas };
  } catch (e) {
    return { duplicatas: [], erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  BUSCA NO ACERVO DIGITAL (Índice Drive)
// ───────────────────────────────────────────────────────────────────────────

/**
 * Busca documentos na aba "Índice Drive" por termo(s) no nome.
 * Suporta múltiplas palavras e normalização NFD (acentos).
 * Para perfil GED (departamento = 'GED' ou vazio), retorna todos.
 * Para demais perfis, o filtro de departamento é aplicado quando possível.
 */
function buscarDocumentos(termo, pagina, departamento) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Índice Drive');
    if (!sheet) return { resultados: [], total: 0, erro: 'Índice não encontrado.' };

    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return { resultados: [], total: 0 };

    const dados = sheet.getRange(2, 1, ultimaLinha - 1, 4).getValues();

    const termoNorm = normalizar(termo);
    const palavras  = termoNorm.split(/\s+/).filter(Boolean);

    const isGED = !departamento || departamento === 'GED';

    const filtrados = dados.filter(row => {
      const nome = normalizar(row[0]);
      const termoOk = palavras.length === 0 || palavras.every(p => nome.includes(p));
      if (!termoOk) return false;
      if (!isGED && departamento) {
        const pasta = normalizar(row[3]);
        if (pasta && !pasta.includes(normalizar(departamento))) return false;
      }
      return true;
    });

    const total = filtrados.length;
    const resultados = filtrados.map(row => ({
      nome:  String(row[0]),
      url:   String(row[1]),
      data:  String(row[2]),
      pasta: String(row[3])
    }));

    return { resultados, total };
  } catch (e) {
    return { resultados: [], total: 0, erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  BUSCA NO ACERVO FÍSICO (para Consulta)
// ───────────────────────────────────────────────────────────────────────────

/**
 * Busca todo o acervo físico para fins de CONSULTA (não confundir com
 * buscarDocumentosRetirada, que exclui documentos já bloqueados).
 */
function buscarAcervoFisicoConsulta(departamento, incluirMock) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
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

    const documentos = todasLinhas.map(row => ({
      projeto:   String(row[0] || '').trim(),
      descricao: String(row[1] || '').trim(),
      caixa:     String(row[2] || row[4] || '').trim()
    }));

    return { documentos };
  } catch (e) {
    return { documentos: [], erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  HISTÓRICO DO USUÁRIO
// ───────────────────────────────────────────────────────────────────────────

/**
 * Busca registros do histórico do usuário.
 * @param {'retirada'|'arquivamento'|'descarte'} tipo
 * @param {string} emailUsuario  — filtra pelo e-mail; vazio retorna todos (ADM)
 */
function buscarHistorico(tipo, emailUsuario) {
  try {
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const nomeAba = tipo === 'retirada'
      ? 'Retiradas'
      : tipo === 'descarte'
        ? 'Descartes'
        : 'Arquivamento';

    const sheet = ss.getSheetByName(nomeAba);

    if (!sheet || sheet.getLastRow() < 2)
      return { registros: [], total: 0, cabecalhos: [] };

    const ultimaCol  = sheet.getLastColumn();
    const dados      = sheet.getRange(2, 1, sheet.getLastRow() - 1, ultimaCol).getValues();
    const cabecalhos = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(String);

    const filtrados = emailUsuario
      ? dados.filter(row => String(row[3] || '').trim().toLowerCase() === emailUsuario.toLowerCase())
      : dados;

    filtrados.sort((a, b) => {
      const da = a[1] instanceof Date ? a[1].getTime() : 0;
      const db = b[1] instanceof Date ? b[1].getTime() : 0;
      return db - da;
    });

    const tz = Session.getScriptTimeZone();
    const registros = filtrados.map(row =>
      Object.fromEntries(cabecalhos.map((cab, i) => {
        let val = row[i];
        if (val instanceof Date)
          val = Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm');
        return [cab, String(val == null ? '' : val)];
      }))
    );

    return { registros, total: registros.length, cabecalhos };
  } catch (e) {
    return { registros: [], total: 0, cabecalhos: [], erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  ETIQUETAS
// ───────────────────────────────────────────────────────────────────────────

/**
 * Busca os dados de um arquivamento pelo código para reimprimir etiquetas.
 */
function buscarDadosEtiqueta(codigo) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Arquivamento');
    if (!sheet || sheet.getLastRow() < 2)
      return { erro: 'Aba Arquivamento não encontrada.' };

    const ultimaCol  = sheet.getLastColumn();
    const dados      = sheet.getRange(2, 1, sheet.getLastRow() - 1, ultimaCol).getValues();
    const cabecalhos = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(String);
    const tz         = Session.getScriptTimeZone();

    const linhas = dados
      .filter(row => String(row[0] || '').trim() === codigo.trim())
      .map(row =>
        Object.fromEntries(cabecalhos.map((cab, i) => {
          let val = row[i];
          if (val instanceof Date)
            val = Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm');
          return [cab, String(val == null ? '' : val)];
        }))
      );

    if (linhas.length === 0)
      return { erro: 'Código não encontrado.' };

    return { linhas };
  } catch (e) {
    return { erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  DASHBOARD ADM
// ───────────────────────────────────────────────────────────────────────────

function getDashboard(emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const sheetRet = ss.getSheetByName("Retiradas");
    let pendentes = 0, liberados = 0, cancelados = 0;
    if (sheetRet && sheetRet.getLastRow() > 1) {
      const dados = sheetRet.getRange(2, 1, sheetRet.getLastRow()-1, 13).getValues();
      const lotes = {};
      dados.forEach(row => {
        const cod       = String(row[0]).trim();
        const statusRet = String(row[9]).trim();
        if (!cod) return;
        if (!lotes[cod]) lotes[cod] = statusRet;
      });
      Object.values(lotes).forEach(s => {
        if (s === 'PENDENTE')  pendentes++;
        if (s === 'LIBERADO')  liberados++;
        if (s === 'CANCELADO') cancelados++;
      });
    }

    const sheetArq = ss.getSheetByName("Arquivamento");
    let arquivamentosHoje = 0;
    const hoje = new Date();
    hoje.setHours(0,0,0,0);
    if (sheetArq && sheetArq.getLastRow() > 1) {
      const datas = sheetArq.getRange(2, 2, sheetArq.getLastRow()-1, 1).getValues();
      datas.forEach(row => {
        const d = new Date(row[0]);
        d.setHours(0,0,0,0);
        if (d.getTime() === hoje.getTime()) arquivamentosHoje++;
      });
    }

    const sheetUsr = ss.getSheetByName("Usuários");
    let usuariosAtivos = 0, usuariosTotal = 0;
    if (sheetUsr && sheetUsr.getLastRow() > 1) {
      const dados = sheetUsr.getRange(2, 1, sheetUsr.getLastRow()-1, 6).getValues();
      dados.forEach(row => {
        if (String(row[0]).trim()) {
          usuariosTotal++;
          if (String(row[5]).trim().toLowerCase() === 'sim') usuariosAtivos++;
        }
      });
    }

    const sheetDes = ss.getSheetByName("Descartes");
    let descartesPendentes = 0;
    if (sheetDes && sheetDes.getLastRow() > 1) {
      const dados = sheetDes.getRange(2, 8, sheetDes.getLastRow()-1, 1).getValues();
      dados.forEach(row => {
        if (String(row[0]).trim() === 'PENDENTE') descartesPendentes++;
      });
    }

    return { pendentes, liberados, cancelados, arquivamentosHoje, usuariosAtivos, usuariosTotal, descartesPendentes };
  } catch(e) { return { erro: e.message }; }
}


// ───────────────────────────────────────────────────────────────────────────
//  S.A.D. — Solicitação de Arquivamento de Documentos
// ───────────────────────────────────────────────────────────────────────────

/**
 * Busca os dados de um lote aprovado para renderizar o S.A.D.
 */
function getSAD(codigo, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Arquivamento');
    if (!sheet || sheet.getLastRow() < 2)
      return { ok: false, erro: 'Aba Arquivamento não encontrada ou vazia.' };

    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    const tz    = Session.getScriptTimeZone();

    const linhasDoLote = dados.filter(row => String(row[0]).trim() === codigo);
    if (!linhasDoLote.length)
      return { ok: false, erro: 'Lote não encontrado: ' + codigo };

    const status = String(linhasDoLote[0][11]).trim();
    if (status !== 'APROVADO')
      return { ok: false, erro: 'Lote não está aprovado. Status atual: ' + status };

    const primeiraLinha = linhasDoLote[0];
    const responsavel   = String(primeiraLinha[2]).trim();
    const departamento  = String(primeiraLinha[4]).trim();
    const dataAprovacao = primeiraLinha[1]
      ? Utilities.formatDate(new Date(primeiraLinha[1]), tz, 'dd/MM/yyyy')
      : '';

    const documentos = linhasDoLote.map(row => {
      const codigoId    = String(row[5]).trim();
      const descricao   = String(row[6]).trim();
      const dataDocRaw  = row[7];
      const informacao  = String(row[8]).trim();
      const temporalidade = String(row[10]).trim();

      let dataDoc = '';
      if (dataDocRaw) {
        if (dataDocRaw instanceof Date) {
          dataDoc = Utilities.formatDate(dataDocRaw, tz, 'dd/MM/yyyy');
        } else {
          const str = String(dataDocRaw).trim();
          try {
            const parsed = new Date(str);
            if (!isNaN(parsed)) {
              dataDoc = Utilities.formatDate(parsed, tz, 'dd/MM/yyyy');
            } else {
              dataDoc = str;
            }
          } catch(e) {
            dataDoc = str;
          }
        }
      }

      return { codigoId, descricao, dataDoc, informacao, temporalidade };
    });

    const colunasAtivas = {
      codigoId:      documentos.some(d => d.codigoId),
      descricao:     documentos.some(d => d.descricao),
      dataDoc:       documentos.some(d => d.dataDoc),
      informacao:    documentos.some(d => d.informacao),
      temporalidade: documentos.some(d => d.temporalidade)
    };

    return {
      ok: true,
      codigo,
      responsavel,
      departamento,
      dataAprovacao,
      local: LOCAL_GED,
      documentos,
      colunasAtivas,
      totalDocumentos: documentos.length
    };

  } catch(e) {
    return { ok: false, erro: e.message };
  }
}
