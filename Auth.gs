// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Auth.gs | Autenticação, hashing de senha e guards de segurança
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Gera hash SHA-256 de uma string.
 * Retorna string hexadecimal de 64 caracteres.
 */
function hashSenha(senha) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    senha,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => {
    const hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/** Verifica se um valor já é um hash SHA-256 (64 caracteres hexadecimais) */
function senhaEhHash(valor) {
  return typeof valor === 'string' && /^[0-9a-f]{64}$/.test(valor);
}

/**
 * GUARD DE SEGURANÇA — verifica se o e-mail informado pertence a um usuário
 * ativo com perfil "GED". Usado no início de toda função server-side
 * exclusiva do administrador, para impedir que um usuário comum a chame
 * diretamente via google.script.run / console do GAS.
 *
 * Retorna true se autorizado, false caso contrário.
 * Não lança exceção — quem chama decide o que retornar ao frontend.
 */
function _ehGedAtivo_(email) {
  try {
    if (!email) return false;
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet || sheet.getLastRow() < 2) return false;

    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    const emailNorm = String(email).trim().toLowerCase();

    for (const row of dados) {
      const emailPlanilha = String(row[3] || '').trim().toLowerCase();
      if (emailPlanilha !== emailNorm) continue;
      const perfil = String(row[4] || '').trim();
      const ativo  = String(row[5] || '').trim().toLowerCase();
      return (perfil === 'GED' && ativo === 'sim');
    }
    return false;
  } catch (e) {
    return false;
  }
}

/** Resposta padrão de acesso negado, usada pelos guards das funções ADM. */
function _respostaAcessoNegado_() {
  return { ok: false, sucesso: false, erro: 'Acesso restrito ao administrador GED.', acessoNegado: true };
}


// ───────────────────────────────────────────────────────────────────────────
//  AUTENTICAÇÃO UNIFICADA
// ───────────────────────────────────────────────────────────────────────────

/**
 * Verifica credenciais na aba "Usuários" — válido para TODOS os perfis.
 *
 * Colunas da aba "Usuários":
 *   A  Login
 *   B  Senha (texto puro legado ou hash SHA-256)
 *   C  Nome
 *   D  Email
 *   E  Perfil  ("GED" = administrador; qualquer outro = usuário departamento)
 *   F  Ativo   ("Sim" / "Não")
 *   G  Senha Alterada ("Sim" / "Não")
 *   H  Departamento
 *
 * Retorna: { status, nome, email, perfil, departamento, precisaTrocarSenha }
 */
function verificarLogin(login, senha) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { status: 'erro' };

    const dados = sheet.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const loginPlanilha = String(dados[i][0]).trim().toLowerCase();
      const senhaPlanilha = String(dados[i][1]).trim();
      const nome          = String(dados[i][2]).trim();
      const email         = String(dados[i][3]).trim();
      const perfil        = String(dados[i][4]).trim();
      const ativo         = String(dados[i][5]).trim().toLowerCase();
      const senhaAlterada = String(dados[i][6] || '').trim();
      const departamento  = String(dados[i][7] || '').trim();

      if (loginPlanilha !== login.toLowerCase()) continue;

      // SEGURANÇA: perfil GED nunca pode autenticar com senha em texto puro.
      if (perfil === 'GED' && !senhaEhHash(senhaPlanilha)) {
        return { status: 'requer_migracao' };
      }

      let senhaValida = false;
      if (senhaAlterada === 'Sim' || senhaEhHash(senhaPlanilha)) {
        senhaValida = (hashSenha(senha) === senhaPlanilha);
      } else {
        senhaValida = (senhaPlanilha === senha);
      }

      if (!senhaValida)    return { status: 'erro' };
      if (ativo !== 'sim') return { status: 'inativo' };

      return {
        status: 'ok',
        nome,
        email,
        perfil,
        departamento: perfil === 'GED' ? 'GED' : departamento,
        precisaTrocarSenha: (senhaAlterada !== 'Sim')
      };
    }

    return { status: 'erro' };
  } catch (e) {
    return { status: 'erro' };
  }
}

/**
 * Troca a senha do usuário.
 * Valida a senha atual, hasheia a nova e marca coluna G como "Sim".
 */
function trocarSenha(login, senhaAtual, novaSenha) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { sucesso: false, erro: 'Aba não encontrada.' };

    const dados = sheet.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const emailPlanilha = String(dados[i][3]).trim().toLowerCase();
      const senhaPlanilha = String(dados[i][1]).trim();
      const senhaAlterada = String(dados[i][6] || '').trim();

      if (emailPlanilha !== String(login).trim().toLowerCase()) continue;

      let senhaAtualValida = false;
      if (senhaAlterada === 'Sim' || senhaEhHash(senhaPlanilha)) {
        senhaAtualValida = (hashSenha(senhaAtual) === senhaPlanilha);
      } else {
        senhaAtualValida = (senhaPlanilha === senhaAtual);
      }

      if (!senhaAtualValida) return { sucesso: false, erro: 'Senha atual incorreta.' };
      if (!novaSenha || novaSenha.length < 6)
        return { sucesso: false, erro: 'A nova senha deve ter pelo menos 6 caracteres.' };

      sheet.getRange(i + 1, 2).setValue(hashSenha(novaSenha));
      sheet.getRange(i + 1, 7).setValue('Sim');
      return { sucesso: true };
    }

    return { sucesso: false, erro: 'Usuário não encontrado.' };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}


// ───────────────────────────────────────────────────────────────────────────
//  UTILITÁRIO — MIGRAÇÃO DE SENHAS GED PARA SHA-256
//  Executar UMA VEZ manualmente no editor do Apps Script antes do go-live.
// ───────────────────────────────────────────────────────────────────────────

/**
 * Percorre a aba "Usuários" e, para cada linha com perfil "GED"
 * cuja senha ainda não seja um hash SHA-256, aplica hashSenha()
 * e marca coluna G como "Não" (troca obrigatória no 1º login).
 */
function migrarSenhasGED() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('Aba Usuários não encontrada ou vazia.');
    return { ok: false, migrados: 0, erro: 'Aba Usuários não encontrada ou vazia.' };
  }

  while (sheet.getMaxColumns() < 8) sheet.insertColumnAfter(sheet.getMaxColumns());

  const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
  let migrados = 0;

  dados.forEach((row, idx) => {
    const perfil        = String(row[4]).trim();
    const senhaPlanilha = String(row[1]).trim();
    const senhaAlterada = String(row[6] || '').trim();

    if (perfil !== 'GED') return;
    if (senhaAlterada === 'Sim' || senhaEhHash(senhaPlanilha)) return;

    const linhaReal = idx + 2;
    sheet.getRange(linhaReal, 2).setValue(hashSenha(senhaPlanilha));
    sheet.getRange(linhaReal, 7).setValue('Não');
    migrados++;
    Logger.log(`Migrado: linha ${linhaReal} — login: ${String(row[0]).trim()}`);
  });

  Logger.log(`migrarSenhasGED concluído. ${migrados} senha(s) migrada(s).`);
  return { ok: true, migrados };
}
