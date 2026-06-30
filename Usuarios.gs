// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Usuarios.gs | CRUD de usuários (exclusivo ADM)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Lista usuários da aba "Usuários".
 * SEGURANÇA: o campo senha NÃO é retornado ao frontend.
 * Lê até coluna H (departamento).
 */
function listarUsuarios(emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return Object.assign({ usuarios: [] }, _respostaAcessoNegado_());
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet || sheet.getLastRow() < 2) return { usuarios: [] };

    while (sheet.getMaxColumns() < 8) sheet.insertColumnAfter(sheet.getMaxColumns());

    const dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 8).getValues();
    const usuarios = [];
    dados.forEach((r, i) => {
      if (!String(r[0]).trim()) return;
      usuarios.push({
        linha:         i + 2,
        login:         String(r[0]).trim(),
        // senha omitida intencionalmente
        nome:          String(r[2]).trim(),
        email:         String(r[3]).trim(),
        perfil:        String(r[4]).trim(),
        ativo:         String(r[5]).trim(),
        senhaAlterada: String(r[6]).trim(),
        departamento:  String(r[7]).trim(),
      });
    });
    return { usuarios };
  } catch(e) { return { usuarios: [], erro: e.message }; }
}

/**
 * Cria novo usuário.
 * - Aplica hash SHA-256 na senha antes de gravar.
 * - Marca coluna G como "Não" → troca obrigatória no 1º login.
 * - Grava coluna H (departamento).
 */
function criarUsuario(dados, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { sucesso: false };

    while (sheet.getMaxColumns() < 8) sheet.insertColumnAfter(sheet.getMaxColumns());

    const existentes = sheet.getRange(2, 1, Math.max(sheet.getLastRow()-1,1), 1).getValues();
    for (const r of existentes) {
      if (String(r[0]).trim().toLowerCase() === dados.login.toLowerCase())
        return { sucesso: false, erro: 'Login já existe.' };
    }

    const proxLinha = Math.max(sheet.getLastRow(), 1) + 1;
    const senhaHash = hashSenha(dados.senha);

    sheet.getRange(proxLinha, 1, 1, 8).setValues([[
      dados.login,
      senhaHash,
      dados.nome,
      dados.email,
      dados.perfil,
      dados.ativo || 'Sim',
      'Não',
      dados.departamento || ''
    ]]);
    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

/**
 * Atualiza dados do usuário (exceto senha).
 * Para redefinir senha, use trocarSenha() ou redefinirSenhaAdm().
 */
function atualizarUsuario(linha, dados, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { sucesso: false };

    while (sheet.getMaxColumns() < 8) sheet.insertColumnAfter(sheet.getMaxColumns());

    sheet.getRange(linha, 1).setValue(dados.login);
    sheet.getRange(linha, 3).setValue(dados.nome);
    sheet.getRange(linha, 4).setValue(dados.email);
    sheet.getRange(linha, 5).setValue(dados.perfil);
    sheet.getRange(linha, 6).setValue(dados.ativo);
    sheet.getRange(linha, 8).setValue(dados.departamento || '');

    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

/**
 * Redefine a senha de um usuário pelo ADM.
 * Hasheia a nova senha e marca G como "Não" (troca obrigatória no próximo login).
 */
function redefinirSenhaAdm(linha, novaSenha, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    if (!novaSenha || novaSenha.length < 6)
      return { sucesso: false, erro: 'A senha deve ter pelo menos 6 caracteres.' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { sucesso: false };
    sheet.getRange(linha, 2).setValue(hashSenha(novaSenha));
    sheet.getRange(linha, 7).setValue('Não');
    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

function toggleUsuarioAtivo(linha, novoStatus, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { sucesso: false };
    sheet.getRange(linha, 6).setValue(novoStatus);
    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

function deletarUsuario(linha, emailSolicitante) {
  if (!_ehGedAtivo_(emailSolicitante)) return _respostaAcessoNegado_();
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuários");
    if (!sheet) return { sucesso: false };
    sheet.deleteRow(linha);
    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}
