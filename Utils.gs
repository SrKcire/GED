// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Utils.gs | Funções utilitárias gerais (sem dependência de módulo)
// ═══════════════════════════════════════════════════════════════════════════

/** Remove acentos e normaliza string para comparações */
function normalizar(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .trim();
}

/**
 * Gera um código único sequencial para o registro.
 * Formato: PREFIXO_TIPO.PREFIXO_DEPT-N  (ex: AM.INFRA-42)
 */
function gerarCodigoUnico(nomeAba, prefixoTipo, departamento) {
  try {
    const planilha = SpreadsheetApp.openById(SHEET_ID);
    const sheet    = planilha.getSheetByName(nomeAba);
    const prefDept = DEPT_PREFIXO_MAP[departamento] || departamento.toUpperCase().replace(/\s/g, '');
    const prefixo  = prefixoTipo + '.' + prefDept;

    if (!sheet || sheet.getLastRow() < 2) return prefixo + '-1';

    const codigos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    let maximo = 0;

    codigos.forEach(row => {
      const val = String(row[0]).trim();
      if (val.startsWith(prefixo + '-')) {
        const num = parseInt(val.split('-')[1], 10);
        if (!isNaN(num) && num > maximo) maximo = num;
      }
    });

    return prefixo + '-' + (maximo + 1);
  } catch (e) {
    return 'COD-ERR';
  }
}

/** Calcula dias úteis entre duas datas */
function calcularDiasUteis(dataInicio, dataFim) {
  let count = 0;
  const d = new Date(dataInicio);
  d.setHours(0,0,0,0);
  const fim = new Date(dataFim);
  fim.setHours(0,0,0,0);
  while (d <= fim) {
    const dia = d.getDay();
    if (dia !== 0 && dia !== 6) count++;
    d.setDate(d.getDate() + 1);
  }
  return Math.max(0, count - 1);
}

/** Avança N dias úteis a partir de uma data base */
function adicionarDiasUteis(dataBase, diasUteis) {
  const d = new Date(dataBase);
  let adicionados = 0;
  while (adicionados < diasUteis) {
    d.setDate(d.getDate() + 1);
    const dia = d.getDay();
    if (dia !== 0 && dia !== 6) adicionados++;
  }
  return d;
}
