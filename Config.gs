// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Config.gs | Constantes globais, roteamento e inclusão de partials HTML
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_ID = "1wVxjv_fPNhMHrZ-vNQDcrXXDVXXs3cBtDSwbWCsre2g";

// Mapeamento: departamento → aba de acervo
const DEPT_ABA_MAP = {
  'Infraestrutura': 'DOC.INFRA',
  'Administrativo': 'DOC.ADMIN',
  'Financeiro':     'DOC.FIN',
  'RH':             'DOC.RH',
  'Pedagógico':     'DOC.PED',
  'Comunicação':    'DOC.COM',
  'TI':             'DOC.TI',
};

// Mapeamento: departamento → prefixo de código
const DEPT_PREFIXO_MAP = {
  'Infraestrutura': 'INFRA',
  'Administrativo': 'ADMIN',
  'Financeiro':     'FIN',
  'RH':             'RH',
  'Pedagógico':     'PED',
  'Comunicação':    'COM',
  'TI':             'TI',
};

// Endereços copiados em todos os e-mails automáticos
const CC_GED = 'rafael.spinelli@csa.com.br,erick.mesquita@csa.com.br';

// Local físico do acervo (usado no S.A.D.)
const LOCAL_GED = 'GED (TÉRREO CD67)';


// ───────────────────────────────────────────────────────────────────────────
//  PONTO DE ENTRADA ÚNICO
// ───────────────────────────────────────────────────────────────────────────

/**
 * Único doGet() do sistema unificado.
 * Serve index_UNIFIED.html para todos os perfis.
 * A diferenciação de interface ocorre no frontend após o login.
 */
function doGet() {
  return HtmlService.createTemplateFromFile("index_UNIFIED")
    .evaluate()
    .setTitle("GED")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Inclusão de arquivos HTML partials (usado pelo template GAS).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
