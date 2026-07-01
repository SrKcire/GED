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
//  CABEÇALHOS DE PLANILHA — fonte única de verdade
//  Usados tanto na criação da aba quanto na leitura por posição de coluna
//  (buscarHistorico/buscarDadosEtiqueta), para que uma renomeação manual de
//  cabeçalho na planilha não quebre silenciosamente os campos lidos pelo
//  frontend (que espera sempre esta mesma grafia).
// ───────────────────────────────────────────────────────────────────────────
const COLUNAS_ARQUIVAMENTO = [
  'CÓDIGO', 'DATA', 'RESPONSÁVEL', 'EMAIL', 'DEPARTAMENTO',
  'CÓDIGO / ID', 'DESCRIÇÃO / TIPO', 'DATA DO DOCUMENTO',
  'INFORMAÇÃO', 'ÁREA', 'TEMPORALIDADE',
  'STATUS', 'CAIXA', 'MOTIVO DO CANCELAMENTO', 'DATA DE VENCIMENTO'
];

const COLUNAS_RETIRADAS = [
  'CÓDIGO', 'DATA', 'RESPONSÁVEL', 'EMAIL', 'DEPARTAMENTO',
  'PROJETO', 'DESCRIÇÃO', 'CAIXA', 'MOTIVO',
  'STATUS RETIRADA', 'DATA DE RETIRADA',
  'STATUS DEVOLUÇÃO', 'DATA DE DEVOLUÇÃO',
  'DATA DE ADIAMENTO', 'MOTIVO/JUSTIFICATIVA'
];

const COLUNAS_DESCARTES = [
  'CÓDIGO', 'DATA', 'RESPONSÁVEL', 'EMAIL', 'DEPARTAMENTO',
  'DOCUMENTOS', 'JUSTIFICATIVA', 'STATUS', 'OBSERVAÇÃO ADM'
];


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
