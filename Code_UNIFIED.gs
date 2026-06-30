// ═══════════════════════════════════════════════════════════════════════════
//  GED — Colégio Santo Agostinho
//  Code_UNIFIED.gs  |  ARQUIVO LEGADO — mantido apenas como referência
//
//  Este arquivo foi dividido nos seguintes módulos:
//
//    Config.gs          — Constantes globais (SHEET_ID, CC_GED, LOCAL_GED,
//                          DEPT_ABA_MAP, DEPT_PREFIXO_MAP), doGet(), include()
//    Utils.gs           — normalizar(), gerarCodigoUnico(), calcularDiasUteis(),
//                          adicionarDiasUteis()
//    Auth.gs            — hashSenha(), _ehGedAtivo_(), verificarLogin(),
//                          trocarSenha(), migrarSenhasGED()
//    Consulta.gs        — buscarDocumentos(), buscarAcervoFisicoConsulta(),
//                          buscarHistorico(), getDashboard(), getSAD()
//    Arquivamento.gs    — salvarDados(), getArquivamentos(), aprovarArquivamento(),
//                          cancelarArquivamento(), excluirDocumentoArquivamento()
//    Retirada.gs        — buscarDocumentosRetirada(), salvarRetirada(),
//                          getRetiradas(), atualizarStatusLote(),
//                          verificarVencimentosRetiradas()
//    Descarte.gs        — solicitarDescarte(), getDescartes(), aprovarDescarte(),
//                          recusarDescarte()
//    Digitalizacao.gs   — buscarAluno(), salvarDossie(), registrarDigitalizacao(),
//                          getDigitalizados(), getDashboardDigitalizacao()
//    Usuarios.gs        — listarUsuarios(), criarUsuario(), atualizarUsuario(),
//                          redefinirSenhaAdm(), toggleUsuarioAtivo(), deletarUsuario()
//    Emails.gs          — escHtml(), gerarAssinaturaGed(), enviarEmail*()
//
// ═══════════════════════════════════════════════════════════════════════════
