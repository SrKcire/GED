// ═══════════════════════════════════════════════════════════════════════════
//  GED UNIFICADO — Colégio Santo Agostinho
//  Emails.gs | Templates HTML e envio de e-mails automáticos
// ═══════════════════════════════════════════════════════════════════════════

/** Escapa caracteres especiais HTML para uso seguro em templates de e-mail. */
function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function gerarAssinaturaGed() {
  return `
    <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;margin-top:20px;text-align:center;">
      <p style="margin:5px 0;color:#495057;font-weight:bold;">Atenciosamente,</p>
      <p style="margin:5px 0;color:#495057;"><strong>Colégio Santo Agostinho — Leblon</strong></p>
      <p style="margin:10px 0;color:#007bff;font-weight:bold;font-size:16px;">Gestão Eletrônica de Documentos</p>
      <p style="margin:5px 0;color:#495057;">Rua José Linhares, 88 — Rio de Janeiro RJ | CEP: 22430-220</p>
    </div>`;
}

function gerarRodapeAutomatico() {
  return `
    <div style="background-color:#e9ecef;padding:10px 20px;border-radius:0 0 8px 8px;margin-top:10px;">
      <p style="margin:0;color:#6c757d;font-size:12px;text-align:center;">
        Este e-mail é enviado automaticamente pelo Sistema GED — Colégio Santo Agostinho.<br>
        Por favor, não responda a este e-mail.
      </p>
    </div>`;
}

function gerarTabelaLote(dados) {
  const linhas = (dados.documentos || []).map(d => `
    <tr>
      <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.projeto) || '—'}</td>
      <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.descricao) || '—'}</td>
      <td style="border:1px solid #90caf9;padding:8px;text-align:center;">${escHtml(d.caixa) || '—'}</td>
    </tr>`).join('');

  return `
    <div style="background-color:#e3f2fd;border-radius:6px;padding:15px;margin:20px 0;">
      <h3 style="color:#1976d2;margin-top:0;">Detalhes da Solicitação</h3>
      <table style="width:100%;border-collapse:collapse;margin-bottom:10px;">
        <tr style="background-color:#bbdefb;">
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Código</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Responsável</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Departamento</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Motivo</th>
        </tr>
        <tr>
          <td style="border:1px solid #90caf9;padding:8px;"><strong>${escHtml(dados.codigo)}</strong></td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(dados.responsavel) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(dados.departamento) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(dados.motivo) || '—'}</td>
        </tr>
      </table>
      <table style="width:100%;border-collapse:collapse;">
        <tr style="background-color:#bbdefb;">
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Projeto</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Descrição</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;text-align:center;">Caixa</th>
        </tr>
        ${linhas}
      </table>
    </div>`;
}


// ── E-mail 1: LIBERAÇÃO ──────────────────────────────────────────────────

function enviarEmailLiberacao(codigo, dados) {
  try {
    if (!dados || !dados.email) return;
    const remetente = Session.getActiveUser().getEmail();
    const assunto   = `[GED] Retirada ${codigo} — LIBERADA para retirada`;

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Confirmação de Retirada — GED</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(dados.responsavel)},</strong></p>
          <div style="background-color:#d4edda;border:1px solid #c3e6cb;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#155724;margin:0;font-weight:bold;">Sua solicitação de retirada ${escHtml(codigo)} foi LIBERADA.</p>
            <p style="color:#155724;margin:10px 0 0 0;">Os documentos estão disponíveis para retirada no GED (Térreo, CD67).</p>
          </div>
          ${gerarTabelaLote(dados)}
          <p style="color:#495057;line-height:1.6;">Os documentos devem ser devolvidos ao GED em até <strong>5 dias úteis</strong>.<br>Horário de atendimento: Segunda a Sexta, até as 16:50.</p>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente. Não é necessário responder.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(dados.email, assunto, '', {
      from: remetente,
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail liberação: ' + e.message); }
}


// ── E-mail 2: CANCELAMENTO ───────────────────────────────────────────────

function enviarEmailCancelamento(codigo, dados) {
  try {
    if (!dados || !dados.email) return;
    const remetente = Session.getActiveUser().getEmail();
    const assunto   = `[GED] Retirada ${codigo} — CANCELADA`;

    const motivoHtml = dados.motivoCancelamento
      ? `<div style="background-color:#fff3cd;border:1px solid #ffeaa7;border-radius:6px;padding:12px;margin:16px 0;">
           <p style="color:#856404;margin:0;"><strong>Motivo do Cancelamento:</strong> ${escHtml(dados.motivoCancelamento)}</p>
         </div>`
      : '';

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Cancelamento de Retirada — GED</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(dados.responsavel)},</strong></p>
          <div style="background-color:#f8d7da;border:1px solid #f5c6cb;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#721c24;margin:0;font-weight:bold;">A solicitação de retirada ${escHtml(codigo)} foi CANCELADA.</p>
          </div>
          ${motivoHtml}
          ${gerarTabelaLote(dados)}
          <p style="color:#495057;line-height:1.6;">Caso necessite dos documentos, realize uma nova solicitação pelo sistema GED.</p>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente. Não é necessário responder.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(dados.email, assunto, '', {
      from: remetente,
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail cancelamento: ' + e.message); }
}


// ── E-mail 3: ADIAMENTO ──────────────────────────────────────────────────

function enviarEmailAdiamento(codigo, dados) {
  try {
    if (!dados || !dados.email) return;
    const remetente    = Session.getActiveUser().getEmail();
    const assunto      = `[GED] Retirada ${codigo} — PRAZO ADIADO`;
    const dataFormatada = Utilities.formatDate(
      new Date(dados.dataFinal + 'T00:00:00'),
      Session.getScriptTimeZone(), 'dd/MM/yyyy'
    );

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Confirmação de Adiamento — GED</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(dados.responsavel)},</strong></p>
          <div style="background-color:#d4edda;border:1px solid #c3e6cb;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#155724;margin:0;font-weight:bold;">O prazo para devolução dos documentos da retirada ${escHtml(codigo)} foi estendido.</p>
            <p style="color:#155724;margin:10px 0 0 0;">Nova data de devolução: <strong>${dataFormatada}</strong></p>
          </div>
          ${gerarTabelaLote(dados)}
          ${dados.justificativa ? `
          <div style="background-color:#fff3cd;border:1px solid #ffeaa7;border-radius:6px;padding:12px;margin:16px 0;">
            <p style="color:#856404;margin:0;"><strong>Justificativa:</strong> ${escHtml(dados.justificativa)}</p>
          </div>` : ''}
          <p style="color:#495057;line-height:1.6;">Após essa data, os documentos deverão ser devolvidos ao GED até as 16:50.<br>Não será possível realizar um novo adiamento.</p>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente. Não é necessário responder.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(dados.email, assunto, '', {
      from: remetente,
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail adiamento: ' + e.message); }
}


// ── E-mail 4: COBRANÇA 5 DIAS ÚTEIS ────────────────────────────────────

function enviarEmailCobranca5Dias(email, responsavel, codigo, dataRetirada, diasUteis, documentos) {
  try {
    if (!email) return;
    const assunto = `[GED] Retirada ${codigo} — PRAZO VENCIDO`;
    const dataFormatada = Utilities.formatDate(
      new Date(dataRetirada), Session.getScriptTimeZone(), 'dd/MM/yyyy'
    );

    const linhasDocs = (documentos || []).map(d => `
      <tr>
        <td style="border:1px solid #90caf9;padding:8px;">${d.projeto || '—'}</td>
        <td style="border:1px solid #90caf9;padding:8px;">${d.descricao || '—'}</td>
        <td style="border:1px solid #90caf9;padding:8px;text-align:center;">${d.caixa || '—'}</td>
      </tr>`).join('');

    const tabelaDocs = (documentos && documentos.length > 0) ? `
      <div style="background-color:#e3f2fd;border-radius:6px;padding:15px;margin:16px 0;">
        <h3 style="color:#1976d2;margin-top:0;">Documentos Retirados</h3>
        <table style="width:100%;border-collapse:collapse;">
          <tr style="background-color:#bbdefb;">
            <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Projeto</th>
            <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Descrição</th>
            <th style="border:1px solid #90caf9;padding:8px;text-align:center;color:#1565c0;">Caixa</th>
          </tr>
          ${linhasDocs}
        </table>
      </div>` : '';

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Alerta Automático — Gestão Eletrônica de Documentos</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(responsavel)},</strong></p>
          <div style="background-color:#fff3cd;border:1px solid #ffeaa7;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#856404;margin:0;font-weight:bold;">
              O prazo máximo de 5 dias úteis para a permanência dos documentos da retirada
              <strong>${escHtml(codigo)}</strong> foi atingido.<br>
              Por favor, devolva os documentos ao GED até as 16:50 de hoje.
            </p>
          </div>
          <table style="width:100%;border-collapse:collapse;margin:16px 0;">
            <tr style="background-color:#bbdefb;">
              <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Código</th>
              <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Data de Retirada</th>
              <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Dias Úteis Decorridos</th>
            </tr>
            <tr>
              <td style="border:1px solid #90caf9;padding:8px;"><strong>${codigo}</strong></td>
              <td style="border:1px solid #90caf9;padding:8px;">${dataFormatada}</td>
              <td style="border:1px solid #90caf9;padding:8px;color:#d32f2f;font-weight:bold;">${diasUteis} dias</td>
            </tr>
          </table>
          ${tabelaDocs}
          <p style="color:#495057;line-height:1.6;">
            Em caso de necessidade de prolongamento do prazo, contate os responsáveis do GED:<br>
            <strong>Rafael Spinelli</strong> ou <strong>Erick Mesquita</strong>.
          </p>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente diariamente às 10h até que a devolução seja registrada.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(email, assunto, '', {
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail cobrança 5 dias: ' + e.message); }
}


// ── E-mail 5: COBRANÇA PÓS-ADIAMENTO VENCIDO ────────────────────────────

function enviarEmailCobrancaAdiamentoVencido(email, responsavel, codigo, dataAdiamento, documentos) {
  try {
    if (!email) return;
    const assunto = `[GED] Retirada ${codigo} — PRAZO DE ADIAMENTO VENCIDO`;
    const dataFormatada = Utilities.formatDate(
      new Date(dataAdiamento), Session.getScriptTimeZone(), 'dd/MM/yyyy'
    );

    const linhasDocs = (documentos || []).map(d => `
      <tr>
        <td style="border:1px solid #90caf9;padding:8px;">${d.projeto || '—'}</td>
        <td style="border:1px solid #90caf9;padding:8px;">${d.descricao || '—'}</td>
        <td style="border:1px solid #90caf9;padding:8px;text-align:center;">${d.caixa || '—'}</td>
      </tr>`).join('');

    const tabelaDocs = (documentos && documentos.length > 0) ? `
      <div style="background-color:#e3f2fd;border-radius:6px;padding:15px;margin:16px 0;">
        <h3 style="color:#1976d2;margin-top:0;">Documentos Retirados</h3>
        <table style="width:100%;border-collapse:collapse;">
          <tr style="background-color:#bbdefb;">
            <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Projeto</th>
            <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Descrição</th>
            <th style="border:1px solid #90caf9;padding:8px;text-align:center;color:#1565c0;">Caixa</th>
          </tr>
          ${linhasDocs}
        </table>
      </div>` : '';

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Alerta Automático — Gestão Eletrônica de Documentos</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(responsavel)},</strong></p>
          <div style="background-color:#f8d7da;border:1px solid #f5c6cb;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#721c24;margin:0;font-weight:bold;">
              O prazo de adiamento da retirada <strong>${escHtml(codigo)}</strong> venceu em ${dataFormatada}.<br>
              Por favor, devolva os documentos ao GED hoje até as 16:50.
            </p>
          </div>
          <table style="width:100%;border-collapse:collapse;margin:16px 0;">
            <tr style="background-color:#bbdefb;">
              <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Código</th>
              <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Data Limite do Adiamento</th>
            </tr>
            <tr>
              <td style="border:1px solid #90caf9;padding:8px;"><strong>${codigo}</strong></td>
              <td style="border:1px solid #90caf9;padding:8px;color:#d32f2f;font-weight:bold;">${dataFormatada}</td>
            </tr>
          </table>
          ${tabelaDocs}
          <p style="color:#495057;line-height:1.6;">Caso ainda precise dos documentos, devolva-os e realize uma nova solicitação pelo sistema GED.</p>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente diariamente às 10h até que a devolução seja registrada.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(email, assunto, '', {
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail cobrança adiamento vencido: ' + e.message); }
}


// ── E-mail 6: APROVAÇÃO DE ARQUIVAMENTO ────────────────────────────────

function enviarEmailAprovacaoArquivamento(codigo, dados) {
  try {
    const remetente = Session.getActiveUser().getEmail();
    const assunto   = `[GED] Arquivamento ${codigo} — APROVADO`;

    const isSolicitacao = dados.isSolicitacao;

    const linhasDoc = (dados.documentos || []).map(d => {
      if (isSolicitacao) {
        return `<tr>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.informacao) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.dataDoc) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.descricao) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.codigoId) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.temporalidade) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;text-align:center;font-weight:bold;">${escHtml(d.caixa) || '—'}</td>
        </tr>`;
      } else {
        return `<tr>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.descricao) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.informacao) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.dataDoc) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.area) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;">${escHtml(d.temporalidade) || '—'}</td>
          <td style="border:1px solid #90caf9;padding:8px;text-align:center;font-weight:bold;">${escHtml(d.caixa) || '—'}</td>
        </tr>`;
      }
    }).join('');

    const cabecalhoDoc = isSolicitacao
      ? `<tr style="background-color:#bbdefb;">
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Setor</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Data do Documento</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Tipo de Solicitação</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">ID da Solicitação</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Temporalidade</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:center;color:#1565c0;">Caixa</th>
        </tr>`
      : `<tr style="background-color:#bbdefb;">
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Projeto</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Descrição</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Data do Documento</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Área</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:left;color:#1565c0;">Temporalidade</th>
          <th style="border:1px solid #90caf9;padding:8px;text-align:center;color:#1565c0;">Caixa</th>
        </tr>`;

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Confirmação de Arquivamento — GED</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(dados.responsavel)},</strong></p>
          <div style="background-color:#d4edda;border:1px solid #c3e6cb;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#155724;margin:0;font-weight:bold;">O arquivamento ${escHtml(codigo)} foi APROVADO e os documentos foram registrados no acervo do GED.</p>
          </div>
          <div style="background-color:#e3f2fd;border-radius:6px;padding:15px;margin:20px 0;">
            <h3 style="color:#1976d2;margin-top:0;">Documentos Arquivados</h3>
            <table style="width:100%;border-collapse:collapse;">
              ${cabecalhoDoc}
              ${linhasDoc}
            </table>
          </div>
          <div style="background-color:#fff3cd;border:1px solid #ffeaa7;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#856404;margin:0;font-weight:bold;">Ação necessária: Compareça ao GED para assinar o Documento de Arquivamento.</p>
            <p style="color:#856404;margin:8px 0 0 0;">Horário de atendimento: Segunda a Sexta, até as 16:50.</p>
          </div>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente. Não é necessário responder.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(dados.email, assunto, '', {
      from: remetente,
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail aprovação arquivamento: ' + e.message); }
}


// ── E-mail 7: CANCELAMENTO DE ARQUIVAMENTO ──────────────────────────────

function enviarEmailCancelamentoArquivamento(codigo, dados) {
  try {
    const remetente = Session.getActiveUser().getEmail();
    const assunto   = `[GED] Arquivamento ${codigo} — CANCELADO`;

    const motivoHtml = dados.motivoCancelamento
      ? `<div style="background-color:#fff3cd;border:1px solid #ffeaa7;border-radius:6px;padding:12px;margin:16px 0;">
           <p style="color:#856404;margin:0;"><strong>Motivo do Cancelamento:</strong> ${escHtml(dados.motivoCancelamento)}</p>
         </div>`
      : '';

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background-color:#f8f9fa;padding:20px;border-radius:8px;">
          <h2 style="color:#495057;text-align:center;margin-bottom:20px;">Cancelamento de Arquivamento — GED</h2>
          <p style="color:#495057;line-height:1.6;"><strong>Prezado(a) ${escHtml(dados.responsavel)},</strong></p>
          <div style="background-color:#f8d7da;border:1px solid #f5c6cb;border-radius:6px;padding:15px;margin:20px 0;">
            <p style="color:#721c24;margin:0;font-weight:bold;">O arquivamento ${escHtml(codigo)} foi CANCELADO. Compareça ao GED para mais informações e devolução dos documentos.</p>
          </div>
          ${motivoHtml}
          <p style="color:#495057;line-height:1.6;">Caso necessite arquivar documentos, realize uma nova solicitação pelo sistema GED.</p>
          <p style="color:#666;font-size:14px;font-style:italic;">Este e-mail é gerado automaticamente. Não é necessário responder.</p>
        </div>
        ${gerarAssinaturaGed()}
        ${gerarRodapeAutomatico()}
      </div>`;

    GmailApp.sendEmail(dados.email, assunto, '', {
      from: remetente,
      cc:   CC_GED,
      htmlBody: html,
      name: 'GED — Colégio Santo Agostinho'
    });
  } catch(e) { Logger.log('Erro e-mail cancelamento arquivamento: ' + e.message); }
}


// ── Dispatcher de e-mails de retirada ───────────────────────────────────

function enviarEmailRetirada(codigo, acao, dados) {
  if (acao === 'LIBERAR')  enviarEmailLiberacao(codigo, dados);
  if (acao === 'CANCELAR') enviarEmailCancelamento(codigo, dados);
  if (acao === 'ADIAR')    enviarEmailAdiamento(codigo, dados);
}
