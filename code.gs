// =======================
// CONFIGURAÇÕES
// =======================
const CONFIG = {
  SHEET_NAME: 'Respostas ao formulário 1',
  TEMPLATE_ID: '1ESBK8cvLiv477W4fdYOrgWT9KBHEe31fEvYYYUUU',
  FOLDER_ID: '1Z_0rnfVb0e7tjDx8SH_YwE5QRUV5556K'
};

// Índices das colunas (1 = A, 2 = B, ...)
const COL = {
  DATA_HORA: 1,       // A
  EMAIL: 2,           // B
  PONTUACAO: 3,       // C
  NOME: 4,            // D
  CNPJ: 5,            // E
  ANALISTA: 6,        // F
  CURSO: 17,          // Q
  DATA_INICIO: 18,    // R
  DATA_TERMINO: 19,   // S
  STATUS: 20,         // T
  CERT_GERADO: 21,    // U
  LINK_CERT: 22,      // V
  TOKEN: 23,          // W
  VALIDADE: 24        // X
};

// =======================
// FUNÇÃO PRINCIPAL
// =======================

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  const lastCol = sheet.getLastColumn();
  const valores = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const status = valores[COL.STATUS - 1];
  const certGerado = valores[COL.CERT_GERADO - 1];

  // Só gera certificado se estiver Aprovado e ainda não tiver gerado
  if (status !== 'Aprovado') return;
  if (certGerado === 'SIM') return;

  const dados = {
    linha: row,
    nome: valores[COL.NOME - 1],
    email: valores[COL.EMAIL - 1],
    curso: valores[COL.CURSO - 1],
    dataInicio: valores[COL.DATA_INICIO - 1],
    dataTermino: valores[COL.DATA_TERMINO - 1],
    validade: valores[COL.VALIDADE - 1],
    analista: valores[COL.ANALISTA - 1]
  };

  const token = gerarToken();
  const pdfFile = gerarCertificadoPDF(dados, token);
  const link = pdfFile.getUrl();

  // Atualiza planilha (U, V, W)
  sheet.getRange(row, COL.CERT_GERADO).setValue('SIM');
  sheet.getRange(row, COL.LINK_CERT).setValue(link);
  sheet.getRange(row, COL.TOKEN).setValue(token);

  // Envia e-mail com PDF em anexo
  enviarEmailComCertificado(dados, token, pdfFile);
}

// =======================
// GERAR TOKEN
// =======================

function gerarToken(tamanho = 12) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let token = '';
  for (let i = 0; i < tamanho; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// =======================
// GERAR PDF DO CERTIFICADO
// =======================

function gerarCertificadoPDF(dados, token) {
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

  const nomeArquivoBase = `Certificado - ${dados.nome} - ${dados.curso}`;
  const copia = templateFile.makeCopy(nomeArquivoBase, folder);
  const presentation = SlidesApp.openById(copia.getId());

  const tz = Session.getScriptTimeZone();
  const dataCurso = dados.dataTermino || dados.dataInicio || new Date();
  const dataCursoStr = Utilities.formatDate(new Date(dataCurso), tz, 'dd/MM/yyyy');

  const mapa = {
    '<<NOME>>': dados.nome,
    '<<DATA>>': dataCursoStr,
    '<<INSTRUTOR>>': dados.analista || ''
  };

  presentation.getSlides().forEach(slide => {
    slide.getShapes().forEach(shape => {
      if (!shape.getText) return;
      const textRange = shape.getText();
      let text = textRange.asString();
      let original = text;

      Object.keys(mapa).forEach(chave => {
        if (text.includes(chave)) {
          text = text.replace(chave, mapa[chave]);
        }
      });

      if (text !== original) {
        textRange.setText(text);
      }
    });
  });

  presentation.saveAndClose();

  // Gera PDF
  const blob = copia.getAs('application/pdf');
  const pdfFile = folder.createFile(blob).setName(nomeArquivoBase + '.pdf');

  // Mantém só o PDF
  copia.setTrashed(true);

  return pdfFile;
}

// =======================
// ENVIAR E-MAIL (APENAS PDF EM ANEXO)
// =======================

function enviarEmailComCertificado(dados, token, pdfFile) {
  const tz = Session.getScriptTimeZone();
  const validadeStr = dados.validade
    ? Utilities.formatDate(new Date(dados.validade), tz, 'dd/MM/yyyy')
    : 'indeterminada';

  const assunto = `Certificado - ${dados.curso || 'Treinamento SIEG'}`;

  const corpoHtml = `
    Olá, ${dados.nome}!<br><br>
    Obrigado por participar do treinamento <b>${dados.curso || 'SIEG'}</b>.<br><br>
    Seu certificado em PDF está anexado a este e-mail.<br><br>
    Token: <b>${token}</b><br>
    Validade: <b>${validadeStr}</b><br><br>
    Qualquer dúvida, conte com nosso time.<br>
    Equipe SIEG.
  `;

  MailApp.sendEmail({
    to: dados.email,
    subject: assunto,
    htmlBody: corpoHtml,
    attachments: [pdfFile.getAs('application/pdf')],
    name: 'Equipe SIEG'
  });
}
