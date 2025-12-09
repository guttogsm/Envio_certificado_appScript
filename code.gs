// =======================
// CONFIGURAÇÕES
// =======================
const CONFIG = {
  SHEET_NAME: 'Form_Responses',
  TEMPLATE_ID: 'COLOQUE_AQUI_O_ID_DO_MODELO_SLIDES',
  FOLDER_ID: 'COLOQUE_AQUI_O_ID_DA_PASTA_PDFS'
};

// Índices das colunas (1 = A, 2 = B, ...)
const COL = {
  DATA_HORA: 1,
  EMAIL: 2,
  PONTUACAO: 3,
  NOME: 4,
  CNPJ: 5,
  ANALISTA: 6,
  CURSO: 17,
  DATA_INICIO: 18,
  DATA_TERMINO: 19,
  STATUS: 20,
  CERT_GERADO: 21,
  LINK_CERT: 22,
  TOKEN: 23,
  VALIDADE: 24
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
  const pdf = gerarCertificadoPDF(dados, token);
  const link = pdf.getUrl();

  sheet.getRange(row, COL.CERT_GERADO).setValue('SIM');
  sheet.getRange(row, COL.LINK_CERT).setValue(link);
  sheet.getRange(row, COL.TOKEN).setValue(token);

  enviarEmailComCertificado(dados, link, token);
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

  const blob = copia.getAs('application/pdf');
  const pdfFile = folder.createFile(blob).setName(nomeArquivoBase + '.pdf');

  copia.setTrashed(true);

  return pdfFile;
}

// =======================
// ENVIAR E-MAIL
// =======================

function enviarEmailComCertificado(dados, link, token) {
  const tz = Session.getScriptTimeZone();
  const validadeStr = dados.validade
    ? Utilities.formatDate(new Date(dados.validade), tz, 'dd/MM/yyyy')
    : 'indeterminada';

  const assunto = `Certificado - ${dados.curso || 'Treinamento SIEG'}`;

  const corpoHtml = `
    Olá, ${dados.nome}!<br><br>
    Obrigado por participar do treinamento <b>${dados.curso || 'SIEG'}</b>.<br><br>
    Seu certificado já está disponível:<br>
    <a href="${link}">Acessar certificado</a><br><br>
    Token: <b>${token}</b><br>
    Validade: <b>${validadeStr}</b><br><br>
    Qualquer dúvida, conte com nosso time.<br>
    Equipe SIEG.
  `;

  MailApp.sendEmail({
    to: dados.email,
    subject: assunto,
    htmlBody: corpoHtml
  });
}
