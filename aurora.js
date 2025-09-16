/*
==================================================
📌 SCRIPT: Monitoramento, Alertas e Geração de Relatórios
==================================================

🧠 OBJETIVO:
Este script é executado automaticamente quando a planilha é editada.
Ele realiza as seguintes funções principais:

1️⃣ Quando o status (coluna A) muda para "ANÁLISE RETORNO":
   - Envia um e-mail automático para o auditor responsável (coluna R)
   - Insere um balão de comentário na célula com o nome do auditor

2️⃣ Quando o status muda para qualquer valor (exceto vazio ou "ANÁLISE"):
   - Atualiza a data de tramitação (coluna P) com a data atual

3️⃣ Verifica se a data de abertura (coluna D) ultrapassa 30 dias após a data de desligamento (coluna N):
   - Se ultrapassar:
     • Colore a célula da data de abertura em vermelho
     • Adiciona um balão de comentário com o aviso: 
       "Abertura feita após 30 dias do desligamento. Verificar pendência ou justificativa."
   - Se estiver dentro do prazo, mantém a cor padrão e remove comentários

4️⃣ Gera relatórios e memorandos automaticamente com botões

✉️ Os e-mails dos auditores estão definidos no objeto `mapaEmails`.

📄 Planilha usada: CONTROLE EXPANSÃO E MOVIMENTAÇÃO DE SERVIDORES

👩🏻‍💻 Responsável pelo script: Luana  
📧 E-mail: luana.41331@santanadeparnaiba.sp.gov.br  
📞 Ramal: 8819

🕐 Última atualização: 16/07/2025
*/

// ============================================================================
// CONFIGURAÇÕES E CONSTANTES - MEMORANDO ADP
// ============================================================================
const MEMORANDO_CONFIG = {
	TEMPLATE_ID: "1EErrs3JO1S2TvMMpOqWHoyPSusVQwBIITW48PRwaRTA",
	FONT_FAMILY: "Calibri",
	FONT_SIZE: 12,
	BORDER_COLOR: "#cccccc",
	HEADER_BG_COLOR: "#f3f3f3",
	HEADER_TEXT_COLOR: "#45818e",
	TABLE_ROW_HEIGHT: 0.00688,
	TABLE_CELL_WIDTH: 0.02302,
	CURRENT_YEAR: new Date().getFullYear(),
	MESES: [
	  "janeiro", "fevereiro", "março", "abril", "maio", "junho",
	  "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
	],
	NUMEROS_EXTENSO: [
	  "zero", "um", "dois", "três", "quatro", "cinco", 
	  "seis", "sete", "oito", "nove", "dez"
	]
  };
  
  // Mapeamento das colunas da planilha para memorando
  const COLUNAS_MEMO = {
	SISGEP: 2,        // B - PROCESSO
	TIPO: 3,          // C - TIPO MOVIMENTAÇÃO
	STATUS: 1,        // A - STATUS (corrigido)
	SECRETARIA: 6,    // F - SECRETARIA
	CARGO: 8,         // H - CARGO
	QUANTIDADE: 9,    // I - QTD SOLICITADA
	NOME_SERVIDOR: 12,    // L - Nome
	PRONTUARIO: 13,       // M - Prontuário
	DESLIGAMENTO: 14,     // N - Desligamento/Retorno
	DEPARTAMENTO: 7,      // G - DEPARTAMENTO (corrigido)
	DETALHAMENTO: 15      // O - DETALHAMENTO
  };
  
  // === GATILHO DE EDIÇÃO ===
  function onEdit(e) {
	  if (!e) return;
	  handleSpreadsheetEdit(e);
  }
  
  // === GATILHO DE TEMPO (ex: a cada 5 min) ===
  function timeDrivenFunction() {
	  notificarAuditor();
  }
  
  // === FUNÇÃO PRINCIPAL DE EDIÇÃO ===
  function handleSpreadsheetEdit(e) {
	  var sheet = e.source.getActiveSheet();
	  var range = e.range;
	  var linha = range.getRow();
	  var coluna = range.getColumn();
	  var valorSelecionado = range.getValue();
	  var sheetName = sheet.getName();
  
	  // Utilitário: normaliza texto (lowercase + remoção de acentos + trim)
	  function normalizeText(value) {
		  if (value === null || typeof value === "undefined") return "";
		  return value
			  .toString()
			  .normalize('NFD')
			  .replace(/[\u0300-\u036f]/g, '')
			  .toLowerCase()
			  .trim();
	  }
  
	  // COLUNAS IMPORTANTES - ATUALIZADAS CONFORME NOVA ORDEM
	  var colunaStatus = 1; // A - STATUS
	  var colunaProcesso = 2; // B - PROCESSO
	  var colunaTipo = 3; // C - TIPO MOVIMENTAÇÃO
	  var colunaAbertura = 4; // D - DATA ABERTURA
	  var colunaEnvioADP = 5; // E - ENVIO ADP
	  var colunaSecretaria = 6; // F - SECRETARIA
	  var colunaDepartamento = 7; // G - DEPARTAMENTO
	  var colunaCargo = 8; // H - CARGO
	  var colunaQuantidade = 9; // I - QTD SOLICITADA
	  var colunaSalario = 10; // J - SALÁRIO MENSAL
	  var colunaCustoAnual = 11; // K - CUSTO ANUAL
	  var colunaNome = 12; // L - Nome
	  var colunaProntuario = 13; // M - Prontuário
	  var colunaDesligamento = 14; // N - Desligamento/ Retorno
	  var colunaDetalhamento = 15; // O - DETALHAMENTO
	  var colunaDataTramitacao = 16; // P - DATA TRAMITAÇÃO
	  var colunaFluxo = 17; // Q - FLUXO
	  var colunaAuditor = 18; // R - AUDITOR
	  var colunaMemo = 19; // S - MEMO
	  var colunaRelatorio = 20; // T - RELATÓRIO
  
	  // BLOCO 1 – Atualiza data na coluna P se o status mudou (exceto "ANÁLISE")
	  if (sheetName === "CONTROLE 2025" && coluna === colunaStatus && linha > 1) {
		  var cellData = sheet.getRange(linha, colunaDataTramitacao);
		  var statusAtualNorm = normalizeText(valorSelecionado);
		  var statusAntigoNorm = typeof e.oldValue !== "undefined" ? normalizeText(e.oldValue) : null;
  
		  // Só prosseguir se houve alteração real do valor de status
		  if (statusAntigoNorm === null || statusAntigoNorm !== statusAtualNorm) {
			  if (!sheet.isRowHiddenByFilter(linha)) {
				  if (statusAtualNorm === "" || statusAtualNorm === "analise") {
					  if (cellData.getValue() !== "") {
						  cellData.setValue("");
					  }
				  } else {
					  cellData.setValue(new Date());
					  cellData.setNumberFormat('dd/MM/yyyy hh:mm');
				  }
			  }
		  }
	  }
  
	  // BLOCO 2 – Verifica se abertura foi após 30 dias do desligamento
	  var dataDesligamento = sheet.getRange(linha, colunaDesligamento).getValue();
	  var dataAbertura = sheet.getRange(linha, colunaAbertura).getValue();
	  var cellAbertura = sheet.getRange(linha, colunaAbertura);
  
	  if (dataDesligamento instanceof Date && dataAbertura instanceof Date) {
		  var prazoLimite = new Date(dataDesligamento);
		  prazoLimite.setDate(prazoLimite.getDate() + 30);
  
		  if (!sheet.isRowHiddenByFilter(linha)) {
			  if (dataAbertura > prazoLimite) {
				  if (cellAbertura.getFontColor() !== "red") {
					  cellAbertura.setFontColor("red");
				  }
				  var msg = "Abertura feita após 30 dias do desligamento. Verificar pendência ou justificativa.";
				  if (cellAbertura.getComment() !== msg) {
					  cellAbertura.setComment(msg);
				  }
			  } else {
				  if (cellAbertura.getFontColor() !== "black") {
					  cellAbertura.setFontColor("black");
				  }
				  if (cellAbertura.getComment()) {
					  cellAbertura.setComment("");
				  }
			  }
		  }
	  }
  
	  // BLOCO 3 – Armazena linha se status for "ANÁLISE RETORNO"
	  var scriptProps = PropertiesService.getScriptProperties();
	  if (coluna === colunaStatus && valorSelecionado === "ANÁLISE RETORNO") {
		  var linhasStr = scriptProps.getProperty("linhasNotificar") || "[]";
		  var linhas = JSON.parse(linhasStr);
  
		  if (!linhas.includes(linha)) {
			  linhas.push(linha);
			  scriptProps.setProperty("linhasNotificar", JSON.stringify(linhas));
			  console.log("Linha adicionada para notificação: " + linha);
		  }
	  }
  
	  // BLOCO 4 – Limpeza e verificação de prontuários duplicados (coluna M = 13)
	  var abasPermitidas = ["CONTROLE 2025"];
  
	  if (
		  abasPermitidas.includes(sheetName) &&
		  valorSelecionado !== "" &&
		  coluna === colunaProntuario
	  ) {
		  verificarProntuariosDuplicados(sheet, linha, colunaProntuario);
	  }
  }
  
  // === FUNÇÃO PARA VERIFICAR PRONTUÁRIOS DUPLICADOS ===
  function verificarProntuariosDuplicados(sheet, linha, colunaProntuario) {
	  var range = sheet.getRange(linha, colunaProntuario);
	  var valorSelecionado = range.getValue();
	  var valorLimpo = String(valorSelecionado).replace(/[.,]/g, "").trim();
  
	  if (valorLimpo !== String(valorSelecionado)) {
		  range.setValue(valorLimpo);
	  }
  
	  var prontuariosAtuais = valorLimpo.split(/\s+|\n+/);
	  var duplicado = false;
  
	  for (var i = 2; i < linha; i++) {
		  var valorAnterior = sheet.getRange(i, colunaProntuario).getValue();
		  var valorAnteriorLimpo = String(valorAnterior).replace(/[.,]/g, "").trim();
		  if (!valorAnteriorLimpo) continue;
  
		  var prontuariosAnteriores = valorAnteriorLimpo.split(/\s+|\n+/);
		  for (var atual of prontuariosAtuais) {
			  if (prontuariosAnteriores.includes(atual)) {
				  duplicado = true;
				  break;
			  }
		  }
		  if (duplicado) break;
	  }
  
	  if (duplicado) {
		  range.setFontColor("red");
		  range.setComment(
			  "⚠️ Este prontuário (ou parte dele) já foi usado acima. Verifique possível duplicidade.",
		  );
	  } else {
		  range.setFontColor("black");
		  range.setComment("");
	  }
  }
  
  // === FUNÇÃO DE NOTIFICAÇÃO POR EMAIL (GATILHO DE TEMPO) ===
  function notificarAuditor() {
	  var ss = SpreadsheetApp.getActiveSpreadsheet();
	  var sheet = ss.getSheetByName("CONTROLE 2025");
	  if (!sheet) {
		  console.error("Aba 'CONTROLE 2025' não encontrada.");
		  return;
	  }
  
	  var scriptProps = PropertiesService.getScriptProperties();
	  var linhasStr = scriptProps.getProperty("linhasNotificar") || "[]";
	  var linhas = JSON.parse(linhasStr);
	  if (linhas.length === 0) return;
  
	  var mapaEmails = {
		  Luana: "luana.41331@santanadeparnaiba.sp.gov.br",
		  Natalice: "natalice.36293@santanadeparnaiba.sp.gov.br",
		  // Adicione outros aqui conforme necessário
	  };
  
	  var novasLinhas = [];
  
	  for (var i = 0; i < linhas.length; i++) {
		  var linha = linhas[i];
		  if (linha < 2) continue;
  
		  var nomeAuditor = sheet.getRange(linha, 18).getValue(); // Coluna R - AUDITOR
		  var processo = sheet.getRange(linha, 2).getValue(); // Coluna B - PROCESSO
		  var secretaria = sheet.getRange(linha, 6).getValue(); // Coluna F - SECRETARIA
		  var statusAtual = sheet.getRange(linha, 1).getValue(); // Coluna A - STATUS
  
		  if (statusAtual !== "ANÁLISE RETORNO") {
			  console.log(
				  `Linha ${linha} – status não é mais 'ANÁLISE RETORNO'. Pulando.`,
			  );
			  continue;
		  }
  
		  if (!nomeAuditor) {
			  console.warn(
				  `Linha ${linha} – auditor em branco. Mantendo para nova tentativa.`,
			  );
			  novasLinhas.push(linha);
			  continue;
		  }
  
		  nomeAuditor = nomeAuditor.trim();
		  var emailAuditor = mapaEmails[nomeAuditor];
  
		  if (emailAuditor) {
			  var assunto = `Processo ${processo} retornou para análise`;
			  var mensagem = `Olá ${nomeAuditor},\n\nO processo ${processo} da secretaria ${secretaria} foi atualizado com o status "ANÁLISE RETORNO".\n\nPor favor, verifique se há pendências ou se pode dar continuidade à análise.`;
  
			  try {
				  MailApp.sendEmail(emailAuditor, assunto, mensagem);
				  console.log(
					  `E-mail enviado para ${nomeAuditor} sobre processo ${processo}`,
				  );
			  } catch (erro) {
				  console.error(`Erro ao enviar e-mail (linha ${linha}): ${erro}`);
				  novasLinhas.push(linha);
			  }
		  } else {
			  console.warn(
				  `Linha ${linha} – e-mail não encontrado para auditor: ${nomeAuditor}`,
			  );
			  novasLinhas.push(linha); // Tenta novamente em outro ciclo se o nome for corrigido
		  }
	  }
  
	  scriptProps.setProperty("linhasNotificar", JSON.stringify(novasLinhas));
  }
  
  // ID da pasta onde os documentos serão salvos
  const PASTA_DOCUMENTOS_ID = "1OBHunABxlCl0WHsBKFse-6icL8Aat4Py";
  
  // ==================================================
  // 🗂️ FUNÇÃO PARA MOVER ARQUIVO PARA PASTA ESPECÍFICA
  // ==================================================
  function moverArquivoParaPasta(docId, nomeArquivo) {
	  try {
		  var arquivo = DriveApp.getFileById(docId);
		  var pastaDestino = DriveApp.getFolderById(PASTA_DOCUMENTOS_ID);
  
		  // Remove o arquivo da pasta raiz (se estiver lá)
		  var pastasOriginais = arquivo.getParents();
		  while (pastasOriginais.hasNext()) {
			  var pastaOriginal = pastasOriginais.next();
			  pastaOriginal.removeFile(arquivo);
		  }
  
		  // Adiciona o arquivo à pasta de destino
		  pastaDestino.addFile(arquivo);
  
		  Logger.log(`Arquivo "${nomeArquivo}" movido para a pasta com sucesso.`);
		  return true;
	  } catch (error) {
		  Logger.log(`Erro ao mover arquivo "${nomeArquivo}": ${error.toString()}`);
		  return false;
	  }
  }
  
  // ==================================================
  // 🔗 FUNÇÃO PARA ADICIONAR LINK NA ABA 'Controle de Memos'
  // ==================================================
  function adicionarLinkControleMemos(tipo, numeroDoc, secretaria, cargo, processo, url) {
	  try {
		  var ss = SpreadsheetApp.getActiveSpreadsheet();
		  var sheetMemos = ss.getSheetByName("Controle de Memos");
		  
		  if (!sheetMemos) {
			  Logger.log("Aba 'Controle de Memos' não encontrada.");
			  return false;
		  }
  
		  // Encontra a próxima linha vazia
		  var ultimaLinha = sheetMemos.getLastRow();
		  var proximaLinha = ultimaLinha + 1;
  
		  // Data atual formatada
		  var hoje = new Date();
		  var dataFormatada = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM");
  
		  // Limpa aspas duplas da URL e do número do documento para evitar conflitos
		  var urlLimpa = url.toString().replace(/"/g, '');
		  var numeroDocLimpo = numeroDoc.toString().replace(/"/g, '');
  
		  if (tipo === "memorando") {
			  // Coluna B - Memo - CORREÇÃO: usando aspas simples na fórmula
			  sheetMemos.getRange(proximaLinha, 2).setFormula('=HYPERLINK("' + urlLimpa + '";"' + numeroDocLimpo + '")');
			  // Coluna C - Data
			  sheetMemos.getRange(proximaLinha, 3).setValue(dataFormatada);
			  // Coluna D - Secretaria
			  sheetMemos.getRange(proximaLinha, 4).setValue(secretaria);
			  // Coluna E - Cargo
			  sheetMemos.getRange(proximaLinha, 5).setValue(cargo);
			  // Coluna G - Processo
			  sheetMemos.getRange(proximaLinha, 7).setValue(processo);
		  } else if (tipo === "relatorio") {
			  // Coluna F - Relatórios - CORREÇÃO: usando aspas simples na fórmula
			  sheetMemos.getRange(proximaLinha, 6).setFormula('=HYPERLINK("' + urlLimpa + '";"' + numeroDocLimpo + '")');
			  // Coluna C - Data
			  sheetMemos.getRange(proximaLinha, 3).setValue(dataFormatada);
			  // Coluna D - Secretaria
			  sheetMemos.getRange(proximaLinha, 4).setValue(secretaria);
			  // Coluna E - Cargo
			  sheetMemos.getRange(proximaLinha, 5).setValue(cargo);
			  // Coluna G - Processo
			  sheetMemos.getRange(proximaLinha, 7).setValue(processo);
		  }
  
		  Logger.log(`Link do ${tipo} adicionado na aba 'Controle de Memos' com sucesso.`);
		  return true;
	  } catch (error) {
		  Logger.log(`Erro ao adicionar link na aba 'Controle de Memos': ${error.toString()}`);
		  return false;
	  }
  }
  
  // ============================================================================
  // 📄 GERADOR DE MEMORANDO ADP MELHORADO
  // ============================================================================
  
  // Validação e extração de dados
  function validarLinhaSelecionada(linha) {
	if (linha <= 1) {
	  SpreadsheetApp.getUi().alert(
		"Erro", 
		"Por favor, selecione uma linha com dados válidos.", 
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	  return false;
	}
	return true;
  }
  
  function validarDados(dados) {
	const camposObrigatorios = ['secretaria', 'cargo', 'tipo', 'status'];
	const camposFaltantes = camposObrigatorios.filter(campo => !dados[campo]);
	
	if (camposFaltantes.length > 0) {
	  SpreadsheetApp.getUi().alert(
		"Dados Incompletos", 
		`Os seguintes campos são obrigatórios: ${camposFaltantes.join(', ')}`, 
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	  return false;
	}
	
	return true;
  }
  
  function extrairDadosDaPlanilha(sheet, linha) {
	const lerCelula = (coluna, padrao = "") => {
	  const valor = sheet.getRange(linha, coluna).getValue();
	  return valor ? valor.toString().trim() : padrao;
	};
  
	return {
	  sisgep: lerCelula(COLUNAS_MEMO.SISGEP),
	  secretaria: lerCelula(COLUNAS_MEMO.SECRETARIA),
	  tipo: lerCelula(COLUNAS_MEMO.TIPO).toUpperCase(),
	  cargo: lerCelula(COLUNAS_MEMO.CARGO),
	  quantidade: sheet.getRange(linha, COLUNAS_MEMO.QUANTIDADE).getValue() || 1,
	  nomeServidor: lerCelula(COLUNAS_MEMO.NOME_SERVIDOR),
	  prontuario: lerCelula(COLUNAS_MEMO.PRONTUARIO),
	  desligamento: sheet.getRange(linha, COLUNAS_MEMO.DESLIGAMENTO).getValue(),
	  departamento: lerCelula(COLUNAS_MEMO.DEPARTAMENTO),
	  status: lerCelula(COLUNAS_MEMO.STATUS).toUpperCase(),
	  justificativa: lerCelula(COLUNAS_MEMO.DETALHAMENTO) // Usando DETALHAMENTO como justificativa
	};
  }
  
  function formatarData(data) {
	if (data instanceof Date) {
	  return Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
	}
	return data || "";
  }
  
  function gerarNumeroMemo() {
	return Math.floor(Math.random() * 9000) + 1000;
  }
  
  function formatarDataExtenso(data = new Date()) {
	return `${data.getDate()} de ${MEMORANDO_CONFIG.MESES[data.getMonth()]} de ${data.getFullYear()}`;
  }
  
  function numeroParaExtenso(num) {
	return (num >= 0 && num < MEMORANDO_CONFIG.NUMEROS_EXTENSO.length) 
	  ? MEMORANDO_CONFIG.NUMEROS_EXTENSO[num] 
	  : num.toString();
  }
  
  function criarDocumentoComTemplate(dados) {
	const numeroMemo = gerarNumeroMemo();
	const dataFormatada = formatarDataExtenso();
	const desligamentoFormatado = formatarData(dados.desligamento);
	const nomeDoc = `MEMORANDO nº ${numeroMemo}/${MEMORANDO_CONFIG.CURRENT_YEAR} - ADP`;
  
	const docFile = DriveApp.getFileById(MEMORANDO_CONFIG.TEMPLATE_ID).makeCopy(nomeDoc);
	const doc = DocumentApp.openById(docFile.getId());
	const body = doc.getBody();
  
	return {
	  doc,
	  body,
	  numeroMemo,
	  dataFormatada,
	  desligamentoFormatado,
	  nomeDoc
	};
  }
  
  function configurarDocumento(body) {
	body.clear();
	
	const normalStyle = {
	  [DocumentApp.Attribute.FONT_FAMILY]: MEMORANDO_CONFIG.FONT_FAMILY,
	  [DocumentApp.Attribute.FONT_SIZE]: MEMORANDO_CONFIG.FONT_SIZE,
	  [DocumentApp.Attribute.BOLD]: false
	};
	
	body.setAttributes(normalStyle);
  }
  
  function adicionarCabecalho(body, numeroMemo, dataFormatada) {
	// Título (sem quebra de linha depois)
	const titulo = body.appendParagraph(`MEMORANDO Nº ${numeroMemo}/${MEMORANDO_CONFIG.CURRENT_YEAR} - ADP`);
	titulo.setBold(true);
	titulo.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
	// Data (logo na linha seguinte, sem espaço)
	const data = body.appendParagraph(`Santana de Parnaíba, ${dataFormatada}`);
	data.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
	body.appendParagraph("");
  }
  
  function adicionarInformacoesPrincipais(body, dados) {
	const adicionarLinha = (label, valor) => {
	  const p = body.appendParagraph("");
	  p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
	  const text = p.editAsText();
	  text.appendText(label).setBold(true);
	  text.appendText(valor);
	};
  
	adicionarLinha("De: ", "Secretaria Municipal de Administração - ADP");
	adicionarLinha("Para: ", "Sr. Secretário José Roberto Martins Santos");
  
	body.appendParagraph("");
	const refTexto = gerarTextoReferencia(dados);
	adicionarLinha("Ref.: ", refTexto);
	body.appendParagraph("");
  
	// Título centralizado (verifica se é Banco de Talentos)
	const tituloAnalise = body.appendParagraph(
	  dados.status.includes("DEFERIDO BT") 
		? "Análise de Demanda de Pessoal - Banco de Talentos" 
		: "Análise de Demanda de Pessoal"
	);
	tituloAnalise.setBold(true);
	tituloAnalise.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
	body.appendParagraph("");
  
	// Saudação (SEM negrito)
	const saudacao = body.appendParagraph("Senhor Secretário,");
	saudacao.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
	saudacao.setBold(false);
	body.appendParagraph("");
  }
  
  function gerarTextoReferencia(dados) {
	const tiposReferencia = {
	  "PROCESSO SELETIVO": () => `Substituição por Processo Seletivo de ${dados.cargo}`,
	  "AMPLIAÇÃO": () => `${dados.tipo} ${dados.cargo}`,
	  "PERMUTA": () => `${dados.tipo} ${dados.cargo}`,
	  "default": () => `Substituição de ${dados.cargo}`
	};
  
	for (const [tipo, gerador] of Object.entries(tiposReferencia)) {
	  if (tipo !== 'default' && dados.tipo.includes(tipo)) {
		return gerador();
	  }
	}
	
	return tiposReferencia.default();
  }
  
  function adicionarConteudoPorTipo(body, dados, desligamentoFormatado) {
	// Primeiro verifica tipos específicos
	if (dados.tipo.includes("PERMUTA")) {
	  processarPermuta(body, dados);
	  return;
	}
	
	if (dados.tipo.includes("AMPLIAÇÃO")) {
	  processarAmpliacao(body, dados);
	  return;
	}
	
	if (dados.tipo.includes("PROCESSO SELETIVO")) {
	  processarProcessoSeletivo(body, dados, desligamentoFormatado);
	  return;
	}
	
	// Se não é um tipo específico, usa o status para decidir
	if (dados.status.includes("INDEFERIDO")) {
	  processarIndeferido(body, dados, desligamentoFormatado);
	  return;
	}
	
	if (dados.status.includes("DEFERIDO BT")) {
	  processarDeferidoBT(body, dados, desligamentoFormatado);
	  return;
	}
	
	// Caso padrão: usa DEFERIDO comum
	processarDeferido(body, dados, desligamentoFormatado);
  }
  
  function processarPermuta(body, dados) {
	const paragrafo = body.appendParagraph("Encaminhamos o presente expediente referente à solicitação de permuta entre os cargos:");
	paragrafo.setBold(false);
	body.appendParagraph("");
	
	const placeholder = body.appendParagraph("[PREENCHER COM OS DADOS MANUALMENTE]");
	placeholder.setItalic(true);
	placeholder.setBold(false);
	body.appendParagraph("");
  
	adicionarConclusaoComDestaque(body, "deferimento", " da solicitação. Encaminho para demais providências.");
  }
  
  function processarAmpliacao(body, dados) {
	const quantidade = dados.quantidade || 1;
	const plural = quantidade > 1 ? "s" : "";
	
	const paragrafo = body.appendParagraph(
	  `Encaminhamos o presente expediente referente à solicitação de ampliação de ${quantidade} ` +
	  `(${numeroParaExtenso(quantidade)}) ${dados.cargo}${plural}, conforme relatório em anexo.`
	);
	paragrafo.setBold(false);
	body.appendParagraph("");
  
	adicionarConclusaoComDestaque(body, "deferimento", " da solicitação. Assim, encaminhamos para demais providências.");
  }
  
  function processarProcessoSeletivo(body, dados, desligamentoFormatado) {
	const quantidade = dados.quantidade || 1;
	
	const paragrafo = body.appendParagraph(
	  `Encaminhamos o presente expediente referente à solicitação de substituição por meio de ` +
	  `Processo Seletivo de ${quantidade} (${numeroParaExtenso(quantidade)}) servidor(a) no cargo de ${dados.cargo}, ` +
	  `conforme detalhado abaixo:`
	);
	paragrafo.setBold(false);
	body.appendParagraph("");
  
	criarTabelaServidorDesligado(body, dados, desligamentoFormatado);
	body.appendParagraph("");
  
	adicionarConclusaoComDestaque(body, "deferimento", " da solicitação. Assim, encaminhamos para demais providências.");
  }
  
  function processarIndeferido(body, dados, desligamentoFormatado) {
	const quantidade = dados.quantidade || 1;
	const plural = quantidade > 1 ? "s" : "";
	
	const paragrafo1 = body.appendParagraph(
	  `Encaminhamos o presente expediente referente à solicitação de substituição de ${quantidade} ` +
	  `(${numeroParaExtenso(quantidade)}) ${dados.cargo}${plural}, conforme detalhamento abaixo:`
	);
	paragrafo1.setBold(false);
	body.appendParagraph("");
  
	criarTabelaServidorDesligado(body, dados, desligamentoFormatado);
	body.appendParagraph("");
  
	const justificativa = dados.justificativa || 'PREENCHER COM A JUSTIFICATIVA';
	adicionarConclusaoComDestaque(
	  body, 
	  "indeferimento", 
	  ` da solicitação, considerando que, [${justificativa}]`
	);
  }
  
  function processarDeferidoBT(body, dados, desligamentoFormatado) {
	const quantidade = dados.quantidade || 1;
	
	const paragrafo3 = body.appendParagraph(
	  `Encaminhamos o presente expediente referente à solicitação de substituição de ${quantidade} ` +
	  `(${numeroParaExtenso(quantidade)}) servidor(a) no cargo de ${dados.cargo}, conforme detalhado abaixo:`
	);
	paragrafo3.setBold(false);
	body.appendParagraph("");
  
	// Tabela 1: Servidor desligado
	criarTabelaServidorDesligado(body, dados, desligamentoFormatado);
	body.appendParagraph("");
  
	const conclusao1 = body.appendParagraph("Após a devida análise, manifestamos parecer favorável ao ");
	conclusao1.setBold(false);
	const conclText1 = conclusao1.editAsText();
	const inicio1 = conclText1.getText().length;
	conclText1.appendText("deferimento");
	conclText1.setBold(inicio1, conclText1.getText().length - 1, true);
	conclText1.setUpperCase(inicio1, conclText1.getText().length - 1, true);
	conclText1.appendText(" da solicitação, com atendimento por meio da indicação de servidor(a) disponível no Banco de Talentos, conforme detalhado a seguir:");
	conclusao1.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
	body.appendParagraph("");
  
	// Tabela 2: Banco de Talentos (VAZIA para preenchimento manual)
	criarTabelaBancoTalentosVazia(body);
  }
  
  function processarDeferido(body, dados, desligamentoFormatado) {
	const quantidade = dados.quantidade || 1;
	
	const paragrafo2 = body.appendParagraph(
	  `Encaminhamos o presente expediente referente à solicitação de substituição de ${quantidade} ` +
	  `(${numeroParaExtenso(quantidade)}) servidor(a) no cargo de ${dados.cargo}, conforme detalhado abaixo:`
	);
	paragrafo2.setBold(false);
	body.appendParagraph("");
  
	criarTabelaServidorDesligado(body, dados, desligamentoFormatado);
	body.appendParagraph("");
  
	adicionarConclusaoComDestaque(body, "deferimento", " da solicitação. Assim, encaminhamos para demais providências.");
  }
  
  function adicionarConclusaoComDestaque(body, palavraDestaque, textoComplementar) {
	const conclusao = body.appendParagraph("Após a devida análise, manifestamos parecer favorável ao ");
	conclusao.setBold(false);
	const texto = conclusao.editAsText();
	const inicio = texto.getText().length;
	
	texto.appendText(palavraDestaque);
	texto.setBold(inicio, texto.getText().length - 1, true);
	texto.setUpperCase(inicio, texto.getText().length - 1, true);
	texto.appendText(textoComplementar);
	
	conclusao.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  }
  
  function criarTabelaServidorDesligado(body, dados, desligamentoFormatado) {
	const tabela = body.appendTable([
	  ["Secretaria", "Nome", "Prontuário", "Desligamento", "Departamento"],
	  [
		dados.secretaria,
		dados.nomeServidor,
		dados.prontuario,
		desligamentoFormatado,
		dados.departamento
	  ]
	]);
	formatarTabela(tabela, true);
  }
  
  function criarTabelaBancoTalentosVazia(body) {
	const tabela = body.appendTable([
	  ["Secretaria", "Nome", "Prontuário"],
	  ["", "", ""] // Linha vazia para preenchimento manual
	]);
	formatarTabela(tabela, true);
  }
  
  function formatarTabela(tabela, comCabecalho = false) {
	tabela.setBorderWidth(1);
	tabela.setBorderColor(MEMORANDO_CONFIG.BORDER_COLOR);
  
	for (let i = 0; i < tabela.getNumRows(); i++) {
	  const row = tabela.getRow(i);
	  row.setMinimumHeight(MEMORANDO_CONFIG.TABLE_ROW_HEIGHT);
	  
	  for (let j = 0; j < row.getNumCells(); j++) {
		const cell = row.getCell(j);
		cell.setWidth(MEMORANDO_CONFIG.TABLE_CELL_WIDTH);
		cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
		cell.getChild(0).asParagraph().setSpacingAfter(0.07);
	  }
	}
  
	if (comCabecalho && tabela.getNumRows() > 0) {
	  const headerRow = tabela.getRow(0);
	  for (let j = 0; j < headerRow.getNumCells(); j++) {
		const cell = headerRow.getCell(j);
		cell.setBackgroundColor(MEMORANDO_CONFIG.HEADER_BG_COLOR);
		const text = cell.editAsText();
		text.setBold(true);
		text.setForegroundColor(MEMORANDO_CONFIG.HEADER_TEXT_COLOR);
	  }
	}
  }
  
  function adicionarAssinatura(body) {
	body.appendParagraph("");
	const atenciosamente = body.appendParagraph("Atenciosamente,");
	atenciosamente.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
	atenciosamente.setBold(false);
	
	const assinatura = body.appendParagraph("Secretaria Municipal de Administração");
	assinatura.setBold(true);
	assinatura.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  
  function abrirDocumento(url) {
	const htmlOutput = HtmlService.createHtmlOutput(`
	  <script>
		window.open('${url}', '_blank');
		google.script.host.close();
	  </script>
	`);
	SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Memorando ADP Gerado com Sucesso!");
  }
  
  function tratarErro(error) {
	Logger.log(`Erro ao gerar memorando: ${error.toString()}`);
	console.error("Stack trace:", error.stack);
	
	SpreadsheetApp.getUi().alert(
	  "Erro", 
	  `Ocorreu um erro ao gerar o memorando: ${error.message}. Verifique os dados e tente novamente.`, 
	  SpreadsheetApp.getUi().ButtonSet.OK
	);
  }
  
  // FUNÇÃO PRINCIPAL MELHORADA DO MEMORANDO ADP
  function gerarMemorandoADP() {
	try {
	  const sheet = SpreadsheetApp.getActiveSheet();
	  const linha = sheet.getActiveRange().getRow();
  
	  // Validação inicial
	  if (!validarLinhaSelecionada(linha)) {
		return;
	  }
  
	  // Extrair e validar dados
	  const dados = extrairDadosDaPlanilha(sheet, linha);
	  if (!validarDados(dados)) {
		return;
	  }
  
	  // Gerar documento
	  const docInfo = criarDocumentoComTemplate(dados);
	  const body = docInfo.body;
  
	  // Configurar estilos e estrutura
	  configurarDocumento(body);
	  adicionarCabecalho(body, docInfo.numeroMemo, docInfo.dataFormatada);
	  adicionarInformacoesPrincipais(body, dados);
	  adicionarConteudoPorTipo(body, dados, docInfo.desligamentoFormatado);
	  adicionarAssinatura(body);
  
	  // Finalizar documento
	  docInfo.doc.saveAndClose();
  
	  // Mover arquivo para pasta específica
	  if (PASTA_DOCUMENTOS_ID) {
		moverArquivoParaPasta(docInfo.doc.getId(), docInfo.nomeDoc);
	  }
  
	  // Adicionar link na aba 'Controle de Memos' - coluna B
	  const url = docInfo.doc.getUrl();
	  const numeroDocFormatado = `${docInfo.numeroMemo}/${MEMORANDO_CONFIG.CURRENT_YEAR}`;
	  adicionarLinkControleMemos("memorando", numeroDocFormatado, dados.secretaria, dados.cargo, dados.sisgep, url);
  
	  // Abrir documento
	  abrirDocumento(url);
  
	} catch (error) {
	  tratarErro(error);
	}
  }