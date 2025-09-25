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

🕐 Última atualização: 24/09/2025
*/

// ============================================================================
// CONFIGURAÇÕES E CONSTANTES - SISTEMA DE TEMPLATES E NUMERAÇÃO
// ============================================================================
const MEMORANDO_CONFIG = {
	SHEET_NAME: "CONTROLE 2025",
	SOURCE_SHEET_FULLNAME: "CONTROLE EXPANSÃO E MOVIMENTAÇÃO DE SERVIDORES",
	MEMOS_SHEET_NAME: "Controle de Memos",
	PLANILHA_MEMOS_ID: "1vdQa93PB1CyZP0PSAN9AHc5ukmc6L09WUP2AnLagaUE",
	ABA_NUMERACAO: "Página1",
	TEMPLATES: {
	  SUBSTITUICAO: "1uIg38a9mXbZMnFv2Z65Cn2jOVWyWmuUOXIHrjD-3KRo",
	  DEFERIDO_BT: "1tGAcqT4x1kzd5s3tYO1sl7l8EFGOib-15lASeVykKxo",
	  AMPLIACAO: "12KlR22gNE53V3p833EkBGz9ffw0POXgU5xqAzibowiQ",
	  PERMUTA: "17x78OuQZSlSUUIM0Rleh_Tl_mKuUIZAAqgY7LPtI7AQ",
	  PROCESSO_SELETIVO: "18_ExWv5EIGbZyOdGUjpc-T76WjHzLPrW8rkUF2zAF8U"
	},
	PASTA_DOCUMENTOS_ID: "1OBHunABxlCl0WHsBKFse-6icL8Aat4Py",
	MAPA_EMAILS: {
	  Luana: "luana.41331@santanadeparnaiba.sp.gov.br",
	  Natalice: "natalice.36293@santanadeparnaiba.sp.gov.br"
	},
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
	STATUS: 1,        // A - STATUS
	SECRETARIA: 6,    // F - SECRETARIA
	CARGO: 8,         // H - CARGO
	QUANTIDADE: 9,    // I - QTD SOLICITADA
	NOME_SERVIDOR: 12, // L - Nome
	PRONTUARIO: 13,   // M - Prontuário
	DESLIGAMENTO: 14, // N - Desligamento/Retorno
	DEPARTAMENTO: 7,  // G - DEPARTAMENTO
	DETALHAMENTO: 15, // O - DETALHAMENTO
	AUDITOR: 18       // R - AUDITOR
  };
  
  // ============================================================================
  // GATILHOS DE EXECUÇÃO
  // ============================================================================
  function onEdit(e) {
	if (!e) return;
	handleSpreadsheetEdit(e);
  }
  
  function timeDrivenFunction() {
	notificarAuditor();
  }
  
  // ============================================================================
  // FUNÇÃO PRINCIPAL DE EDIÇÃO
  // ============================================================================
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
  
  // ============================================================================
  // FUNÇÃO PARA VERIFICAR PRONTUÁRIOS DUPLICADOS
  // ============================================================================
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
		"⚠️ Este prontuário (ou parte dele) já foi usado acima. Verifique possível duplicidade."
	  );
	} else {
	  range.setFontColor("black");
	  range.setComment("");
	}
  }
  
  // ============================================================================
  // FUNÇÃO DE NOTIFICAÇÃO POR EMAIL (GATILHO DE TEMPO)
  // ============================================================================
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
		  `Linha ${linha} – status não é mais 'ANÁLISE RETORNO'. Pulando.`
		);
		continue;
	  }
  
	  if (!nomeAuditor) {
		console.warn(
		  `Linha ${linha} – auditor em branco. Mantendo para nova tentativa.`
		);
		novasLinhas.push(linha);
		continue;
	  }
  
	  nomeAuditor = nomeAuditor.trim();
	  var emailAuditor = MEMORANDO_CONFIG.MAPA_EMAILS[nomeAuditor];
  
	  if (emailAuditor) {
		var assunto = `Processo ${processo} retornou para análise`;
		var mensagem = `Olá ${nomeAuditor},\n\nO processo ${processo} da secretaria ${secretaria} foi atualizado com o status "ANÁLISE RETORNO".\n\nPor favor, verifique se há pendências ou se pode dar continuidade à análise.`;
  
		try {
		  MailApp.sendEmail(emailAuditor, assunto, mensagem);
		  console.log(
			`E-mail enviado para ${nomeAuditor} sobre processo ${processo}`
		  );
		} catch (erro) {
		  console.error(`Erro ao enviar e-mail (linha ${linha}): ${erro}`);
		  novasLinhas.push(linha);
		}
	  } else {
		console.warn(
		  `Linha ${linha} – e-mail não encontrado para auditor: ${nomeAuditor}`
		);
		novasLinhas.push(linha); // Tenta novamente em outro ciclo se o nome for corrigido
	  }
	}
  
	scriptProps.setProperty("linhasNotificar", JSON.stringify(novasLinhas));
  }
  
  // ============================================================================
  // FUNÇÕES AUXILIARES PARA GERENCIAMENTO DE ARQUIVOS
  // ============================================================================
  
  // Função para mover arquivo para pasta específica
  function moverArquivoParaPasta(docId, nomeArquivo) {
	try {
	  var arquivo = DriveApp.getFileById(docId);
	  var pastaDestino = DriveApp.getFolderById(MEMORANDO_CONFIG.PASTA_DOCUMENTOS_ID);
  
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
  
  // Função para adicionar link na aba 'Controle de Memos'
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
		// Coluna B - Memo
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
		// Coluna F - Relatórios
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
  // SISTEMA DE NUMERAÇÃO PARA MEMORANDOS
  // ============================================================================
  
  // Função para pegar próximo número de memorando
  function pegarProximoNumeroMemo() {
	try {
	  // Log para debug
	  Logger.log("Tentando acessar planilha ID: " + MEMORANDO_CONFIG.PLANILHA_MEMOS_ID);
	  Logger.log("Tentando acessar aba: " + MEMORANDO_CONFIG.ABA_NUMERACAO);
	  
	  const planilha = SpreadsheetApp.openById(MEMORANDO_CONFIG.PLANILHA_MEMOS_ID);
	  const sheet = planilha.getSheetByName(MEMORANDO_CONFIG.ABA_NUMERACAO);
	  
	  if (!sheet) {
		throw new Error(`Aba '${MEMORANDO_CONFIG.ABA_NUMERACAO}' não encontrada na planilha.`);
	  }
	  
	  const range = sheet.getDataRange();
	  const values = range.getValues();
	  const backgrounds = range.getBackgrounds();
  
	  Logger.log("Procurando por números disponíveis (sem cor de fundo)...");
  
	  // Percorre linha por linha, coluna por coluna
	  for (let i = 0; i < values.length; i++) {
		for (let j = 0; j < values[i].length; j++) {
		  const numero = values[i][j];
		  const cor = backgrounds[i][j];
		  
		  // Verifica se a célula tem número válido
		  if (numero && numero.toString().trim() !== "") {
			
			// Debug do que está sendo encontrado
			Logger.log(`Célula [${i+1},${j+1}]: Valor="${numero}", Cor="${cor}"`);
			
			// Procura por células SEM cor de fundo (disponíveis)
			// Células vazias/brancas podem ter cor "" ou "#ffffff" ou null
			const celulaSemCor = (
			  cor === "" || 
			  cor === "#ffffff" || 
			  cor === "#FFFFFF" || 
			  cor === null || 
			  cor === undefined ||
			  cor.toLowerCase() === "#ffffff"
			);
			
			if (celulaSemCor) {
			  const cell = sheet.getRange(i + 1, j + 1);
			  
			  // Pinta de amarelo para marcar como usado
			  cell.setBackground("#FFFF00");
			  
			  Logger.log(`Número encontrado e marcado como usado: ${numero}`);
			  return numero.toString().trim();
			}
		  }
		}
	  }
	  
	  throw new Error("Nenhum número disponível na planilha de numeração. Todos os números já foram utilizados.");
	  
	} catch (error) {
	  Logger.log("Erro detalhado ao buscar número: " + error.toString());
	  Logger.log("Stack: " + error.stack);
	  
	  // Se der erro, usar backup mas alertar o usuário
	  const numeroBackup = "BACKUP-" + Math.floor(Math.random() * 9000 + 1000);
	  
	  SpreadsheetApp.getUi().alert(
		"Aviso - Numeração",
		`Não foi possível acessar a planilha de numeração.\nErro: ${error.message}\n\nUsando número backup: ${numeroBackup}`,
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	  
	  return numeroBackup;
	}
  }
  
  // ============================================================================
  // SISTEMA DE TEMPLATES PARA MEMORANDOS
  // ============================================================================
  
  // Função para escolher template baseado no tipo e status
  function escolherTemplate(dados) {
	if (dados.tipo.includes("PERMUTA")) {
	  return MEMORANDO_CONFIG.TEMPLATES.PERMUTA;
	}
	if (dados.tipo.includes("AMPLIAÇÃO")) {
	  return MEMORANDO_CONFIG.TEMPLATES.AMPLIACAO;
	}
	if (dados.tipo.includes("PROCESSO SELETIVO")) {
	  return MEMORANDO_CONFIG.TEMPLATES.PROCESSO_SELETIVO;
	}
	if (dados.status.includes("DEFERIDO BT")) {
	  return MEMORANDO_CONFIG.TEMPLATES.DEFERIDO_BT;
	}
	return MEMORANDO_CONFIG.TEMPLATES.SUBSTITUICAO; // Template padrão
  }
  
  // ============================================================================
  // FUNÇÕES DE FORMATAÇÃO E VALIDAÇÃO
  // ============================================================================
  
  // Função para formatar quantidade em extenso
  function formatarQuantidade(qtd) {
	const num = parseInt(qtd) || 1;
	return `${num.toString().padStart(2, "0")} (${numeroParaExtenso(num)})`;
  }
  
  function numeroParaExtenso(num) {
	return (num >= 0 && num < MEMORANDO_CONFIG.NUMEROS_EXTENSO.length) 
	  ? MEMORANDO_CONFIG.NUMEROS_EXTENSO[num] 
	  : num.toString();
  }
  
  function formatarData(data) {
	if (data instanceof Date) {
	  return Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
	}
	return data ? data.toString() : "";
  }
  
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
	  justificativa: lerCelula(COLUNAS_MEMO.DETALHAMENTO)
	};
  }
  
  // ============================================================================
  // PREENCHIMENTO DE TEMPLATE
  // ============================================================================
  function preencherTemplate(doc, dados, numeroMemo) {
	const body = doc.getBody();
	const hoje = new Date();
	const dia = hoje.getDate().toString();
	const mes = MEMORANDO_CONFIG.MESES[hoje.getMonth()];
  
	// Substituições básicas
	body.replaceText("\\[NUMERO\\]", numeroMemo);
	body.replaceText("\\[DIA\\]", dia);
	body.replaceText("\\[MES\\]", mes);
	body.replaceText("\\[ANO\\]", MEMORANDO_CONFIG.CURRENT_YEAR.toString());
	body.replaceText("\\[CARGO\\]", dados.cargo);
	body.replaceText("\\[QUANTIDADE\\]", formatarQuantidade(dados.quantidade));
  
	// Preenchimento de tabela (se existir)
	const tabelas = body.getTables();
	if (tabelas.length > 0) {
	  const tabela = tabelas[0];
	  if (tabela.getNumRows() > 1) {
		tabela.getCell(1, 0).setText(dados.secretaria || "");
		tabela.getCell(1, 1).setText(dados.nomeServidor || "");
		tabela.getCell(1, 2).setText(dados.prontuario || "");
		tabela.getCell(1, 3).setText(formatarData(dados.desligamento));
		if (tabela.getRow(1).getNumCells() > 4) {
		  tabela.getCell(1, 4).setText(dados.departamento || "");
		}
	  }
	}
  }
  
  // ============================================================================
  // FUNÇÃO DE INTERFACE
  // ============================================================================
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
  
  // ============================================================================
  // FUNÇÃO PRINCIPAL PARA GERAR MEMORANDO ADP
  // ============================================================================
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
  
	  // Obter número do memorando
	  const numeroMemo = pegarProximoNumeroMemo();
	  
	  // Escolher template apropriado
	  const templateId = escolherTemplate(dados);
	  
	  // Criar documento baseado no template
	  const nomeDoc = `MEMORANDO nº ${numeroMemo}/${MEMORANDO_CONFIG.CURRENT_YEAR} - ADP`;
	  const docFile = DriveApp.getFileById(templateId).makeCopy(nomeDoc);
	  const doc = DocumentApp.openById(docFile.getId());
  
	  // Preencher template com dados
	  preencherTemplate(doc, dados, numeroMemo);
	  
	  // Salvar documento
	  doc.saveAndClose();
  
	  // Mover arquivo para pasta específica
	  if (MEMORANDO_CONFIG.PASTA_DOCUMENTOS_ID) {
		moverArquivoParaPasta(docFile.getId(), nomeDoc);
	  }
  
	  // Adicionar link na aba de controle
	  const url = doc.getUrl();
	  const numeroDocFormatado = `${numeroMemo}/${MEMORANDO_CONFIG.CURRENT_YEAR}`;
	  adicionarLinkControleMemos("memorando", numeroDocFormatado, dados.secretaria, dados.cargo, dados.sisgep, url);
  
	  // Abrir documento
	  abrirDocumento(url);
  
	} catch (error) {
	  tratarErro(error);
	}
  }
  
  // ============================================================================
  // FUNÇÕES DE MENU E INTERFACE
  // ============================================================================
  
  // Função chamada pelo menu personalizado
  function menuGerarMemorando() {
	gerarMemorandoADP();
  }
  
  // Função para criar menu personalizado na planilha
  function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('📋 ADP Ferramentas')
	  .addItem('📝 Gerar Memorando', 'menuGerarMemorando')
	  .addItem('📊 Gerar Relatório', 'menuGerarRelatorio')
	  .addSeparator()
	  .addItem('📧 Teste Notificação', 'testeNotificacao')
	  .addToUi();
  }
  
  // Função para gerar relatório (placeholder)
  function menuGerarRelatorio() {
	SpreadsheetApp.getUi().alert(
	  'Função em Desenvolvimento',
	  'A função de geração de relatórios será implementada em breve.',
	  SpreadsheetApp.getUi().ButtonSet.OK
	);
  }
  
  // Função para teste de notificação
  function testeNotificacao() {
	try {
	  notificarAuditor();
	  SpreadsheetApp.getUi().alert(
		'Teste Concluído',
		'Verificação de notificações executada. Consulte os logs para detalhes.',
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	} catch (error) {
	  SpreadsheetApp.getUi().alert(
		'Erro no Teste',
		`Erro durante o teste: ${error.message}`,
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	}
  }
  
  // ============================================================================
  // FUNÇÕES AUXILIARES ADICIONAIS
  // ============================================================================
  
  // Função para limpar propriedades do script (utilitário)
  function limparPropriedades() {
	const props = PropertiesService.getScriptProperties();
	props.deleteProperty("linhasNotificar");
	Logger.log("Propriedades do script limpas.");
  }
  
  // Função para debug - mostrar dados da linha selecionada
  function debugLinhaSelecionada() {
	try {
	  const sheet = SpreadsheetApp.getActiveSheet();
	  const linha = sheet.getActiveRange().getRow();
	  
	  if (linha <= 1) {
		SpreadsheetApp.getUi().alert('Erro', 'Selecione uma linha com dados válidos.', SpreadsheetApp.getUi().ButtonSet.OK);
		return;
	  }
	  
	  const dados = extrairDadosDaPlanilha(sheet, linha);
	  const dadosStr = JSON.stringify(dados, null, 2);
	  
	  Logger.log("Dados da linha " + linha + ":");
	  Logger.log(dadosStr);
	  
	  SpreadsheetApp.getUi().alert(
		'Debug - Dados Extraídos',
		'Dados registrados no log. Verifique Extensões > Apps Script > Execuções para ver os detalhes.',
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	} catch (error) {
	  Logger.log("Erro no debug: " + error.toString());
	  SpreadsheetApp.getUi().alert('Erro', 'Erro durante debug: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
	}
  }
  
  // ============================================================================
  // FUNÇÕES DE MANUTENÇÃO E LIMPEZA
  // ============================================================================
  
  // Função para resetar formatações de prontuários duplicados
  function resetarFormatacaoProntuarios() {
	try {
	  const sheet = SpreadsheetApp.getActiveSheet();
	  const ultimaLinha = sheet.getLastRow();
	  const colunaProntuario = 13; // Coluna M
	  
	  for (let linha = 2; linha <= ultimaLinha; linha++) {
		const cell = sheet.getRange(linha, colunaProntuario);
		if (cell.getFontColor() === "#ff0000") { // Se estiver vermelho
		  cell.setFontColor("black");
		  cell.setComment("");
		}
	  }
	  
	  SpreadsheetApp.getUi().alert(
		'Limpeza Concluída',
		'Formatações de prontuários duplicados foram resetadas.',
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	} catch (error) {
	  SpreadsheetApp.getUi().alert('Erro', 'Erro na limpeza: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
	}
  }
  
  // Função para verificar status da planilha de numeração
  function verificarPlanilhaNumeracao() {
	try {
	  const sheet = SpreadsheetApp.openById(MEMORANDO_CONFIG.PLANILHA_MEMOS_ID)
								  .getSheetByName(MEMORANDO_CONFIG.ABA_NUMERACAO);
	  const range = sheet.getDataRange();
	  const values = range.getValues();
	  const backgrounds = range.getBackgrounds();
	  
	  let disponiveis = 0;
	  let usados = 0;
	  
	  for (let i = 0; i < values.length; i++) {
		for (let j = 0; j < values[i].length; j++) {
		  const numero = values[i][j];
		  const cor = backgrounds[i][j].toLowerCase();
		  
		  if (numero && numero.toString().trim() !== "") {
			if (cor === "#ffff00") {
			  usados++;
			} else {
			  disponiveis++;
			}
		  }
		}
	  }
	  
	  SpreadsheetApp.getUi().alert(
		'Status da Numeração',
		`Números disponíveis: ${disponiveis}\nNúmeros usados: ${usados}`,
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	  
	} catch (error) {
	  SpreadsheetApp.getUi().alert(
		'Erro',
		'Não foi possível acessar a planilha de numeração: ' + error.message,
		SpreadsheetApp.getUi().ButtonSet.OK
	  );
	}
  }