/*
==================================================
🔌 SCRIPT: Monitoramento, Alertas e Geração de Relatórios
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
			sheetMemos.getRange(proximaLinha, 2).setFormula('=HYPERLINK("' + urlLimpa + '","' + numeroDocLimpo + '")');
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
			sheetMemos.getRange(proximaLinha, 6).setFormula('=HYPERLINK("' + urlLimpa + '","' + numeroDocLimpo + '")');
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

// ==================================================
// 📄 GERADOR DE Relatório Técnico
// ==================================================
function gerarRelatorioTecnico() {
	try {
		var sheet = SpreadsheetApp.getActiveSheet();
		var activeRange = sheet.getActiveRange();
		var linha = activeRange.getRow();

		if (linha <= 1) {
			SpreadsheetApp.getUi().alert(
				"Erro",
				"Por favor, selecione uma linha com dados (não o cabeçalho).",
				SpreadsheetApp.getUi().ButtonSet.OK,
			);
			return;
		}

		var dados = {
			sisgep: sheet.getRange(linha, 2).getValue() || "N/A", // B - PROCESSO
			secretaria: sheet.getRange(linha, 6).getValue() || "N/A", // F - SECRETARIA
			tipo: sheet.getRange(linha, 3).getValue() || "N/A", // C - TIPO MOVIMENTAÇÃO
			dataAbertura: sheet.getRange(linha, 4).getValue() || new Date(), // D - DATA ABERTURA
			cargo: sheet.getRange(linha, 8).getValue() || "N/A", // H - CARGO
			quantidade: sheet.getRange(linha, 9).getValue() || 1, // I - QTD SOLICITADA
			nomeServidor: sheet.getRange(linha, 12).getValue() || "", // L - Nome
			prontuario: sheet.getRange(linha, 13).getValue() || "", // M - Prontuário
			desligamento: sheet.getRange(linha, 14).getValue() || "", // N - Desligamento/ Retorno
			justificativa: sheet.getRange(linha, 15).getValue() || "", // O - DETALHAMENTO
		};

		var numeroRelatorio = Math.floor(Math.random() * 9000) + 1000;
		var meses = [
			"janeiro",
			"fevereiro",
			"março",
			"abril",
			"maio",
			"junho",
			"julho",
			"agosto",
			"setembro",
			"outubro",
			"novembro",
			"dezembro",
		];
		var hoje = new Date();
		var dataAtual =
			hoje.getDate() +
			" de " +
			meses[hoje.getMonth()] +
			" de " +
			hoje.getFullYear();
		var nomeDoc = `ADP RELATÓRIO TÉCNICO Nº ${numeroRelatorio}_2025 - ${dados.cargo} - Sisgep ${dados.sisgep}`;

		// Criação do documento
		var doc = DocumentApp.create(nomeDoc);
		var docId = doc.getId();
		doc.saveAndClose();

		// Reabre o documento para manipulação
		var reopenedDoc = DocumentApp.openById(docId);
		var body = reopenedDoc.getBody();
		var header = reopenedDoc.getHeader();

		// ✅ CORREÇÃO: Configuração das margens usando o método correto
		body.setMarginTop(0); // 1 cm = 28.35 pontos
		body.setMarginBottom(28.35);
		body.setMarginLeft(28.35);
		body.setMarginRight(28.35);

		// Configuração do cabeçalho
		if (!header) {
			header = reopenedDoc.addHeader();
		}
		header.clear();

		// INSERÇÃO DA IMAGEM NO CABEÇALHO
		try {
			var imageFileId = "1rosg8f8K9E4VQCRkCYtI_eO2PcXJJDvR";
			var image = DriveApp.getFileById(imageFileId).getBlob();
			var headerImg = header.appendImage(image);

			// Configuração do tamanho da imagem (680x76 px = ~18cmx2cm)
			headerImg.setWidth(680);
			headerImg.setHeight(76);

			// Centralizar a imagem
			var headerParagraph = headerImg.getParent();
			headerParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
		} catch (e) {
			Logger.log("Erro ao inserir imagem no cabeçalho: " + e.toString());
		}

		// Limpa o corpo do documento
		body.clear();

		// ESTILO GLOBAL para o corpo
		var normalStyle = {};
		normalStyle[DocumentApp.Attribute.FONT_FAMILY] = "Calibri";
		normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
		normalStyle[DocumentApp.Attribute.BOLD] = false;
		normalStyle[DocumentApp.Attribute.UNDERLINE] = false;
		body.setAttributes(normalStyle);

		// TÍTULO DO RELATÓRIO (negrito e sublinhado)
		var titulo = body.appendParagraph(
			`ADP RELATÓRIO TÉCNICO Nº ${numeroRelatorio}/2025`,
		);
		titulo.setBold(true);
		titulo.setUnderline(true);
		titulo.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

		// DATA (formatação normal)
		var dataLocal = body.appendParagraph(`Santana de Parnaíba, ${dataAtual}`);
		dataLocal.setAttributes(normalStyle);
		dataLocal.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
		body.appendParagraph(""); // Espaço

		// FUNÇÃO PARA LINHAS DE INFORMAÇÃO
		function addInfoLine(title, value) {
			var p = body.appendParagraph("");
			p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
			p.appendText(title).setBold(true); // Apenas o título em negrito
			p.appendText(value).setBold(false); // Conteúdo em fonte normal
		}

		addInfoLine("Secretaria: ", dados.secretaria);
		addInfoLine("Tipo de Solicitação: ", dados.tipo);
		addInfoLine("Processo SISGEP: ", dados.sisgep);
		addInfoLine("Assunto: ", dados.cargo + " - Processo Seletivo");

		// Espaços
		body.appendParagraph("");
		body.appendParagraph("");

		// SEÇÃO 1
		body
			.appendParagraph("1. Justificativa da Secretaria")
			.setBold(true)
			.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

		var textoJustificativa =
			dados.justificativa ||
			`A ${dados.secretaria} solicita a ${dados.tipo.toLowerCase()} de ${dados.quantidade} servidor${dados.quantidade > 1 ? "es" : ""} para o cargo de ${dados.cargo}, tendo em vista a necessidade de fortalecimento da equipe para o adequado funcionamento dos serviços públicos municipais.`;

		body
			.appendParagraph(textoJustificativa)
			.setAttributes(normalStyle)
			.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

		body.appendParagraph("");

		// SEÇÃO 2
		body
			.appendParagraph("2. Comprovação da Demanda")
			.setBold(true)
			.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

		var textoComprovacao = `Após análise da solicitação apresentada pela ${dados.secretaria}, verifica-se que a demanda está fundamentada na necessidade de ${dados.tipo.toLowerCase()} de profissional para o cargo de ${dados.cargo}.`;

		if (dados.tipo.toString().toUpperCase().includes("SUBSTITUIÇÃO")) {
			textoComprovacao = `A solicitação refere-se à substituição de servidor(a) do cargo de ${dados.cargo}`;
			if (dados.nomeServidor) textoComprovacao += `, ${dados.nomeServidor}`;
			if (dados.prontuario)
				textoComprovacao += ` (Prontuário ${dados.prontuario})`;
			if (dados.desligamento) {
				var dataDesligamento =
					dados.desligamento instanceof Date
						? Utilities.formatDate(
								dados.desligamento,
								Session.getScriptTimeZone(),
								"dd/MM/yyyy",
							)
						: dados.desligamento.toString();
				textoComprovacao += `, desligado(a) em ${dataDesligamento}`;
			}
			textoComprovacao += `. A substituição se mostra necessária para manter a continuidade dos serviços prestados pela pasta.`;
		}

		textoComprovacao += `\n\nDiante do exposto, manifesta-se parecer técnico favorável ao deferimento da solicitação, considerando a necessidade demonstrada para o adequado funcionamento dos serviços públicos municipais.`;

		body
			.appendParagraph(textoComprovacao)
			.setAttributes(normalStyle)
			.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

		reopenedDoc.saveAndClose();

		// Move o arquivo para a pasta específica
		if (PASTA_DOCUMENTOS_ID) {
			moverArquivoParaPasta(docId, nomeDoc);
		}

		// Adiciona o link na aba 'Controle de Memos'
		var url = reopenedDoc.getUrl();
		var numeroDocFormatado = `RT ${numeroRelatorio}/2025`;
		adicionarLinkControleMemos("relatorio", numeroDocFormatado, dados.secretaria, dados.cargo, dados.sisgep, url);

		var htmlOutput = HtmlService.createHtmlOutput(`
        <script>
          window.open('${url}', '_blank');
          google.script.host.close();
        </script>
      `);

		SpreadsheetApp.getUi().showModalDialog(
			htmlOutput,
			"Relatório Técnico Gerado com Sucesso!",
		);
	} catch (error) {
		Logger.log("Erro ao gerar relatório: " + error.toString());
		SpreadsheetApp.getUi().alert(
			"Erro",
			"Erro ao gerar relatório: " + error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK,
		);
	}
}

// ==================================================
// 📄 GERADOR DE MEMORANDO ADP
// ==================================================
function gerarMemorandoADP() {
	try {
		var sheet = SpreadsheetApp.getActiveSheet();
		var linha = sheet.getActiveRange().getRow();

		if (linha <= 1) {
			SpreadsheetApp.getUi().alert(
				"Erro",
				"Por favor, selecione uma linha com dados.",
				SpreadsheetApp.getUi().ButtonSet.OK,
			);
			return;
		}

		var dados = {
			sisgep: sheet.getRange(linha, 2).getValue() || "N/A", // B - PROCESSO
			secretaria: sheet.getRange(linha, 6).getValue() || "N/A", // F - SECRETARIA
			tipo: sheet.getRange(linha, 3).getValue() || "N/A", // C - TIPO MOVIMENTAÇÃO
			cargo: sheet.getRange(linha, 8).getValue() || "N/A", // H - CARGO
			quantidade: sheet.getRange(linha, 9).getValue() || 1, // I - QTD SOLICITADA
			nomeServidor: sheet.getRange(linha, 12).getValue() || "", // L - Nome
			prontuario: sheet.getRange(linha, 13).getValue() || "", // M - Prontuário
		};

		var numeroMemo = Math.floor(Math.random() * 9000) + 1000;
		var meses = [
			"janeiro",
			"fevereiro",
			"março",
			"abril",
			"maio",
			"junho",
			"julho",
			"agosto",
			"setembro",
			"outubro",
			"novembro",
			"dezembro",
		];
		var hoje = new Date();
		var dataAtual =
			hoje.getDate() +
			" de " +
			meses[hoje.getMonth()] +
			" de " +
			hoje.getFullYear();
		var nomeDoc = `MEMORANDO nº ${numeroMemo}_2025 - ADP - Resposta ao Processo Sisgep ${dados.sisgep}`;

		// Cria documento
		var doc = DocumentApp.create(nomeDoc);
		var docId = doc.getId();
		doc.saveAndClose();

		var reopenedDoc = DocumentApp.openById(docId);
		var body = reopenedDoc.getBody();
		var header = reopenedDoc.getHeader();

		// Configuração das margens - IGUAL AO RELATÓRIO TÉCNICO
		body.setMarginTop(0);
		body.setMarginBottom(28.35);
		body.setMarginLeft(28.35);
		body.setMarginRight(28.35);

		// Configuração do cabeçalho
		if (!header) {
			header = reopenedDoc.addHeader();
		}
		header.clear();

		// INSERÇÃO DA IMAGEM NO CABEÇALHO
		try {
			var imageFileId = "1rosg8f8K9E4VQCRkCYtI_eO2PcXJJDvR";
			var image = DriveApp.getFileById(imageFileId).getBlob();
			var headerImg = header.appendImage(image);

			// Configuração do tamanho da imagem
			headerImg.setWidth(680);
			headerImg.setHeight(76);

			// Centralizar a imagem
			var headerParagraph = headerImg.getParent();
			headerParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
		} catch (e) {
			Logger.log("Erro ao inserir imagem no cabeçalho: " + e.toString());
		}

		// Limpa o corpo do documento
		body.clear();

		// ESTILO GLOBAL para o corpo - Arial 12
		var normalStyle = {};
		normalStyle[DocumentApp.Attribute.FONT_FAMILY] = "Arial";
		normalStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
		normalStyle[DocumentApp.Attribute.BOLD] = false;
		normalStyle[DocumentApp.Attribute.UNDERLINE] = false;
		body.setAttributes(normalStyle);

		// TÍTULO DO MEMORANDO
		var titulo = body.appendParagraph(`MEMORANDO Nº ${numeroMemo}/2025 - ADP`);
		titulo.setBold(true);
		titulo.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

		// DATA
		var dataLocal = body.appendParagraph(`Santana de Parnaíba, ${dataAtual}`);
		dataLocal.setAttributes(normalStyle);
		dataLocal.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

		body.appendParagraph(""); // Espaço

		// FUNÇÃO PARA LINHAS DE INFORMAÇÃO
		function addInfoLine(title, value) {
			var p = body.appendParagraph("");
			p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
			p.appendText(title).setBold(true); // Apenas o título em negrito
			p.appendText(value).setBold(false); // Conteúdo em fonte normal
		}

		addInfoLine("De: ", "Análise de Demanda de Pessoal");
		addInfoLine("Para: ", "Secretaria Municipal de Administração");
		addInfoLine("Sr.: ", "José Roberto Martins Santos");
		body.appendParagraph("");
		addInfoLine("Ref.: ", `${dados.tipo} ${dados.cargo.toLowerCase()}`);

		body.appendParagraph("");

		// TÍTULO CENTRALIZADO
		var tituloAnalise = body.appendParagraph("Análise de Demanda de Pessoal");
		tituloAnalise.setBold(true);
		tituloAnalise.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

		body.appendParagraph("");

		// SAUDAÇÃO
		var saudacao = body.appendParagraph("Senhor Secretário,");
		saudacao.setAttributes(normalStyle);
		saudacao.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
		body.appendParagraph("");

		// CONTEÚDO PRINCIPAL
		var textoMemo = "";
		if (dados.tipo.toString().toUpperCase().includes("SUBSTITUIÇÃO")) {
			textoMemo = `Encaminhamos o presente expediente referente à solicitação de substituição de servidor(a) do cargo de ${dados.cargo}`;
			if (dados.nomeServidor) textoMemo += `, ${dados.nomeServidor}`;
			if (dados.prontuario) textoMemo += ` (Prontuário ${dados.prontuario})`;
			textoMemo += ` por meio de processo seletivo.`;
		} else {
			textoMemo = `Encaminhamos o presente expediente referente à solicitação de ${dados.tipo.toLowerCase()} de ${dados.quantidade} ${dados.cargo}${dados.quantidade > 1 ? "s" : ""} por meio de processo seletivo.`;
		}

		var paragrafoTexto = body.appendParagraph(textoMemo);
		paragrafoTexto.setAttributes(normalStyle);
		paragrafoTexto.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
		body.appendParagraph("");

		// CONCLUSÃO
		var conclusao = body.appendParagraph(
			"Após a devida análise, manifestamos parecer favorável ao DEFERIMENTO da solicitação. Assim, encaminhamos para demais providências.",
		);
		conclusao.setAttributes(normalStyle);
		conclusao.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

		body.appendParagraph("");
		body.appendParagraph("");

		// DESPEDIDA
		var despedida = body.appendParagraph("Atenciosamente,");
		despedida.setAttributes(normalStyle);
		despedida.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

		reopenedDoc.saveAndClose();

		// Move o arquivo para a pasta específica
		if (PASTA_DOCUMENTOS_ID) {
			moverArquivoParaPasta(docId, nomeDoc);
		}

		// Adiciona o link na aba 'Controle de Memos'
		var url = reopenedDoc.getUrl();
		var numeroDocFormatado = `${numeroMemo}/2025`;
		adicionarLinkControleMemos("memorando", numeroDocFormatado, dados.secretaria, dados.cargo, dados.sisgep, url);

		var htmlOutput = HtmlService.createHtmlOutput(`
        <script>
          window.open('${url}', '_blank');
          google.script.host.close();
        </script>
      `);
		SpreadsheetApp.getUi().showModalDialog(
			htmlOutput,
			"Memorando ADP Gerado com Sucesso!",
		);
	} catch (error) {
		Logger.log("Erro ao gerar memorando: " + error.toString());
		SpreadsheetApp.getUi().alert(
			"Erro",
			"Erro ao gerar memorando: " + error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK,
		);
	}
}

// ==================================================
// 🎛️ MENU PERSONALIZADO
// ==================================================

function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu("📋 Relatórios ADP")
		.addItem("📝 Gerar Relatório Técnico", "gerarRelatorioTecnico")
		.addItem("📄 Gerar Memorando", "gerarMemorandoADP")
		.addSeparator()
		.addItem("ℹ️ Sobre", "mostrarSobre")
		.addToUi();
}

function mostrarSobre() {
	SpreadsheetApp.getUi().alert(
		"📋 Aurora - Gestão Inteligente - Versão 2.4",
		"Desenvolvido por: Luana Halcsik Leite\n\n" +
			"📸 Funcionalidades:\n" +
			"• Monitoramento automático de processos\n" +
			"• Controle de usuários e permissões\n" +
			"• Alertas por e-mail\n" +
			"• Geração automática de relatórios\n" +
			"• Geração automática de memorandos\n" +
			"• Salvamento automático em pasta específica\n" +
			"• Links automáticos na aba 'Controle de Memos'\n\n" +
			"📧 Suporte: luana.41331@santanadeparnaiba.sp.gov.br\n" +
			"📞 Ramal: 8819\n\n" +
			"🆕 Última atualização: 16/09/2025",
		SpreadsheetApp.getUi().ButtonSet.OK,
	);
}

// ==================================================
// 📅 ATUALIZAÇÃO AUTOMÁTICA DE DATA NA COLUNA P
// (Unificada no handleSpreadsheetEdit e único onEdit no topo)
// ==================================================

// ==================================================
// 🔧 RECRIAR FILTRO DA ABA 'CONTROLE 2025'
// ==================================================
function corrigirFiltro() {
    const sh = SpreadsheetApp.getActive().getSheetByName('CONTROLE 2025');
    if (!sh) return;
    const filter = sh.getFilter();
    if (filter) filter.remove();
    // Cria filtro cobrindo todas as linhas existentes da aba
    sh.getRange(1, 1, sh.getMaxRows(), sh.getLastColumn()).createFilter();
}