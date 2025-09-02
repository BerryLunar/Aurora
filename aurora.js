/*
==================================================
📌 SCRIPT: Monitoramento, Alertas e Geração de Relatórios
==================================================

🧠 OBJETIVO:
Este script é executado automaticamente quando a planilha é editada.
Ele realiza as seguintes funções principais:

1️⃣ Quando o status (coluna A) muda para "ANÁLISE RETORNO":
   - Envia um e-mail automático para o auditor responsável (coluna T)
   - Insere um balão de comentário na célula com o nome do auditor

2️⃣ Quando o status muda para qualquer valor (exceto vazio ou "ANÁLISE"):
   - Atualiza a data de tramitação (coluna R) com a data atual

3️⃣ Verifica se a data de abertura (coluna D) ultrapassa 30 dias após a data de desligamento (coluna O):
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

🕒 Última atualização: 16/07/2025
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

	// COLUNAS IMPORTANTES
	var colunaStatus = 1; // A
	var colunaProcesso = 2; // B
	var colunaAbertura = 4; // D
	var colunaSecretaria = 6; // F
	var colunaDesligamento = 15; // O
	var colunaData = 17; // Q
	var colunaAuditor = 19; // S

	// BLOCO 1 – Atualiza data na coluna R se o status mudou (exceto "ANÁLISE")
	if (coluna === colunaStatus && linha > 1) {
		var cellData = sheet.getRange(linha, colunaData);
		if (valorSelecionado === "") {
			cellData.setValue("");
		} else if (valorSelecionado !== "ANÁLISE") {
			cellData.setValue(new Date());
		}
	}

	// BLOCO 2 – Verifica se abertura foi após 30 dias do desligamento
	var dataDesligamento = sheet.getRange(linha, colunaDesligamento).getValue();
	var dataAbertura = sheet.getRange(linha, colunaAbertura).getValue();
	var cellAbertura = sheet.getRange(linha, colunaAbertura);

	if (dataDesligamento instanceof Date && dataAbertura instanceof Date) {
		var prazoLimite = new Date(dataDesligamento);
		prazoLimite.setDate(prazoLimite.getDate() + 30);

		if (dataAbertura > prazoLimite) {
			cellAbertura.setFontColor("red");
			cellAbertura.setComment(
				"Abertura feita após 30 dias do desligamento. Verificar pendência ou justificativa.",
			);
		} else {
			cellAbertura.setFontColor("black");
			cellAbertura.setComment("");
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

	// BLOCO 4 – Limpeza e verificação de prontuários duplicados (coluna N = 14)
	var abasPermitidas = ["CONTROLE 2025"];
	var colunaProntuario = 14;

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

		var nomeAuditor = sheet.getRange(linha, 18).getValue();
		var processo = sheet.getRange(linha, 2).getValue();
		var secretaria = sheet.getRange(linha, 6).getValue();
		var statusAtual = sheet.getRange(linha, 1).getValue();

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
			sisgep: sheet.getRange(linha, 2).getValue() || "N/A",
			secretaria: sheet.getRange(linha, 6).getValue() || "N/A",
			tipo: sheet.getRange(linha, 3).getValue() || "N/A",
			dataAbertura: sheet.getRange(linha, 4).getValue() || new Date(),
			cargo: sheet.getRange(linha, 8).getValue() || "N/A",
			quantidade: sheet.getRange(linha, 9).getValue() || 1,
			nomeServidor: sheet.getRange(linha, 13).getValue() || "",
			prontuario: sheet.getRange(linha, 14).getValue() || "",
			desligamento: sheet.getRange(linha, 15).getValue() || "",
			justificativa: sheet.getRange(linha, 16).getValue() || "",
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

		// Move o arquivo para a pasta específica (se PASTA_DOCUMENTOS_ID estiver definido)
		if (
			PASTA_DOCUMENTOS_ID &&
			PASTA_DOCUMENTOS_ID !== "1OBHunABxlCl0WHsBKFse-6icL8Aat4Py"
		) {
			moverArquivoParaPasta(docId, nomeDoc);
		}

		// Adiciona o link do relatório na coluna U (21) - CORREÇÃO DO HYPERLINK
		var url = reopenedDoc.getUrl();
		try {
			var cellRelatorio = sheet.getRange(linha, 21); // Coluna U
			cellRelatorio.setFormula('=HYPERLINK("' + url + '","RT")');
		} catch (linkError) {
			Logger.log(
				"Erro ao inserir hyperlink do relatório: " + linkError.toString(),
			);
			// Fallback: inserir apenas o URL
			sheet.getRange(linha, 22).setValue(url);
		}
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
			sisgep: sheet.getRange(linha, 2).getValue() || "N/A",
			secretaria: sheet.getRange(linha, 6).getValue() || "N/A",
			tipo: sheet.getRange(linha, 3).getValue() || "N/A",
			cargo: sheet.getRange(linha, 8).getValue() || "N/A",
			quantidade: sheet.getRange(linha, 9).getValue() || 1,
			nomeServidor: sheet.getRange(linha, 13).getValue() || "",
			prontuario: sheet.getRange(linha, 14).getValue() || "",
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

		// Move o arquivo para a pasta específica (se PASTA_DOCUMENTOS_ID estiver definido)
		if (
			PASTA_DOCUMENTOS_ID &&
			PASTA_DOCUMENTOS_ID !== "1OBHunABxlCl0WHsBKFse-6icL8Aat4Py"
		) {
			moverArquivoParaPasta(docId, nomeDoc);
		}

		// Adiciona o link do memorando na coluna T (20) - CORREÇÃO DO HYPERLINK
		var url = reopenedDoc.getUrl();
		try {
			var cellMemo = sheet.getRange(linha, 20); // Coluna T
			cellMemo.setFormula('=HYPERLINK("' + url + '","Memo")');
		} catch (linkError) {
			Logger.log(
				"Erro ao inserir hyperlink do memorando: " + linkError.toString(),
			);
			// Fallback: inserir apenas o URL
			sheet.getRange(linha, 21).setValue(url);
		}
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
		"📋 Aurora - Gestão Inteligente - Versão 2.3",
		"Desenvolvido por: Luana Halcsik Leite\n\n" +
			"🔸 Funcionalidades:\n" +
			"• Monitoramento automático de processos\n" +
			"• Controle de usuários e permissões\n" +
			"• Alertas por e-mail\n" +
			"• Geração automática de relatórios\n" +
			"• Geração automática de memorandos\n" +
			"• Salvamento automático em pasta específica\n" +
			"• Links automáticos na planilha\n\n" +
			"📧 Suporte: luana.41331@santanadeparnaiba.sp.gov.br\n" +
			"📞 Ramal: 8819\n\n" +
			"🆕 Última atualização: 16/07/2025",
		SpreadsheetApp.getUi().ButtonSet.OK,
	);
}

// ==================================================
// 📅 ATUALIZAÇÃO AUTOMÁTICA DE DATA NA COLUNA Q
// ==================================================

/**
 * Função para atualizar automaticamente a data na coluna Q
 * quando o status for diferente de 'análise' e 'vazio'
 * Funciona APENAS na planilha '**CONTROLE 2025'
 */
function onEdit(e) {
  // Verificar se o evento existe
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  
  // RESTRIÇÃO: Executar apenas na planilha específica
  if (sheetName !== '**CONTROLE 2025') {
    return;
  }
  
  // Verificar se a edição foi feita na coluna do status
  // Assumindo que o status está em uma coluna específica (ajuste conforme necessário)
  const statusColumn = getStatusColumn(); // Você precisa definir qual coluna contém o status
  
  if (range.getColumn() !== statusColumn) {
    return;
  }
  
  const editedRow = range.getRow();
  const statusValue = range.getValue();
  
  // Verificar se o status é diferente de 'análise' e não está vazio
  if (shouldUpdateDate(statusValue)) {
    const dateColumn = 17; // Coluna Q (17ª coluna)
    const dateCell = sheet.getRange(editedRow, dateColumn);
    
    // Atualizar com a data e hora atual
    const now = new Date();
    dateCell.setValue(now);
    
    // Opcional: Formatar a célula de data
    dateCell.setNumberFormat('dd/mm/yyyy hh:mm');
  }
}

/**
 * Determina qual coluna contém o status
 * AJUSTE ESTA FUNÇÃO conforme sua planilha
 */
function getStatusColumn() {
  // Exemplo: se o status está na coluna P (16ª coluna)
  return 16;
  
  // Ou você pode buscar dinamicamente pelo cabeçalho:
  /*
  const sheet = SpreadsheetApp.getActiveSheet();
  const headerRow = 1; // Assumindo que os cabeçalhos estão na linha 1
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().toLowerCase().includes('status')) {
      return i + 1; // +1 porque getColumn() é 1-indexed
    }
  }
  return null;
  */
}

/**
 * Verifica se a data deve ser atualizada baseada no valor do status
 */
function shouldUpdateDate(statusValue) {
  if (!statusValue) return false;
  
  const status = statusValue.toString().toLowerCase().trim();
  
  // Não atualizar se for 'análise' ou vazio
  if (status === 'análise' || status === '') {
    return false;
  }
  
  return true;
}

/**
 * Função alternativa caso você queira usar um trigger específico
 * em vez do onEdit global
 */
function setupSpecificTrigger() {
  // Deletar triggers existentes para evitar duplicação
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditSpecific') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Criar novo trigger
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditSpecific')
    .onEdit()
    .create();
}

/**
 * Handler específico com mais controles
 */
function onEditSpecific(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Log para debug (remover em produção)
    console.log(`Editando planilha: ${sheetName}`);
    
    // RESTRIÇÃO PRINCIPAL: apenas planilha específica
    if (sheetName !== '**CONTROLE 2025') {
      console.log(`Ignorando edição em: ${sheetName}`);
      return;
    }
    
    const range = e.range;
    const editedRow = range.getRow();
    const editedColumn = range.getColumn();
    const newValue = range.getValue();
    
    // Definir qual coluna contém o status (ajuste conforme necessário)
    const statusColumn = getStatusColumn();
    
    if (editedColumn !== statusColumn) {
      return;
    }
    
    // Verificar condições para atualização
    if (shouldUpdateDate(newValue)) {
      const dateColumn = 17; // Coluna Q
      const dateCell = sheet.getRange(editedRow, dateColumn);
      
      const currentDate = new Date();
      dateCell.setValue(currentDate);
      dateCell.setNumberFormat('dd/mm/yyyy hh:mm:ss');
      
      console.log(`Data atualizada na linha ${editedRow}: ${currentDate}`);
    }
    
  } catch (error) {
    console.error('Erro na função onEditSpecific:', error);
  }
}
