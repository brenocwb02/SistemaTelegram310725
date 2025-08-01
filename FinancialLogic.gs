/**
 * @file FinancialLogic.gs
 * @description Este arquivo contém a lógica de negócio central do bot financeiro.
 * Inclui interpretação de mensagens, cálculos financeiros, categorização e atualização de saldos.
 */

// As constantes de estado do tutorial (TUTORIAL_STATE_WAITING_DESPESA, etc.) foram movidas para Management.gs
// para evitar redeclaração e garantir um ponto único de verdade.

// Variáveis globais para os dados da planilha que são acessados frequentemente
// Serão populadas e armazenadas em cache.
let cachedPalavrasChave = null;
let cachedCategorias = null;
let cachedContas = null;
let cachedConfig = null;

/**
 * Obtém dados de uma aba da planilha e os armazena em cache.
 * @param {string} sheetName O nome da aba.
 * @param {string} cacheKey A chave para o cache.
 * @param {number} [expirationInSeconds=300] Tempo de expiração do cache em segundos.
 * @returns {Array<Array<any>>} Os dados da aba (incluindo cabeçalhos).
 */
function getSheetDataWithCache(sheetName, cacheKey, expirationInSeconds = 300) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    // logToSheet(`Dados da aba '${sheetName}' recuperados do cache.`, "DEBUG");
    return JSON.parse(cachedData);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    logToSheet(`ERRO: Aba '${sheetName}' não encontrada.`, "ERROR");
    throw new Error(`Aba '${sheetName}' não encontrada.`);
  }

  const data = sheet.getDataRange().getValues();
  cache.put(cacheKey, JSON.stringify(data), expirationInSeconds);
  // logToSheet(`Dados da aba '${sheetName}' lidos da planilha e armazenados em cache.`, "DEBUG");
  return data;
}

/**
 * Interpreta uma mensagem do Telegram para extrair informações de transação.
 * @param {string} mensagem O texto da mensagem recebida.
 * @param {string} usuario O nome do usuário que enviou a mensagem.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Um objeto contendo os detalhes da transação ou uma mensagem de erro/status.
 */
function interpretarMensagemTelegram(mensagem, usuario, chatId) {
  logToSheet(`Interpretando mensagem: "${mensagem}" para usuário: ${usuario}`, "INFO");

  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

  const textoNormalizado = normalizarTexto(mensagem);
  logToSheet(`Texto normalizado: "${textoNormalizado}"`, "DEBUG");

  // --- 1. Detectar Tipo (Despesa, Receita, Transferência) ---
  const tipoInfo = detectarTipoTransacao(textoNormalizado, dadosPalavras);
  if (!tipoInfo) {
    return { errorMessage: "Não consegui identificar se é uma despesa, receita ou transferência. Tente ser mais claro." };
  }
  const tipoTransacao = tipoInfo.tipo;
  const keywordTipo = tipoInfo.keyword;
  logToSheet(`Tipo de transação detectado: ${tipoTransacao} (keyword: ${keywordTipo})`, "DEBUG");

  // 2. Extrair Valor
  const valor = extrairValor(textoNormalizado);
  logToSheet(`Valor extraído: ${valor}`, "DEBUG");
  if (isNaN(valor) || valor <= 0) {
    return { errorMessage: "Não consegui identificar o valor. Por favor, inclua um número válido (ex: 50, 12,50)." };
  }

  // --- Lógica especial para transferências ---
  if (tipoTransacao === "Transferência") {
      return interpretarTransferencia(textoNormalizado, valor, usuario, chatId, dadosContas, dadosPalavras);
  }

  // 3. Extrair Conta/Cartão e Método de Pagamento
  const { conta, infoConta, metodoPagamento, keywordConta, keywordMetodo } = extrairContaMetodoPagamento(textoNormalizado, dadosContas, dadosPalavras);
  logToSheet(`Conta/Cartão extraída: ${conta} (Tipo: ${infoConta ? infoConta.tipo : 'N/A'}), Método de Pagamento: ${metodoPagamento}`, "DEBUG");
  
  if (!infoConta) {
      logToSheet(`Nenhuma informacao de conta encontrada para "${conta}".`, "WARN");
  }

  // 4. Extrair Categoria e Subcategoria
  const { categoria, subcategoria, keywordCategoria } = extrairCategoriaSubcategoria(textoNormalizado, tipoTransacao, dadosPalavras);
  logToSheet(`Categoria: ${categoria}, Subcategoria: ${subcategoria} (keyword: ${keywordCategoria})`, "DEBUG");

  // 5. Extrair Descrição (o que sobrou)
  // CORREÇÃO: Não remover a palavra-chave da categoria, pois ela geralmente é a própria descrição.
  const keywordsToRemove = [keywordTipo, keywordConta, keywordMetodo];
  const descricaoFinal = extrairDescricao(textoNormalizado, String(valor), keywordsToRemove);
  logToSheet(`Descricao Final: "${descricaoFinal}"`, "DEBUG");

  // 6. Extrair Parcelas
  const parcelasTotais = extrairParcelas(textoNormalizado);
  logToSheet(`Parcelas: ${parcelasTotais}`, "DEBUG");

  // 7. Calcular Data de Vencimento para Cartões de Crédito (se aplicável)
  let dataVencimento = new Date(); // Padrão: data da transação
  let isCreditCardTransaction = false;

  if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
    isCreditCardTransaction = true;
    dataVencimento = calcularVencimentoCartao(infoConta, new Date(), dadosContas); 
    logToSheet(`Transação em cartão de crédito. Data de vencimento calculada: ${dataVencimento}`, "DEBUG");
  } else {
    logToSheet(`Transação não é em cartão de crédito. Data de vencimento será a data da transação: ${dataVencimento}`, "DEBUG");
  }

  const transactionId = Utilities.getUuid(); // ID único para a transação

  const transacaoData = {
    id: transactionId,
    data: new Date(), // Data da transação é a data atual
    descricao: descricaoFinal,
    categoria: categoria,
    subcategoria: subcategoria,
    tipo: tipoTransacao,
    valor: valor,
    metodoPagamento: metodoPagamento,
    conta: conta,
    parcelasTotais: parcelasTotais,
    parcelaAtual: 1, // Sempre começa na parcela 1 para novas transações
    dataVencimento: dataVencimento,
    usuario: usuario,
    status: "Ativo", // Status inicial
    dataRegistro: new Date(),
    isCreditCardTransaction: isCreditCardTransaction // Indica se é transação de cartão
  };

  // Se houver parcelas, prepara para solicitar confirmação
  if (parcelasTotais > 1) {
    return prepararConfirmacaoParcelada(transacaoData, chatId);
  } else {
    return prepararConfirmacaoSimples(transacaoData, chatId);
  }
}

/**
 * ATUALIZADO: Detecta o tipo de transação e a palavra-chave que o acionou.
 * @param {string} mensagemCompleta O texto da mensagem normalizada.
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba "PalavrasChave".
 * @returns {Object|null} Um objeto {tipo, keyword} ou null se não for detectado.
 */
function detectarTipoTransacao(mensagemCompleta, dadosPalavras) {
  logToSheet(`[detectarTipoTransacao] Mensagem Completa: "${mensagemCompleta}"`, "DEBUG");

  const palavrasReceitaFixas = ['recebi', 'salario', 'rendeu', 'pix recebido', 'transferencia recebida', 'deposito', 'entrada', 'renda', 'pagamento recebido', 'reembolso', 'cashback'];
  const palavrasDespesaFixas = ['gastei', 'paguei', 'comprei', 'saida', 'débito', 'debito'];

  for (let palavraRec of palavrasReceitaFixas) {
    if (mensagemCompleta.includes(palavraRec)) {
      logToSheet(`[detectarTipoTransacao] Receita detectada pela palavra fixa: "${palavraRec}"`, "DEBUG");
      return { tipo: "Receita", keyword: palavraRec };
    }
  }

  for (let palavraDes of palavrasDespesaFixas) {
    if (mensagemCompleta.includes(palavraDes)) {
      logToSheet(`[detectarTipoTransacao] Despesa detectada pela palavra fixa: "${palavraDes}"`, "DEBUG");
      return { tipo: "Despesa", keyword: palavraDes };
    }
  }

  for (let i = 1; i < dadosPalavras.length; i++) {
    const tipoPalavra = (dadosPalavras[i][0] || "").toString().trim().toLowerCase();
    const chave = normalizarTexto(dadosPalavras[i][1] || "");
    const valorInterpretado = (dadosPalavras[i][2] || "").toString().trim();

    if (tipoPalavra === "tipo_transacao" && chave) {
      const regex = new RegExp(`\\b${chave}\\b`);
      if (regex.test(mensagemCompleta)) {
        logToSheet(`[detectarTipoTransacao] Tipo detectado da planilha: "${valorInterpretado}" pela palavra: "${chave}"`, "DEBUG");
        return { tipo: valorInterpretado, keyword: chave };
      }
    }
  }

  logToSheet("[detectarTipoTransacao] Nenhum tipo especifico detectado. Retornando null.", "WARN");
  return null;
}

/**
 * Extrai o valor numérico da mensagem.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @returns {number} O valor numérico extraído, ou NaN.
 */
function extrairValor(textoNormalizado) {
  const regex = /(\d[\d\.,]*)/; 
  const match = textoNormalizado.match(regex);
  if (match) {
    return parseBrazilianFloat(match[1]); 
  }
  return NaN;
}

/**
 * ATUALIZADO: Extrai a conta, método de pagamento e as palavras-chave correspondentes.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba 'PalavrasChave'.
 * @returns {Object} Objeto com conta, infoConta, metodoPagamento, keywordConta e keywordMetodo.
 */
function extrairContaMetodoPagamento(textoNormalizado, dadosContas, dadosPalavras) {
  let contaEncontrada = "Não Identificada";
  let metodoPagamentoEncontrado = "Não Identificado";
  let melhorInfoConta = null;
  let maiorSimilaridadeConta = 0;
  let melhorPalavraChaveConta = "";
  let melhorPalavraChaveMetodo = "";

  // 1. Encontrar a melhor conta/cartão
  for (let i = 1; i < dadosContas.length; i++) {
    const nomeContaPlanilha = (dadosContas[i][0] || "").toString().trim();
    const nomeContaNormalizado = normalizarTexto(nomeContaPlanilha);
    const palavrasChaveConta = (dadosContas[i][3] || "").toString().trim().split(',').map(s => normalizarTexto(s.trim()));
    palavrasChaveConta.push(nomeContaNormalizado);

    for (const palavraChave of palavrasChaveConta) {
        if (!palavraChave) continue;
        if (textoNormalizado.includes(palavraChave)) {
            const similarity = calculateSimilarity(textoNormalizado, palavraChave);
            const currentSimilarity = (palavraChave === nomeContaNormalizado) ? similarity * 1.5 : similarity; 
            if (currentSimilarity > maiorSimilaridadeConta) {
                maiorSimilaridadeConta = currentSimilarity;
                contaEncontrada = nomeContaPlanilha;
                melhorInfoConta = obterInformacoesDaConta(nomeContaPlanilha, dadosContas);
                melhorPalavraChaveConta = palavraChave;
            }
        }
    }
  }

  // 2. Extrair Método de Pagamento
  let maiorSimilaridadeMetodo = 0;
  for (let i = 1; i < dadosPalavras.length; i++) {
    const tipo = (dadosPalavras[i][0] || "").toString().trim().toLowerCase();
    const palavraChave = (dadosPalavras[i][1] || "").toString().trim().toLowerCase();
    const valorInterpretado = (dadosPalavras[i][2] || "").toString().trim();

    if (tipo === "meio_pagamento" && palavraChave && textoNormalizado.includes(palavraChave)) {
        const similarity = calculateSimilarity(textoNormalizado, palavraChave);
        if (similarity > maiorSimilaridadeMetodo) {
          maiorSimilaridadeMetodo = similarity;
          metodoPagamentoEncontrado = valorInterpretado;
          melhorPalavraChaveMetodo = palavraChave;
        }
    }
  }

  // 3. Lógica de fallback para método de pagamento
  if (melhorInfoConta && normalizarTexto(melhorInfoConta.tipo) === "cartao de credito") {
    if (normalizarTexto(metodoPagamentoEncontrado) === "nao identificado" || normalizarTexto(metodoPagamentoEncontrado) === "debito") {
      metodoPagamentoEncontrado = "Crédito";
      logToSheet(`[ExtrairContaMetodo] Conta e cartao de credito, metodo de pagamento ajustado para "Credito".`, "DEBUG");
    }
  }
  
  return { 
      conta: contaEncontrada, 
      infoConta: melhorInfoConta, 
      metodoPagamento: metodoPagamentoEncontrado,
      keywordConta: melhorPalavraChaveConta,
      keywordMetodo: melhorPalavraChaveMetodo
  };
}


/**
 * ATUALIZADO: Extrai categoria, subcategoria e a palavra-chave correspondente.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @param {string} tipoTransacao O tipo de transação (Despesa, Receita).
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba 'PalavrasChave'.
 * @returns {Object} Objeto com categoria, subcategoria e keywordCategoria.
 */
function extrairCategoriaSubcategoria(textoNormalizado, tipoTransacao, dadosPalavras) {
  let categoriaEncontrada = "Não Identificada";
  let subcategoriaEncontrada = "Não Identificada";
  let melhorScoreSubcategoria = -1;
  let melhorPalavraChaveCategoria = "";

  for (let i = 1; i < dadosPalavras.length; i++) {
    const tipoPalavraChave = (dadosPalavras[i][0] || "").toString().trim().toLowerCase();
    const palavraChave = (dadosPalavras[i][1] || "").toString().trim().toLowerCase();
    const valorInterpretado = (dadosPalavras[i][2] || "").toString().trim();

    if (tipoPalavraChave === "subcategoria" && palavraChave && textoNormalizado.includes(palavraChave)) {
        const similarity = calculateSimilarity(textoNormalizado, palavraChave); 
        if (similarity > melhorScoreSubcategoria) { 
          if (valorInterpretado.includes(">")) {
            const partes = valorInterpretado.split(">");
            const categoria = partes[0].trim();
            const subcategoria = partes[1].trim();
            const tipoCategoria = (dadosPalavras[i][3] || "").toString().trim().toLowerCase();
            
            if (!tipoCategoria || normalizarTexto(tipoCategoria) === normalizarTexto(tipoTransacao)) {
              categoriaEncontrada = categoria;
              subcategoriaEncontrada = subcategoria;
              melhorScoreSubcategoria = similarity;
              melhorPalavraChaveCategoria = palavraChave;
            }
          }
        }
    }
  }
  return { 
      categoria: categoriaEncontrada, 
      subcategoria: subcategoriaEncontrada,
      keywordCategoria: melhorPalavraChaveCategoria
  };
}


/**
 * ATUALIZADO: Extrai a descrição final da transação, removendo os dados já identificados.
 * @param {string} textoParaLimpar O texto normalizado da mensagem do usuário.
 * @param {string} valor O valor extraído (como string).
 * @param {Array<string>} keywordsToRemove As palavras-chave a serem removidas.
 * @returns {string} A descrição limpa.
 */
function extrairDescricao(textoParaLimpar, valor, keywordsToRemove) {
  let descricao = textoParaLimpar;

  // Remove o valor
  descricao = descricao.replace(new RegExp(`\\b${valor.replace(/\./g, '\\.').replace(/,/g, '[\\.,]')}\\b`, 'gi'), '');

  // Remove as palavras-chave de metadados
  keywordsToRemove.forEach(keyword => {
    if (keyword) {
      descricao = descricao.replace(new RegExp(`\\b${keyword.replace(/ /g, '\\s+')}\\b`, "gi"), '');
    }
  });

  // Remove termos comuns de parcelamento
  descricao = descricao.replace(/\b(em\s+\d+\s*x)\b/gi, "");
  descricao = descricao.replace(/\b(\d+\s*x)\b/gi, "");
  descricao = descricao.replace(/\b((\d+)\s*(vezes|x))\b/gi, "");

  // Limpa múltiplos espaços e preposições que sobraram
  descricao = descricao.replace(/\s+/g, " ").trim();
  const preposicoes = ['de', 'da', 'do', 'dos', 'das', 'e', 'ou', 'a', 'o', 'no', 'na', 'nos', 'nas', 'com', 'em', 'para', 'por'];
  preposicoes.forEach(prep => {
    descricao = descricao.replace(new RegExp(`^${prep}\\s+|\\s+${prep}$`, 'gi'), "").trim();
  });
  descricao = descricao.replace(/\s+/g, " ").trim();
  
  if (descricao.length < 3) {
    return "Lançamento Geral";
  }

  return capitalize(descricao);
}

/**
 * Extrai o número total de parcelas da mensagem.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @returns {number} O número de parcelas (padrão 1 se não for encontrado).
 */
function extrairParcelas(textoNormalizado) {
  const regex = /(\d+)\s*(?:x|vezes)/;
  const match = textoNormalizado.match(regex);
  return match ? parseInt(match[1], 10) : 1;
}

/**
 * Prepara e envia uma mensagem de confirmação para transações simples (não parceladas).
 * Armazena os dados da transação em cache.
 * @param {Object} transacaoData Os dados da transação.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Status de confirmação pendente.
 */
function prepararConfirmacaoSimples(transacaoData, chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${transacaoData.id}`;
  cache.put(cacheKey, JSON.stringify(transacaoData), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS); // Expira em 5 minutos

  let mensagem = `✅ Confirme seu Lançamento:\n\n`;
  mensagem += `*Tipo:* ${escapeMarkdown(transacaoData.tipo)}\n`;
  mensagem += `*Descricao:* ${escapeMarkdown(transacaoData.descricao)}\n`;
  mensagem += `*Valor:* ${formatCurrency(transacaoData.valor)}\n`;
  mensagem += `*Conta:* ${escapeMarkdown(transacaoData.conta)}\n`;
  mensagem += `*Metodo:* ${escapeMarkdown(transacaoData.metodoPagamento)}\n`;
  mensagem += `*Categoria:* ${escapeMarkdown(transacaoData.categoria)}\n`;
  mensagem += `*Subcategoria:* ${escapeMarkdown(transacaoData.subcategoria)}\n`;

  const teclado = {
    inline_keyboard: [
      [{ text: "✅ Confirmar", callback_data: `confirm_${transacaoData.id}` }],
      [{ text: "❌ Cancelar", callback_data: `cancel_${transacaoData.id}` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  return { status: "PENDING_CONFIRMATION", transactionId: transacaoData.id };
}

/**
 * Prepara e envia uma mensagem de confirmação para transações parceladas.
 * Armazena os dados da transação em cache.
 * @param {Object} transacaoData Os dados da transação.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Status de confirmação pendente.
 */
function prepararConfirmacaoParcelada(transacaoData, chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${transacaoData.id}`;
  cache.put(cacheKey, JSON.stringify(transacaoData), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS); // Expira em 5 minutos

  let mensagem = `✅ Confirme seu Lançamento Parcelado:\n\n`;
  mensagem += `*Tipo:* ${escapeMarkdown(transacaoData.tipo)}\n`;
  mensagem += `*Descricao:* ${escapeMarkdown(transacaoData.descricao)}\n`;
  mensagem += `*Valor Total:* ${formatCurrency(transacaoData.valor)}\n`;
  mensagem += `*Parcelas:* ${transacaoData.parcelasTotais}x de ${formatCurrency(transacaoData.valor / transacaoData.parcelasTotais)}\n`;
  mensagem += `*Conta:* ${escapeMarkdown(transacaoData.conta)}\n`;
  mensagem += `*Metodo:* ${escapeMarkdown(transacaoData.metodoPagamento)}\n`;
  mensagem += `*Categoria:* ${escapeMarkdown(transacaoData.categoria)}\n`;
  mensagem += `*Subcategoria:* ${escapeMarkdown(transacaoData.subcategoria)}\n`;
  mensagem += `*Primeiro Vencimento:* ${Utilities.formatDate(transacaoData.dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy")}\n`;


  const teclado = {
    inline_keyboard: [
      [{ text: "✅ Confirmar Parcelamento", callback_data: `confirm_${transacaoData.id}` }],
      [{ text: "❌ Cancelar", callback_data: `cancel_${transacaoData.id}` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  return { status: "PENDING_CONFIRMATION", transactionId: transacaoData.id };
}

/**
 * ATUALIZADO: Registra a transação confirmada na planilha.
 * @param {Object} transacaoData Os dados da transação (pode ser um objeto único ou um array para transferências).
 * @param {string} usuario O nome do usuário que confirmou.
 * @param {string} chatId O ID do chat do Telegram.
 */
function registrarTransacaoConfirmada(transacaoData, usuario, chatId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);

    if (!transacoesSheet || !contasSheet) {
      enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Transacoes' ou 'Contas' não encontrada para registrar.");
      return;
    }
    
    const transacoesParaRegistrar = Array.isArray(transacaoData) ? transacaoData : [transacaoData];
    let transferCounter = 1;

    for (const transacao of transacoesParaRegistrar) {
        const infoConta = obterInformacoesDaConta(transacao.conta, contasSheet.getDataRange().getValues()); 
        const valorParcela = transacao.valor / transacao.parcelasTotais;
        
        const dataVencimentoBase = new Date(transacao.dataVencimento);
        const dataTransacaoBase = new Date(transacao.data);
        const dataRegistroBase = new Date(transacao.dataRegistro);

        for (let i = 0; i < transacao.parcelasTotais; i++) {
          let dataVencimentoParcela = new Date(dataVencimentoBase);
          dataVencimentoParcela.setMonth(dataVencimentoBase.getMonth() + i);

          if (dataVencimentoParcela.getDate() !== dataVencimentoBase.getDate()) {
              const lastDayOfMonth = new Date(dataVencimentoParcela.getFullYear(), dataVencimentoParcela.getMonth() + 1, 0).getDate();
              dataVencimentoParcela.setDate(Math.min(dataVencimentoBase.getDate(), lastDayOfMonth));
          }

          if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
            dataVencimentoParcela = calcularVencimentoCartaoParaParcela(infoConta, dataVencimentoBase, i + 1, transacao.parcelasTotais, contasSheet.getDataRange().getValues());
          }

          // CORREÇÃO: Lógica de ID simplificada
          let idFinal;
          if (transacoesParaRegistrar.length > 1) { // Transferência
              idFinal = `${transacao.id}-${transferCounter}`;
          } else if (transacao.parcelasTotais > 1) { // Parcelado
              idFinal = `${transacao.id}-${i + 1}`;
          } else { // À vista
              idFinal = transacao.id;
          }

          registrarTransacaoNaPlanilha(
            dataTransacaoBase,
            transacao.descricao,
            transacao.categoria,
            transacao.subcategoria,
            transacao.tipo,
            valorParcela,
            transacao.metodoPagamento,
            transacao.conta,
            transacao.parcelasTotais,
            i + 1,
            dataVencimentoParcela,
            usuario,
            transacao.status,
            idFinal,
            dataRegistroBase
          );
        }
        if (transacoesParaRegistrar.length > 1) {
            transferCounter++;
        }
    }
    
    if (transacoesParaRegistrar.length > 1) {
        enviarMensagemTelegram(chatId, `✅ Transferência de *${formatCurrency(transacoesParaRegistrar[0].valor)}* registrada com sucesso!`);
    } else {
        enviarMensagemTelegram(chatId, `✅ Lançamento de *${formatCurrency(transacoesParaRegistrar[0].valor)}* (${transacoesParaRegistrar[0].parcelasTotais}x) registrado com sucesso!`);
    }
    
    logToSheet(`Transacao ${transacaoData.id || transacaoData[0].id} confirmada e registrada por ${usuario}.`, "INFO");
    atualizarSaldosDasContas();

  } catch (e) {
    logToSheet(`ERRO ao registrar transacao confirmada: ${e.message} na linha ${e.lineNumber}. Stack: ${e.stack}`, "ERROR");
    enviarMensagemTelegram(chatId, `❌ Houve um erro ao registrar sua transação: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancela uma transação pendente.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} transactionId O ID da transação pendente.
 */
function cancelarTransacaoPendente(chatId, transactionId) {
  enviarMensagemTelegram(chatId, "❌ Lançamento cancelado.");
  logToSheet(`Transacao ${transactionId} cancelada por ${chatId}.`, "INFO");
}


/**
 * ATUALIZADO: Calcula a data de vencimento da fatura do cartão de crédito para uma transação.
 * @param {Object} infoConta O objeto de informações da conta (do 'Contas.gs').
 * @param {Date} transactionDate A data da transacao a ser usada como referencia.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @returns {Date} A data de vencimento calculada.
 */
function calcularVencimentoCartao(infoConta, transactionDate, dadosContas) {
    const diaTransacao = transactionDate.getDate();
    const mesTransacao = transactionDate.getMonth();
    const anoTransacao = transactionDate.getFullYear();

    const diaFechamento = infoConta.diaFechamento;
    const diaVencimento = infoConta.vencimento;
    const tipoFechamento = infoConta.tipoFechamento || "padrao";

    logToSheet(`[CalcVencimento] Calculando vencimento para ${infoConta.nomeOriginal}. Transacao em: ${transactionDate.toLocaleDateString()}, Dia Fechamento: ${diaFechamento}, Dia Vencimento: ${diaVencimento}, Tipo Fechamento: ${tipoFechamento}`, "DEBUG");

    let mesFechamento;
    let anoFechamento;

    if (tipoFechamento === "padrao" || tipoFechamento === "fechamento-mes") {
        if (diaTransacao <= diaFechamento) {
            mesFechamento = mesTransacao;
            anoFechamento = anoTransacao;
        } else {
            mesFechamento = mesTransacao + 1;
            anoFechamento = anoTransacao;
        }
    } else if (tipoFechamento === "fechamento-anterior") {
        mesFechamento = mesTransacao;
        anoFechamento = anoTransacao;
    } else {
        logToSheet(`[CalcVencimento] Tipo de fechamento desconhecido: ${tipoFechamento}. Assumindo padrao.`, "WARN");
        if (diaTransacao <= diaFechamento) {
            mesFechamento = mesTransacao;
            anoFechamento = anoTransacao;
        } else {
            mesFechamento = mesTransacao + 1;
            anoFechamento = anoTransacao;
        }
    }

    let vencimentoAno = anoFechamento;
    let vencimentoMes = mesFechamento + 1;

    if (vencimentoMes > 11) {
        vencimentoMes -= 12;
        vencimentoAno++;
    }

    let dataVencimento = new Date(vencimentoAno, vencimentoMes, diaVencimento);

    if (dataVencimento.getMonth() !== vencimentoMes) {
        dataVencimento = new Date(vencimentoAno, vencimentoMes + 1, 0);
    }
    
    logToSheet(`[CalcVencimento] Data de Vencimento Final Calculada: ${dataVencimento.toLocaleDateString()}`, "DEBUG");
    return dataVencimento;
}

/**
 * NOVO: Calcula a data de vencimento da fatura do cartão de crédito para uma PARCELA específica.
 * Essencial para garantir que cada parcela tenha a data de vencimento correta.
 * @param {Object} infoConta O objeto de informações da conta (do 'Contas.gs').
 * @param {Date} dataPrimeiraParcelaVencimento A data de vencimento da primeira parcela (já calculada por calcularVencimentoCartao).
 * @param {number} numeroParcela O número da parcela atual (1, 2, 3...).
 * @param {number} totalParcelas O número total de parcelas.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @returns {Date} A data de vencimento calculada para a parcela.
 */
function calcularVencimentoCartaoParaParcela(infoConta, dataPrimeiraParcelaVencimento, numeroParcela, totalParcelas, dadosContas) {
    if (numeroParcela === 1) {
        return dataPrimeiraParcelaVencimento;
    }

    // Começa com a data de vencimento da primeira parcela
    let dataVencimentoParcela = new Date(dataPrimeiraParcelaVencimento);

    // Adiciona o número de meses correspondente à parcela
    dataVencimentoParcela.setMonth(dataVencimentoParcela.getMonth() + (numeroParcela - 1));

    // Ajuste para garantir que o dia do vencimento não "pule" para o mês seguinte
    // se o dia do vencimento original for maior que o número de dias no mês atual
    // (ex: 31 de janeiro -> 31 de março, mas fevereiro não tem dia 31).
    if (dataVencimentoParcela.getDate() !== dataPrimeiraParcelaVencimento.getDate()) {
        const lastDayOfMonth = new Date(dataVencimentoParcela.getFullYear(), dataVencimentoParcela.getMonth() + 1, 0).getDate();
        dataVencimentoParcela.setDate(Math.min(dataVencimentoParcela.getDate(), lastDayOfMonth)); // Use o dia atual da parcela, não o dia original
    }
    logToSheet(`[CalcVencimentoParcela] Calculado vencimento para parcela ${numeroParcela} de ${infoConta.nomeOriginal}: ${dataVencimentoParcela.toLocaleDateString()}`, "DEBUG");
    return dataVencimentoParcela;
}

// --- CORREÇÃO ---
// Lógica de `atualizarSaldosDasContas` foi reestruturada para maior clareza e precisão,
// especialmente no cálculo de faturas consolidadas.
/**
 * ATUALIZADO: Atualiza os saldos de todas as contas na planilha 'Contas'
 * e os armazena na variável global `globalThis.saldosCalculados`.
 * Esta é uma função crucial para manter os dados do dashboard e dos comandos do bot atualizados.
 */
function atualizarSaldosDasContas() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 

  try {
    logToSheet("Iniciando atualizacao de saldos das contas.", "INFO");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    
    if (!contasSheet || !transacoesSheet) {
      logToSheet("Erro: Aba 'Contas' ou 'Transacoes' não encontrada.", "ERROR");
      return;
    }

    const dadosContas = contasSheet.getDataRange().getValues();
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    
    globalThis.saldosCalculados = {}; // Limpa os saldos anteriores

    // --- PASSO 1: Inicializa todas as contas ---
    for (let i = 1; i < dadosContas.length; i++) {
      const linha = dadosContas[i];
      const nomeOriginal = (linha[0] || "").toString().trim();
      if (!nomeOriginal) continue;

      const nomeNormalizado = normalizarTexto(nomeOriginal);
      globalThis.saldosCalculados[nomeNormalizado] = {
        nomeOriginal: nomeOriginal,
        nomeNormalizado: nomeNormalizado,
        tipo: (linha[1] || "").toString().toLowerCase().trim(),
        saldo: parseBrazilianFloat(String(linha[3] || '0')), // Saldo Inicial
        limite: parseBrazilianFloat(String(linha[5] || '0')),
        vencimento: parseInt(linha[6]) || null,
        diaFechamento: parseInt(linha[9]) || null,
        tipoFechamento: (linha[10] || "").toString().trim(),
        contaPaiAgrupador: normalizarTexto((linha[12] || "").toString().trim()),
        faturaAtual: 0, // Gastos do ciclo que vence no próximo mês
        saldoTotalPendente: 0 // Saldo devedor total
      };
    }
    logToSheet("[AtualizarSaldos] Passo 1/4: Contas inicializadas.", "DEBUG");


    // --- PASSO 2: Processa transações para calcular saldos individuais ---
    const today = new Date();
    let nextCalendarMonth = today.getMonth() + 1;
    let nextCalendarYear = today.getFullYear();
    if (nextCalendarMonth > 11) {
        nextCalendarMonth = 0;
        nextCalendarYear++;
    }

    for (let i = 1; i < dadosTransacoes.length; i++) {
      const linha = dadosTransacoes[i];
      const tipoTransacao = (linha[4] || "").toString().toLowerCase().trim();
      const valor = parseBrazilianFloat(String(linha[5] || '0'));
      const contaNormalizada = normalizarTexto(linha[7] || "");
      const categoria = normalizarTexto(linha[2] || "");
      const subcategoria = normalizarTexto(linha[3] || "");
      const dataVencimento = parseData(linha[10]);

      if (!globalThis.saldosCalculados[contaNormalizada]) continue;

      const infoConta = globalThis.saldosCalculados[contaNormalizada];

      if (infoConta.tipo === "conta corrente" || infoConta.tipo === "dinheiro físico") {
        if (tipoTransacao === "receita") infoConta.saldo += valor;
        else if (tipoTransacao === "despesa") infoConta.saldo -= valor;
      } else if (infoConta.tipo === "cartão de crédito") {
        const isPayment = (categoria === "contas a pagar" && subcategoria === "pagamento de fatura");
        if (isPayment) {
          infoConta.saldoTotalPendente -= valor;
        } else if (tipoTransacao === "despesa") {
          infoConta.saldoTotalPendente += valor;
          if (dataVencimento && dataVencimento.getMonth() === nextCalendarMonth && dataVencimento.getFullYear() === nextCalendarYear) {
            infoConta.faturaAtual += valor;
          }
        }
      }
    }
    logToSheet("[AtualizarSaldos] Passo 2/4: Saldos individuais calculados.", "DEBUG");


    // --- PASSO 3: Consolida saldos de cartões em 'Faturas Consolidadas' ---
    for (const nomeNormalizado in globalThis.saldosCalculados) {
      const infoConta = globalThis.saldosCalculados[nomeNormalizado];
      if (infoConta.tipo === "cartão de crédito" && infoConta.contaPaiAgrupador) {
        const agrupadorNormalizado = infoConta.contaPaiAgrupador;
        if (globalThis.saldosCalculados[agrupadorNormalizado] && globalThis.saldosCalculados[agrupadorNormalizado].tipo === "fatura consolidada") {
          const agrupador = globalThis.saldosCalculados[agrupadorNormalizado];
          agrupador.saldoTotalPendente += infoConta.saldoTotalPendente;
          agrupador.faturaAtual += infoConta.faturaAtual;
        }
      }
    }
    logToSheet("[AtualizarSaldos] Passo 3/4: Saldos consolidados.", "DEBUG");


    // --- PASSO 4: Atualiza a planilha 'Contas' com os novos saldos ---
    const saldosParaPlanilha = [];
    for (let i = 1; i < dadosContas.length; i++) {
      const nomeOriginal = (dadosContas[i][0] || "").toString().trim();
      const nomeNormalizado = normalizarTexto(nomeOriginal);
      if (globalThis.saldosCalculados[nomeNormalizado]) {
        const infoConta = globalThis.saldosCalculados[nomeNormalizado];
        let saldoFinal;
        if (infoConta.tipo === "fatura consolidada" || infoConta.tipo === "cartão de crédito") {
          saldoFinal = infoConta.saldoTotalPendente;
        } else {
          saldoFinal = infoConta.saldo;
        }
        saldosParaPlanilha.push([round(saldoFinal, 2)]);
      } else {
        saldosParaPlanilha.push([dadosContas[i][4]]); // Mantém o valor antigo se a conta não foi encontrada
      }
    }

    if (saldosParaPlanilha.length > 0) {
      // Coluna E (índice 4) é a 'Saldo Atualizado'
      contasSheet.getRange(2, 5, saldosParaPlanilha.length, 1).setValues(saldosParaPlanilha);
    }
    logToSheet("[AtualizarSaldos] Passo 4/4: Planilha 'Contas' atualizada.", "INFO");

  } catch (e) {
    logToSheet(`ERRO FATAL em atualizarSaldosDasContas: ${e.message} na linha ${e.lineNumber}. Stack: ${e.stack}`, "ERROR");
  } finally {
    lock.releaseLock();
  }
}


/**
 * NOVO: Gera as contas recorrentes para o próximo mês com base na aba 'Contas_a_Pagar'.
 * Evita duplicatas e ajusta o dia de vencimento se o dia original não existir no próximo mês.
 * Esta função é acionada por um gatilho de tempo ou manualmente.
 */
function generateRecurringBillsForNextMonth() {
    logToSheet("Iniciando geracao de contas recorrentes para o proximo mes.", "INFO");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
    
    if (!contasAPagarSheet) {
        logToSheet("Erro: Aba 'Contas_a_Pagar' nao encontrada para gerar contas recorrentes.", "ERROR");
        throw new Error("Aba 'Contas_a_Pagar' não encontrada.");
    }

    const dadosContasAPagar = contasAPagarSheet.getDataRange().getValues();
    const headers = dadosContasAPagar[0];

    const colID = headers.indexOf('ID');
    const colDescricao = headers.indexOf('Descricao');
    const colCategoria = headers.indexOf('Categoria');
    const colValor = headers.indexOf('Valor');
    const colDataVencimento = headers.indexOf('Data de Vencimento');
    const colStatus = headers.indexOf('Status');
    const colRecorrente = headers.indexOf('Recorrente');
    const colContaSugeria = headers.indexOf('Conta de Pagamento Sugerida');
    const colObservacoes = headers.indexOf('Observacoes');
    const colIDTransacaoVinculada = headers.indexOf('ID Transacao Vinculada');

    if ([colID, colDescricao, colCategoria, colValor, colDataVencimento, colStatus, colRecorrente, colContaSugeria, colObservacoes, colIDTransacaoVinculada].some(idx => idx === -1)) {
        logToSheet("Erro: Colunas essenciais faltando na aba 'Contas_a_Pagar' para geracao de contas recorrentes.", "ERROR");
        throw new Error("Colunas essenciais faltando na aba 'Contas_a_Pagar'.");
    }

    const today = new Date();
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const nextMonthNum = nextMonth.getMonth(); // 0-indexed
    const nextYearNum = nextMonth.getFullYear();

    logToSheet(`Gerando contas recorrentes para: ${getNomeMes(nextMonthNum)}/${nextYearNum}`, "DEBUG");

    const newBills = [];
    const existingBillsInNextMonth = new Set(); // Para evitar duplicatas

    // Primeiro, verifica as contas já existentes para o próximo mês
    for (let i = 1; i < dadosContasAPagar.length; i++) {
        const row = dadosContasAPagar[i];
        const dataVencimentoExistente = parseData(row[colDataVencimento]);
        if (dataVencimentoExistente &&
            dataVencimentoExistente.getMonth() === nextMonthNum &&
            dataVencimentoExistente.getFullYear() === nextYearNum) {
            existingBillsInNextMonth.add(normalizarTexto(row[colDescricao] + row[colValor] + row[colCategoria]));
        }
    }
    logToSheet(`Contas existentes no proximo mes: ${existingBillsInNextMonth.size}`, "DEBUG");


    // Processa contas do mês atual para gerar para o próximo
    for (let i = 1; i < dadosContasAPagar.length; i++) {
        const row = dadosContasAPagar[i];
        const recorrente = (row[colRecorrente] || "").toString().trim().toLowerCase();
        
        if (recorrente === "verdadeiro") {
            const currentDescricao = (row[colDescricao] || "").toString().trim();
            const currentValor = parseBrazilianFloat(String(row[colValor]));
            const currentCategoria = (row[colCategoria] || "").toString().trim();
            const currentDataVencimento = parseData(row[colDataVencimento]);
            const currentContaSugeria = (row[colContaSugeria] || "").toString().trim();
            const currentObservacoes = (row[colObservacoes] || "").toString().trim();
            
            // Cria uma chave única para a conta baseada em seus atributos principais
            const billKey = normalizarTexto(currentDescricao + currentValor + currentCategoria);

            // Verifica se a conta já existe para o próximo mês
            if (existingBillsInNextMonth.has(billKey)) {
                logToSheet(`Conta recorrente "${currentDescricao}" ja existe para ${getNomeMes(nextMonthNum)}/${nextYearNum}. Pulando.`, "DEBUG");
                continue;
            }

            if (currentDataVencimento) {
                let newDueDate = new Date(currentDataVencimento);
                newDueDate.setMonth(newDueDate.getMonth() + 1); // Avança um mês

                // Ajusta o dia para o último dia do mês se o dia original não existir no novo mês
                // Ex: 31 de janeiro -> 28/29 de fevereiro
                if (newDueDate.getDate() !== currentDataVencimento.getDate()) {
                    newDueDate = new Date(newDueDate.getFullYear(), newDueDate.getMonth() + 1, 0); // Último dia do mês
                }

                const newRow = [
                    Utilities.getUuid(), // Novo ID único
                    currentDescricao,
                    currentCategoria,
                    currentValor,
                    Utilities.formatDate(newDueDate, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                    "Pendente", // Status inicial
                    "Verdadeiro", // Continua sendo recorrente
                    currentContaSugeria,
                    currentObservacoes,
                    "" // ID Transacao Vinculada (vazio)
                ];
                newBills.push(newRow);
                logToSheet(`Conta recorrente "${currentDescricao}" gerada para ${getNomeMes(newDueDate.getMonth())}/${newDueDate.getFullYear()}.`, "INFO");
            }
        }
    }

    if (newBills.length > 0) {
        contasAPagarSheet.getRange(contasAPagarSheet.getLastRow() + 1, 1, newBills.length, newBills[0].length).setValues(newBills);
        logToSheet(`Total de ${newBills.length} contas recorrentes adicionadas.`, "INFO");
    } else {
        logToSheet("Nenhuma nova conta recorrente para adicionar para o proximo mes.", "INFO");
    }
}

/**
 * NOVO: Processa o comando /marcar_pago vindo do Telegram.
 * Marca uma conta a pagar como "Pago" na planilha e tenta vincular a uma transação existente.
 * Se não encontrar uma transação, pergunta se o usuário deseja registrar uma agora.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} textoRecebido O texto completo do comando (/marcar_pago_<ID_CONTA>).
 * @param {string} usuario O nome do usuário.
 */
function processarMarcarPago(chatId, textoRecebido, usuario) {
  const idContaAPagar = textoRecebido.substring("/marcar_pago_".length);
  logToSheet(`[MarcarPago] Processando marcar pago para ID: ${idContaAPagar}`, "INFO");

  const contaAPagarInfo = obterInformacoesDaContaAPagar(idContaAPagar);

  if (!contaAPagarInfo) {
    enviarMensagemTelegram(chatId, `❌ Conta a Pagar com ID *${escapeMarkdown(idContaAPagar)}* não encontrada.`);
    logToSheet(`Erro: Conta a Pagar ID ${idContaAPagar} não encontrada para marcar como pago.`, "WARN");
    return;
  }

  if (normalizarTexto(contaAPagarInfo.status) === "pago") {
    enviarMensagemTelegram(chatId, `ℹ️ A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* já está paga.`);
    logToSheet(`Conta a Pagar ID ${idContaAPagar} já está paga.`, "INFO");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const dadosTransacoes = transacoesSheet.getDataRange().getValues();

  // Tenta encontrar uma transação correspondente para vincular
  let transacaoVinculada = null;
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const linha = dadosTransacoes[i];
    const dataTransacao = parseData(linha[0]);
    const descricaoTransacao = normalizarTexto(linha[1]);
    const valorTransacao = parseBrazilianFloat(String(linha[5]));
    const idTransacao = linha[13];

    // Verifica se a transação é do mês atual, do tipo despesa,
    // e se a descrição e o valor são semelhantes
    if (dataTransacao && dataTransacao.getMonth() === mesAtual && dataTransacao.getFullYear() === anoAtual &&
        normalizarTexto(linha[4]) === "despesa" &&
        calculateSimilarity(descricaoTransacao, normalizarTexto(contaAPagarInfo.descricao)) > SIMILARITY_THRESHOLD &&
        Math.abs(valorTransacao - contaAPagarInfo.valor) < 0.01) { // Margem de erro para o valor
        transacaoVinculada = idTransacao;
        logToSheet(`[MarcarPago] Transacao existente (ID: ${idTransacao}) encontrada para vincular a conta ${idContaAPagar}.`, "INFO");
        break;
    }
  }

  if (transacaoVinculada) {
    vincularTransacaoAContaAPagar(chatId, idContaAPagar, transacaoVinculada);
  } else {
    // Se não encontrou transação existente, pergunta se quer registrar uma agora
    const mensagem = `A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* (R$ ${contaAPagarInfo.valor.toFixed(2).replace('.', ',')}) será marcada como paga.`;
    const teclado = {
      inline_keyboard: [
        [{ text: "✅ Marcar como Pago (sem registrar transação)", callback_data: `confirm_marcar_pago_sem_transacao_${idContaAPagar}` }],
        [{ text: "📝 Registrar e Marcar como Pago", callback_data: `confirm_marcar_pago_e_registrar_${idContaAPagar}` }],
        [{ text: "❌ Cancelar", callback_data: `cancel_${idContaAPagar}` }]
      ]
    };
    enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
    logToSheet(`[MarcarPago] Nenhuma transacao existente encontrada para ${idContaAPagar}. Solicitando acao do usuario.`, "INFO");
  }
}

/**
 * NOVO: Função para lidar com a confirmação de marcar conta a pagar.
 * Esta função é chamada a partir de um callback_query.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} action O tipo de ação (sem_transacao ou e_registrar).
 * @param {string} idContaAPagar O ID da conta a pagar.
 * @param {string} usuario O nome do usuário.
 */
function handleMarcarPagoConfirmation(chatId, action, idContaAPagar, usuario) {
  logToSheet(`[MarcarPagoConfirm] Acão: ${action}, ID Conta: ${idContaAPagar}, Usuario: ${usuario}`, "INFO");

  const contaAPagarInfo = obterInformacoesDaContaAPagar(idContaAPagar);

  if (!contaAPagarInfo) {
    enviarMensagemTelegram(chatId, `❌ Conta a Pagar com ID *${escapeMarkdown(idContaAPagar)}* não encontrada.`);
    logToSheet(`Erro: Conta a Pagar ID ${idContaAPagar} não encontrada para confirmação de marcar como pago.`, "WARN");
    return;
  }

  if (normalizarTexto(contaAPagarInfo.status) === "pago") {
    enviarMensagemTelegram(chatId, `ℹ️ A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* já está paga.`);
    logToSheet(`Conta a Pagar ID ${idContaAPagar} já está paga.`, "INFO");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
  const colStatus = contaAPagarInfo.headers.indexOf('Status') + 1;
  const colIDTransacaoVinculada = contaAPagarInfo.headers.indexOf('ID Transacao Vinculada') + 1;

  if (action === "sem_transacao") {
    try {
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colStatus).setValue("Pago");
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colIDTransacaoVinculada).setValue("MARCADO_MANUALMENTE"); // Indica que foi pago manualmente
      enviarMensagemTelegram(chatId, `✅ Conta *${escapeMarkdown(contaAPagarInfo.descricao)}* marcada como paga (sem registro de transação).`);
      logToSheet(`Conta a Pagar ${idContaAPagar} marcada como paga manualmente.`, "INFO");
      atualizarSaldosDasContas();
    } catch (e) {
      logToSheet(`ERRO ao marcar conta a pagar ${idContaAPagar} sem transacao: ${e.message}`, "ERROR");
      enviarMensagemTelegram(chatId, `❌ Erro ao marcar conta como paga: ${e.message}`);
    }
  } else if (action === "e_registrar") {
    try {
      // Cria uma transação com os dados da conta a pagar
      const transacaoData = {
        id: Utilities.getUuid(),
        data: new Date(),
        descricao: `Pagamento de ${contaAPagarInfo.descricao}`,
        categoria: contaAPagarInfo.categoria,
        subcategoria: "Pagamento de Fatura" || "", // Se não houver, padrão para Pagamento de Fatura
        tipo: "Despesa",
        valor: contaAPagarInfo.valor,
        metodoPagamento: contaAPagarInfo.contaDePagamentoSugeria || "Débito", // Usa conta sugerida ou padrão
        conta: contaAPagarInfo.contaDePagamentoSugeria || "Não Identificada", // Usa conta sugerida ou padrão
        parcelasTotais: 1,
        parcelaAtual: 1,
        dataVencimento: contaAPagarInfo.dataVencimento,
        usuario: usuario,
        status: "Ativo",
        dataRegistro: new Date()
      };
      
      registrarTransacaoConfirmada(transacaoData, usuario, chatId); // Registra a transação
      
      // Vincula a nova transação à conta a pagar
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colStatus).setValue("Pago");
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colIDTransacaoVinculada).setValue(transacaoData.id);
      logToSheet(`Conta a Pagar ${idContaAPagar} marcada como paga e vinculada a nova transacao ${transacaoData.id}.`, "INFO");
      enviarMensagemTelegram(chatId, `✅ Transação de *${formatCurrency(transacaoData.valor)}* para *${escapeMarkdown(contaAPagarInfo.descricao)}* registrada e conta marcada como paga!`);
      atualizarSaldosDasContas();
    } catch (e) {
      logToSheet(`ERRO ao registrar e marcar conta a pagar ${idContaAPagar}: ${e.message}`, "ERROR");
      enviarMensagemTelegram(chatId, `❌ Erro ao registrar e marcar conta como paga: ${e.message}`);
    }
  }
}

/**
 * NOVO: Interpreta uma mensagem de transferência.
 * @param {string} textoNormalizado O texto da mensagem normalizada.
 * @param {number} valor O valor da transferência.
 * @param {string} usuario O nome do usuário.
 * @param {string} chatId O ID do chat.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba 'PalavrasChave'.
 * @returns {Object} Um objeto de confirmação ou erro.
 */
function interpretarTransferencia(textoNormalizado, valor, usuario, chatId, dadosContas, dadosPalavras) {
    const match = textoNormalizado.match(/(?:de|do)\s(.+?)\s(?:para|pra)\s(.+)/);
    if (!match) {
        return { errorMessage: "Para transferências, use o formato 'transferi [valor] de [conta origem] para [conta destino]'." };
    }

    const textoContaOrigem = match[1].trim();
    const textoContaDestino = match[2].trim();

    const { conta: contaOrigem } = extrairContaMetodoPagamento(textoContaOrigem, dadosContas, dadosPalavras);
    const { conta: contaDestino } = extrairContaMetodoPagamento(textoContaDestino, dadosContas, dadosPalavras);

    if (contaOrigem === "Não Identificada" || contaDestino === "Não Identificada") {
        return { errorMessage: `Não consegui identificar as contas da transferência. Origem encontrada: "${contaOrigem}", Destino encontrado: "${contaDestino}".` };
    }

    const transactionId = Utilities.getUuid();
    
    const transacaoSaida = {
        id: transactionId, // Usa o mesmo ID base
        data: new Date(),
        descricao: `Transferência para ${contaDestino}`,
        // CORREÇÃO: Adiciona o emoji à categoria
        categoria: "🔄 Transferências",
        subcategoria: "Entre Contas",
        tipo: "Despesa",
        valor: valor,
        metodoPagamento: "Transferência",
        conta: contaOrigem,
        parcelasTotais: 1,
        parcelaAtual: 1,
        dataVencimento: new Date(),
        usuario: usuario,
        status: "Ativo",
        dataRegistro: new Date()
    };

    const transacaoEntrada = {
        id: transactionId, // Usa o mesmo ID base
        data: new Date(),
        descricao: `Transferência de ${contaOrigem}`,
        // CORREÇÃO: Adiciona o emoji à categoria
        categoria: "🔄 Transferências",
        subcategoria: "Entre Contas",
        tipo: "Receita",
        valor: valor,
        metodoPagamento: "Transferência",
        conta: contaDestino,
        parcelasTotais: 1,
        parcelaAtual: 1,
        dataVencimento: new Date(),
        usuario: usuario,
        status: "Ativo",
        dataRegistro: new Date()
    };
    
    const transacaoData = [transacaoSaida, transacaoEntrada];
    
    return prepararConfirmacaoTransferencia(transacaoData, chatId);
}

/**
 * NOVO: Prepara e envia uma mensagem de confirmação para transferências.
 * @param {Array<Object>} transacoes Array com as duas transações (saída e entrada).
 * @param {string} chatId O ID do chat.
 * @returns {Object} Status de confirmação pendente.
 */
function prepararConfirmacaoTransferencia(transacoes, chatId) {
    const cache = CacheService.getScriptCache();
    const transactionId = transacoes[0].id; // ID base
    const cacheKey = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${transactionId}`;
    cache.put(cacheKey, JSON.stringify(transacoes), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS);

    const saida = transacoes[0];
    const entrada = transacoes[1];

    let mensagem = `✅ Confirme sua Transferência:\n\n`;
    mensagem += `*Valor:* ${formatCurrency(saida.valor)}\n`;
    mensagem += `*De:* ${escapeMarkdown(saida.conta)}\n`;
    mensagem += `*Para:* ${escapeMarkdown(entrada.conta)}\n`;

    const teclado = {
        inline_keyboard: [
            [{ text: "✅ Confirmar", callback_data: `confirm_${transactionId}` }],
            [{ text: "❌ Cancelar", callback_data: `cancel_${transactionId}` }]
        ]
    };

    enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
    return { status: "PENDING_CONFIRMATION", transactionId: transactionId };
}
