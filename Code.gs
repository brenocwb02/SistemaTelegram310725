/**
 * @file Code.gs
 * @description Este arquivo contém a função principal `doPost` que atua como o webhook do Telegram,
 * processando as mensagens e callbacks recebidas.
 */

// Variável global para armazenar os saldos calculados.
// Usar `globalThis` é uma boa prática para garantir que ela seja acessível em diferentes arquivos .gs.
// É populada pela função `atualizarSaldosDasContas` em FinancialLogic.gs.
globalThis.saldosCalculados = {};

/**
 * NOVO: Gera e envia um link de acesso seguro e temporário para o Dashboard.
 * Esta função foi movida para este arquivo para corrigir o erro 'not defined'.
 * @param {string} chatId O ID do chat do Telegram.
 */
function enviarLinkDashboard(chatId) {
  logToSheet(`[Dashboard] Gerando link de acesso seguro para o chatId: ${chatId}`, "INFO");
  const cache = CacheService.getScriptCache();
  
  // 1. Gera um token único e aleatório.
  const token = Utilities.getUuid();

  // 2. Armazena o token no cache, associando-o ao chatId do usuário.
  // A chave é o token, o valor é o chatId. A validade é definida na constante.
  const cacheKey = `${CACHE_KEY_DASHBOARD_TOKEN}_${token}`;
  cache.put(cacheKey, chatId.toString(), CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS);
  logToSheet(`[Dashboard] Token '${token}' armazenado no cache para o chatId '${chatId}' por ${CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS} segundos.`, "DEBUG");

  // 3. Obtém a URL do Web App.
  const webAppUrl = ScriptApp.getService().getUrl();

  // 4. Constrói a URL segura com o token como parâmetro.
  const secureUrl = `${webAppUrl}?token=${token}`;

  // 5. Envia a mensagem para o usuário.
  const mensagem = `Aqui está o seu link de acesso seguro ao Dashboard Financeiro. \n\n` +
                   `*Atenção:* Este link é de uso único e expira em ${CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS / 60} minutos.\n\n` +
                   `[Clique aqui para abrir o Dashboard](${secureUrl})`;
  
  // Envia com parse_mode 'Markdown' para garantir a formatação do link.
  enviarMensagemTelegram(chatId, mensagem, { parse_mode: 'Markdown' });
}



/**
 * Função principal que é acionada pelo webhook do Telegram.
 * Processa as mensagens e callbacks recebidas.
 * @param {Object} e O objeto de evento do webhook.
 */
function doPost(e) {
  // Bloco try...catch para capturar qualquer erro inesperado durante a execução.
  try {
    if (!e || !e.postData || !e.postData.contents) {
      logToSheet("doPost recebido com dados vazios ou invalidos (e.postData.contents esta vazio). Ignorando.", "WARN");
      return;
    }

    currentLogLevel = getLogLevelConfig();
    logToSheet(`Nível de log configurado para esta execução: ${currentLogLevel}`, "INFO");

    const data = JSON.parse(e.postData.contents || '{}');
    const chatId = data.message?.chat?.id || data.callback_query?.message?.chat?.id;
    let textoRecebido = (data.message?.text || data.callback_query?.data || "").trim();

    // --- Carrega dados essenciais no início ---
    const configData = getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
    const usuario = getUsuarioPorChatId(chatId, configData);
    
    // NOVO: Verifica se existe um estado de edição ativo.
    const editState = getEditState(chatId);

    const updateId = data.update_id;
    if (updateId) {
      const cache = CacheService.getScriptCache();
      const cachedUpdate = cache.get(updateId.toString());
      if (cachedUpdate) {
        logToSheet(`Update ID ${updateId} ja processado. Ignorando execucao duplicada.`, "WARN");
        return;
      }
      cache.put(updateId.toString(), "processed", 60);
    }

    let comandoBase;
    let complemento;

    if (data.callback_query) {
      answerCallbackQuery(data.callback_query.id);
      logToSheet(`Callback query ID ${data.callback_query.id} reconhecida.`, "DEBUG");

      // CORREÇÃO: Priorizar o tratamento de callbacks de edição se houver um editState ativo
      if (editState && textoRecebido.startsWith('edit_')) {
          comandoBase = "/editar_campo_callback"; // Comando base interno para callbacks de edição
          complemento = textoRecebido.substring("edit_".length); // Apenas o nome do campo
          logToSheet(`[doPost] Callback de edição detectado com editState ativo. Campo: ${complemento}`, "INFO");
      } else if (textoRecebido.startsWith('confirm_')) {
        comandoBase = "/confirm";
        complemento = textoRecebido.substring('confirm_'.length);
      } else if (textoRecebido.startsWith('cancel_')) {
        comandoBase = "/cancel";
        complemento = textoRecebido.substring('cancel_'.length);
      }
      else if (textoRecebido.startsWith('/tutorial_')) {
        comandoBase = textoRecebido;
        complemento = "";
      }
      else if (textoRecebido.startsWith("/marcar_pago_")) {
        comandoBase = "/marcar_pago";
        complemento = textoRecebido.substring("/marcar_pago_".length);
      } else if (textoRecebido.startsWith("/excluir_")) {
        comandoBase = "/excluir";
        complemento = textoRecebido.substring("/excluir_".length);
      } else if (textoRecebido.startsWith("/extrato_usuario_")) {
        comandoBase = "/extrato_usuario";
        complemento = textoRecebido.substring("/extrato_usuario_".length);
      }
      else if (textoRecebido === "cancelar_edicao") { // Callback para cancelar edição
          comandoBase = "/cancelar_edicao";
          complemento = "";
      }
      else {
        comandoBase = textoRecebido.startsWith("/") ? textoRecebido : "/" + textoRecebido;
        complemento = "";
      }
    }
    else if (data.message) {
      const textoLimpo = textoRecebido.trim();
      const partesTexto = textoLimpo.split(/\s+/);
      const primeiraPalavra = partesTexto[0].toLowerCase();
      
      // CORREÇÃO: Lógica de detecção de comando refeita para maior robustez.
      // 1. Normaliza o comando removendo a barra inicial, se houver.
      const comandoNormalizado = primeiraPalavra.startsWith('/') ? primeiraPalavra.substring(1) : primeiraPalavra;

      // 2. Define a lista de comandos conhecidos SEM a barra.
      // REMOVIDO: "editar_campo" e seus variantes não são comandos diretos, são callbacks.
      const comandosConhecidosSemBarra = ["start", "dashboard", "resumo", "saldo", "extrato", "proximasfaturas", "contasapagar", "metas", "ajuda", "editar", "vincular_conta", "tutorial", "adicionar_conta", "listar_contas", "adicionar_categoria", "listar_categorias", "listar_subcategorias"];

      // 3. Verifica se o comando normalizado está na lista.
      if (comandosConhecidosSemBarra.includes(comandoNormalizado)) {
          comandoBase = `/${comandoNormalizado}`; // Adiciona a barra de volta para o switch
          complemento = partesTexto.slice(1).join(" ");
      } 
      // NOVO: Se há um estado de edição ativo E a mensagem NÃO é um comando,
      // então é a entrada do usuário para o campo que está sendo editado.
      else if (editState && !textoLimpo.startsWith('/')) {
          logToSheet(`[doPost] Estado de edição detectado para ${chatId}. Processando entrada de edição.`, "INFO");
          processarEdicaoFinal(chatId, usuario, textoRecebido, editState, dadosContas);
          return; // Importante: retorna para não continuar processando como lançamento
      }
      else {
          comandoBase = "/lancamento"; // Se não é comando conhecido e não é entrada de edição, tenta como lançamento
          complemento = textoLimpo;
      }
    } else {
      logToSheet("Webhook recebido, mas sem mensagem ou callback query reconhecida.", "INFO");
      return;
    }

    logToSheet(`doPost - Chat ID: ${chatId}, Texto Recebido: "${textoRecebido}", Comando Base: "${comandoBase}", Complemento: "${complemento}"`, "INFO");

    const debugTutorialState = getTutorialState(chatId);
    logToSheet(`[DEBUG doPost Start] ChatID: ${chatId}, ComandoBase: "${comandoBase}", TextoRecebido: "${textoRecebido}", TutorialState: ${JSON.stringify(debugTutorialState)}`, "DEBUG");

    if (usuario === "Desconhecido") {
      enviarMensagemTelegram(chatId, "❌ Voce não está autorizado a usar este bot.");
      logToSheet(`Usuario ${chatId} não autorizado.`, "WARN");
      return;
    }

    const { month: targetMonth, year: targetYear } = parseMonthAndYear(complemento);
    logToSheet(`doPost - Mes Alvo: ${targetMonth}, Ano Alvo: ${targetYear}`, "DEBUG");

    if (debugTutorialState && debugTutorialState.currentStep > 0 &&
        !comandoBase.startsWith("/tutorial_") &&
        comandoBase !== "/confirm" &&
        comandoBase !== "/cancel") {

      logToSheet(`[doPost] Usuario ${chatId} esta no tutorial (Passo ${debugTutorialState.currentStep}, Acao Esperada: ${debugTutorialState.expectedAction}). Tentando processar input pelo tutorial.`, "INFO");
      const handledByTutorial = processTutorialInput(chatId, usuario, textoRecebido, debugTutorialState);
      if (handledByTutorial) {
        logToSheet(`[doPost] Mensagem tratada pelo tutorial.`, "INFO");
        return;
      }
    }

    // --- Processamento dos comandos ---
    switch (comandoBase) {
      case "/confirm":
        logToSheet(`Comando /confirm detectado para transacao ID: ${complemento}`, "INFO");
        const cacheConfirm = CacheService.getScriptCache();
        const cacheKeyConfirm = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${complemento}`;
        const cachedTransactionDataConfirm = cacheConfirm.get(cacheKeyConfirm);

        if (cachedTransactionDataConfirm) {
          const transacaoData = JSON.parse(cachedTransactionDataConfirm);
          registrarTransacaoConfirmada(transacaoData, usuario, chatId);
          cacheConfirm.remove(cacheKeyConfirm);
        } else {
          enviarMensagemTelegram(chatId, "⚠️ Esta confirmação expirou ou já foi processada.");
          logToSheet(`CallbackQuery para transacao ID ${complemento} recebida, mas dados nao encontrados no cache (confirm).`, "WARN");
        }
        return;

      case "/cancel":
        logToSheet(`Comando /cancel detectado para transacao ID: ${complemento}`, "INFO");
        const cacheCancel = CacheService.getScriptCache();
        const cacheKeyCancel = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${complemento}`;
        const cachedTransactionDataCancel = cacheCancel.get(cacheKeyCancel);

        if (cachedTransactionDataCancel) {
          cancelarTransacaoPendente(chatId, complemento);
          cacheCancel.remove(cacheKeyCancel);
        } else {
          enviarMensagemTelegram(chatId, "⚠️ Este cancelamento expirou ou já foi processada.");
          logToSheet(`CallbackQuery para transacao ID ${complemento} recebida, mas dados nao encontrados no cache (cancel).`, "WARN");
        }
        return;

      // NOVO: Adicionado o case para o comando /dashboard
      case "/dashboard":
        logToSheet(`Comando /dashboard detectado.`, "INFO");
        enviarLinkDashboard(chatId);
        return;

      case "/adicionar_conta":
          logToSheet(`Comando /adicionar_conta detectado. Complemento: "${complemento}"`, "INFO");
          adicionarNovaConta(chatId, usuario, complemento);
          return;
      case "/listar_contas":
          logToSheet(`Comando /listar_contas detectado.`, "INFO");
          listarContas(chatId, usuario);
          return;
      case "/adicionar_categoria":
          logToSheet(`Comando /adicionar_categoria detectado. Complemento: "${complemento}"`, "INFO");
          adicionarNovaCategoria(chatId, usuario, complemento);
          return;
      case "/listar_categorias":
          logToSheet(`Comando /listar_categorias detectado.`, "INFO");
          listarCategorias(chatId);
          return;
      case "/listar_subcategorias":
          logToSheet(`Comando /listar_subcategorias detectado. Complemento: "${complemento}"`, "INFO");
          if (complemento) {
            listarSubcategorias(chatId, complemento);
          } else {
            enviarMensagemTelegram(chatId, "❌ Por favor, forneça o nome da categoria principal. Ex: `/listar_subcategorias Alimentação`");
          }
          return;
      case "/tutorial":
      case "/tutorial_start":
          logToSheet(`Comando /tutorial ou /tutorial_start detectado.`, "INFO");
          clearTutorialState(chatId);
          handleTutorialStep(chatId, usuario, 1);
          return;
      case "/tutorial_next":
          logToSheet(`Comando /tutorial_next detectado.`, "INFO");
          let tutorialStateNext = getTutorialState(chatId);
          if (tutorialStateNext && tutorialStateNext.currentStep > 0 && tutorialStateNext.currentStep < 6) {
            handleTutorialStep(chatId, usuario, tutorialStateNext.currentStep + 1);
          } else if (tutorialStateNext && tutorialStateNext.currentStep === 6) {
            handleTutorialStep(chatId, usuario, 6);
          } else {
            enviarMensagemTelegram(chatId, "🤔 Não há tutorial em andamento. Digite /tutorial para começar!");
            clearTutorialState(chatId);
          }
          return;
      case "/tutorial_prev":
          logToSheet(`Comando /tutorial_prev detectado.`, "INFO");
          let tutorialStatePrev = getTutorialState(chatId);
          if (tutorialStatePrev && tutorialStatePrev.currentStep > 1) {
            handleTutorialStep(chatId, usuario, tutorialStatePrev.currentStep - 1);
          } else {
            enviarMensagemTelegram(chatId, "Você já está no início do tutorial. Digite /tutorial para reiniciar.");
            clearTutorialState(chatId);
          }
          return;
      case "/tutorial_skip":
          logToSheet(`Comando /tutorial_skip detectado.`, "INFO");
          enviarMensagemTelegram(chatId, "Tutorial pulado. Se precisar de ajuda, digite /ajuda a qualquer momento.");
          clearTutorialState(chatId);
          return;
      case "/editar":
            if(normalizarTexto(complemento) === 'ultimo' || normalizarTexto(complemento) === 'último'){
                iniciarEdicaoUltimo(chatId, usuario);
            } else {
                enviarMensagemTelegram(chatId, "Comando de edição inválido. Use `/editar ultimo`.");
            }
            return;
      // NOVO: Case para processar o callback_data dos botões de edição
      case "/editar_campo_callback":
            const campoParaEditar = complemento; // O complemento já é o nome do campo (ex: 'data', 'valor')
            solicitarNovoValorParaEdicao(chatId, campoParaEditar);
            return;
      case "/cancelar_edicao":
            clearEditState(chatId);
            enviarMensagemTelegram(chatId, "Edição cancelada.");
            return;
      case "/start":
          enviarMensagemTelegram(chatId, `Olá ${escapeMarkdown(usuario)}! Bem-vindo ao seu assistente financeiro. Eu posso te ajudar a registrar gastos e receitas, ver seu saldo, extrato e mais.\n\nPara começar, tente algo como:\n- "gastei 50 no mercado com Cartao Nubank Breno"\n- "paguei 50 de uber no debito do Santander"\n- "recebi 100 de salário no Itaú via PIX"\n- "transferi 20 do Nubank para o Itaú"\n\nOu use comandos como:\n- /resumo (para ver seu resumo do mês)\n- / /saldo (para ver o saldo das suas contas)\n- /ajuda (para ver todos os comandos e exemplos)\n\nSe precisar de ajuda, digite /ajuda`);
          return;
      case "/extrato":
          logToSheet(`Comando /extrato detectado. Complemento: "${complemento}"`, "INFO");
          if (!complemento) {
            mostrarMenuExtrato(chatId);
          } else {
            enviarExtrato(chatId, usuario, complemento);
          }
          return;
      case "/extrato_tudo":
          logToSheet(`Comando /extrato_tudo detectado.`, "INFO");
          enviarExtrato(chatId, usuario, "tudo");
          return;
      case "/extrato_receitas":
          logToSheet(`Comando /extrato_receitas detectado.`, "INFO");
          enviarExtrato(chatId, usuario, "receitas");
          return;
      case "/extrato_despesas":
          logToSheet(`Comando /extrato_despesas detectado.`, "INFO");
          enviarExtrato(chatId, usuario, "despesas");
          return;
      case "/extrato_pessoa":
          logToSheet(`Comando /extrato_pessoa detectado.`, "INFO");
          mostrarMenuPorPessoa(chatId, configData); // Usa configData pré-carregado
          return;
      case "/resumo":
          const allUserNames = getAllUserNames(configData);
          const targetUser = findUserNameInText(complemento, allUserNames);
          const { month: targetMonthResumo, year: targetYearResumo } = parseMonthAndYear(complemento);

          if (targetUser) {
            logToSheet(`Comando /resumo por pessoa detectado para ${targetUser}.`, "INFO");
            enviarResumoPorPessoa(chatId, usuario, targetUser, targetMonthResumo, targetYearResumo);
          } else {
            logToSheet(`Comando /resumo geral detectado.`, "INFO");
            enviarResumo(chatId, usuario, targetMonthResumo, targetYearResumo);
          }
          return;
      case "/saldo":
          logToSheet(`Comando /saldo detectado.`, "INFO");
          enviarSaldo(chatId, usuario);
          return;
      case "/proximasfaturas":
          logToSheet(`Comando /proximasfaturas detectado.`, "INFO");
          enviarFaturasFuturas(chatId, usuario);
          return;
      case "/contasapagar":
          logToSheet(`Comando /contasapagar detectado. Mes: ${targetMonth}, Ano: ${targetYear}`, "INFO");
          enviarContasAPagar(chatId, usuario, targetMonth, targetYear);
          return;
      case "/marcar_pago":
          logToSheet(`Comando /marcar_pago detectado. ID da Conta: "${complemento}"`, "INFO");
          processarMarcarPago(chatId, textoRecebido, usuario);
          return;
      case "/excluir":
          logToSheet(`Comando /excluir detectado para ID: ${complemento}`, "INFO");
          excluirLancamentoPorId(complemento, chatId);
          return;
      case "/extrato_usuario":
          logToSheet(`Comando /extrato_usuario detectado para usuario: ${complemento}`, "INFO");
          enviarExtrato(chatId, usuario, complemento);
          return;
      case "/vincular_conta":
          logToSheet(`Comando /vincular_conta detectado. Complemento: "${complemento}"`, "INFO");
          const lastSpaceIndex = complemento.lastIndexOf(' ');
          if (lastSpaceIndex !== -1) {
            const idContaAPagar = complemento.substring(0, lastSpaceIndex).trim();
            const idTransacao = complemento.substring(lastSpaceIndex + 1).trim();
            if (idContaAPagar && idTransacao) {
              vincularTransacaoAContaAPagar(chatId, idContaAPagar, idTransacao);
            } else {
              enviarMensagemTelegram(chatId, "❌ Formato invalido para vincular. Use: `/vincular_conta <ID_CONTA_A_PAGAR> <ID_TRANSACAO>`");
            }
          } else {
            enviarMensagemTelegram(chatId, "❌ Formato invalido para vincular. Use: `/vincular_conta <ID_CONTA_A_PAGAR> <ID_TRANSACAO>`");
          }
          return;
      case "/ajuda":
          logToSheet(`Comando /ajuda detectado.`, "INFO");
          enviarAjuda(chatId);
          return;
      case "/metas":
          logToSheet(`Comando /metas detectado. Mes: ${targetMonth}, Ano: ${targetYear}`, "INFO");
          enviarMetas(chatId, usuario, targetMonth, targetYear);
          return;

      default:
        const palavrasConsulta = ["quanto", "qual", "quais", "listar", "mostrar", "total"];
        const primeiraPalavraConsulta = textoRecebido.toLowerCase().split(' ')[0];

        if (palavrasConsulta.includes(primeiraPalavraConsulta)) {
            logToSheet(`Consulta em linguagem natural detectada: "${textoRecebido}".`, "INFO");
            processarConsultaLinguagemNatural(chatId, usuario, textoRecebido);
            return;
        }

        logToSheet(`Comando '${comandoBase}' não reconhecido como comando direto. Tentando interpretar como lançamento.`, "INFO");
        const resultadoLancamento = interpretarMensagemTelegram(textoRecebido, usuario, chatId);

        if (resultadoLancamento && resultadoLancamento.handled) {
          logToSheet("Mensagem ja tratada e resposta enviada por funcao interna. Nenhuma acao adicional necessaria.", "INFO");
        } else if (resultadoLancamento && resultadoLancamento.message) {
          logToSheet(`Mensagem interpretada com sucesso: ${resultadoLancamento.message}`, "INFO");
        } else if (resultadoLancamento && resultadoLancamento.errorMessage) {
          logToSheet(`Erro na interpretação da mensagem: ${resultadoLancamento.errorMessage}`, "WARN");
        } else if (resultadoLancamento && resultadoLancamento.status === 'PENDING_CONFIRMATION') {
          logToSheet(`Confirmacao de transacao pendente para ID: ${resultadoLancamento.transactionId}`, "INFO");
        } else {
          enviarMensagemTelegram(chatId, "❌ Não entendi seu comando ou lançamento. Digite /ajuda para ver o que posso fazer.");
        }
        return;
    }
  } catch (err) {
    const chatIdForError = e?.postData?.contents ? JSON.parse(e.postData.contents)?.message?.chat?.id || JSON.parse(e.postData.contents)?.callback_query?.message?.chat?.id : null;
    logToSheet(`ERRO FATAL E INESPERADO EM doPost: ${err.message}. Stack: ${err.stack}`, "ERROR");
    if (chatIdForError) {
        enviarMensagemTelegram(chatIdForError, "❌ Ocorreu um erro crítico no sistema. O administrador foi notificado. Por favor, tente novamente mais tarde.");
    }
  }
}

/**
 * Cria o menu personalizado quando a planilha é aberta.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Gasto Certo')
      .addItem('Configuração Inicial', 'showSetupUI')
      .addSeparator()
      .addItem('Gerar Contas Recorrentes (Próximo Mês)', 'triggerGenerateRecurringBills')
      .addToUi();
}

/**
 * Mostra a UI de configuração (caixa de diálogo).
 */
function showSetupUI() {
  const html = HtmlService.createHtmlOutputFromFile('SetupDialog.html')
      .setWidth(400)
      .setHeight(500); // Aumentei a altura para caber o novo campo
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuração Inicial do Gasto Certo');
}

/**
 * Salva as credenciais do Telegram e a URL do Web App nas Propriedades do Script
 * e configura o webhook do bot. Chamada pelo SetupDialog.html.
 * @param {string} token O token do bot do Telegram.
 * @param {string} chatId O ID do chat principal.
 * @param {string} webAppUrl A URL do Web App para o webhook.
 * @returns {Object} Um objeto indicando o sucesso ou falha da operação.
 */
function saveCredentialsAndSetupWebhook(token, chatId, webAppUrl) {
  try {
    if (!token || !chatId || !webAppUrl) {
      throw new Error("O Token, o ID do Chat e a URL do App da Web são obrigatórios.");
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(TELEGRAM_TOKEN_PROPERTY_KEY, token);
    scriptProperties.setProperty(ADMIN_CHAT_ID_PROPERTY_KEY, chatId);
    scriptProperties.setProperty(WEB_APP_URL_PROPERTY_KEY, webAppUrl); // Salva a URL
    
    logToSheet("Configurações de Token, Chat ID e URL salvas com sucesso.", "INFO");

    const webhookResult = setupWebhook();

    if (webhookResult && webhookResult.ok) {
        initializeSheets(); // Inicializa as abas após a configuração bem-sucedida
        logToSheet("Configuração e inicialização concluídas com sucesso.", "INFO");
        return { success: true, message: "Credenciais salvas e bot configurado com sucesso!" };
    } else {
        const errorDescription = webhookResult ? webhookResult.description : "Resposta inválida da API do Telegram.";
        throw new Error(`Falha ao configurar o webhook: ${errorDescription}`);
    }

  } catch (e) {
    logToSheet(`Erro durante a configuração: ${e.message}`, "ERROR");
    return { success: false, message: e.message };
  }
}


/**
 * Função para configurar o webhook do Telegram.
 * Agora lê a URL do Web App diretamente das Propriedades do Script, que é mais confiável.
 * @returns {Object} Um objeto com o resultado da API do Telegram.
 */
function setupWebhook() {
  try {
    const token = getTelegramBotToken();
    // A URL é lida das propriedades, onde foi salva pela caixa de diálogo.
    const webhookUrl = PropertiesService.getScriptProperties().getProperty(WEB_APP_URL_PROPERTY_KEY);

    if (!webhookUrl) {
      const errorMessage = "URL do Web App não encontrada nas Propriedades do Script. Execute a 'Configuração Inicial' e forneça a URL correta.";
      throw new Error(errorMessage);
    }

    const url = `https://api.telegram.org/bot${token}/setWebhook?url=${webhookUrl}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseText = response.getContentText();
    logToSheet(`Resposta da configuração do webhook: ${responseText}`, "INFO");
    return JSON.parse(responseText);

  } catch (e) {
    logToSheet(`Erro ao configurar o webhook: ${e.message}`, "ERROR");
    return { ok: false, description: e.message };
  }
}


/**
 * NOVO: Adiciona ou atualiza a configuração do usuário administrador na aba 'Configuracoes'.
 * @param {string} adminChatId O Chat ID do administrador.
 */
function updateAdminConfig(adminChatId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_CONFIGURACOES);
    const data = configSheet.getDataRange().getValues();
    let adminRowFound = false;

    // Procura por uma linha de admin existente para atualizar
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'chatId') {
            configSheet.getRange(i + 1, 2).setValue(adminChatId);
            configSheet.getRange(i + 1, 3).setValue('Admin'); // Define um nome padrão
            adminRowFound = true;
            break;
        }
    }
    
    // Se não encontrou, adiciona uma nova linha
    if (!adminRowFound) {
        configSheet.appendRow(['chatId', adminChatId, 'Admin', 'Default']);
    }
}


/**
 * Adiciona um novo usuário ao sistema.
 * @param {string} chatId O ID do chat do novo usuário.
 * @param {string} userName O nome do usuário.
 */
function addNewUser(chatId, userName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEETS.USERS);
  if (sheet) {
    // Verifica se o usuário já existe
    const existingUser = findRowByValue(SHEETS.USERS, 1, chatId);
    if (!existingUser) {
      sheet.appendRow([chatId, userName, new Date()]);
      Logger.log(`Novo usuário adicionado: ${userName} (${chatId})`);
    }
  }
}

/**
 * Inicializa todas as abas necessárias da planilha com base no objeto HEADERS.
 * Garante que o ambiente do usuário seja criado corretamente.
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Itera sobre o objeto HEADERS para criar cada aba com seus respectivos cabeçalhos.
  for (const sheetName in HEADERS) {
    if (Object.prototype.hasOwnProperty.call(HEADERS, sheetName)) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        const headers = HEADERS[sheetName];
        if (headers && headers.length > 0) {
          sheet.appendRow(headers);
          logToSheet(`Aba '${sheetName}' criada com sucesso.`, "INFO");
        }
      }
    }
  }
  
  // Garante que a aba de logs também seja criada.
  if (!ss.getSheetByName(SHEET_LOGS_SISTEMA)) {
      const logSheet = ss.insertSheet(SHEET_LOGS_SISTEMA);
      logSheet.appendRow(["timestamp", "level", "message"]);
      logToSheet(`Aba de sistema '${SHEET_LOGS_SISTEMA}' criada com sucesso.`, "INFO");
  }
}
