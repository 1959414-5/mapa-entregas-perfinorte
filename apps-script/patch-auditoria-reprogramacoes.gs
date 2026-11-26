/**
 * PATCH - AUDITORIA DE PRAZOS E REPROGRAMACOES
 *
 * Aplicacao:
 * 1) No objeto MAPA_CFG, atualize ULTIMA_COLUNA para 173 (FQ).
 * 2) No doPost(e), adicione ao switch:
 *      case "concluirEntrega":
 *        return jsonOutput_(concluirEntrega(body.linha, body.motivo, body.observacao));
 * 3) No onEdit(e) existente, logo depois de confirmar que a aba e HIST e row >= 2,
 *    adicione:
 *      if (col === MAPA_CFG.COL_PREVISAO) {
 *        registrarEdicaoPrevisaoManual_(e);
 *        return;
 *      }
 *      if (col === MAPA_CFG.COL_STATUS &&
 *          String(e.range.getValue() || "").trim().toUpperCase() === "ENTREGUE") {
 *        registrarConclusaoManual_(sheet, row);
 *        return;
 *      }
 * 4) Cole TODO este arquivo no final do projeto Apps Script.
 * 5) Rode inicializarAuditoriaEntregas() uma vez pelo editor, autorize e confira
 *    a aba Historico_Reprogramacoes. Depois publique uma NOVA VERSAO da implantacao.
 *
 * Colunas novas reservadas:
 * FM (169) data/hora real da conclusao
 * FN (170) motivo do atraso
 * FO (171) observacao do atraso
 * FP (172) previsao original (imutavel)
 * FQ (173) identificador interno estavel da entrega
 *
 * A auditoria historica fica na aba Historico_Reprogramacoes.
 * Em registros anteriores a implantacao, a previsao original sera inicializada
 * com a data que estiver cadastrada no momento da migracao. Reprogramacoes
 * anteriores nao podem ser reconstruidas porque nao existe historico delas.
 */

const AUDITORIA_ENTREGAS_CFG_ = {
  COL_CONCLUSAO: 169,
  COL_MOTIVO: 170,
  COL_OBSERVACAO: 171,
  COL_PREVISAO_ORIGINAL: 172,
  COL_ID: 173,
  ABA_REPROGRAMACOES: "Historico_Reprogramacoes"
};

const MOTIVOS_ATRASO_ENTREGA_ = [
  "Atraso na produção",
  "Falta de matéria-prima ou estoque",
  "Atraso de fornecedor",
  "Reprogramação da produção",
  "Problema de transporte/logística",
  "Indisponibilidade do cliente",
  "Alteração solicitada pelo cliente",
  "Informações pendentes do pedido",
  "Problema operacional",
  "Condições externas",
  "Erro no planejamento do prazo",
  "Outros"
];

const EXEMPLOS_MOTIVOS_ATRASO_ = {
  "Atraso na produção": "Pedido não ficou pronto dentro do prazo.",
  "Falta de matéria-prima ou estoque": "Material indisponível para produção.",
  "Atraso de fornecedor": "Material recebido fora do prazo.",
  "Reprogramação da produção": "Alteração de prioridades ou capacidade produtiva.",
  "Problema de transporte/logística": "Indisponibilidade de veículo, rota ou atraso no carregamento.",
  "Indisponibilidade do cliente": "Cliente não estava disponível para receber.",
  "Alteração solicitada pelo cliente": "Cliente solicitou mudança na data de entrega.",
  "Informações pendentes do pedido": "Especificações, dados ou documentação incompletos.",
  "Problema operacional": "Erro, retrabalho ou falha no processo.",
  "Condições externas": "Clima, interdição ou imprevistos externos.",
  "Erro no planejamento do prazo": "Prazo estimado inadequadamente.",
  "Outros": "Situação não contemplada nas opções anteriores."
};

function getOrCreateHistoricoReprogramacoes_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(AUDITORIA_ENTREGAS_CFG_.ABA_REPROGRAMACOES);
  if (!sh) {
    sh = ss.insertSheet(AUDITORIA_ENTREGAS_CFG_.ABA_REPROGRAMACOES);
    sh.getRange(1, 1, 1, 8).setValues([[
      "ID_ENTREGA",
      "CODIGO_CLIENTE",
      "CLIENTE",
      "PREVISAO_ANTERIOR",
      "NOVA_PREVISAO",
      "ALTERADO_EM",
      "MOTIVO",
      "OBSERVACAO"
    ]]);
    sh.setFrozenRows(1);
    sh.getRange("D:E").setNumberFormat("dd/mm/yyyy");
    sh.getRange("F:F").setNumberFormat("dd/mm/yyyy hh:mm:ss");
  }
  return sh;
}

function normalizarDataAuditoria_(valor) {
  if (valor === null || valor === undefined || valor === "") return "";
  const tz = Session.getScriptTimeZone();

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, tz, "yyyy-MM-dd");
  }

  const texto = String(valor).trim();
  if (!texto) return "";

  let m = texto.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];

  m = texto.match(/^(\d{2})[\/-](\d{2})[\/-](\d{4})/);
  if (m) return m[3] + "-" + m[2] + "-" + m[1];

  const data = new Date(texto);
  if (!isNaN(data.getTime())) return Utilities.formatDate(data, tz, "yyyy-MM-dd");
  return "";
}

function dataAuditoriaParaDate_(iso) {
  if (!iso) return "";
  const p = String(iso).split("-");
  if (p.length !== 3) return "";
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
}

function dataHoraAuditoriaIso_(valor) {
  if (!(valor instanceof Date) || isNaN(valor.getTime())) return "";
  return Utilities.formatDate(
    valor,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}

function obterIdEstavelEntrega_(hist, linha) {
  const cfg = AUDITORIA_ENTREGAS_CFG_;
  let id = String(hist.getRange(linha, cfg.COL_ID).getValue() || "").trim();
  if (!id) {
    id = Utilities.getUuid();
    hist.getRange(linha, cfg.COL_ID).setValue(id);
  }
  return id;
}

function inicializarAuditoriaEntregas() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) throw new Error("Aba HISTORICO_ENTREGAS não encontrada.");

  const lastRow = encontrarUltimaLinhaComDados_(hist);
  if (lastRow < 2) {
    getOrCreateHistoricoReprogramacoes_();
    return { ok: true, inicializadas: 0 };
  }

  const total = lastRow - 1;
  const clientes = hist.getRange(2, MAPA_CFG.COL_FANTASIA, total, 1).getValues();
  const dados = hist.getRange(2, MAPA_CFG.COL_PREVISAO, total,
    AUDITORIA_ENTREGAS_CFG_.COL_ID - MAPA_CFG.COL_PREVISAO + 1).getValues();

  let alterou = false;
  const metadados = dados.map(function(r, i) {
    const temEntrega = String(clientes[i][0] || "").trim() !== "";
    const original = r[AUDITORIA_ENTREGAS_CFG_.COL_PREVISAO_ORIGINAL - MAPA_CFG.COL_PREVISAO];
    const idAtual = r[AUDITORIA_ENTREGAS_CFG_.COL_ID - MAPA_CFG.COL_PREVISAO];

    let previsaoOriginal = original;
    let id = idAtual;

    if (temEntrega && !previsaoOriginal && r[0]) {
      previsaoOriginal = r[0];
      alterou = true;
    }
    if (temEntrega && !id) {
      id = Utilities.getUuid();
      alterou = true;
    }

    return [
      r[AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO - MAPA_CFG.COL_PREVISAO],
      r[AUDITORIA_ENTREGAS_CFG_.COL_MOTIVO - MAPA_CFG.COL_PREVISAO],
      r[AUDITORIA_ENTREGAS_CFG_.COL_OBSERVACAO - MAPA_CFG.COL_PREVISAO],
      previsaoOriginal,
      id
    ];
  });

  if (alterou) {
    hist.getRange(2, AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO, total, 5).setValues(metadados);
    SpreadsheetApp.flush();
  }

  getOrCreateHistoricoReprogramacoes_();
  return { ok: true, inicializadas: metadados.filter(r => r[4]).length };
}

function garantirAuditoriaLinha_(hist, linha) {
  const cfg = AUDITORIA_ENTREGAS_CFG_;
  const atual = hist.getRange(linha, cfg.COL_PREVISAO_ORIGINAL, 1, 2).getValues()[0];
  let original = atual[0];
  let id = atual[1];

  const previsaoAtual = hist.getRange(linha, MAPA_CFG.COL_PREVISAO).getValue();
  if (!original && previsaoAtual) {
    original = previsaoAtual;
    hist.getRange(linha, cfg.COL_PREVISAO_ORIGINAL).setValue(original);
  }
  if (!id) {
    id = Utilities.getUuid();
    hist.getRange(linha, cfg.COL_ID).setValue(id);
  }
  return { previsaoOriginal: original, id: id };
}

function contarReprogramacoesPorId_() {
  const sh = getOrCreateHistoricoReprogramacoes_();
  const lastRow = sh.getLastRow();
  const mapa = {};
  if (lastRow < 2) return mapa;

  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  ids.forEach(function(id) {
    const chave = String(id || "").trim();
    if (chave) mapa[chave] = (mapa[chave] || 0) + 1;
  });
  return mapa;
}

function registrarReprogramacao_(hist, linha, anterior, nova, motivo, observacao) {
  const antigoIso = normalizarDataAuditoria_(anterior);
  const novoIso = normalizarDataAuditoria_(nova);
  if (!antigoIso || !novoIso || antigoIso === novoIso) return 0;

  const meta = garantirAuditoriaLinha_(hist, linha);
  const sh = getOrCreateHistoricoReprogramacoes_();
  const codigo = hist.getRange(linha, CFG.COL_ID_HIST).getValue();
  const cliente = hist.getRange(linha, MAPA_CFG.COL_FANTASIA).getValue();

  sh.appendRow([
    meta.id,
    codigo || "",
    cliente || "",
    dataAuditoriaParaDate_(antigoIso),
    dataAuditoriaParaDate_(novoIso),
    new Date(),
    motivo || "Alteração de previsão",
    observacao || ""
  ]);
  return 1;
}

function entregaEstaAtrasadaPelaPrevisaoOriginal_(hist, linha) {
  const meta = garantirAuditoriaLinha_(hist, linha);
  const previsaoIso = normalizarDataAuditoria_(meta.previsaoOriginal);
  if (!previsaoIso) {
    return { atrasada: false, previsaoOriginal: "" };
  }

  // Não foram definidos horários-limite diferentes para MANHÃ e TARDE.
  // Até essa definição, a regra é por dia: atraso começa no dia seguinte.
  const hojeIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return {
    atrasada: hojeIso > previsaoIso,
    previsaoOriginal: previsaoIso
  };
}

function concluirEntrega(linha, motivo, observacao) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida. Atualize a página e tente novamente." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const verificacao = entregaEstaAtrasadaPelaPrevisaoOriginal_(hist, linhaNum);
    const motivoTexto = String(motivo || "").trim();
    const observacaoTexto = String(observacao || "").trim();

    if (verificacao.atrasada) {
      if (MOTIVOS_ATRASO_ENTREGA_.indexOf(motivoTexto) === -1) {
        return { ok: false, erro: "Selecione um dos motivos disponíveis." };
      }
      if (motivoTexto === "Outros" && !observacaoTexto) {
        return { ok: false, erro: "Descreva o motivo quando selecionar Outros." };
      }
    }

    const agora = new Date();
    hist.getRange(linhaNum, MAPA_CFG.COL_STATUS).setValue("ENTREGUE");
    hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO).setValue(agora);
    hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_MOTIVO).setValue(
      verificacao.atrasada ? motivoTexto : ""
    );
    hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_OBSERVACAO).setValue(
      verificacao.atrasada ? observacaoTexto : ""
    );
    SpreadsheetApp.flush();

    return {
      ok: true,
      dataHoraConclusao: dataHoraAuditoriaIso_(agora),
      atrasada: verificacao.atrasada,
      motivoAtraso: verificacao.atrasada ? motivoTexto : ""
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Substitui atualizarCampoEntrega() existente.
 * Registros de data são feitos sob lock e nunca sobrescrevem a previsão original.
 */
function atualizarCampoEntrega(linha, campo, valor) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida (a planilha pode ter mudado — atualize a página)." };
  }

  const colunasPorCampo = {
    motorista: MAPA_CFG.COL_MOTORISTA,
    status: MAPA_CFG.COL_STATUS,
    periodo: MAPA_CFG.COL_PERIODO,
    endereco: MAPA_CFG.COL_ENDERECO,
    observacoes: MAPA_CFG.COL_OBS,
    data: MAPA_CFG.COL_PREVISAO,
    representante: 15,
    ok: 165
  };
  const coluna = colunasPorCampo[campo];
  if (!coluna) return { ok: false, erro: "Campo não editável: " + campo };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (campo === "status" && String(valor || "").trim().toUpperCase() === "ENTREGUE") {
      const verificacao = entregaEstaAtrasadaPelaPrevisaoOriginal_(hist, linhaNum);
      if (verificacao.atrasada) {
        return {
          ok: true,
          requerJustificativa: true,
          atrasada: true,
          previsaoOriginal: verificacao.previsaoOriginal,
          cliente: hist.getRange(linhaNum, MAPA_CFG.COL_FANTASIA).getValue(),
          codigo: hist.getRange(linhaNum, CFG.COL_ID_HIST).getValue()
        };
      }

      const agora = new Date();
      hist.getRange(linhaNum, MAPA_CFG.COL_STATUS).setValue("ENTREGUE");
      hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO).setValue(agora);
      hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_MOTIVO, 1, 2).clearContent();
      SpreadsheetApp.flush();
      return { ok: true, dataHoraConclusao: dataHoraAuditoriaIso_(agora), atrasada: false };
    }

    let valorFinal = valor;
    if (campo === "data") {
      const atual = hist.getRange(linhaNum, MAPA_CFG.COL_PREVISAO).getValue();
      const novaIso = normalizarDataAuditoria_(valor);
      const atualIso = normalizarDataAuditoria_(atual);
      const meta = garantirAuditoriaLinha_(hist, linhaNum);

      // Se o registro ainda não tinha uma previsão original, a data atual
      // anterior à edição passa a ser a linha de base.
      if (!meta.previsaoOriginal) {
        const base = atual || (novaIso ? dataAuditoriaParaDate_(novaIso) : "");
        if (base) hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_PREVISAO_ORIGINAL).setValue(base);
      }

      if (atualIso && novaIso && atualIso !== novaIso) {
        registrarReprogramacao_(hist, linhaNum, atual, novaIso, "Alteração de previsão", "");
      }
      valorFinal = novaIso ? dataAuditoriaParaDate_(novaIso) : "";
    }

    if (campo === "ok") valorFinal = (valor === true || valor === "true");

    hist.getRange(linhaNum, coluna).setValue(valorFinal);
    SpreadsheetApp.flush();

    if (campo === "endereco") {
      const texto = String(valor || "").trim();
      if (enderecoNoPadrao_(texto)) {
        geocodificarEndereco(texto);
        const codigoCliente = hist.getRange(linhaNum, CFG.COL_ID_HIST).getValue();
        const fantasiaCliente = hist.getRange(linhaNum, MAPA_CFG.COL_FANTASIA).getValue();
        salvarEnderecoRotaCliente_(fantasiaCliente, texto);
        salvarEnderecoNoHistorico_(codigoCliente, texto);
      }
    }

    const metaAtual = campo === "data" ? garantirAuditoriaLinha_(hist, linhaNum) : null;
    const totalReprogramacoes = metaAtual ? (contarReprogramacoesPorId_()[metaAtual.id] || 0) : undefined;
    return {
      ok: true,
      qtdReprogramacoes: totalReprogramacoes
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Função chamada pelo onEdit(e) para datas alteradas diretamente na planilha.
 * Adicione a chamada indicada no cabeçalho deste patch ao onEdit existente.
 */
function registrarEdicaoPrevisaoManual_(e) {
  if (!e || !e.range) return;
  const hist = e.range.getSheet();
  const linha = e.range.getRow();
  if (linha < 2 || hist.getName() !== CFG.HIST) return;

  const novoValor = e.range.getValue();
  const valorAnterior = e.oldValue || "";
  const anteriorIso = normalizarDataAuditoria_(valorAnterior);
  const novoIso = normalizarDataAuditoria_(novoValor);

  const meta = garantirAuditoriaLinha_(hist, linha);
  if (!meta.previsaoOriginal) {
    const base = anteriorIso || novoIso;
    if (base) hist.getRange(linha, AUDITORIA_ENTREGAS_CFG_.COL_PREVISAO_ORIGINAL).setValue(dataAuditoriaParaDate_(base));
  }

  if (anteriorIso && novoIso && anteriorIso !== novoIso) {
    registrarReprogramacao_(hist, linha, valorAnterior, novoValor, "Alteração manual na planilha", "");
  }
}

/**
 * Registra a data/hora em alterações manuais de status na planilha.
 * A validação com modal é aplicada no site; edições diretas na planilha
 * recebem o timestamp, mas não abrem modal.
 */
function registrarConclusaoManual_(hist, linha) {
  const status = String(hist.getRange(linha, MAPA_CFG.COL_STATUS).getValue() || "").trim().toUpperCase();
  if (status !== "ENTREGUE") return;
  const celulaConclusao = hist.getRange(linha, AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO);
  if (!celulaConclusao.getValue()) {
    garantirAuditoriaLinha_(hist, linha);
    celulaConclusao.setValue(new Date());
    SpreadsheetApp.flush();
  }
}

/**
 * Substitui obterEntregasValidas_() existente para expor os campos de auditoria.
 */
function obterEntregasValidas_(filtroMotorista, filtroData) {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return [];

  const lastRow = encontrarUltimaLinhaComDados_(hist);
  if (lastRow < 2) return [];

  // Migra linhas antigas sem sobrescrever campos já gravados.
  inicializarAuditoriaEntregas();

  const fm = filtroMotorista ? String(filtroMotorista).trim().toUpperCase() : null;
  const fd = filtroData ? String(filtroData).trim() : null;

  const blocoInicio = hist.getRange(2, CFG.COL_ID_HIST, lastRow - 1, 7).getValues();
  const largura = AUDITORIA_ENTREGAS_CFG_.COL_ID - MAPA_CFG.COL_ENDERECO + 1;
  const blocoCol = hist.getRange(2, MAPA_CFG.COL_ENDERECO, lastRow - 1, largura).getValues();
  const cache = lerCacheGeo_();
  const contagemReprogramacoes = contarReprogramacoesPorId_();

  const entregas = [];
  const OFF_FANTASIA_INICIO = MAPA_CFG.COL_FANTASIA - CFG.COL_ID_HIST;
  const OFF_REPR_INICIO = 15 - CFG.COL_ID_HIST;

  const OFF_MOTORISTA = MAPA_CFG.COL_MOTORISTA - MAPA_CFG.COL_ENDERECO;
  const OFF_PREVISAO = MAPA_CFG.COL_PREVISAO - MAPA_CFG.COL_ENDERECO;
  const OFF_ENTRADA = MAPA_CFG.COL_ENTRADA - MAPA_CFG.COL_ENDERECO;
  const OFF_STATUS = MAPA_CFG.COL_STATUS - MAPA_CFG.COL_ENDERECO;
  const OFF_PERIODO = MAPA_CFG.COL_PERIODO - MAPA_CFG.COL_ENDERECO;
  const OFF_OBS = MAPA_CFG.COL_OBS - MAPA_CFG.COL_ENDERECO;
  const OFF_OK = 165 - MAPA_CFG.COL_ENDERECO;

  const OFF_CONCLUSAO = AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO - MAPA_CFG.COL_ENDERECO;
  const OFF_MOTIVO = AUDITORIA_ENTREGAS_CFG_.COL_MOTIVO - MAPA_CFG.COL_ENDERECO;
  const OFF_OBSERVACAO = AUDITORIA_ENTREGAS_CFG_.COL_OBSERVACAO - MAPA_CFG.COL_ENDERECO;
  const OFF_ORIGINAL = AUDITORIA_ENTREGAS_CFG_.COL_PREVISAO_ORIGINAL - MAPA_CFG.COL_ENDERECO;
  const OFF_ID = AUDITORIA_ENTREGAS_CFG_.COL_ID - MAPA_CFG.COL_ENDERECO;

  for (let i = 0; i < blocoCol.length; i++) {
    const codigo = blocoInicio[i][0];
    const fantasia = blocoInicio[i][OFF_FANTASIA_INICIO];
    const representante = blocoInicio[i][OFF_REPR_INICIO];
    const dadosLinha = blocoCol[i];

    const entrada = dadosLinha[OFF_ENTRADA];
    const endereco = dadosLinha[0];
    const motorista = dadosLinha[OFF_MOTORISTA];
    const previsao = dadosLinha[OFF_PREVISAO];
    const status = dadosLinha[OFF_STATUS];
    const periodo = dadosLinha[OFF_PERIODO];
    const obs = dadosLinha[OFF_OBS];
    const ok = dadosLinha[OFF_OK];

    if (!fantasia) continue;

    const textoEndereco = String(endereco || "").trim();
    const ehAviso = textoEndereco.toUpperCase() === "AVISO";
    if (fm && String(motorista || "").trim().toUpperCase() !== fm) continue;

    const dataIso = _toIsoDateString(previsao) || "";
    if (dataIso && dataIso < MAPA_CFG.DATA_CORTE) continue;
    if (fd && dataIso !== fd) continue;

    const coords = (!ehAviso && textoEndereco) ? cache[textoEndereco.toUpperCase()] : null;
    const idEntrega = String(dadosLinha[OFF_ID] || "").trim();

    entregas.push({
      dataEntrada: _toIsoDateOnly_(entrada),
      dataEntradaHora: dataHoraAuditoriaIso_(entrada),
      dataHoraConclusao: dataHoraAuditoriaIso_(dadosLinha[OFF_CONCLUSAO]),
      motivoAtraso: dadosLinha[OFF_MOTIVO] || "",
      observacaoAtraso: dadosLinha[OFF_OBSERVACAO] || "",
      previsaoOriginal: _toIsoDateString(dadosLinha[OFF_ORIGINAL]) || "",
      qtdReprogramacoes: idEntrega ? (contagemReprogramacoes[idEntrega] || 0) : 0,
      linha: i + 2,
      codigo: codigo || "",
      cliente: fantasia,
      representante: representante || "",
      motorista: motorista || "",
      data: dataIso,
      periodo: periodo || "",
      status: status || "",
      ok: ok === true,
      cidade: (!ehAviso && textoEndereco) ? extrairCidadeUf_(textoEndereco) : "",
      endereco: ehAviso ? "" : textoEndereco,
      observacoes: obs || "",
      lat: coords ? coords.lat : null,
      lng: coords ? coords.lng : null
    });
  }

  return entregas;
}
