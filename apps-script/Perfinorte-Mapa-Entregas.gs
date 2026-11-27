// =====================================================
// TRIGGER AUTOMÁTICO — PROCESSAMENTO EM BATCH A CADA 30s
// =====================================================

/*************************************
 * CONFIGURAÇÕES FIXAS
 *************************************/
const CFG = {
  HIST: "HISTORICO_ENTREGAS",
  ENTREGAS: "ENTREGAS",
  PEDIDOS: "PEDIDOS",
  ROTAS: "ROTAS_CLIENTES",
  MOTORISTAS: "MOTORISTAS_PEDIDO",

  COL_ID_HIST: 9,          // I
  COL_STATUS_HIST: 164,    // FH
  COL_DATA_HIST: 162       // FF
};

/*************************************
 * ONEDIT — MARCA EM PROCESSANDO (SEM LOOP)
 *************************************/
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  geocodificarAoEditarEndereco_(e);

  if (sheet.getName() !== CFG.HIST || row < 2) return;

  if (col === MAPA_CFG.COL_PREVISAO) {
  registrarEdicaoPrevisaoManual_(e);
  return;
}
if (col === MAPA_CFG.COL_STATUS &&
    String(e.range.getValue() || "").trim().toUpperCase() === "ENTREGUE") {
  registrarConclusaoManual_(sheet, row);
  return;
}

  // 🔸 Se editou a coluna ID (I)
  if (col === CFG.COL_ID_HIST) {
    const id = String(e.range.getValue() || "").trim();

    // Se apagou o ID → limpa tudo
    if (!id) {
      sheet.getRange(row, 11).clearContent();  // K
      sheet.getRange(row, 15).clearContent();  // O
      sheet.getRange(row, 157).clearContent(); // FA
      sheet.getRange(row, 160).clearContent(); // FD
      sheet.getRange(row, CFG.COL_STATUS_HIST).clearContent();
      sheet.getRange(row, CFG.COL_DATA_HIST).clearContent();
      return;
    }

    // 🔹 Só marca PROCESSANDO se ainda não estiver
    const statusAtual = String(sheet.getRange(row, CFG.COL_STATUS_HIST).getValue() || "");
    if (!statusAtual.includes("PROCESSANDO")) {
      sheet.getRange(row, CFG.COL_STATUS_HIST).setValue("⏳ PROCESSANDO");
      sheet.getRange(row, CFG.COL_DATA_HIST).setValue(new Date());
    }

    // Chama a função, mas agora sem risco de loop
    if (typeof registrarDataHoraAoDigitarID === "function") {
      registrarDataHoraAoDigitarID(e);
    }
    return;
  }

  // 🔸 Se editou manualmente o STATUS
  if (col === CFG.COL_STATUS_HIST) {
    const status = String(e.range.getValue() || "").toUpperCase();

    // Se mudou de "❌ NÃO ENCONTRADO" para "EM FILA", NÃO LIMPA NADA
    if (status === "EM FILA") {
      return; // mantém dados manuais
    }
  }
}







/*************************************
 * PROCESSA FILA AUTOMATICAMENTE
 *************************************/
function processarFilaHistoricoEmBatch() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return;

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return;

  const ids = hist.getRange(2, CFG.COL_ID_HIST, lastRow - 1, 1).getValues().flat().map(v => String(v).trim());
  const status = hist.getRange(2, CFG.COL_STATUS_HIST, lastRow - 1, 1).getValues().flat();

  const filas = [];
  ids.forEach((id, i) => {
    const stat = String(status[i] || "").toUpperCase();
    if (id && stat.includes("PROCESSANDO")) {
      filas.push({ id, row: 2 + i });
    }
  });

  if (!filas.length) return;

  const ent = ss.getSheetByName(CFG.ENTREGAS);
  const ped = ss.getSheetByName(CFG.PEDIDOS);
  const rot = ss.getSheetByName(CFG.ROTAS);

  const dadosEnt = ent && ent.getLastRow() >= 4
    ? ent.getRange(4, 1, ent.getLastRow() - 3, ent.getLastColumn()).getValues()
    : [];

  const dadosPed = ped && ped.getLastRow() >= 2
    ? ped.getRange(2, 1, ped.getLastRow() - 1, ped.getLastColumn()).getValues()
    : [];

  const mapaRotas = {};
  if (rot && rot.getLastRow() >= 2) {
    const dadosRot = rot.getRange(2, 2, rot.getLastRow() - 1, 3).getValues();
    dadosRot.forEach(r => {
      const fantasia = String(r[0] || "").trim().toUpperCase();
      const rota = r[2];
      if (fantasia) mapaRotas[fantasia] = rota;
    });
  }

  const cad = ss.getSheetByName("CADASTRO_CLIENTES");

const dadosCad = cad && cad.getLastRow() >= 2
  ? cad.getRange(2, 1, cad.getLastRow() - 1, 3).getValues()
  : [];

processarLinhasHistorico(
  filas,
  ss,
  hist,
  dadosEnt,
  dadosPed,
  mapaRotas,
  dadosCad
);

}


/*************************************
 * PROCESSAMENTO EM LOOP INTERNO (30s)
 *************************************/
function processarFilaHistoricoIntervalo() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return;

  const start = new Date().getTime();
  const DURACAO = 60000;
  const INTERVALO = 30000;

  while (new Date().getTime() - start < DURACAO) {
    const lastRow = hist.getLastRow();
    if (lastRow < 2) break;

    const ids = hist.getRange(2, CFG.COL_ID_HIST, lastRow - 1, 1).getValues().flat().map(v => String(v).trim());
    const status = hist.getRange(2, CFG.COL_STATUS_HIST, lastRow - 1, 1).getValues().flat();

    const filas = [];
    ids.forEach((id, i) => {
      const stat = String(status[i] || "").toUpperCase();
      if (id && stat.includes("PROCESSANDO")) {
        filas.push({ id, row: 2 + i });
      }
    });

    if (!filas.length) break;

processarLinhasHistorico(
  filas,
  ss,
  hist,
  dadosEnt,
  dadosPed,
  mapaRotas,
  dadosCad
);

    Utilities.sleep(INTERVALO);
  }
}


/*************************************
 * PROCESSA LINHAS DO HISTÓRICO
 *************************************/
function processarLinhasHistorico(
  filas,
  ss,
  hist,
  dadosEnt,
  dadosPed,
  mapaRotas,
  dadosCad
)

 {

  const updatesK = [];
  const updatesO = [];
  const updatesFA = [];
  const updatesFD = [];
  const updatesStatus = [];
  const linhas = [];

  const mapaEntregas = {};
  dadosEnt.forEach(r => {
    const id = String(r[8] || "").trim();
    if (id) mapaEntregas[id] = r;
  });

  const mapaPedidos = {};
  dadosPed.forEach(r => {
    const id = String(r[8] || "").trim();
    if (id) mapaPedidos[id] = r;
  });

const mapaCadastro = {};

dadosCad.forEach(r => {
  const id = String(r[0] || "").trim(); // coluna A
  if (id) {
    mapaCadastro[id] = r[2]; // coluna C (Nome)
  }
});

  filas.forEach(f => {

    const id = f.id;
    const row = f.row;

    let fantasia = "";
    let vendedor = "";
    let rota = "";
    let entrega = "";
    let status = "";

    if (mapaEntregas[id]) {

      const r = mapaEntregas[id];
      fantasia = r[10] || "";
      vendedor = r[14] || "";
      rota = r[156] || "";
      entrega = r[159] || "";
      status = "EM FILA";

    } else if (mapaPedidos[id]) {

      const r = mapaPedidos[id];
      fantasia = r[10] || "";
      vendedor = r[156] || "";
      rota = mapaRotas[fantasia.trim().toUpperCase()] || "";
      status = "EM FILA";

    } else if (mapaCadastro[id]) {

  fantasia = mapaCadastro[id];

  // Procura a rota pelo nome do cliente
  rota = mapaRotas[String(fantasia).trim().toUpperCase()] || "";

  status = "EM FILA";

} else {

  status = "❌ NÃO ENCONTRADO";

}

    linhas.push(row);
    updatesK.push([fantasia]);
    updatesO.push([vendedor]);
    updatesFA.push([rota]);
    updatesFD.push([entrega]);
    updatesStatus.push([status]);

  });

  linhas.forEach((row,i)=>{

    hist.getRange(row,11).setValue(updatesK[i][0]);
    hist.getRange(row,15).setValue(updatesO[i][0]);
    hist.getRange(row,157).setValue(updatesFA[i][0]);
    hist.getRange(row,160).setValue(updatesFD[i][0]);
    hist.getRange(row,164).setValue(updatesStatus[i][0]);

  });

}






/*************************************
 * CRIAR TRIGGER AUTOMÁTICO 1 MINUTO
 *************************************/
function criarTriggerFilaIntervalo() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "processarFilaHistoricoIntervalo") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("processarFilaHistoricoIntervalo")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("Trigger criado: processarFilaHistoricoIntervalo a cada 1 minuto (loop interno 30s)");
}








/*************************************
 * 1️⃣ ENVIO PARA HISTÓRICO (CHECKBOX)
 *************************************/
function enviarParaHistorico(e) {
  if (e.range.getValue() !== true) return;

  const ss = e.source;
  const ent = ss.getSheetByName(CFG.ENTREGAS);
  const hist = ss.getSheetByName(CFG.HIST);
  if (!ent || !hist) return;

  const row = e.range.getRow();
  const id = String(ent.getRange(row, CFG.COL_ID).getValue()).trim();

  if (!id) {
    e.range.setValue(false);
    return;
  }

  // 🔒 Evita duplicar no histórico
  const lastHist = hist.getLastRow();
  const idsHist = lastHist >= CFG.LINHA_HIST
    ? hist.getRange(CFG.LINHA_HIST, CFG.COL_ID, lastHist - 1, 1)
        .getValues()
        .flat()
        .map(v => String(v).trim())
    : [];

  if (idsHist.includes(id)) {
    e.range.setValue(false);
    return;
  }

  // 📋 Copia a linha inteira
  const totalCols = ent.getLastColumn();
  const dados = ent.getRange(row, 1, 1, totalCols).getValues()[0];

  // Garante coluna de status
  while (dados.length < CFG.COL_STATUS_HIST) dados.push("");
  dados[CFG.COL_STATUS_HIST - 1] = "EM ANDAMENTO";

  // ➕ Insere no topo do histórico
  hist.insertRowsAfter(1, 1);
  hist.getRange(2, 1, 1, dados.length).setValues([dados]);

  // 🕒 DATA E HORA FIXA (AUTOMÁTICA) – coluna FF (162)
  hist.getRange(2, 162).setValue(new Date());

  // 🔁 Desmarca o checkbox
  e.range.setValue(false);

  // 🎨 Reaplica formatação
  restaurarFormatacaoCondicionalHistorico();
}



/*************************************
 * 2️⃣ RESTAURAR FORMATAÇÃO CONDICIONAL
 *************************************/
function restaurarFormatacaoCondicionalHistorico() {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("HISTORICO_ENTREGAS");
  if (!sheet) return;

  const intervalo = sheet.getRange("A2:FL1000");
  const regras = [];

  // 🔴 CANCELAR (prioridade máxima)
  regras.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$FH2="CANCELAR"')
      .setBackground("#F4CCCC")
      .setFontColor("#333333")
      .setRanges([intervalo])
      .build()
  );

  // 🟢 ENTREGUE
  regras.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$FH2="ENTREGUE"')
      .setBackground("#d4edbc")
      .setFontColor("#333333")
      .setRanges([intervalo])
      .build()
  );


// 🔵 CHECKBOX FI → AZUL
regras.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$FI2=TRUE')
    .setBackground("#cfe2f3")
    .setFontColor("#333333")
    .setRanges([intervalo])
    .build()
);

    // 🟡 EM ANDAMENTO
  regras.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$FH2="EM ANDAMENTO"')
      .setBackground("#fff2cc")
      .setFontColor("#333333")
      .setRanges([intervalo])
      .build()
  );

// 🟣 EM FILA
regras.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$FH2="EM FILA"')
    .setBackground("#ead1dc")
    .setFontColor("#333333")
    .setRanges([intervalo])
    .build()
);




  sheet.setConditionalFormatRules(regras);
}



/*************************************
 * 2️⃣ STATUS DO HISTÓRICO
 *************************************/
function tratarStatusHistorico(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  if (row < CFG.LINHA_HIST) return;

  const status = String(e.range.getValue()).toUpperCase().trim();


  if (status === "CANCELAR") sheet.deleteRow(row);
}





/*************************************
 * 4️⃣ FUNÇÕES AUXILIARES DE MOVIMENTO
 *************************************/

function removerDoHistoricoApenas(sheet, id) {
  const ids = sheet.getRange(2, CFG.COL_ID, sheet.getLastRow() - 1, 1)
    .getValues().flat().map(v => String(v).trim());
  const pos = ids.indexOf(id);
  if (pos !== -1) sheet.deleteRow(2 + pos);
}

/*************************************
 * 5️⃣ MOTORISTAS
 *************************************/
function monitorarMotoristasPedido(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.MOTORISTAS) return;

  const row = e.range.getRow();
  const id = sheet.getRange(row, 1).getValue();
  const motorista = sheet.getRange(row, 2).getValue();
  if (!id || !motorista) return;

  const ent = SpreadsheetApp.getActive().getSheetByName(CFG.ENTREGAS);
  const ids = ent.getRange(3, CFG.COL_ID, ent.getLastRow() - 2, 1)
    .getValues().flat();

  const pos = ids.indexOf(id);
  if (pos !== -1) ent.getRange(pos + 3, 158).setValue(motorista);
}

function atualizarPedidosEMotoristasAtualizado() {
  const ss = SpreadsheetApp.getActive();
  const mot = ss.getSheetByName(CFG.MOTORISTAS);
  const ent = ss.getSheetByName(CFG.ENTREGAS);
  if (!mot || !ent) return;

  const dados = mot.getRange(3, 1, mot.getLastRow() - 2, 2).getValues();
  const mapa = {};

  dados.forEach(r => {
    String(r[0]).split(";").map(x => x.trim()).forEach(id => {
      if (id) mapa[id] = r[1];
    });
  });

  const idsEnt = ent.getRange(3, CFG.COL_ID, ent.getLastRow() - 2, 1).getValues().flat();
  const out = idsEnt.map(id => [mapa[String(id).trim()] || ""]);

  ent.getRange(3, 158, out.length, 1).setValues(out);
}

function registrarDataHoraHistorico(linha) {
  const hist = SpreadsheetApp.getActive().getSheetByName(CFG.HIST);
  if (!hist || linha < CFG.LINHA_HIST) return;

  const celula = hist.getRange(linha, CFG.COL_DATA_HIST);
  if (!celula.getValue()) {
    celula.setValue(new Date());
  }
}

function LIMPAR_TODAS_FORMATACOES_HISTORICO() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("HISTORICO_ENTREGAS");
  if (!sheet) return;

  sheet.clearConditionalFormatRules();
}

function alinharCoresComHistoricoCorreto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetHist = ss.getSheetByName("HISTORICO_ENTREGAS");
  const sheetEnt  = ss.getSheetByName("ENTREGAS");
  if (!sheetHist || !sheetEnt) return;

  const PRIMEIRA_LINHA_HIST = 2;
  const PRIMEIRA_LINHA_ENT  = 4;

  const COL_ID   = 2;    // B
  const COL_DATA = 160;  // FD

  const lastHist = sheetHist.getLastRow();
  const lastEnt  = sheetEnt.getLastRow();
  if (lastHist < PRIMEIRA_LINHA_HIST || lastEnt < PRIMEIRA_LINHA_ENT) return;

  const lastColEnt  = sheetEnt.getLastColumn();
  const lastColHist = sheetHist.getLastColumn();

  // =====================================================
  // 🔹 LÊ HISTÓRICO (IDS, DATAS E CORES)
  // =====================================================
  const dadosHist = sheetHist
    .getRange(
      PRIMEIRA_LINHA_HIST,
      1,
      lastHist - PRIMEIRA_LINHA_HIST + 1,
      lastColHist
    )
    .getValues();

  const coresHist = sheetHist
    .getRange(
      PRIMEIRA_LINHA_HIST,
      1,
      lastHist - PRIMEIRA_LINHA_HIST + 1,
      lastColHist
    )
    .getBackgrounds();

  // 🔑 Mapa: "ID|DATA_ISO" → cor da linha
  const mapaCores = {};

  for (let i = 0; i < dadosHist.length; i++) {
    const id = String(dadosHist[i][COL_ID - 1] || "").trim();
    const data = dadosHist[i][COL_DATA - 1];

    if (!id || !(data instanceof Date)) continue;

    const chave = id + "|" + Utilities.formatDate(
      data,
      ss.getSpreadsheetTimeZone(),
      "yyyy-MM-dd"
    );

    // Usa a cor da PRIMEIRA coluna como referência da linha
    mapaCores[chave] = coresHist[i][0];
  }

  // =====================================================
  // 🔹 LÊ ENTREGAS E APLICA CORES
  // =====================================================
  const dadosEnt = sheetEnt
    .getRange(
      PRIMEIRA_LINHA_ENT,
      1,
      lastEnt - PRIMEIRA_LINHA_ENT + 1,
      lastColEnt
    )
    .getValues();

  const backgrounds = [];

  for (let i = 0; i < dadosEnt.length; i++) {
    const id = String(dadosEnt[i][COL_ID - 1] || "").trim();
    const data = dadosEnt[i][COL_DATA - 1];

    let cor = "#FFFFFF";

    if (id && data instanceof Date) {
      const chave = id + "|" + Utilities.formatDate(
        data,
        ss.getSpreadsheetTimeZone(),
        "yyyy-MM-dd"
      );

      if (mapaCores[chave]) {
        cor = mapaCores[chave];
      }
    }

    backgrounds.push(new Array(lastColEnt).fill(cor));
  }

  // =====================================================
  // 🔹 APLICA EM BATCH
  // =====================================================
  sheetEnt
    .getRange(PRIMEIRA_LINHA_ENT, 1, backgrounds.length, lastColEnt)
    .setBackgrounds(backgrounds);

  SpreadsheetApp.flush();
}


function atualizarFJ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetHist = ss.getSheetByName("HISTORICO_ENTREGAS");
  const sheetCad = ss.getSheetByName("CADASTRO_CLIENTES");

  // Pega dados
  const dadosHist = sheetHist.getRange(2, 9, sheetHist.getLastRow()-1, 1).getValues(); // coluna I = 9
  const dadosCad = sheetCad.getRange(2, 1, sheetCad.getLastRow()-1, 4).getValues();   // col A:D

  // Cria mapa de ID -> valor coluna D
  const mapaCad = {};
  dadosCad.forEach(linha => {
    const id = linha[0];
    const valor = linha[3]; // coluna D
    if (id) mapaCad[id] = valor;
  });

  // Monta valores para FJ
  const valoresFJ = dadosHist.map(linha => {
    const id = linha[0];
    return [id && mapaCad[id] ? mapaCad[id] : ""];
  });

  // Escreve na coluna FJ (166)
  sheetHist.getRange(2, 166, valoresFJ.length, 1).setValues(valoresFJ);
}

/**********************************************************************************************
 * ATUALIZAR FD = (H + 1 DIA ÚTIL)
 * - Pega a data da coluna H da aba ENTREGAS
 * - Soma 1 dia
 * - Avança até o próximo dia útil (ignora sábados, domingos e feriados)
 **********************************************************************************************/
function atualizarFDapartirDeH() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEnt = ss.getSheetByName("ENTREGAS");
  const sheetFeriados = ss.getSheetByName("FERIADOS");
  if (!sheetEnt) return;

  const PRIMEIRA_LINHA = 4;
  const COL_H = 8;    // Coluna H
  const COL_FD = 160; // Coluna FD
  const ultima = sheetEnt.getLastRow();
  if (ultima < PRIMEIRA_LINHA) return;

  // ====== LER FERIADOS ======
  const feriados = [];
  if (sheetFeriados) {
    const ultimaF = sheetFeriados.getLastRow();
    if (ultimaF >= 3) {
      const raw = sheetFeriados.getRange(3, 1, ultimaF - 2, 1).getValues().flat();
      const tz = ss.getSpreadsheetTimeZone();
      raw.forEach(v => {
        if (v instanceof Date) {
          feriados.push(Utilities.formatDate(v, tz, "yyyy-MM-dd"));
        }
      });
    }
  }

  // ====== FUNÇÃO AUXILIAR: PEGA PRÓXIMO DIA ÚTIL ======
  function proximoDiaUtil(data) {
    const tz = ss.getSpreadsheetTimeZone();
    let d = new Date(data.getTime());
    d.setDate(d.getDate() + 1); // soma o 1º dia
    while (true) {
      let diaSemana = d.getDay(); // 0 domingo, 6 sábado
      let formato = Utilities.formatDate(d, tz, "yyyy-MM-dd");
      if (diaSemana === 0 || diaSemana === 6 || feriados.includes(formato)) {
        d.setDate(d.getDate() + 1);
        continue;
      }
      return d;
    }
  }

  // ====== LER COLUNA H ======
  const datasH = sheetEnt
    .getRange(PRIMEIRA_LINHA, COL_H, ultima - PRIMEIRA_LINHA + 1, 1)
    .getValues();

  const novasDatasFD = [];
  datasH.forEach(linha => {
    const dataH = linha[0];
    if (dataH instanceof Date) {
      novasDatasFD.push([proximoDiaUtil(dataH)]);
    } else {
      novasDatasFD.push([""]);
    }
  });

  // ====== GRAVAR COLUNA FD ======
  sheetEnt
    .getRange(PRIMEIRA_LINHA, COL_FD, novasDatasFD.length, 1)
    .setValues(novasDatasFD);

  SpreadsheetApp.flush();
  Logger.log("✔ FD atualizado a partir da coluna H com dias úteis.");
}

/**********************************************************************************************
 * 1️⃣ CALCULAR PRÓXIMO DIA ÚTIL — SEMPRE +1 DIA, DEPOIS AJUSTA
 **********************************************************************************************/
function calcularProximoDiaUtil(data, feriados = []) {
  const tz = Session.getScriptTimeZone();
  // Sempre adiciona +1 dia primeiro
  let d = new Date(data);
  d.setDate(d.getDate() + 1);
  while (true) {
    const iso = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    const diaSemana = d.getDay(); // 0 = domingo, 6 = sábado
    // Se não for final de semana e não for feriado → válido
    if (diaSemana !== 0 && diaSemana !== 6 && !feriados.includes(iso)) {
      return d;
    }
    // Caso contrário: avança mais um dia
    d.setDate(d.getDate() + 1);
  }
}

/**********************************************************************************************
 * 2️⃣ ATUALIZAR COLUNA FD (VERSÃO FINAL)
 * - Lê: ID da entrega (coluna B)
 * - Lê: data original (aba MOTORISTAS_PEDIDO)
 * - Calcula PRÓXIMO DIA ÚTIL respeitando feriados
 * - Atualiza somente células que mudaram
 **********************************************************************************************/
function atualizarColunaFD() {
  const inicio = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEntregas = ss.getSheetByName("ENTREGAS");
  const sheetMotoristas = ss.getSheetByName("MOTORISTAS_PEDIDO");
  const sheetFeriados = ss.getSheetByName("FERIADOS");
  if (!sheetEntregas || !sheetMotoristas || !sheetFeriados) return;

  const PRIMEIRA_LINHA = 4;
  const COL_ID_ENTREGAS = 2;
  const COL_FD_ENTREGAS = 160;
  const COL_DATA_MOTORISTAS = 3;

  const ultimaLinhaEntregas = sheetEntregas.getLastRow();
  if (ultimaLinhaEntregas < PRIMEIRA_LINHA) return;

  // IDs na aba ENTREGAS
  const idsEntregas = sheetEntregas
    .getRange(PRIMEIRA_LINHA, COL_ID_ENTREGAS, ultimaLinhaEntregas - PRIMEIRA_LINHA + 1, 1)
    .getValues()
    .flat()
    .map(id => String(id).trim());

  // Datas existentes na FD (antes)
  const datasAtuais = sheetEntregas
    .getRange(PRIMEIRA_LINHA, COL_FD_ENTREGAS, idsEntregas.length, 1)
    .getValues()
    .map(r => r[0]);

  // ---------------------------------------------------------------------
  // 📌 LÊ FERIADOS (a partir da linha 2, ignorando cabeçalho)
  // ---------------------------------------------------------------------
  const lastRowF = sheetFeriados.getLastRow();
  let feriados = [];
  if (lastRowF >= 2) {
    const rawF = sheetFeriados.getRange(2, 1, lastRowF - 1, 1).getValues().flat();
    const isoSet = {};
    rawF.forEach(v => {
      const iso = _toIsoDateString(v);
      if (iso) isoSet[iso] = true;
    });
    feriados = Object.keys(isoSet);
  }

  // ---------------------------------------------------------------------
  // 📌 LÊ dados da aba MOTORISTAS_PEDIDO
  // ---------------------------------------------------------------------
  const ultimaLinhaMotoristas = sheetMotoristas.getLastRow();
  if (ultimaLinhaMotoristas < 4) return;
  const dadosMotoristas = sheetMotoristas
    .getRange(4, 1, ultimaLinhaMotoristas - 3, 3)
    .getValues();

  const mapaDatas = {};
  dadosMotoristas.forEach(l => {
    const ids = String(l[0]).split(";").map(x => x.trim()).filter(x => x);
    const dataH = l[2];
    if (dataH instanceof Date && !isNaN(dataH.getTime())) {
      ids.forEach(id => mapaDatas[id] = dataH);
    }
  });

  // ---------------------------------------------------------------------
  // 📌 Atualiza apenas as células que realmente mudaram
  // ---------------------------------------------------------------------
  let alteradas = 0;
  for (let i = 0; i < idsEntregas.length; i++) {
    const id = idsEntregas[i];
    const linhaFD = sheetEntregas.getRange(PRIMEIRA_LINHA + i, COL_FD_ENTREGAS);
    const atual = datasAtuais[i];

    // 🟦 1) SE NÃO TIVER ID → LIMPA A COLUNA FD
    if (!id) {
      if (atual !== "") {
        linhaFD.clearContent();
        alteradas++;
      }
      continue;
    }

    // 🟩 2) ID EXISTE → CALCULA NOVA DATA
    const dataBase = mapaDatas[id];
    let novaData = "";
    if (dataBase instanceof Date) {
      novaData = calcularProximoDiaUtil(dataBase, feriados);
    }

    // 🟨 3) Só grava se realmente mudou
    if (String(novaData) !== String(atual)) {
      linhaFD.setValue(novaData);
      alteradas++;
    }
  }

  SpreadsheetApp.flush();
  const tempo = (new Date().getTime() - inicio) / 1000;
  Logger.log(`✅ FD atualizado (${alteradas} células) em ${tempo.toFixed(2)}s`);
}

/**********************************************************************************************
 * Converte qualquer formato de data para "yyyy-MM-dd" — robusto
 **********************************************************************************************/
function _toIsoDateString(val) {
  const tz = Session.getScriptTimeZone();
  if (!val && val !== 0) return null;

  // Se já é objeto Date
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, tz, "yyyy-MM-dd");
  }

  // Se for número (serial do Sheets)
  if (typeof val === "number") {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }

  // Se for string
  if (typeof val === "string") {
    const s = val.trim();
    // yyyy-mm-dd ou yyyy/mm/dd
    let m = s.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})$/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    // dd/mm/yyyy ou dd-mm-yyyy
    m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    // fallback - tenta converter
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }

  return null;
}

/**********************************************************************************************
 * 8️⃣ ATUALIZAR ROTAS
 **********************************************************************************************/
function atualizarRotasNaEntregas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRotas = ss.getSheetByName("ROTAS_CLIENTES");
  const sheetEntregas = ss.getSheetByName("ENTREGAS");
  if (!sheetRotas || !sheetEntregas) return;

  const PRIMEIRA_LINHA = 3;
  if (sheetRotas.getLastRow() < PRIMEIRA_LINHA || sheetEntregas.getLastRow() < PRIMEIRA_LINHA) return;

  // ====== LER DADOS DE ROTAS ======
  const dadosRotas = sheetRotas
    .getRange(PRIMEIRA_LINHA, 1, sheetRotas.getLastRow() - PRIMEIRA_LINHA + 1, 4)
    .getValues();

  const mapaRotasPedido = {};
  const mapaRotasCliente = {};

  dadosRotas.forEach(r => {
    if (r[0]) mapaRotasPedido[String(r[0]).trim()] = r[3] ? String(r[3]).trim() : "";
    if (r[1]) mapaRotasCliente[String(r[1]).trim().toUpperCase()] = r[3] ? String(r[3]).trim() : "";
  });

  // ====== LER DADOS DE ENTREGAS ======
  const rangeEnt = sheetEntregas
    .getRange(PRIMEIRA_LINHA, 2, sheetEntregas.getLastRow() - PRIMEIRA_LINHA + 1, 10)
    .getValues();

  // ====== ATUALIZAR ROTAS ======
  for (let i = 0; i < rangeEnt.length; i++) {
    const pedido = rangeEnt[i][0] ? String(rangeEnt[i][0]).trim() : null;
    const cliente = rangeEnt[i][9] ? String(rangeEnt[i][9]).trim().toUpperCase() : null;

    // Atualiza coluna 157 (rota cliente)
    if (cliente && mapaRotasCliente[cliente]) {
      sheetEntregas.getRange(PRIMEIRA_LINHA + i, 157).setValue(mapaRotasCliente[cliente]);
    } else {
      sheetEntregas.getRange(PRIMEIRA_LINHA + i, 157).clearContent();
    }

    // Atualiza coluna 159 (rota pedido)
    if (pedido && mapaRotasPedido[pedido]) {
      sheetEntregas.getRange(PRIMEIRA_LINHA + i, 159).setValue(mapaRotasPedido[pedido]);
    } else {
      sheetEntregas.getRange(PRIMEIRA_LINHA + i, 159).clearContent();
    }
  }

  Logger.log("✔ Rotas atualizadas na aba ENTREGAS.");
}

function registrarDataHoraAoDigitarID(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheet.getName() !== "HISTORICO_ENTREGAS" || col !== 9 || row < 2) return;

  const valor = e.range.getValue();
  const celulaData = sheet.getRange(row, 162); // FF

  if (valor) {
    const agoraSP = Utilities.formatDate(
      new Date(),
      "America/Sao_Paulo",
      "dd/MM/yyyy HH:mm:ss"
    );
    celulaData.setValue(agoraSP);
  } else {
    celulaData.clearContent();
  }
}

/*******************************************************************************
 * MAPA DE ENTREGAS — GEOCODIFICAÇÃO + API (JSON)
 * -----------------------------------------------------------------------------
 * Cole este bloco no FINAL do seu code.gs atual (não mexe em nada que já existe).
 *
 * O que ele faz:
 * 1) Lê a coluna ROTAS (FA = coluna 157) da aba HISTORICO_ENTREGAS — a mesma
 *    coluna que já existe, sem criar nada novo.
 * 2) SÓ processa a linha se o texto da coluna ROTAS estiver no padrão de
 *    endereço combinado:
 *       RUA/AVENIDA, NÚMERO (ou S/N) - BAIRRO - CIDADE/UF
 *    Exemplos válidos:
 *       AVENIDA BRASIL, 257 - LOTE B - MARINGÁ/PR
 *       RUA RODOLFO CREMM, S/N - JARDIM INDUSTRIAL - MARINGÁ/PR
 *       AV. COLOMBO, 5790 - MARINGÁ/PR
 *    Se o texto ainda estiver como está hoje ("PRÓXIMO A PERFINORTE",
 *    "CONTORNO SUL" etc.) a linha é simplesmente IGNORADA — não geocodifica,
 *    não aparece no mapa, não dá erro. Assim que vocês padronizarem aquela
 *    linha, ela passa a aparecer automaticamente.
 * 3) Geocodifica (transforma o endereço em lat/lng) usando o Google
 *    Geocoding API — precisa de uma chave com faturamento ativado no Google
 *    Cloud (o volume de vocês fica bem dentro da faixa gratuita).
 * 4) Guarda o resultado numa aba "Cache_Geocoding" pra nunca geocodificar o
 *    mesmo endereço duas vezes (evita estourar o limite de 1 req/segundo).
 * 5) Expõe um endpoint web (doGet) que devolve as entregas válidas em JSON,
 *    prontas pro mapa consumir.
 *
 * NÃO é preciso criar nenhuma coluna nova — só ir trocando aos poucos o texto
 * da coluna ROTAS pelo padrão de endereço acima.
 *******************************************************************************/

const MAPA_CFG = {
  COL_FANTASIA:  11,   // FANTASIA (nome do cliente)
  COL_ENTRADA: 162, // FF - ENTRADA
  COL_ENDERECO:  157,  // ROTAS (coluna FA) — usada como endereço quando padronizada
  COL_MOTORISTA: 158,  // MOTORISTA
  COL_PREVISAO:  160,  // PREVISÃO (data)
  COL_STATUS:    164,  // STATUS
  COL_CIDADE:    166,  // CIDADE
  COL_PERIODO:   167,  // PERÍODO (MANHÃ/TARDE)
  COL_OBS:       168,  // OBSERVAÇÕES
  ULTIMA_COLUNA: 173,  // maior índice de coluna usado acima (define o range de leitura)
  SHEET_CACHE_GEO: "Cache_Geocoding",

  // 🔑 Chave do Google Geocoding API — crie em console.cloud.google.com
  // (precisa de um projeto com faturamento ativado, mas o volume de vocês
  // fica bem dentro da faixa gratuita).
  GOOGLE_GEOCODING_KEY: PropertiesService.getScriptProperties().getProperty("GOOGLE_GEOCODING_KEY") || "",

  // 🔑 Chave do OpenRouteService — usada pra sugerir a melhor ordem de visita
  // das entregas com "OK" marcado. Grátis em openrouteservice.org/dev/#/signup.
  ORS_API_KEY: PropertiesService.getScriptProperties().getProperty("ORS_API_KEY") || "",

  // 📍 Endereço da Perfinorte — ponto de partida e chegada nas rotas sugeridas
  ENDERECO_ORIGEM: "RODOVIA PR 317, 4736 - GLEBA RIBEIRAO PINGUIM - MARINGÁ/PR",

  // 📅 Corte de data: o sistema ignora entregas com PREVISÃO anterior a essa
  // data (histórico antigo não interessa mais pro mapa/tabela) — entregas
  // sem nenhuma data preenchida continuam aparecendo normalmente.
  DATA_CORTE: "2026-08-07"
};

/*************************************
 * VALIDAÇÃO DE PADRÃO
 * Confere se o texto está em:
 *   RUA, NÚMERO (ou S/N) - BAIRRO - CIDADE/UF
 *************************************/
function enderecoNoPadrao_(texto) {
  if (!texto) return false;
  const t = String(texto).trim();
  if (!t) return false;

  // precisa ter pelo menos uma vírgula separando a rua do número
  const partesVirgula = t.split(',');
  if (partesVirgula.length < 2) return false;

  const resto = partesVirgula.slice(1).join(',').trim();

  // divide o restante pelos hifens (número - bairro - cidade/uf)
  const segmentos = resto.split('-').map(s => s.trim()).filter(Boolean);
  if (segmentos.length < 2) return false;

  const numero = segmentos[0];
  const cidadeUf = segmentos[segmentos.length - 1];

  const numeroOk = /^(\d+[A-Za-z]?|S\/?N)$/i.test(numero);
  const cidadeUfOk = /^.+\/[A-Za-z]{2}$/.test(cidadeUf);

  return numeroOk && cidadeUfOk;
}

/**
 * Extrai a "CIDADE/UF" de dentro do próprio texto do endereço padronizado
 * (o último pedaço depois do último hífen) — usada em vez da coluna CIDADE
 * da planilha, que pode estar desatualizada ou não bater com o endereço real.
 */
function extrairCidadeUf_(endereco) {
  const t = String(endereco || "").trim();
  const partesVirgula = t.split(',');
  if (partesVirgula.length < 2) return "";
  const resto = partesVirgula.slice(1).join(',').trim();
  const segmentos = resto.split('-').map(s => s.trim()).filter(Boolean);
  if (!segmentos.length) return "";
  return segmentos[segmentos.length - 1];
}

/*************************************
 * GEOCODIFICAÇÃO (com cache)
 *************************************/
function getOrCreateCacheGeoSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(MAPA_CFG.SHEET_CACHE_GEO);
  if (!sheet) {
    sheet = ss.insertSheet(MAPA_CFG.SHEET_CACHE_GEO);
    sheet.appendRow(["ENDEREÇO", "LAT", "LNG", "ATUALIZADO EM"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function lerCacheGeo_() {
  const sheet = getOrCreateCacheGeoSheet_();
  const lastRow = sheet.getLastRow();
  const mapa = {};
  if (lastRow < 2) return mapa;
  const dados = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  dados.forEach(r => {
    const chave = String(r[0] || "").trim().toUpperCase();
    if (chave) mapa[chave] = { lat: r[1], lng: r[2] };
  });
  return mapa;
}

/**
 * Geocodifica UM endereço (usa cache se já existir; senão chama o Google
 * Geocoding API e salva). A checagem de padrão fica a cargo de quem chama
 * esta função (geocodificarAoEditarEndereco_ e geocodificarPendentes).
 * Retorna {lat, lng} ou null se não encontrou.
 */
function geocodificarEndereco(endereco) {
  const chave = String(endereco || "").trim();
  if (!chave) return null;

  const temChave = MAPA_CFG.GOOGLE_GEOCODING_KEY && MAPA_CFG.GOOGLE_GEOCODING_KEY.trim().length > 10;
  if (!temChave) {
    Logger.log("❌ Configure MAPA_CFG.GOOGLE_GEOCODING_KEY com sua chave do Google Cloud antes de geocodificar.");
    return null;
  }

  const sheet = getOrCreateCacheGeoSheet_();
  const cache = lerCacheGeo_();
  const existente = cache[chave.toUpperCase()];
  if (existente) return existente;

  const url = "https://maps.googleapis.com/maps/api/geocode/json"
    + "?address=" + encodeURIComponent(chave)
    + "&region=br&components=country:BR"
    + "&key=" + encodeURIComponent(MAPA_CFG.GOOGLE_GEOCODING_KEY);

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const codigo = resp.getResponseCode();
    const texto = resp.getContentText() || "{}";

    if (codigo !== 200) {
      Logger.log("❌ Google Geocoding retornou HTTP " + codigo + " para '" + chave + "': " + texto);
      return null;
    }

    const json = JSON.parse(texto);

    if (json.status === "ZERO_RESULTS") {
      Logger.log("⚠ Endereço não encontrado: " + chave);
      return null;
    }
    if (json.status === "OVER_QUERY_LIMIT") {
      Logger.log("⏳ Limite de requisições do Google Geocoding atingido — tente novamente em instantes: " + chave);
      return null;
    }
    if (json.status !== "OK") {
      Logger.log("❌ Google Geocoding retornou status " + json.status + " para '" + chave + "': " + (json.error_message || ""));
      return null;
    }

    const resultado = json.results[0];
    const lat = resultado.geometry.location.lat;
    const lng = resultado.geometry.location.lng;
    sheet.appendRow([chave, lat, lng, new Date()]);
    return { lat: lat, lng: lng };
  } catch (err) {
    Logger.log("❌ Erro geocodificando '" + chave + "': " + err);
  }
  return null;
}

/**
 * Salva/atualiza manualmente a posição de um endereço no cache de
 * geocodificação — usado quando a pessoa arrasta o pino no mini-mapa da
 * "Nova entrega" pra corrigir uma localização que o Google Geocoding
 * acertou torto. Da próxima vez que esse MESMO endereço aparecer, já
 * vem certo, sem precisar corrigir de novo.
 */
function salvarCorrecaoGeocodificacao_(endereco, lat, lng) {
  const chave = String(endereco || "").trim();
  if (!chave) return;
  const latNum = parseFloat(lat);
  const lngNum = parseFloat(lng);
  if (isNaN(latNum) || isNaN(lngNum)) return;

  const sheet = getOrCreateCacheGeoSheet_();
  const lastRow = sheet.getLastRow();
  let linhaExistente = -1;
  if (lastRow >= 2) {
    const enderecos = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    linhaExistente = enderecos.findIndex(function (e) {
      return String(e || "").trim().toUpperCase() === chave.toUpperCase();
    });
  }

  const registro = [chave, latNum, lngNum, new Date()];
  if (linhaExistente >= 0) {
    sheet.getRange(linhaExistente + 2, 1, 1, 4).setValues([registro]);
  } else {
    sheet.appendRow(registro);
  }
}

/*******************************************************************************
 * MEMÓRIA DE ENDEREÇOS POR CLIENTE
 * -----------------------------------------------------------------------------
 * Toda vez que ela salva uma entrega com um endereço já no padrão certo, esse
 * endereço fica guardado — tanto na ROTAS_CLIENTES (coluna D, como já era),
 * quanto numa aba própria que junta o HISTÓRICO de todos os endereços já
 * usados por aquele código de cliente (pode ter mais de um: matriz, filial,
 * obra, etc). Da próxima vez que ela buscar esse cliente, o sistema mostra
 * os endereços já usados pra ela escolher, em vez de digitar tudo de novo.
 *******************************************************************************/
function getOrCreateHistoricoEnderecosSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("Historico_Enderecos_Cliente");
  if (!sheet) {
    sheet = ss.insertSheet("Historico_Enderecos_Cliente");
    sheet.getRange(1, 1, 1, 3).setValues([["CÓDIGO", "ENDEREÇO", "ÚLTIMO USO"]]);
  }
  return sheet;
}

/**
 * Guarda (ou atualiza a data de uso) de um endereço já usado por esse
 * código de cliente — não duplica se for o mesmo endereço de novo.
 */
function salvarEnderecoNoHistorico_(codigo, endereco) {
  const cod = String(codigo || "").trim();
  const end = String(endereco || "").trim();
  if (!cod || !end) return;

  const sheet = getOrCreateHistoricoEnderecosSheet_();
  const lastRow = sheet.getLastRow();
  let linhaExistente = -1;

  if (lastRow >= 2) {
    const dados = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    linhaExistente = dados.findIndex(function (r) {
      return String(r[0]).trim() === cod && String(r[1]).trim().toUpperCase() === end.toUpperCase();
    });
  }

  if (linhaExistente >= 0) {
    sheet.getRange(linhaExistente + 2, 3).setValue(new Date()); // só atualiza a data de uso
  } else {
    sheet.appendRow([cod, end, new Date()]);
  }
}

/**
 * Lista todos os endereços já usados por um código de cliente, do mais
 * recente pro mais antigo — pra mostrar como opção quando ela buscar esse
 * cliente de novo.
 */
function listarEnderecosCliente(codigo) {
  const cod = String(codigo || "").trim();
  if (!cod) return [];

  const sheet = getOrCreateHistoricoEnderecosSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const dados = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return dados
    .filter(function (r) { return String(r[0]).trim() === cod; })
    .sort(function (a, b) { return new Date(b[2]) - new Date(a[2]); })
    .map(function (r) { return { endereco: r[1] }; });
}

/**
 * Salva o endereço na linha da ROTAS_CLIENTES correspondente à fantasia
 * (coluna B) — coluna D (ROTA). Se a fantasia ainda não tiver linha lá,
 * cria uma nova.
 */
function salvarEnderecoRotaCliente_(fantasia, endereco) {
  const fant = String(fantasia || "").trim();
  const end = String(endereco || "").trim();
  if (!fant || !end) return;

  const ss = SpreadsheetApp.getActive();
  const rot = ss.getSheetByName("ROTAS_CLIENTES");
  if (!rot) return;

  const lastRow = rot.getLastRow();
  let linhaExistente = -1;

  if (lastRow >= 2) {
    const fantasias = rot.getRange(2, 2, lastRow - 1, 1).getValues(); // coluna B
    linhaExistente = fantasias.findIndex(function (r) {
      return String(r[0]).trim().toUpperCase() === fant.toUpperCase();
    });
  }

  if (linhaExistente >= 0) {
    rot.getRange(linhaExistente + 2, 4).setValue(end); // coluna D — ROTA
  } else {
    const novaLinha = lastRow + 1;
    rot.getRange(novaLinha, 2).setValue(fant); // coluna B — FANTASIA
    rot.getRange(novaLinha, 4).setValue(end);  // coluna D — ROTA
  }
}

/*******************************************************************************
 * NOVO CADASTRO DE CLIENTE (aba CADASTRO_CLIENTES)
 * -----------------------------------------------------------------------------
 * Coluna A = código do cliente, coluna C = nome fantasia. O código sugerido
 * é sempre o maior número já cadastrado + 1 (ela pode editar se quiser).
 *******************************************************************************/

/**
 * Olha a coluna A inteira da CADASTRO_CLIENTES, acha o maior número já
 * usado, e sugere o próximo (maior + 1). Se não achar nenhum número
 * válido, sugere 1.
 */
function obterProximoCodigoCliente_() {
  const ss = SpreadsheetApp.getActive();
  const cad = ss.getSheetByName("CADASTRO_CLIENTES");
  if (!cad) return { ok: false, erro: "Aba CADASTRO_CLIENTES não encontrada." };

  const lastRow = cad.getLastRow();
  let maiorCodigo = 0;

  if (lastRow >= 2) {
    const codigos = cad.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    codigos.forEach(function (c) {
      const num = parseInt(c, 10);
      if (!isNaN(num) && num > maiorCodigo) maiorCodigo = num;
    });
  }

  return { ok: true, proximoCodigo: maiorCodigo + 1 };
}

/**
 * Grava um novo cliente na CADASTRO_CLIENTES — coluna A (código) e coluna C
 * (nome fantasia). Não deixa duplicar um código já existente.
 */
function salvarNovoCadastroCliente(codigo, fantasia) {
  const cod = String(codigo || "").trim();
  const nome = String(fantasia || "").trim();
  if (!cod) return { ok: false, erro: "Informe o código do cliente." };
  if (!nome) return { ok: false, erro: "Informe o nome fantasia." };

  const ss = SpreadsheetApp.getActive();
  const cad = ss.getSheetByName("CADASTRO_CLIENTES");
  if (!cad) return { ok: false, erro: "Aba CADASTRO_CLIENTES não encontrada." };

  const lastRow = cad.getLastRow();
  if (lastRow >= 2) {
    const codigosExistentes = cad.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const jaExiste = codigosExistentes.some(function (c) {
      return String(c).trim() === cod;
    });
    if (jaExiste) {
      return { ok: false, erro: "Já existe um cliente cadastrado com o código " + cod + "." };
    }
  }

  const novaLinha = lastRow + 1;
  cad.getRange(novaLinha, 1).setValue(cod);  // coluna A — CÓDIGO
  cad.getRange(novaLinha, 3).setValue(nome); // coluna C — NOME FANTASIA
  SpreadsheetApp.flush();

  return { ok: true, codigo: cod, fantasia: nome };
}

/**
 * Roda em lote sobre a coluna ROTAS da HISTORICO_ENTREGAS. Só geocodifica as
 * linhas cujo texto já está no padrão de endereço; o resto é ignorado (fica
 * como está até alguém padronizar). Rode manualmente sempre que quiser
 * atualizar o cache com endereços novos/recém-padronizados.
 */
function geocodificarPendentes() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return;

  const lastRow = encontrarUltimaLinhaComDados_(hist);
  if (lastRow < 2) return;

  const rotas = hist.getRange(2, MAPA_CFG.COL_ENDERECO, lastRow - 1, 1)
    .getValues().flat();

  const cache = lerCacheGeo_();
  let novos = 0;
  let falhas = 0;
  let ignorados = 0;

  rotas.forEach(rota => {
    const texto = String(rota || "").trim();
    if (!texto) return;

    if (!enderecoNoPadrao_(texto)) { ignorados++; return; }
    if (cache[texto.toUpperCase()]) return; // já geocodificado

    const resultado = geocodificarEndereco(texto);
    if (resultado) {
      novos++;
    } else {
      falhas++;
    }
    Utilities.sleep(120); // o Google Geocoding aguenta bem mais que isso, é só uma folga de segurança
  });

  Logger.log("✔ Geocodificação concluída — " + novos + " endereço(s) novo(s) com sucesso, "
    + falhas + " falha(s), " + ignorados + " linha(s) ainda fora do padrão (ignoradas).");
}

/**
 * Geocodifica automaticamente assim que alguém edita a coluna ROTAS — mas só
 * se o novo texto já estiver no padrão de endereço. Chame esta função dentro
 * do seu onEdit(e) já existente (veja instrução no chat).
 */
function geocodificarAoEditarEndereco_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.HIST) return;
  if (e.range.getColumn() !== MAPA_CFG.COL_ENDERECO) return;
  if (e.range.getRow() < 2) return;

  const texto = String(e.range.getValue() || "").trim();
  if (!texto || !enderecoNoPadrao_(texto)) return; // fora do padrão: não faz nada

  geocodificarEndereco(texto);
}

/*************************************
 * ENDPOINT WEB — API pura (GET pra ler, POST pra escrever)
 * -----------------------------------------------------------------------------
 * O Mapa.html agora mora FORA do Apps Script (GitHub Pages) e fala com esse
 * projeto só por essas duas rotas. Nada de HtmlService aqui — só JSON.
 *
 * GET  ?formato=dados                        → todas as entregas válidas
 * GET  ?formato=motoristas                    → lista de motoristas
 * GET  ?formato=buscarCliente&codigo=X        → busca cliente por código
 * GET  ?formato=geocodificarPreview&endereco=X → geocodifica um endereço
 * POST { acao: "salvarNovaEntrega", dados }
 * POST { acao: "atualizarCampo", linha, campo, valor }
 * POST { acao: "atualizarCodigo", linha, novoCodigo }
 * POST { acao: "removerEntrega", linha }
 * POST { acao: "sugerirRota", pontos }
 * POST { acao: "marcarEntregaMotorista", linha, status, motorista }
 * POST { acao: "confirmarEntrega", linha }
 * POST { acao: "confirmarTodas" }
 *************************************/

/*******************************************************************************
 * CADASTRO DE VEÍCULOS E REFERÊNCIA DE PREÇOS DE COMBUSTÍVEL
 * ---------------------------------------------------------------------------
 * O cadastro fica numa aba própria do Google Sheets. A frota própria e os
 * transportadores terceirizados são identificados separadamente.
 * O primeiro acesso sem dados importa os veículos da planilha veiculos.xlsx.
 *******************************************************************************/

const VEICULOS_SHEET_NAME_ = "CADASTRO_VEICULOS";
const VEICULOS_HEADERS_ = [
  "ID","VINCULO_FROTA","COD_TRANSPORTADOR","TRANSPORTADOR","NOME_EXIBICAO",
  "PLACA","UF","NOME_UF","CEP","COD_MARCA","MARCA","COD_MODELO","MODELO",
  "ANO","COD_COR","COR","COD_TIPO_VEICULO","TIPO_VEICULO","PESO_MAXIMO_KG",
  "CERTIFICADO","CHASSI","VIA_TRANSPORTE","DESC_VIA_TRANSPORTE",
  "CLASSIFICACAO","DESC_CLASSIFICACAO","DATA_COMPRA","DATA_VENDA",
  "SITUACAO_DOCUMENTAL","COMBUSTIVEL_MOTOR","COMBUSTIVEL_CALCULO",
  "CONSUMO_KM_L","PRECO_COMBUSTIVEL_MANUAL","CUSTO_MANUTENCAO_KM",
  "ATIVO","OBSERVACOES","ATUALIZADO_EM"
];

function obterVeiculosIniciais_(){
  const perfinorte = "PERFINORTE INDÚSTRIA E COMÉRCIO DE PERFILADOS LTDA.";
  const data = "00/00/0000";
  return [
    ["FROTA-001","TERCEIRIZADO","4","JOAO BATISTIOL","Ford F4000","ABT6773","PR","Paraná","0","3","FORD","4","F4000","1984","3","MARROM","1","CAR/CAMINHAO/C. ABERTA",6000,"014493247769","LA7GDP68956","TER","Terrestre","C","Cavalo",data,data,"A","DIESEL","DIESEL","","","",true,"Transportador terceirizado conforme planilha veiculos.xlsx.",new Date()],
    ["FROTA-002","TERCEIRIZADO","4","JOAO BATISTIOL","Mercedes-Benz L1113 · João Batistiol","ACG5235","PR","Paraná","0","4","MERCEDES BENZ","5","L1113","1974","4","VERMELHA","1","CAR/CAMINHAO/C. ABERTA",0,"014225749360","34403212061643","TER","Terrestre","C","Cavalo",data,data,"A","DIESEL","DIESEL","","","",true,"Transportador terceirizado conforme planilha veiculos.xlsx. Peso máximo não importado: o arquivo apresenta 999999999,99, valor inválido para capacidade.",new Date()],
    ["FROTA-003","PROPRIO","13",perfinorte,"F1000","ACJ2991","PR","Paraná","0","3","FORD","3","F1000S","1984","2","PRETA","2","CAR/CAMIONETE/C. ABERTA",800,"013821216157","LA7NDM67239","TER","Terrestre","C","Cavalo",data,data,"A","DIESEL","DIESEL","","","",true,"Importado da planilha veiculos.xlsx.",new Date()],
    ["FROTA-004","PROPRIO","13",perfinorte,"Caminhão vermelho · Mercedes-Benz L1113","AIJ0E34","PR","Paraná","0","4","MERCEDES BENZ","5","L1113","1979","4","VERMELHA","1","CAR/CAMINHAO/C. ABERTA",11000,"014832784362","34403212442932","ROD","Rodoviário","C","Cavalo",data,data,"A","DIESEL","DIESEL","","","",true,"A planilha anexada identifica este caminhão como Mercedes-Benz L1113 vermelho. Confirmar divergência com a descrição informada de caminhão Volkswagen vermelho.",new Date()],
    ["FROTA-005","PROPRIO","13",perfinorte,"Volkswagen 15.180","ASZ4656","PR","Paraná","87065005","1","VOLKSWAGEN","1","15.180","2010","1","BRANCA","1","CAR/CAMINHAO/C. ABERTA",15000,"012242091419","953468235AR053341","ROD","Rodoviário","C","Cavalo",data,data,"A","DIESEL","DIESEL","","","",true,"A planilha anexada informa Volkswagen 15.180 na cor branca. Confirmar se há outro caminhão Volkswagen vermelho.",new Date()],
    ["FROTA-006","PROPRIO","13",perfinorte,"Strada Working","AVC0284","PR","Paraná","0","2","FIAT","2","STRADA WORKING","2012","1","BRANCA","2","CAR/CAMIONETE/C. ABERTA",700,"012242091419","9BD27805MC7506031","TER","Terrestre","C","Cavalo",data,data,"A","FLEX","GASOLINA","","","",true,"Modelo adicional identificado na planilha veiculos.xlsx.",new Date()],
    ["FROTA-007","PROPRIO","13",perfinorte,"Strada Freedom","RHO8D30","PR","Paraná","0","2","FIAT","6","STRADA FREEDOM 13CS","2021/2022","1","BRANCA","2","CAR/CAMIONETE/C. ABERTA",720,"213289601803","9BD281A31NYW76036","ROD","Rodoviário","C","Cavalo",data,data,"A","FLEX","GASOLINA","","","",true,"Modelo solicitado e identificado na planilha veiculos.xlsx. Referência inicial de custo em gasolina; alterável no cadastro.",new Date()],
    ["FROTA-008","PROPRIO","13",perfinorte,"HR HDB","TBQ5F75","PR","Paraná","0","5","HYUNDAI","7","HR HDB","2024/2025","1","BRANCA","2","CAR/CAMIONETE/C. ABERTA",1300,"254405089523","95PZBN7KPSB098176","ROD","Rodoviário","C","Cavalo",data,data,"A","DIESEL","DIESEL","","","",true,"Modelo solicitado e identificado na planilha veiculos.xlsx.",new Date()]
  ];
}

function migrarCadastroVeiculosAntigo_(sh){
  const ultimaColuna=sh.getLastColumn();
  if(ultimaColuna!==21||sh.getLastRow()<1)return false;
  const cab=sh.getRange(1,1,1,21).getValues()[0];
  if(String(cab[0]||"")!=="ID")return false;
  const antigos=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,21).getValues():[];
  const migrados=antigos.map(function(r){
    return [
      r[0],r[1],r[2],r[3],r[4],r[5],r[6],"","","",r[7],"",r[8],r[9],"",r[10],"",r[11],r[12],
      "","","","","","","","A",r[13],r[14],r[15],r[16],r[17],r[18],r[19],r[20]
    ];
  });
  sh.clearContents();
  sh.getRange(1,1,1,VEICULOS_HEADERS_.length).setValues([VEICULOS_HEADERS_]);
  if(migrados.length)sh.getRange(2,1,migrados.length,VEICULOS_HEADERS_.length).setValues(migrados);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,VEICULOS_HEADERS_.length).setBackground("#1b2b3d").setFontColor("#ffffff").setFontWeight("bold");
  return true;
}

function getOrCreateVeiculosSheet_(ambiente){
  const ss=SpreadsheetApp.getActive(), env=normalizarAmbienteMapa_(ambiente);
  const nomeAba=env==="TESTE"?"CADASTRO_VEICULOS_TESTE":VEICULOS_SHEET_NAME_;
  let sh=ss.getSheetByName(nomeAba);
  if(env==="TESTE"&&!sh){
    const prod=getOrCreateVeiculosSheet_("PRODUCAO"); sh=ss.insertSheet(nomeAba);
    const vals=prod.getDataRange().getValues();
    if(vals.length&&vals[0].length){sh.getRange(1,1,vals.length,vals[0].length).setValues(vals);sh.setFrozenRows(1);sh.getRange(1,1,1,VEICULOS_HEADERS_.length).setBackground("#1b2b3d").setFontColor("#fff").setFontWeight("bold");}
    return sh;
  }
  if(env==="TESTE"&&sh&&sh.getLastRow()<2){
    const prod=getOrCreateVeiculosSheet_("PRODUCAO"), vals=prod.getDataRange().getValues();
    if(vals.length>1)sh.getRange(1,1,vals.length,vals[0].length).setValues(vals);
    return sh;
  }
  if(!sh){
    sh=ss.insertSheet(VEICULOS_SHEET_NAME_);
    sh.getRange(1,1,1,VEICULOS_HEADERS_.length).setValues([VEICULOS_HEADERS_]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,VEICULOS_HEADERS_.length)
      .setBackground("#1b2b3d").setFontColor("#ffffff").setFontWeight("bold");
    sh.setColumnWidth(1,110);
    sh.setColumnWidth(4,260);
    sh.setColumnWidth(5,220);
    sh.setColumnWidth(20,180);
    sh.setColumnWidth(21,220);
    sh.setColumnWidth(35,420);
  }else{
    migrarCadastroVeiculosAntigo_(sh);
    if(sh.getLastRow()===0){
      sh.getRange(1,1,1,VEICULOS_HEADERS_.length).setValues([VEICULOS_HEADERS_]);
      sh.setFrozenRows(1);
    }
  }
  if(sh.getLastRow()<2){
    const iniciais=obterVeiculosIniciais_();
    sh.getRange(2,1,iniciais.length,VEICULOS_HEADERS_.length).setValues(iniciais);
  }
  return sh;
}

function numeroVeiculoSheet_(valor){
  if(valor===null||valor===undefined||valor==="")return "";
  const n=Number(String(valor).replace(",","."));
  return Number.isFinite(n)&&n>0?n:"";
}

function listarVeiculos_(ambiente){
  const sh=getOrCreateVeiculosSheet_(ambiente);
  const last=sh.getLastRow();
  if(last<2)return [];
  const rows=sh.getRange(2,1,last-1,VEICULOS_HEADERS_.length).getValues();
  return rows.filter(function(r){return String(r[0]||"").trim()!=="";}).map(function(r){
    return {
      id:String(r[0]||""),propriedade:String(r[1]||"PROPRIO"),
      transportadorId:String(r[2]||""),transportador:String(r[3]||""),
      nomeExibicao:String(r[4]||""),placa:String(r[5]||""),uf:String(r[6]||"PR"),
      ufNome:String(r[7]||""),cep:String(r[8]||""),marcaId:String(r[9]||""),
      marca:String(r[10]||""),modeloId:String(r[11]||""),modelo:String(r[12]||""),
      ano:String(r[13]||""),corId:String(r[14]||""),cor:String(r[15]||""),
      tipoVeiculoId:String(r[16]||""),tipoVeiculo:String(r[17]||""),
      pesoMaximoKg:numeroVeiculoSheet_(r[18]),certificado:String(r[19]||""),
      chassi:String(r[20]||""),viaTransporte:String(r[21]||""),
      viaTransporteDesc:String(r[22]||""),classificacao:String(r[23]||""),
      classificacaoDesc:String(r[24]||""),dataCompra:String(r[25]||""),
      dataVenda:String(r[26]||""),situacaoDocumental:String(r[27]||""),
      combustivelMotor:String(r[28]||"DIESEL"),combustivelCalculo:String(r[29]||"DIESEL"),
      consumoKmL:numeroVeiculoSheet_(r[30]),precoCombustivelManual:numeroVeiculoSheet_(r[31]),
      custoManutencaoKm:numeroVeiculoSheet_(r[32])||0,ativo:r[33]!==false,
      observacoes:String(r[34]||""),atualizadoEm:r[35] instanceof Date?r[35].toISOString():String(r[35]||"")
    };
  });
}

function salvarVeiculo_(v,ambiente){
  if(!v||typeof v!=="object")return {ok:false,erro:"Dados do veículo não recebidos."};
  const placa=String(v.placa||"").trim().toUpperCase().replace(/\s+/g,"");
  if(!placa)return {ok:false,erro:"Informe a placa do veículo."};
  if(!String(v.marca||"").trim()||!String(v.modelo||"").trim())return {ok:false,erro:"Informe marca e modelo."};
  const propriedade=String(v.propriedade||"PROPRIO").toUpperCase();
  if(["PROPRIO","TERCEIRIZADO"].indexOf(propriedade)<0)return {ok:false,erro:"Vínculo de frota inválido."};
  const motor=String(v.combustivelMotor||"DIESEL").toUpperCase();
  const calculo=String(v.combustivelCalculo||"DIESEL").toUpperCase();
  if(["DIESEL","FLEX","GASOLINA"].indexOf(motor)<0)return {ok:false,erro:"Tipo de motor inválido."};
  if(["DIESEL","GASOLINA","ETANOL"].indexOf(calculo)<0)return {ok:false,erro:"Combustível de cálculo inválido."};
  if(motor==="DIESEL"&&calculo!=="DIESEL")return {ok:false,erro:"Veículo diesel precisa usar diesel no cálculo."};
  if(motor==="GASOLINA"&&calculo!=="GASOLINA")return {ok:false,erro:"Veículo a gasolina precisa usar gasolina no cálculo."};
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh=getOrCreateVeiculosSheet_(ambiente),last=sh.getLastRow();
    const rows=last>=2?sh.getRange(2,1,last-1,VEICULOS_HEADERS_.length).getValues():[];
    let index=-1;const id=String(v.id||"").trim();
    for(let i=0;i<rows.length;i++){
      if(id&&String(rows[i][0])===id){index=i;continue;}
      if(String(rows[i][5]||"").trim().toUpperCase().replace(/\s+/g,"")===placa)
        return {ok:false,erro:"Já existe um veículo cadastrado com esta placa."};
    }
    const registro=[
      id||( "VEICULO-"+Utilities.getUuid()),propriedade,String(v.transportadorId||"").trim(),
      String(v.transportador||"").trim(),String(v.nomeExibicao||"").trim(),placa,
      String(v.uf||"PR").trim().toUpperCase(),String(v.ufNome||"").trim(),String(v.cep||"").trim(),
      String(v.marcaId||"").trim(),String(v.marca||"").trim(),String(v.modeloId||"").trim(),
      String(v.modelo||"").trim(),String(v.ano||"").trim(),String(v.corId||"").trim(),
      String(v.cor||"").trim(),String(v.tipoVeiculoId||"").trim(),String(v.tipoVeiculo||"").trim(),
      numeroVeiculoSheet_(v.pesoMaximoKg),String(v.certificado||"").trim(),String(v.chassi||"").trim(),
      String(v.viaTransporte||"").trim(),String(v.viaTransporteDesc||"").trim(),
      String(v.classificacao||"").trim(),String(v.classificacaoDesc||"").trim(),
      String(v.dataCompra||"").trim(),String(v.dataVenda||"").trim(),String(v.situacaoDocumental||"A").trim(),
      motor,calculo,numeroVeiculoSheet_(v.consumoKmL),numeroVeiculoSheet_(v.precoCombustivelManual),
      numeroVeiculoSheet_(v.custoManutencaoKm),v.ativo!==false,String(v.observacoes||"").trim(),new Date()
    ];
    if(index>=0)sh.getRange(index+2,1,1,VEICULOS_HEADERS_.length).setValues([registro]);
    else sh.appendRow(registro);
    return {ok:true,id:registro[0]};
  }finally{lock.releaseLock();}
}

function alterarSituacaoVeiculo_(id,ativo,ambiente){
  const chave=String(id||"").trim();
  if(!chave)return {ok:false,erro:"Identificador do veículo não informado."};
  const sh=getOrCreateVeiculosSheet_(ambiente),last=sh.getLastRow();
  if(last<2)return {ok:false,erro:"Cadastro de veículos vazio."};
  const ids=sh.getRange(2,1,last-1,1).getValues().flat();
  const idx=ids.findIndex(function(x){return String(x)===chave;});
  if(idx<0)return {ok:false,erro:"Veículo não encontrado."};
  sh.getRange(idx+2,34).setValue(ativo===true);
  sh.getRange(idx+2,36).setValue(new Date());
  return {ok:true,id:chave,ativo:ativo===true};
}

/**
 * Preços estaduais de gasolina e diesel da API Combustível.
 * A resposta é armazenada por até duas horas no cache do Apps Script.
 * Etanol não é fornecido por essa API e requer preço manual no cadastro.
 */
function obterPrecosCombustivel_(){
  const cache=CacheService.getScriptCache();
  const chave="COMBUSTIVEL_API_PRECOS_V1";
  const guardado=cache.get(chave);
  if(guardado){
    try{return JSON.parse(guardado);}catch(e){}
  }
  try{
    const resp=UrlFetchApp.fetch("https://combustivelapi.com.br/api/precos/",{
      method:"get",muteHttpExceptions:true,followRedirects:true
    });
    if(resp.getResponseCode()!==200){
      return {ok:false,erro:"API de combustível respondeu HTTP "+resp.getResponseCode()+"."};
    }
    const json=JSON.parse(resp.getContentText()||"{}");
    if(json.error!==false||!json.precos){
      return {ok:false,erro:"A API não retornou dados de preços válidos."};
    }
    const resultado={
      ok:true,
      dataColeta:String(json.data_coleta||""),
      fonte:String(json.fonte||"https://combustivelapi.com.br/"),
      moeda:String(json.moeda||"BRL"),
      precos:json.precos
    };
    cache.put(chave,JSON.stringify(resultado),7200);
    return resultado;
  }catch(err){
    return {ok:false,erro:"Não foi possível consultar a API de combustível: "+String(err)};
  }
}


/*******************************************************************************
 * MÓDULO DE VIAGENS E RATEIO LOGÍSTICO
 * Cada viagem reúne um veículo, motorista e várias entregas.
 * O custo contratado do terceirizado é registrado uma vez por viagem e rateado
 * igualmente apenas entre as entregas concluídas nessa viagem.
 *******************************************************************************/
const VIAGENS_SHEET_NAME_ = "CADASTRO_VIAGENS";
const VIAGEM_ENTREGAS_SHEET_NAME_ = "VIAGEM_ENTREGAS";
const VIAGEM_OCORRENCIAS_SHEET_NAME_ = "VIAGEM_OCORRENCIAS";

const VIAGENS_HEADERS_ = [
  "ID","NUMERO","DATA","PERIODO","VEICULO_ID","VEICULO_NOME","PLACA",
  "VINCULO_FROTA","TRANSPORTADOR","MOTORISTA","STATUS","DISTANCIA_PREVISTA_KM",
  "DURACAO_PREVISTA_MIN","CUSTO_KM_ESTIMADO","CUSTO_ESTIMADO","VALOR_NEGOCIADO",
  "PEDAGIOS","DESPESAS_EXTRAS","HODOMETRO_INICIAL","HODOMETRO_FINAL",
  "COMBUSTIVEL_REAL","CUSTO_TOTAL_FINAL","DISTANCIA_REAL_KM","OBSERVACOES",
  "CRIADO_EM","ATUALIZADO_EM","ENCERRADO_EM","VALOR_NF_TOTAL","PESO_LIQUIDO_TOTAL_KG"
];
const VIAGEM_ENTREGAS_HEADERS_ = [
  "ID","VIAGEM_ID","ID_ENTREGA","LINHA_SNAPSHOT","ORDEM","CODIGO_CLIENTE",
  "CLIENTE","ENDERECO","CIDADE","DATA_PREVISTA","PERIODO","NF_NUMEROS",
  "STATUS","MOTIVO","OBSERVACOES","CUSTO_RATEADO","ENTREGUE_EM",
  "CRIADO_EM","ATUALIZADO_EM"
];
const VIAGEM_OCORRENCIAS_HEADERS_ = [
  "ID","VIAGEM_ID","ID_ENTREGA","TIPO","DESCRICAO","DATA_HORA","REGISTRADO_POR"
];
const VIAGEM_NF_HEADERS_ = [
  "ID","VIAGEM_ID","PARADA_ID","ID_ENTREGA","CHAVE_ACESSO","NUMERO","SERIE",
  "DATA_EMISSAO","EMITENTE_CNPJ","EMITENTE_NOME","DESTINATARIO_CNPJ","DESTINATARIO_NOME",
  "MUNICIPIO_DEST","UF_DEST","VALOR_TOTAL","PESO_LIQUIDO_KG","XML_FILE_ID","XML_NOME",
  "IMPORTADO_POR","IMPORTADO_EM","AMBIENTE","STATUS"
];

function getOrCreateLogisticaSheet_(nome,headers){
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(nome);
  if(!sh){
    sh=ss.insertSheet(nome);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length)
      .setBackground("#183149").setFontColor("#ffffff").setFontWeight("bold");
  }else if(sh.getLastRow()===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }else{
    const atuais=sh.getRange(1,1,1,Math.min(sh.getLastColumn(),headers.length)).getValues()[0];
    if(String(atuais[0]||"")!==headers[0]){
      throw new Error("A aba "+nome+" já existe, mas não tem o cabeçalho esperado. Confira antes de continuar.");
    }
    const ultimaColuna=sh.getLastColumn();
    if(ultimaColuna<headers.length){
      sh.getRange(1,ultimaColuna+1,1,headers.length-ultimaColuna)
        .setValues([headers.slice(ultimaColuna)]);
      sh.getRange(1,1,1,headers.length)
        .setBackground("#183149").setFontColor("#ffffff").setFontWeight("bold");
    }
  }
  return sh;
}

function getOrCreateEstruturaViagens_(ambiente){
  const teste=normalizarAmbienteMapa_(ambiente)==="TESTE";
  return {viagens:getOrCreateLogisticaSheet_(teste?"CADASTRO_VIAGENS_TESTE":VIAGENS_SHEET_NAME_,VIAGENS_HEADERS_),
    entregas:getOrCreateLogisticaSheet_(teste?"VIAGEM_ENTREGAS_TESTE":VIAGEM_ENTREGAS_SHEET_NAME_,VIAGEM_ENTREGAS_HEADERS_),
    ocorrencias:getOrCreateLogisticaSheet_(teste?"VIAGEM_OCORRENCIAS_TESTE":VIAGEM_OCORRENCIAS_SHEET_NAME_,VIAGEM_OCORRENCIAS_HEADERS_),
    notas:getOrCreateLogisticaSheet_(teste?"VIAGEM_NOTAS_FISCAIS_TESTE":"VIAGEM_NOTAS_FISCAIS",VIAGEM_NF_HEADERS_)};
}


function filhoXmlPorNomeNFe_(elemento,nome){
  if(!elemento||!elemento.getChildren)return null;
  const filhos=elemento.getChildren();
  for(let i=0;i<filhos.length;i++)if(filhos[i].getName()===nome)return filhos[i];
  return null;
}
function caminhoXmlNFe_(elemento,caminho){
  const partes=String(caminho||"").split("/").filter(Boolean);
  let atual=elemento;
  for(let i=0;i<partes.length;i++){atual=filhoXmlPorNomeNFe_(atual,partes[i]);if(!atual)return null;}
  return atual;
}
function textoXmlNFe_(elemento,caminho){
  const alvo=caminhoXmlNFe_(elemento,caminho);
  return alvo?String(alvo.getText()||"").trim():"";
}
function normalizarDocumentoNFe_(valor){
  return String(valor||"").replace(/\D/g,"");
}
function numeroXmlNFe_(valor){
  if(valor===null||valor===undefined||String(valor).trim()==="")return null;
  const n=Number(String(valor).trim().replace(",","."));
  return Number.isFinite(n)&&n>=0?n:null;
}
function extrairDadosNFeXml_(xmlTexto){
  try{
    const doc=XmlService.parse(String(xmlTexto||""));
    const raiz=doc.getRootElement();
    const nfe=raiz.getName()==="NFe"?raiz:filhoXmlPorNomeNFe_(raiz,"NFe")||raiz;
    const inf=filhoXmlPorNomeNFe_(nfe,"infNFe");
    if(!inf)throw new Error("O XML não contém o grupo infNFe da NF-e.");
    const ide=filhoXmlPorNomeNFe_(inf,"ide");
    const emit=filhoXmlPorNomeNFe_(inf,"emit");
    const dest=filhoXmlPorNomeNFe_(inf,"dest");
    const total=filhoXmlPorNomeNFe_(inf,"total");
    const transp=filhoXmlPorNomeNFe_(inf,"transp");
    if(!ide||!emit||!dest||!total)throw new Error("A estrutura da NF-e está incompleta (ide, emit, dest ou total).");
    const atributoId=inf.getAttribute("Id");
    const chave=normalizarDocumentoNFe_(atributoId?atributoId.getValue().replace(/^NFe/i,""):"");
    const chaveFinal=chave||normalizarDocumentoNFe_(textoXmlNFe_(raiz,"protNFe/infProt/chNFe"));
    if(!/^\d{44}$/.test(chaveFinal))throw new Error("Não foi possível identificar uma chave de acesso válida de 44 dígitos.");
    const numero=textoXmlNFe_(ide,"nNF"),serie=textoXmlNFe_(ide,"serie");
    if(!numero||!serie)throw new Error("Número ou série da NF-e não encontrado.");
    let valor=numeroXmlNFe_(textoXmlNFe_(total,"vNFTot"));
    if(valor===null)valor=numeroXmlNFe_(textoXmlNFe_(total,"ICMSTot/vNF"));
    if(valor===null)throw new Error("Valor total da nota não encontrado em total/vNFTot nem em total/ICMSTot/vNF.");
    // No leiaute NF-e, cada grupo <vol> é filho direto de <transp>.
    // Pode haver mais de um volume; somamos pesoL uma única vez por grupo.
    const listaVolumes=transp?transp.getChildren().filter(x=>x.getName()==="vol"):[];
    let peso=0, volumesComPeso=0;
    listaVolumes.forEach(v=>{
      const p=numeroXmlNFe_(textoXmlNFe_(v,"pesoL"));
      if(p!==null){peso+=p;volumesComPeso++;}
    });
    if(!volumesComPeso)throw new Error("Peso líquido não encontrado em transp/vol/pesoL.");
    const emitCnpj=normalizarDocumentoNFe_(textoXmlNFe_(emit,"CNPJ"));
    const emitCpf=normalizarDocumentoNFe_(textoXmlNFe_(emit,"CPF"));
    const destCnpj=normalizarDocumentoNFe_(textoXmlNFe_(dest,"CNPJ"));
    const destCpf=normalizarDocumentoNFe_(textoXmlNFe_(dest,"CPF"));
    const emitNome=textoXmlNFe_(emit,"xNome");
    const destNome=textoXmlNFe_(dest,"xNome");
    if(!emitNome||!destNome)throw new Error("Nome do emitente ou destinatário não encontrado.");
    return {ok:true,chaveAcesso:chaveFinal,numero:numero,serie:serie,dataEmissao:textoXmlNFe_(ide,"dhEmi")||textoXmlNFe_(ide,"dEmi"),
      emitenteCnpj:emitCnpj||emitCpf,emitenteNome:emitNome,destinatarioCnpj:destCnpj||destCpf,
      destinatarioNome:destNome,municipioDestino:textoXmlNFe_(dest,"enderDest/xMun"),
      ufDestino:textoXmlNFe_(dest,"enderDest/UF"),valorTotal:valor,pesoLiquidoKg:Number(peso.toFixed(3)),
      quantidadeVolumes:volumesComPeso,versao:String(inf.getAttribute("versao")?inf.getAttribute("versao").getValue():""),
      modelo:textoXmlNFe_(ide,"mod")};
  }catch(err){return {ok:false,erro:String(err&&err.message?err.message:err)};}
}
function decodificarXmlNFeBase64_(base64){
  const texto=String(base64||"").replace(/^data:[^,]*;base64,/i,"").replace(/\s/g,"");
  if(!texto)return "";
  return Utilities.newBlob(Utilities.base64Decode(texto),"application/xml").getDataAsString("UTF-8");
}
function obterPastaXmlNFe_(ambiente){
  const teste=normalizarAmbienteMapa_(ambiente)==="TESTE";
  const nome=teste?"PERFINORTE_NFE_XML_TESTE":"PERFINORTE_NFE_XML";
  const pastas=DriveApp.getFoldersByName(nome);
  return pastas.hasNext()?pastas.next():DriveApp.createFolder(nome);
}
function recalcularTotaisNFeViagem_(estrutura,viagemId){
  const linhas=listarLinhasViagens_(estrutura.notas).filter(r=>String(r[1])===String(viagemId)&&String(r[21]||"ATIVA")==="ATIVA");
  const valor=linhas.reduce((s,r)=>s+(numeroLogistica_(r[14])||0),0);
  const peso=linhas.reduce((s,r)=>s+(numeroLogistica_(r[15])||0),0);
  const tripLast=estrutura.viagens.getLastRow();
  if(tripLast<2)return {valorTotal:valor,pesoLiquidoKg:peso,quantidadeNFs:linhas.length};
  const tripIds=estrutura.viagens.getRange(2,1,tripLast-1,1).getValues().flat();
  const idx=tripIds.findIndex(x=>String(x)===String(viagemId));
  if(idx<0)throw new Error("Viagem não encontrada ao recalcular os totais das NFs.");
  estrutura.viagens.getRange(idx+2,28,1,2).setValues([[Number(valor.toFixed(2)),Number(peso.toFixed(3))]]);
  const paradas=listarLinhasViagens_(estrutura.entregas);
  paradas.forEach((r,i)=>{
    if(String(r[1])!==String(viagemId))return;
    const nums=linhas.filter(n=>String(n[2])===String(r[0])).map(n=>String(n[5])).filter(Boolean);
    const unicos=[...new Set(nums)];
    estrutura.entregas.getRange(i+2,12).setValue(unicos.join(", "));
  });
  return {valorTotal:Number(valor.toFixed(2)),pesoLiquidoKg:Number(peso.toFixed(3)),quantidadeNFs:linhas.length};
}
function analisarXmlNFeViagem_(dados,ambiente,papel){
  const role=normalizarPapelMapa_(papel);
  if(["ADMIN","FATURAMENTO"].indexOf(role)<0)return {ok:false,erro:"Somente ADMIN ou FATURAMENTO pode analisar XML de NF-e."};
  if(!dados||!dados.xmlBase64)return {ok:false,erro:"Arquivo XML não recebido."};
  const xml=decodificarXmlNFeBase64_(dados.xmlBase64);
  if(!xml)return {ok:false,erro:"O arquivo XML está vazio."};
  return extrairDadosNFeXml_(xml);
}
function salvarNFeViagem_(dados,ambiente,papel,usuario){
  const role=normalizarPapelMapa_(papel),env=normalizarAmbienteMapa_(ambiente);
  if(["ADMIN","FATURAMENTO"].indexOf(role)<0)return {ok:false,erro:"Somente ADMIN ou FATURAMENTO pode vincular notas fiscais."};
  if(!dados||!dados.viagemId||!dados.idEntrega||!dados.xmlBase64)return {ok:false,erro:"Informe viagem, entrega e arquivo XML."};
  const xml=decodificarXmlNFeBase64_(dados.xmlBase64);
  const nf=extrairDadosNFeXml_(xml);
  if(!nf.ok)return nf;
  const estrutura=getOrCreateEstruturaViagens_(env);
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  let arquivo=null;
  try{
    const tripLast=estrutura.viagens.getLastRow();
    const tripIds=tripLast>=2?estrutura.viagens.getRange(2,1,tripLast-1,1).getValues().flat():[];
    const tripIdx=tripIds.findIndex(x=>String(x)===String(dados.viagemId));
    if(tripIdx<0)return {ok:false,erro:"Viagem não encontrada neste ambiente."};
    const tripRow=tripIdx+2;
    const statusTrip=String(estrutura.viagens.getRange(tripRow,11).getValue()||"");
    if(statusTrip!=="RASCUNHO"&&!(role==="ADMIN"&&["PROGRAMADA","EM_ANDAMENTO"].indexOf(statusTrip)>=0))
      return {ok:false,erro:"As NFs só podem ser vinculadas a um rascunho (ou viagem ativa com permissão ADMIN)."};
    const stopLast=estrutura.entregas.getLastRow();
    const stopRows=stopLast>=2?estrutura.entregas.getRange(2,1,stopLast-1,Math.max(3,estrutura.entregas.getLastColumn())).getValues():[];
    const stopIdx=stopRows.findIndex(r=>String(r[1])===String(dados.viagemId)&&String(r[2])===String(dados.idEntrega));
    if(stopIdx<0)return {ok:false,erro:"A entrega selecionada não pertence a esta viagem."};
    const duplicate=listarLinhasViagens_(estrutura.notas).some(r=>String(r[4])===nf.chaveAcesso&&String(r[21]||"ATIVA")==="ATIVA");
    if(duplicate)return {ok:false,erro:"Esta chave de acesso já está vinculada a uma viagem neste ambiente."};
    const stopId=String(stopRows[stopIdx][0]||"");
    const nomeOriginal=String(dados.nomeArquivo||("NFe-"+nf.chaveAcesso+".xml")).replace(/[\\/:*?"<>|]/g,"_").slice(0,140);
    const blob=Utilities.newBlob(Utilities.base64Decode(String(dados.xmlBase64).replace(/^data:[^,]*;base64,/i,"").replace(/\s/g,"")),"application/xml",nomeOriginal);
    arquivo=obterPastaXmlNFe_(env).createFile(blob);
    const agora=new Date();
    estrutura.notas.appendRow([
      "NFE-"+Utilities.getUuid(),String(dados.viagemId),stopId,String(dados.idEntrega),
      nf.chaveAcesso,nf.numero,nf.serie,nf.dataEmissao,nf.emitenteCnpj,nf.emitenteNome,
      nf.destinatarioCnpj,nf.destinatarioNome,nf.municipioDestino,nf.ufDestino,
      nf.valorTotal,nf.pesoLiquidoKg,arquivo.getId(),nomeOriginal,String(usuario||""),agora,env,"ATIVA"
    ]);
    const totais=recalcularTotaisNFeViagem_(estrutura,dados.viagemId);
    return {ok:true,notaFiscal:nf,totais:totais,arquivoId:arquivo.getId()};
  }catch(err){
    if(arquivo){try{arquivo.setTrashed(true);}catch(ignore){}}
    return {ok:false,erro:String(err&&err.message?err.message:err)};
  }finally{lock.releaseLock();}
}
function liberarViagemFaturamento_(dados,ambiente,papel){
  const role=normalizarPapelMapa_(papel),env=normalizarAmbienteMapa_(ambiente);
  if(["ADMIN","FATURAMENTO"].indexOf(role)<0)return {ok:false,erro:"Somente ADMIN ou FATURAMENTO pode liberar viagens."};
  if(!dados||!dados.viagemId)return {ok:false,erro:"Viagem não identificada."};
  const estrutura=getOrCreateEstruturaViagens_(env), lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const last=estrutura.viagens.getLastRow(),ids=last>=2?estrutura.viagens.getRange(2,1,last-1,1).getValues().flat():[];
    const idx=ids.findIndex(x=>String(x)===String(dados.viagemId));if(idx<0)return {ok:false,erro:"Viagem não encontrada."};
    const row=idx+2,status=String(estrutura.viagens.getRange(row,11).getValue()||"");
    if(status!=="RASCUNHO")return {ok:false,erro:"Somente viagens em rascunho podem ser liberadas."};
    const stops=listarLinhasViagens_(estrutura.entregas).filter(r=>String(r[1])===String(dados.viagemId));
    if(!stops.length)return {ok:false,erro:"A viagem não possui entregas."};
    const notas=listarLinhasViagens_(estrutura.notas).filter(r=>String(r[1])===String(dados.viagemId)&&String(r[21]||"ATIVA")==="ATIVA");
    const semNota=stops.filter(s=>!notas.some(n=>String(n[2])===String(s[0])));
    if(semNota.length)return {ok:false,erro:"Cada entrega precisa ter ao menos uma NF-e vinculada antes da liberação. Pendentes: "+semNota.map(s=>String(s[6]||s[2])).join(", ")};
    recalcularTotaisNFeViagem_(estrutura,dados.viagemId);
    estrutura.viagens.getRange(row,11).setValue("PROGRAMADA");
    estrutura.viagens.getRange(row,26).setValue(new Date());
    return {ok:true,status:"PROGRAMADA",numero:String(estrutura.viagens.getRange(row,2).getValue()||""),totalEntregas:stops.length,totalNFs:notas.length};
  }finally{lock.releaseLock();}
}

function listarLinhasViagens_(sheet){
  const last=sheet.getLastRow();
  return last<2?[]:sheet.getRange(2,1,last-1,sheet.getLastColumn()).getValues();
}

function viagemDataIso_(valor){
  if(valor instanceof Date)return Utilities.formatDate(valor,Session.getScriptTimeZone(),"yyyy-MM-dd");
  return _toIsoDateString(valor)||String(valor||"").trim();
}

function numeroLogistica_(valor){
  if(valor===null||valor===undefined||valor==="")return null;
  const n=Number(String(valor).replace(",","."));
  return Number.isFinite(n)&&n>=0?n:null;
}

function precoCombustivelServidor_(veiculo){
  const manual=numeroLogistica_(veiculo.precoCombustivelManual);
  if(manual!==null&&manual>0)return manual;
  const motor=String(veiculo.combustivelCalculo||"DIESEL").toUpperCase();
  if(motor==="ETANOL")return null;
  const pacote=obterPrecosCombustivel_();
  if(!pacote||!pacote.ok||!pacote.precos)return null;
  const grupo=motor==="GASOLINA"?pacote.precos.gasolina:pacote.precos.diesel;
  const valor=grupo&&grupo[String(veiculo.uf||"PR").toLowerCase()];
  return numeroLogistica_(valor);
}

function custoKmServidor_(veiculo){
  const consumo=numeroLogistica_(veiculo.consumoKmL);
  const preco=precoCombustivelServidor_(veiculo);
  if(!consumo||!preco)return null;
  return preco/consumo+(numeroLogistica_(veiculo.custoManutencaoKm)||0);
}

function mapearStopsViagens_(ambiente){
  const estrutura=getOrCreateEstruturaViagens_(ambiente);
  const rows=listarLinhasViagens_(estrutura.entregas);
  const mapa={};
  rows.forEach(function(r,i){
    const id=String(r[0]||"").trim();
    const viagemId=String(r[1]||"").trim();
    const entregaId=String(r[2]||"").trim();
    if(!id||!viagemId||!entregaId)return;
    const stop={
      id:id,viagemId:viagemId,idEntrega:entregaId,linhaSnapshot:Number(r[3]||0),
      ordem:Number(r[4]||0),codigo:String(r[5]||""),cliente:String(r[6]||""),
      endereco:String(r[7]||""),cidade:String(r[8]||""),data:viagemDataIso_(r[9]),
      periodo:String(r[10]||""),nfs:String(r[11]||""),status:String(r[12]||"PENDENTE"),
      motivo:String(r[13]||""),observacoes:String(r[14]||""),
      custoRateado:numeroLogistica_(r[15]),entregueEm:r[16] instanceof Date?r[16].toISOString():String(r[16]||""),
      atualizadoEm:r[18] instanceof Date?r[18].getTime():new Date(r[18]||0).getTime(),
      row:i+2
    };
    if(!mapa[entregaId])mapa[entregaId]=[];
    mapa[entregaId].push(stop);
  });
  Object.keys(mapa).forEach(id=>mapa[id].sort((a,b)=>a.atualizadoEm-b.atualizadoEm));
  return {rows:rows,mapa:mapa};
}

function listarDadosViagens_(ambiente,papel){
  const env=normalizarAmbienteMapa_(ambiente), role=normalizarPapelMapa_(papel);
  const estrutura=getOrCreateEstruturaViagens_(env);
  const tripRows=listarLinhasViagens_(estrutura.viagens);
  const stopData=mapearStopsViagens_(env);
  const veiculos=listarVeiculos_(env);
  const motoristas=listarMotoristas();
  const entregas=obterEntregasValidas_().filter(function(d){
    return String(d.idEntrega||"").trim()!=="";
  });

  const nfRows=listarLinhasViagens_(estrutura.notas);
  const notasPorViagem={};
  nfRows.forEach(r=>{
    if(String(r[21]||"ATIVA")!=="ATIVA")return;
    const tripId=String(r[1]||"");if(!tripId)return;
    if(!notasPorViagem[tripId])notasPorViagem[tripId]=[];
    notasPorViagem[tripId].push({id:String(r[0]||""),idEntrega:String(r[3]||""),chaveAcesso:String(r[4]||""),numero:String(r[5]||""),serie:String(r[6]||""),dataEmissao:String(r[7]||""),emitenteCnpj:String(r[8]||""),emitenteNome:String(r[9]||""),destinatarioCnpj:String(r[10]||""),destinatarioNome:String(r[11]||""),municipioDestino:String(r[12]||""),ufDestino:String(r[13]||""),valorTotal:numeroLogistica_(r[14])||0,pesoLiquidoKg:numeroLogistica_(r[15])||0,nomeArquivo:String(r[17]||"")});
  });
  const ocorrenciaRows=listarLinhasViagens_(estrutura.ocorrencias);
  const ocorrenciasPorViagem={};
  ocorrenciaRows.forEach(function(r){
    const tripId=String(r[1]||"");
    if(!tripId)return;
    if(!ocorrenciasPorViagem[tripId])ocorrenciasPorViagem[tripId]=[];
    ocorrenciasPorViagem[tripId].push({
      tipo:String(r[3]||""),descricao:String(r[4]||""),
      dataHora:r[5] instanceof Date?Utilities.formatDate(r[5],Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm"):String(r[5]||"")
    });
  });
  const viagens=tripRows.filter(function(r){return role!=="OPERACIONAL"||String(r[10]||"")!=="RASCUNHO";}).map(function(r,i){
    const id=String(r[0]||"");
    const stops=stopData.rows.filter(s=>String(s[1]||"")===id).map(function(s,j){
      return {
        id:String(s[0]||""),viagemId:id,idEntrega:String(s[2]||""),
        linhaSnapshot:Number(s[3]||0),ordem:Number(s[4]||j+1),
        codigo:String(s[5]||""),cliente:String(s[6]||""),endereco:String(s[7]||""),
        cidade:String(s[8]||""),data:viagemDataIso_(s[9]),periodo:String(s[10]||""),
        nfs:String(s[11]||""),status:String(s[12]||"PENDENTE"),motivo:String(s[13]||""),
        observacoes:String(s[14]||""),custoRateado:numeroLogistica_(s[15]),
        entregueEm:s[16] instanceof Date?s[16].toISOString():String(s[16]||"")
      };
    });
    return {
      id:id,numero:String(r[1]||("V-"+(i+1))),data:viagemDataIso_(r[2]),
      periodo:String(r[3]||""),veiculoId:String(r[4]||""),veiculoNome:String(r[5]||""),
      placa:String(r[6]||""),propriedade:String(r[7]||""),transportador:String(r[8]||""),
      motorista:String(r[9]||""),status:String(r[10]||"PROGRAMADA"),
      distanciaPrevistaKm:numeroLogistica_(r[11]),duracaoPrevistaMin:numeroLogistica_(r[12]),
      custoKmEstimado:numeroLogistica_(r[13]),custoEstimado:numeroLogistica_(r[14]),
      valorNegociado:numeroLogistica_(r[15]),pedagios:numeroLogistica_(r[16])||0,
      despesasExtras:numeroLogistica_(r[17])||0,hodometroInicial:numeroLogistica_(r[18]),
      hodometroFinal:numeroLogistica_(r[19]),combustivelReal:numeroLogistica_(r[20]),
      custoTotalFinal:numeroLogistica_(r[21]),distanciaRealKm:numeroLogistica_(r[22]),
      observacoes:String(r[23]||""),paradas:stops,ocorrencias:ocorrenciasPorViagem[id]||[],
      criadoEm:r[24] instanceof Date?r[24].toISOString():String(r[24]||""),
      atualizadoEm:r[25] instanceof Date?r[25].toISOString():String(r[25]||""),
      encerradoEm:r[26] instanceof Date?r[26].toISOString():String(r[26]||""),
      valorNfTotal:numeroLogistica_(r[27])||0,pesoLiquidoTotalKg:numeroLogistica_(r[28])||0,
      notasFiscais:notasPorViagem[id]||[]
    };
  }).sort((a,b)=>String(b.data).localeCompare(String(a.data))||String(b.criadoEm).localeCompare(String(a.criadoEm)));

  const ultimaParada={};
  Object.keys(stopData.mapa).forEach(id=>{
    const arr=stopData.mapa[id];
    if(arr.length)ultimaParada[id]=arr[arr.length-1];
  });
  const atribuidasAtivas=new Set();
  Object.keys(ultimaParada).forEach(id=>{
    const s=ultimaParada[id];
    if(["PENDENTE","EM_ROTA"].indexOf(s.status)>=0)atribuidasAtivas.add(id);
  });
  const entregasDisponiveis=entregas.filter(function(d){
    const status=String(d.status||"").trim().toUpperCase();
    return d.codigo&&d.endereco&&status!=="ENTREGUE"&&!atribuidasAtivas.has(d.idEntrega);
  }).map(function(d){
    const ultima=ultimaParada[d.idEntrega];
    return Object.assign({},d,{
      podeReprogramar:!!ultima&&["REPROGRAMAR","NAO_ENTREGUE"].indexOf(ultima.status)>=0,
      viagemAnterior:ultima?ultima.viagemId:""
    });
  });

  return {
    ok:true,viagens:viagens,veiculos:veiculos.filter(v=>v.ativo!==false),
    motoristas:motoristas,entregasDisponiveis:entregasDisponiveis
  };
}

function criarViagem_(dados,ambiente,papel){
  if(!dados||typeof dados!=="object")return {ok:false,erro:"Dados da viagem não recebidos."};
  const data=String(dados.data||"").trim();
  const periodo=String(dados.periodo||"").trim().toUpperCase();
  const motorista=String(dados.motorista||"").trim();
  const ids=Array.isArray(dados.entregas)?dados.entregas:[];
  if(!data||!motorista)return {ok:false,erro:"Informe data e motorista."};
  if(["MANHÃ","TARDE"].indexOf(periodo)<0)return {ok:false,erro:"Selecione manhã ou tarde para a viagem."};
  if(!ids.length)return {ok:false,erro:"Selecione ao menos uma entrega."};

  const env=normalizarAmbienteMapa_(ambiente), role=normalizarPapelMapa_(papel);
  const estrutura=getOrCreateEstruturaViagens_(env);
  const veiculo=listarVeiculos_(env).find(v=>String(v.id)===String(dados.veiculoId||""));
  if(!veiculo||veiculo.ativo===false)return {ok:false,erro:"Selecione um veículo ativo."};
  const valorNegociado=numeroLogistica_(dados.valorNegociado)||0;
  const terceiros=String(veiculo.propriedade).toUpperCase()==="TERCEIRIZADO";
  if(terceiros&&valorNegociado<=0)return {ok:false,erro:"Informe o valor total negociado com o transportador para esta viagem."};

  const validas=obterEntregasValidas_();
  const porLinha={};
  validas.forEach(d=>{
    const id=String(d.idEntrega||"").trim();
    if(id)porLinha[id]=d;
  });
  const stopData=mapearStopsViagens_(env);
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const selecionadas=[];
    ids.forEach(function(item){
      const id=String(item.idEntrega||"").trim();
      const d=porLinha[id];
      if(!d)throw new Error("Uma das entregas selecionadas não foi encontrada. Atualize a lista e tente novamente.");
      if(String(d.status||"").trim().toUpperCase()==="ENTREGUE")throw new Error("A entrega "+d.cliente+" já está marcada como entregue.");
      const historico=stopData.mapa[id]||[];
      const ultima=historico.length?historico[historico.length-1]:null;
      if(ultima&&["PENDENTE","EM_ROTA"].indexOf(ultima.status)>=0){
        throw new Error("A entrega "+d.cliente+" já está vinculada a uma viagem em aberto.");
      }
      selecionadas.push({item:item,d:d,ordem:Number(item.ordem||selecionadas.length+1)});
    });
    const idViagem="VIAGEM-"+Utilities.getUuid();
    const numero="V-"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyMMdd")+"-"+String(estrutura.viagens.getLastRow()).padStart(3,"0");
    const distancia=numeroLogistica_(dados.distanciaPrevistaKm);
    const duracao=numeroLogistica_(dados.duracaoPrevistaMin);
    const custoKm=terceiros?null:custoKmServidor_(veiculo);
    const pedagios=numeroLogistica_(dados.pedagios)||0;
    const extras=numeroLogistica_(dados.despesasExtras)||0;
    const estimado=terceiros?valorNegociado+pedagios+extras:
      (distancia!==null&&custoKm!==null?distancia*custoKm+pedagios+extras:null);
    const agora=new Date();
    estrutura.viagens.appendRow([
      idViagem,numero,parseDataIsoParaDate_(data),String(dados.periodo||""),
      veiculo.id,veiculo.nomeExibicao||[veiculo.marca,veiculo.modelo].filter(Boolean).join(" "),
      veiculo.placa,String(veiculo.propriedade),veiculo.transportador||"",motorista,
      role==="FATURAMENTO"?"RASCUNHO":"PROGRAMADA",distancia,duracao,custoKm,estimado,terceiros?valorNegociado:"",
      pedagios,extras,numeroLogistica_(dados.hodometroInicial)||"",
      "",numeroLogistica_(dados.combustivelReal)||"","",null,
      String(dados.observacoes||""),agora,agora,""
    ]);
    selecionadas.forEach(function(x,i){
      const d=x.d,item=x.item;
      estrutura.entregas.appendRow([
        "PARADA-"+Utilities.getUuid(),idViagem,d.idEntrega,d.linha,
        x.ordem||i+1,d.codigo,d.cliente,d.endereco,d.cidade||"",
        d.data?parseDataIsoParaDate_(d.data):"",d.periodo||"",
        String(item.nfs||""),"PENDENTE","","", "", "",agora,agora
      ]);
    });
    SpreadsheetApp.flush();
    return {ok:true,id:idViagem,numero:numero,totalEntregas:selecionadas.length,custoEstimado:estimado};
  }finally{lock.releaseLock();}
}

function atualizarParadaViagem_(dados,ambiente){
  if(!dados||!dados.id)return {ok:false,erro:"Parada não identificada."};
  const env=normalizarAmbienteMapa_(ambiente), estrutura=getOrCreateEstruturaViagens_(env);
  const last=estrutura.entregas.getLastRow();
  if(last<2)return {ok:false,erro:"Nenhuma parada cadastrada."};
  const ids=estrutura.entregas.getRange(2,1,last-1,1).getValues().flat();
  const idx=ids.findIndex(x=>String(x)===String(dados.id));
  if(idx<0)return {ok:false,erro:"Parada não encontrada."};
  const row=idx+2;
  const viagemId=String(estrutura.entregas.getRange(row,2).getValue()||"");
  const viagens=estrutura.viagens;
  const lastV=viagens.getLastRow();
  const tripIds=lastV>=2?viagens.getRange(2,1,lastV-1,1).getValues().flat():[];
  const tripIdx=tripIds.findIndex(x=>String(x)===viagemId);
  if(tripIdx<0)return {ok:false,erro:"Viagem não encontrada."};
  const tripStatus=String(viagens.getRange(tripIdx+2,11).getValue()||"");
  if(["CONCLUIDA","CONCLUIDA_PARCIAL","CANCELADA"].indexOf(tripStatus)>=0)
    return {ok:false,erro:"Esta viagem já foi encerrada."};

  const status=String(dados.status||"").toUpperCase();
  const permitidos=["ENTREGUE","NAO_ENTREGUE","REPROGRAMAR","EM_ROTA","PENDENTE"];
  if(permitidos.indexOf(status)<0)return {ok:false,erro:"Situação de parada inválida."};

  if(status==="ENTREGUE"&&env!=="TESTE"){
    const idEntrega=String(estrutura.entregas.getRange(row,3).getValue()||"");
    const hist=SpreadsheetApp.getActive().getSheetByName(CFG.HIST);
    const historicoIds=hist.getRange(2,AUDITORIA_ENTREGAS_CFG_.COL_ID,Math.max(1,hist.getLastRow()-1),1).getValues().flat();
    const histIdx=historicoIds.findIndex(x=>String(x)===idEntrega);
    if(histIdx<0)return {ok:false,erro:"Não localizei a entrega original na planilha. Ela pode ter sido removida."};
    const linhaHistorico=histIdx+2;
    const statusOrigem=String(hist.getRange(linhaHistorico,MAPA_CFG.COL_STATUS).getValue()||"").trim().toUpperCase();
    if(statusOrigem!=="ENTREGUE"){
      const resultado=concluirEntrega(linhaHistorico,dados.motivoAtraso||"",dados.observacoes||"");
      if(!resultado||!resultado.ok)return resultado||{ok:false,erro:"Não foi possível atualizar a entrega."};
    }
    estrutura.entregas.getRange(row,17).setValue(new Date());
  }
  estrutura.entregas.getRange(row,13).setValue(status);
  estrutura.entregas.getRange(row,14).setValue(String(dados.motivoAtraso||dados.motivo||""));
  estrutura.entregas.getRange(row,15).setValue(String(dados.observacoes||""));
  estrutura.entregas.getRange(row,12).setValue(String(dados.nfs||estrutura.entregas.getRange(row,12).getValue()||""));
  estrutura.entregas.getRange(row,19).setValue(new Date());
  SpreadsheetApp.flush();
  return {ok:true,id:dados.id,status:status};
}

function registrarOcorrenciaViagem_(dados,ambiente){
  if(!dados||!dados.viagemId||!String(dados.descricao||"").trim())
    return {ok:false,erro:"Informe a viagem e descreva a ocorrência."};
  const estrutura=getOrCreateEstruturaViagens_(ambiente);
  const last=estrutura.viagens.getLastRow();
  if(last<2)return {ok:false,erro:"Viagem não encontrada."};
  const ids=estrutura.viagens.getRange(2,1,last-1,1).getValues().flat();
  const idx=ids.findIndex(x=>String(x)===String(dados.viagemId));
  if(idx<0)return {ok:false,erro:"Viagem não encontrada."};
  const tipo=String(dados.tipo||"OUTRA").toUpperCase();
  const agora=new Date();
  estrutura.ocorrencias.appendRow([
    "OCORRENCIA-"+Utilities.getUuid(),dados.viagemId,String(dados.idEntrega||""),
    tipo,String(dados.descricao).trim(),agora,String(dados.registradoPor||"Administrativo")
  ]);
  if(["PANE","ACIDENTE","VEICULO_INDISPONIVEL","BLOQUEIO"].indexOf(tipo)>=0){
    estrutura.viagens.getRange(idx+2,11).setValue("INTERROMPIDA");
    estrutura.viagens.getRange(idx+2,26).setValue(agora);
  }
  SpreadsheetApp.flush();
  return {ok:true};
}

function atualizarStatusViagem_(dados,ambiente){
  if(!dados||!dados.id)return {ok:false,erro:"Viagem não identificada."};
  const estrutura=getOrCreateEstruturaViagens_(ambiente);
  const last=estrutura.viagens.getLastRow();
  if(last<2)return {ok:false,erro:"Viagem não encontrada."};
  const ids=estrutura.viagens.getRange(2,1,last-1,1).getValues().flat();
  const idx=ids.findIndex(x=>String(x)===String(dados.id));
  if(idx<0)return {ok:false,erro:"Viagem não encontrada."};
  const row=idx+2;
  const atual=String(estrutura.viagens.getRange(row,11).getValue()||"PROGRAMADA");
  const status=String(dados.status||"").toUpperCase();
  if(status==="EM_ANDAMENTO"){
    if(["PROGRAMADA","INTERROMPIDA"].indexOf(atual)<0)return {ok:false,erro:"Só é possível iniciar ou retomar uma viagem programada/interrompida."};
    estrutura.viagens.getRange(row,11).setValue(status);
    estrutura.viagens.getRange(row,26).setValue(new Date());
    return {ok:true,status:status};
  }
  if(status==="INTERROMPIDA"){
    if(atual!=="EM_ANDAMENTO")return {ok:false,erro:"Somente uma viagem em andamento pode ser interrompida."};
    estrutura.viagens.getRange(row,11).setValue(status);
    estrutura.viagens.getRange(row,26).setValue(new Date());
    return {ok:true,status:status};
  }
  if(status!=="CONCLUIR")return {ok:false,erro:"Situação de viagem inválida."};
  if(["EM_ANDAMENTO","INTERROMPIDA","PROGRAMADA"].indexOf(atual)<0)return {ok:false,erro:"Esta viagem não pode ser encerrada."};

  const rows=listarLinhasViagens_(estrutura.entregas);
  const idxs=[];
  rows.forEach(function(r,i){if(String(r[1])===String(dados.id))idxs.push(i+2);});
  if(!idxs.length)return {ok:false,erro:"Esta viagem não possui entregas vinculadas."};
  const pendentes=idxs.filter(r=>["PENDENTE","EM_ROTA"].indexOf(String(estrutura.entregas.getRange(r,13).getValue()||""))>=0);
  if(pendentes.length)return {ok:false,erro:"Antes de encerrar, marque as entregas restantes como reprogramar ou não entregues."};
  const concluidas=idxs.filter(r=>String(estrutura.entregas.getRange(r,13).getValue()||"")==="ENTREGUE");
  const veiculoId=String(estrutura.viagens.getRange(row,5).getValue()||"");
  const veiculo=listarVeiculos_(ambiente).find(v=>String(v.id)===veiculoId);
  if(!veiculo)return {ok:false,erro:"Veículo vinculado não encontrado."};
  const terceiros=String(veiculo.propriedade).toUpperCase()==="TERCEIRIZADO";
  const valorNegociado=numeroLogistica_(estrutura.viagens.getRange(row,16).getValue())||0;
  const pedagios=numeroLogistica_(dados.pedagios)!==null?numeroLogistica_(dados.pedagios):(numeroLogistica_(estrutura.viagens.getRange(row,17).getValue())||0);
  const extras=numeroLogistica_(dados.despesasExtras)!==null?numeroLogistica_(dados.despesasExtras):(numeroLogistica_(estrutura.viagens.getRange(row,18).getValue())||0);
  const hodIni=numeroLogistica_(dados.hodometroInicial)!==null?numeroLogistica_(dados.hodometroInicial):numeroLogistica_(estrutura.viagens.getRange(row,19).getValue());
  const hodFim=numeroLogistica_(dados.hodometroFinal);
  if(hodIni!==null&&hodFim!==null&&hodFim<hodIni)return {ok:false,erro:"O hodômetro final não pode ser menor que o inicial."};
  const kmReal=hodIni!==null&&hodFim!==null?hodFim-hodIni:null;
  const kmPrev=numeroLogistica_(estrutura.viagens.getRange(row,12).getValue());
  const kmAplicado=kmReal!==null?kmReal:kmPrev;
  const combustivelReal=numeroLogistica_(dados.combustivelReal)!==null?numeroLogistica_(dados.combustivelReal):numeroLogistica_(estrutura.viagens.getRange(row,21).getValue());
  const custoKm=numeroLogistica_(estrutura.viagens.getRange(row,14).getValue());
  const manutencao=numeroLogistica_(veiculo.custoManutencaoKm)||0;
  let custoTotal=null;
  if(terceiros)custoTotal=valorNegociado+pedagios+extras;
  else if(kmAplicado!==null){
    const custoBase=combustivelReal!==null?combustivelReal+kmAplicado*manutencao:
      (custoKm!==null?kmAplicado*custoKm:null);
    if(custoBase!==null)custoTotal=custoBase+pedagios+extras;
  }
  const agora=new Date();
  estrutura.viagens.getRange(row,17,1,6).setValues([[
    pedagios,extras,hodIni===null?"":hodIni,hodFim===null?"":hodFim,
    combustivelReal===null?"":combustivelReal,custoTotal===null?"":custoTotal
  ]]);
  estrutura.viagens.getRange(row,23).setValue(kmReal===null?"":kmReal);
  estrutura.viagens.getRange(row,11).setValue(concluidas.length===idxs.length?"CONCLUIDA":"CONCLUIDA_PARCIAL");
  estrutura.viagens.getRange(row,26).setValue(agora);
  estrutura.viagens.getRange(row,27).setValue(agora);

  if(custoTotal!==null&&concluidas.length){
    const parcela=Math.floor((custoTotal/concluidas.length)*100)/100;
    let acumulado=0;
    concluidas.forEach(function(stopRow,i){
      const rateio=i===concluidas.length-1?Math.round((custoTotal-acumulado)*100)/100:parcela;
      acumulado+=rateio;
      estrutura.entregas.getRange(stopRow,16).setValue(rateio);
      estrutura.entregas.getRange(stopRow,19).setValue(agora);
    });
  }
  SpreadsheetApp.flush();
  return {
    ok:true,status:concluidas.length===idxs.length?"CONCLUIDA":"CONCLUIDA_PARCIAL",
    custoTotal:custoTotal,entregasConcluidas:concluidas.length,
    entregasReprogramadas:idxs.length-concluidas.length,
    kmReal:kmReal,rateioPorEntrega:concluidas.length&&custoTotal!==null?custoTotal/concluidas.length:null
  };
}

function incluirViagemNoModulo_(e){return listarDadosViagens_();}


const CONFIG_MODULOS_FROTA_KEY_ = "PERFINORTE_MODULOS_FROTA_GLOBAL_V1";

function normalizarAmbienteMapa_(ambiente) {
  return String(ambiente || "PRODUCAO").trim().toUpperCase() === "TESTE" ? "TESTE" : "PRODUCAO";
}
function normalizarPapelMapa_(papel) {
  const p = String(papel || "OPERACIONAL").trim().toUpperCase();
  return ["ADMIN", "OPERACIONAL", "FATURAMENTO"].indexOf(p) >= 0 ? p : "OPERACIONAL";
}
function configModulosPadraoMapa_() {
  return {veiculos:{teste:false,producao:false},viagens:{teste:false,producao:false},motivosAtraso:{teste:false,producao:false}};
}
function obterConfigModulosFrota_() {
  const props=PropertiesService.getScriptProperties(), salvo=props.getProperty(CONFIG_MODULOS_FROTA_KEY_), legadoMotivo=props.getProperty(CONFIG_EXIGIR_MOTIVO_ATRASO_);
  const padrao=configModulosPadraoMapa_();
  if(!salvo){padrao.motivosAtraso.producao=legadoMotivo==="true";return {ok:true,config:padrao,padrao:true};}
  try{
    const valor=JSON.parse(salvo)||{}, config=configModulosPadraoMapa_();
    ["veiculos","viagens","motivosAtraso"].forEach(function(nome){
      const atual=valor[nome];
      if(atual&&typeof atual==="object"){config[nome].teste=atual.teste===true;config[nome].producao=atual.producao===true;}
      else if(typeof atual==="boolean")config[nome].producao=atual;
    });
    if(!valor.motivosAtraso||typeof valor.motivosAtraso!=="object")config.motivosAtraso.producao=legadoMotivo==="true";
    return {ok:true,config:config,padrao:false};
  }catch(err){padrao.motivosAtraso.producao=legadoMotivo==="true";return {ok:true,config:padrao,padrao:true};}
}
function salvarConfigModulosFrota_(modulos) {
  if(!modulos||typeof modulos!=="object")return {ok:false,erro:"Configuração dos módulos não recebida."};
  const atual=obterConfigModulosFrota_().config, nova=configModulosPadraoMapa_();
  ["veiculos","viagens","motivosAtraso"].forEach(function(nome){
    const recebido=modulos[nome];
    if(recebido&&typeof recebido==="object"){nova[nome].teste=recebido.teste===true;nova[nome].producao=recebido.producao===true;}
    else if(typeof recebido==="boolean")nova[nome]={teste:atual[nome].teste,producao:recebido};
    else nova[nome]={teste:atual[nome].teste,producao:atual[nome].producao};
  });
  PropertiesService.getScriptProperties().setProperty(CONFIG_MODULOS_FROTA_KEY_,JSON.stringify(nova));
  return {ok:true,config:nova};
}
function moduloFrotaAtivo_(nome,ambiente) {
  const k=normalizarAmbienteMapa_(ambiente)==="TESTE"?"teste":"producao", c=obterConfigModulosFrota_().config;
  return !!(c[nome]&&c[nome][k]===true);
}

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  if (!validarWorkerMapa_(p.secret)) return jsonOutput_({ok:false,erro:"Acesso não autorizado"});
  const papel=normalizarPapelMapa_(p._authRole), ambiente=normalizarAmbienteMapa_(p._ambiente);
  if(ambiente==="TESTE"&&papel!=="ADMIN")return jsonOutput_({ok:false,erro:"O ambiente de teste é exclusivo do perfil ADMIN."});
  switch (p.formato) {
    case "dados":
      return doGetDados_(e);
    case "configModulosFrota": {
      const conf=obterConfigModulosFrota_(), k=ambiente==="TESTE"?"teste":"producao";
      return jsonOutput_({ok:true,config:papel==="ADMIN"?conf.config:null,modulos:{veiculos:conf.config.veiculos[k]===true,viagens:conf.config.viagens[k]===true,motivosAtraso:conf.config.motivosAtraso[k]===true},ambiente:ambiente,papel:papel,padrao:conf.padrao});
    }
    case "veiculos": {
      // FATURAMENTO consulta a frota em modo somente leitura para montar a viagem.
      const podeConsultar=moduloFrotaAtivo_("veiculos",ambiente)||
        (papel==="FATURAMENTO"&&moduloFrotaAtivo_("viagens",ambiente));
      if(!podeConsultar)return jsonOutput_({ok:false,erro:"O módulo Veículos está desativado neste ambiente."});
      return jsonOutput_(listarVeiculos_(ambiente));
    }
    case "viagens":
      if(!moduloFrotaAtivo_("viagens",ambiente))return jsonOutput_({ok:false,erro:"O módulo Viagens está desativado neste ambiente."});
      return jsonOutput_(listarDadosViagens_(ambiente,papel));
    case "precosCombustivel":
      return jsonOutput_(obterPrecosCombustivel_());
    case "motoristas":
      return jsonOutput_(listarMotoristas());
    case "buscarCliente":
      return jsonOutput_(buscarClientePorCodigo(p.codigo));
    case "geocodificarPreview":
      return jsonOutput_(geocodificarPreview(p.endereco));
    case "confirmacoesPendentes":
      return jsonOutput_(listarConfirmacoesPendentes());
    case "coresStatus":
      return jsonOutput_(obterCoresStatus_());
    case "localizacoesMotoristas":
      return jsonOutput_(listarLocalizacoesMotoristas());
    case "etaBase":
      return jsonOutput_(calcularEtaBase(p.lat, p.lng));
    case "proximoDiaUtil":
      return jsonOutput_({ data: proximoDiaUtil_(new Date()) });
    case "rotaEntrePontos":
      return jsonOutput_(calcularRotaEntrePontos(p.lat1, p.lng1, p.lat2, p.lng2));
    case "origemPerfinorte":
      return jsonOutput_(obterOrigemPerfinorte());
    case "dadosMotorista":
      return jsonOutput_(obterDadosMotorista_(p.motorista));
    case "enderecosCliente":
      return jsonOutput_({ enderecos: listarEnderecosCliente(p.codigo) });
    case "proximoCodigoCliente":
      return jsonOutput_(obterProximoCodigoCliente_());
    default:
      return jsonOutput_({ erro: "Use ?formato=dados, motoristas, buscarCliente, geocodificarPreview, confirmacoesPendentes, coresStatus, localizacoesMotoristas, etaBase, proximoDiaUtil, rotaEntrePontos, origemPerfinorte, dadosMotorista, enderecosCliente ou proximoCodigoCliente." });
  }
}

/**
 * Recebe as ações que escrevem na planilha (a página do mapa é externa
 * agora, então essas ações chegam via POST em vez de google.script.run).
 * Corpo esperado: { acao: "...", ...demais campos de cada ação }
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (!validarWorkerMapa_(body._workerSecret)) return jsonOutput_({ok:false,erro:"Acesso não autorizado"});
    const papel=normalizarPapelMapa_(body._authRole), ambiente=normalizarAmbienteMapa_(body._ambiente), acao=String(body.acao||"");
    if(ambiente==="TESTE"&&papel!=="ADMIN")return jsonOutput_({ok:false,erro:"O ambiente de teste é exclusivo do perfil ADMIN."});
    const somenteAdmin=["salvarConfigMotivoAtraso","salvarConfigModulosFrota","salvarVeiculo","alterarSituacaoVeiculo"];
    if(somenteAdmin.indexOf(acao)>=0&&papel!=="ADMIN")return jsonOutput_({ok:false,erro:"Ação permitida somente ao perfil ADMIN."});
    if(["criarViagem","analisarXmlNFe","salvarNFeViagem","liberarViagem"].indexOf(acao)>=0&&["ADMIN","FATURAMENTO"].indexOf(papel)<0)return jsonOutput_({ok:false,erro:"Somente ADMIN ou FATURAMENTO pode executar esta ação."});
    if(["atualizarParadaViagem","atualizarStatusViagem","registrarOcorrenciaViagem"].indexOf(acao)>=0&&["ADMIN","OPERACIONAL"].indexOf(papel)<0)return jsonOutput_({ok:false,erro:"Ação permitida somente a ADMIN ou OPERACIONAL."});
    const permitidasNoTeste=["sugerirRota","obterConfigMotivoAtraso","salvarConfigMotivoAtraso","salvarConfigModulosFrota","salvarVeiculo","alterarSituacaoVeiculo","criarViagem","analisarXmlNFe","salvarNFeViagem","liberarViagem","atualizarParadaViagem","atualizarStatusViagem","registrarOcorrenciaViagem"];
    if(ambiente==="TESTE"&&permitidasNoTeste.indexOf(acao)<0)return jsonOutput_({ok:false,erro:"Esta ação não pode alterar dados reais durante o modo de teste."});

    switch (acao) {
      case "obterConfigMotivoAtraso":
        return jsonOutput_(obterConfigMotivoAtraso_(ambiente));
      case "salvarConfigMotivoAtraso":
        return jsonOutput_(salvarConfigMotivoAtraso_(body.ativo,ambiente));
      case "salvarConfigModulosFrota":
        return jsonOutput_(salvarConfigModulosFrota_(body.modulos));
      case "analisarXmlNFe":
        return jsonOutput_(analisarXmlNFeViagem_(body.dados,ambiente,papel));
      case "salvarNFeViagem":
        return jsonOutput_(salvarNFeViagem_(body.dados,ambiente,papel,body._authUser));
      case "liberarViagem":
        return jsonOutput_(liberarViagemFaturamento_(body.dados,ambiente,papel));
      case "concluirEntrega":
        if(ambiente==="TESTE")return jsonOutput_({ok:false,erro:"A conclusão de entregas reais fica bloqueada no ambiente de teste."});
        return jsonOutput_(concluirEntrega(body.linha, body.motivo, body.observacao));
      case "salvarNovaEntrega":
        return jsonOutput_(salvarNovaEntregaManual(body.dados));
      case "salvarVariasEntregas":
        return jsonOutput_(salvarVariasEntregas(body.itens));
      case "salvarLinhaAtencao":
        return jsonOutput_(salvarLinhaAtencao(body.mensagem, body.data));
      case "salvarNovoCadastroCliente":
        return jsonOutput_(salvarNovoCadastroCliente(body.codigo, body.fantasia));
      case "atualizarCampo":
        return jsonOutput_(atualizarCampoEntrega(body.linha, body.campo, body.valor));
      case "atualizarCodigo":
        return jsonOutput_(atualizarCodigoEntrega(body.linha, body.novoCodigo));
      case "removerEntrega":
        return jsonOutput_(removerEntrega(body.linha));
      case "sugerirRota":
        return jsonOutput_(sugerirRota(body.pontos));
      case "salvarVeiculo":
        if (!moduloFrotaAtivo_("veiculos",ambiente)) return jsonOutput_({ ok: false, erro: "O módulo de veículos está desativado globalmente." });
        return jsonOutput_(salvarVeiculo_(body.veiculo,ambiente));
      case "alterarSituacaoVeiculo":
        if (!moduloFrotaAtivo_("veiculos",ambiente)) return jsonOutput_({ ok: false, erro: "O módulo de veículos está desativado globalmente." });
        return jsonOutput_(alterarSituacaoVeiculo_(body.id,body.ativo,ambiente));
      case "criarViagem":
        if (!moduloFrotaAtivo_("viagens",ambiente)) return jsonOutput_({ ok: false, erro: "O módulo de viagens está desativado globalmente." });
        return jsonOutput_(criarViagem_(body.dados,ambiente,papel));
      case "atualizarParadaViagem":
        if (!moduloFrotaAtivo_("viagens",ambiente)) return jsonOutput_({ ok: false, erro: "O módulo de viagens está desativado globalmente." });
        return jsonOutput_(atualizarParadaViagem_(body.dados,ambiente));
      case "atualizarStatusViagem":
        if (!moduloFrotaAtivo_("viagens",ambiente)) return jsonOutput_({ ok: false, erro: "O módulo de viagens está desativado globalmente." });
        return jsonOutput_(atualizarStatusViagem_(body.dados,ambiente));
      case "registrarOcorrenciaViagem":
        if (!moduloFrotaAtivo_("viagens",ambiente)) return jsonOutput_({ ok: false, erro: "O módulo de viagens está desativado globalmente." });
        return jsonOutput_(registrarOcorrenciaViagem_(body.dados,ambiente));
      case "marcarEntregaMotorista":
        return jsonOutput_(marcarEntregaMotorista(body.linha, body.status, body.motorista));
      case "marcarEmAndamento":
        return jsonOutput_(marcarEmAndamentoPeloMotorista(body.linha, body.motorista));
      case "confirmarEntrega":
        return jsonOutput_(confirmarEntrega(body.linha));
      case "confirmarTodas":
        return jsonOutput_(confirmarTodasEntregas());
      case "atualizarLocalizacao":
        return jsonOutput_(atualizarLocalizacaoMotorista(body.motorista, body.lat, body.lng));
      default:
        return jsonOutput_({ ok: false, erro: "Ação desconhecida: " + acao });
    }
  } catch (err) {
    return jsonOutput_({ ok: false, erro: String(err) });
  }
}

/**
 * Monta a lista de entregas válidas (padrão de endereço ok), sem filtro
 * nenhum de data/período/motorista/status — isso é feito no navegador.
 * Função compartilhada por doGetDados_ (JSON via URL, útil pra depurar) e
 * por buscarTodasEntregasParaMapa (chamada pelo mapa via google.script.run).
 */
/**
 * filtroMotorista e filtroData são opcionais — quando vêm preenchidos,
 * filtra direto aqui dentro do loop (menos dados pra montar e pra mandar
 * pela rede). Usado pelo Motorista.html pra carregar rápido, já que ele só
 * precisa das entregas de UM motorista, não de todo mundo.
 */
function obterEntregasValidas_(filtroMotorista, filtroData) {
  const OFF_ENTRADA = MAPA_CFG.COL_ENTRADA - MAPA_CFG.COL_ENDERECO;
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return [];

  const lastRow = encontrarUltimaLinhaComDados_(hist);
  if (lastRow < 2) return [];

  const fm = filtroMotorista ? String(filtroMotorista).trim().toUpperCase() : null;
  const fd = filtroData ? String(filtroData).trim() : null;

  // Lê em só 2 blocos largos (em vez de 5 leituras separadas) — cada
  // getRange() é uma ida e volta pra planilha; juntar em menos chamadas
  // maiores é bem mais rápido, principalmente com milhares de linhas.
  // (a coluna OK, 165, já vinha duplicada — ela cai dentro do segundo
  // bloco de qualquer forma, então nem precisa de leitura própria)
  const blocoInicio = hist.getRange(2, CFG.COL_ID_HIST, lastRow - 1, 7).getValues(); // 9(CLIENTE)..15(REPR.)
  const largura = MAPA_CFG.COL_OBS - MAPA_CFG.COL_ENDERECO + 1; // 157..168 = 12 colunas
  const blocoCol = hist.getRange(2, MAPA_CFG.COL_ENDERECO, lastRow - 1, largura).getValues();
  const cache = lerCacheGeo_();

  const entregas = [];

  // offsets dentro de blocoInicio (relativos a COL_ID_HIST = 9)
  const OFF_FANTASIA_INICIO = MAPA_CFG.COL_FANTASIA - CFG.COL_ID_HIST; // 11-9=2
  const OFF_REPR_INICIO = 15 - CFG.COL_ID_HIST; // 15-9=6

  // offsets dentro de blocoCol (relativos à COL_ENDERECO = 157)
  const OFF_MOTORISTA = MAPA_CFG.COL_MOTORISTA - MAPA_CFG.COL_ENDERECO;
  const OFF_PREVISAO  = MAPA_CFG.COL_PREVISAO  - MAPA_CFG.COL_ENDERECO;
  const OFF_STATUS    = MAPA_CFG.COL_STATUS    - MAPA_CFG.COL_ENDERECO;
  const OFF_PERIODO   = MAPA_CFG.COL_PERIODO   - MAPA_CFG.COL_ENDERECO;
  const OFF_OBS       = MAPA_CFG.COL_OBS       - MAPA_CFG.COL_ENDERECO;
  const OFF_OK        = 165 - MAPA_CFG.COL_ENDERECO; // 165-157=8

  for (let i = 0; i < blocoCol.length; i++) {
    const codigo     = blocoInicio[i][0];
    const fantasia    = blocoInicio[i][OFF_FANTASIA_INICIO];
    const representante = blocoInicio[i][OFF_REPR_INICIO];
    const dadosLinha = blocoCol[i];
    const entrada    = dadosLinha[OFF_ENTRADA];
    const endereco   = dadosLinha[0];
    const motorista  = dadosLinha[OFF_MOTORISTA];
    const previsao   = dadosLinha[OFF_PREVISAO];
    const status     = dadosLinha[OFF_STATUS];
    const periodo    = dadosLinha[OFF_PERIODO];
    const obs        = dadosLinha[OFF_OBS];
    const ok         = dadosLinha[OFF_OK];

    if (!fantasia) continue;

    // Lê TODAS as linhas agora, mesmo sem endereço padronizado ainda — ela
    // pode adicionar em lote sem endereço e completar depois direto na
    // tabela. O mapa continua só mostrando quem tem coordenada
    // geocodificada de verdade (que só existe pra endereço no padrão),
    // então o comportamento do MAPA não muda em nada.
    const textoEndereco = String(endereco || "").trim();
    const ehAviso = textoEndereco.toUpperCase() === "AVISO"; // linha de atenção — nunca tem endereço de verdade

    // Ignora avisos sem código de entrega na coluna I
    if (ehAviso && !String(codigo || "").trim()) continue;

    if (fm && String(motorista || "").trim().toUpperCase() !== fm) continue;

    // Usa a mesma função _toIsoDateString já existente no seu code.gs
    // (funciona com Date, texto "dd/mm/aaaa", "aaaa-mm-dd" ou número de série do Sheets)
    const dataIso = _toIsoDateString(previsao) || "";

    // 🔒 corte de data: ignora histórico antigo (antes de MAPA_CFG.DATA_CORTE)
    // — mas linhas sem data nenhuma continuam passando normalmente
    if (dataIso && dataIso < MAPA_CFG.DATA_CORTE) continue;

    if (fd && dataIso !== fd) continue;

    const coords = (!ehAviso && textoEndereco) ? cache[textoEndereco.toUpperCase()] : null;

    entregas.push({
      dataEntrada: _toIsoDateOnly_(entrada),
      linha: i + 2, // número da linha real na HISTORICO_ENTREGAS (pra edição na tabela)
      codigo: codigo || "",
      cliente: fantasia,
      representante: representante || "",
      motorista: motorista || "",
      data: dataIso,
      periodo: periodo || "",
      status: status || "",
      ok: ok === true, // checkbox "pedido pronto" (coluna FI)
      cidade: (!ehAviso && textoEndereco) ? extrairCidadeUf_(textoEndereco) : "", // vem do próprio endereço
      endereco: ehAviso ? "" : textoEndereco,
      observacoes: obs || "",
      lat: coords ? coords.lat : null,
      lng: coords ? coords.lng : null
    });
  }

  return entregas;
}

/**
 * Devolve as entregas em JSON via URL (?formato=dados) — útil pra testar/
 * depurar direto no navegador. Aceita ?motorista=X e ?data=aaaa-mm-dd
 * opcionais pra filtrar já no servidor (o Motorista.html usa isso pra
 * carregar rápido, sem baixar as entregas de todo mundo).
 */
function doGetDados_(e) {
 try {
  const p = (e && e.parameter) ? e.parameter : {};
  return jsonOutput_(obterEntregasValidas_(p.motorista, p.data));
 } catch (err) {
  // Em vez de deixar o erro estourar (o que gera aquela página genérica
  // "arquivo não existe" do Drive), devolve o erro real em JSON — assim
  // dá pra ver exatamente o que quebrou.
  return jsonOutput_({ erro: String(err), stack: err.stack || "" });
 }
}

/**
 * Chamada pelo Mapa.html via google.script.run (não passa pelo fetch/URL,
 * evita qualquer bloqueio de origem cruzada entre o domínio da página e o
 * do script). Devolve o array puro, sem envelope de JSON/ContentService.
 */
function buscarTodasEntregasParaMapa() {
  return obterEntregasValidas_();
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/*******************************************************************************
 * NOVA ENTREGA VIA FORMULÁRIO NO MAPA
 * -----------------------------------------------------------------------------
 * Replica o mesmo fluxo que já existe hoje quando alguém digita o código do
 * cliente na coluna I da HISTORICO_ENTREGAS: procura o código em ENTREGAS,
 * depois PEDIDOS, depois CADASTRO_CLIENTES (nessa ordem) e traz FANTASIA,
 * VENDEDOR e uma sugestão de ROTA. A pessoa completa o resto (endereço no
 * padrão, motorista, data, período, observações) direto no mapa.
 * Chamadas a partir do arquivo "Mapa" via google.script.run — não passam
 * pelo doGet, então não aparecem na URL nem precisam de parâmetros extras.
 *******************************************************************************/

/**
 * Busca um cliente pelo código. Chamada pelo botão "Buscar" do formulário.
 */
function buscarClientePorCodigo(codigo) {
  const id = String(codigo || "").trim();
  if (!id) return { encontrado: false, erro: "Informe um código de cliente." };

  const ss = SpreadsheetApp.getActive();
  const ent = ss.getSheetByName(CFG.ENTREGAS);
  const ped = ss.getSheetByName(CFG.PEDIDOS);
  const cad = ss.getSheetByName("CADASTRO_CLIENTES");

  const dadosEnt = ent && ent.getLastRow() >= 4
    ? ent.getRange(4, 1, ent.getLastRow() - 3, ent.getLastColumn()).getValues()
    : [];
  const dadosPed = ped && ped.getLastRow() >= 2
    ? ped.getRange(2, 1, ped.getLastRow() - 1, ped.getLastColumn()).getValues()
    : [];
  const dadosCad = cad && cad.getLastRow() >= 2
    ? cad.getRange(2, 1, cad.getLastRow() - 1, 3).getValues()
    : [];

  const linhaEnt = dadosEnt.find(r => String(r[8] || "").trim() === id);
  const linhaPed = !linhaEnt ? dadosPed.find(r => String(r[8] || "").trim() === id) : null;
  const linhaCad = (!linhaEnt && !linhaPed)
    ? dadosCad.find(r => String(r[0] || "").trim() === id)
    : null;

  let fantasia = "", vendedor = "";

  if (linhaEnt) {
    fantasia = linhaEnt[10] || "";
    vendedor = linhaEnt[14] || "";
  } else if (linhaPed) {
    fantasia = linhaPed[10] || "";
    vendedor = linhaPed[156] || "";
  } else if (linhaCad) {
    fantasia = linhaCad[2] || "";
  } else {
    return { encontrado: false, erro: "Código não encontrado em ENTREGAS, PEDIDOS ou CADASTRO_CLIENTES." };
  }

  // Endereço: primeiro confere se esse CÓDIGO já tem um endereço salvo na
  // "memória" (Historico_Enderecos_Cliente) — é uma busca mais direta e
  // rápida. Só se não tiver nada lá que vai ler a ROTAS_CLIENTES inteira
  // (que é uma aba bem maior) como no jeito de antes.
  let rotaSugerida = "", cep = "", numero = "";
  const enderecosMemoria = listarEnderecosCliente(id);

  if (enderecosMemoria.length) {
    rotaSugerida = enderecosMemoria[0].endereco || "";
  } else {
    const rot = ss.getSheetByName(CFG.ROTAS);
    const mapaRotas = {};
    if (rot && rot.getLastRow() >= 2) {
      const dadosRot = rot.getRange(2, 2, rot.getLastRow() - 1, 5).getValues();
      dadosRot.forEach(r => {
        const f = String(r[0] || "").trim().toUpperCase();
        if (f) mapaRotas[f] = { rota: r[2], cep: r[3], numero: r[4] };
      });
    }
    const infoRota = mapaRotas[String(fantasia).trim().toUpperCase()] || {};
    rotaSugerida = infoRota.rota || "";
    cep = infoRota.cep ? String(infoRota.cep).trim() : "";
    numero = infoRota.numero ? String(infoRota.numero).trim() : "";
  }

  return {
    encontrado: true,
    codigo: id,
    fantasia: fantasia,
    vendedor: vendedor,
    rotaSugerida: rotaSugerida,
    cep: cep,
    numero: numero
  };
}


/**
 * Grava uma nova linha no final da HISTORICO_ENTREGAS (depois da última
 * linha usada), igual ao que acontece quando alguém digita o código direto
 * na planilha — só que aqui já vem junto com o que a pessoa completou
 * manualmente no formulário do mapa.
 * dados = { codigo, endereco, motorista, data (yyyy-mm-dd), periodo, observacoes }
 */
/**
 * Acha a última linha da HISTORICO_ENTREGAS que tem uma entrega de verdade
 * (olhando a coluna FANTASIA) — diferente de hist.getLastRow(), que se
 * confunde porque a planilha tem checkbox/dropdown formatados nas 12 mil
 * linhas inteiras, mesmo nas vazias, fazendo ele "ver" conteúdo onde não tem.
 */
/**
 * Converte "yyyy-mm-dd" num objeto Date de forma segura contra fuso
 * horário — usa meio-dia (12h) em vez de meia-noite, porque meia-noite +
 * fuso negativo (Brasil é UTC-3) "escorrega" pro dia anterior quando salvo
 * na planilha. Meio-dia tem margem de sobra pra qualquer fuso do mundo.
 */
function parseDataIsoParaDate_(iso) {
  if (!iso) return "";
  const partes = String(iso).split("-");
  if (partes.length !== 3) return "";
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]), 12, 0, 0);
}

function encontrarUltimaLinhaComDados_(hist) {
  const valores = hist.getRange(1, 11, hist.getMaxRows(), 1).getValues(); // coluna K = FANTASIA
  for (let i = valores.length - 1; i >= 1; i--) { // pula a linha 0 (cabeçalho)
    if (String(valores[i][0] || "").trim() !== "") {
      return i + 1; // +1 porque a planilha começa em 1, não em 0
    }
  }
  return 1; // só o cabeçalho, nenhuma entrega ainda
}

/**
 * Confere se uma data cai num fim de semana ou está na aba FERIADOS.
 */
function ehFimDeSemanaOuFeriado_(data) {
  const diaSemana = data.getDay(); // 0=domingo, 6=sábado
  if (diaSemana === 0 || diaSemana === 6) return true;

  const ss = SpreadsheetApp.getActive();
  const feriados = ss.getSheetByName("FERIADOS");
  if (!feriados) return false;

  const lastRow = feriados.getLastRow();
  if (lastRow < 1) return false;

  const tz = Session.getScriptTimeZone();
  const dataIso = Utilities.formatDate(data, tz, "yyyy-MM-dd");
  const valores = feriados.getRange(1, 1, lastRow, 1).getValues().flat();

  return valores.some(function (v) {
    if (v instanceof Date) {
      return Utilities.formatDate(v, tz, "yyyy-MM-dd") === dataIso;
    }
    return String(v || "").trim() === dataIso;
  });
}

/**
 * Devolve o próximo dia útil (pulando fins de semana e feriados da aba
 * FERIADOS) a partir de uma data-base — usado na aba "Amanhã" do
 * Motorista.html, pra nunca cair num sábado/domingo/feriado.
 */
function proximoDiaUtil_(dataBase) {
  const d = new Date(dataBase);
  d.setDate(d.getDate() + 1);
  while (ehFimDeSemanaOuFeriado_(d)) {
    d.setDate(d.getDate() + 1);
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Tudo que o Motorista.html precisa (entregas de hoje, de amanhã e as
 * confirmações pendentes) numa chamada SÓ, em vez de três — cada chamada
 * separada ao Apps Script tem um custo próprio de inicialização, então
 * juntar tudo numa execução só é bem mais rápido pro motorista.
 */
function obterDadosMotorista_(motorista) {
  const hojeIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const amanhaIso = proximoDiaUtil_(new Date());

  return {
    hoje: obterEntregasValidas_(motorista, hojeIso),
    amanha: obterEntregasValidas_(motorista, amanhaIso),
    pendentes: listarConfirmacoesPendentes()
  };
}

/**
 * Salva uma "linha de atenção" — um aviso livre na tabela, sem precisar de
 * código de cliente nem endereço. No lugar de CLIENTE, mostra a mensagem
 * digitada; fica com STATUS vazio, o que já ativa sozinho a cor vermelha
 * de atenção (mesma regra que já existe pra qualquer linha sem status).
 */
function salvarLinhaAtencao(mensagem, data) {
  const texto = String(mensagem || "").trim();
  if (!texto) return { ok: false, erro: "Escreva uma mensagem pra linha de atenção." };

  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };

  const linha = encontrarUltimaLinhaComDados_(hist) + 1;

  hist.getRange(linha, 11).setValue(texto); // col 11 (K) — FANTASIA, é o que aparece na coluna Cliente

  const previsao = data ? parseDataIsoParaDate_(data) : "";
  hist.getRange(linha, MAPA_CFG.COL_ENDERECO, 1, 12).setValues([[
    "AVISO",    // 157 FA — marcador especial: linha de atenção, sem endereço de verdade
    "",         // 158 FB — MOTORISTA
    "",         // 159 FC
    previsao,   // 160 FD — PREVISÃO
    "",         // 161 FE
    new Date(), // 162 FF — ENTRADA
    "",         // 163 FG
    "EM FILA",  // 164 FH — STATUS (padrão "Em fila" — a cor de atenção não depende do status, e sim do código vazio)
    false,      // 165 FI — OK
    "",         // 166 FJ — CIDADE
    "",         // 167 FK — PERÍODO
    ""          // 168 FL — OBSERVAÇÕES
  ]]);
  SpreadsheetApp.flush();

  return { ok: true };
}

function salvarNovaEntregaManual(dados) {
  const id = String((dados && dados.codigo) || "").trim();
  if (!id) return { ok: false, erro: "Código do cliente é obrigatório." };

  // A Viviane já buscou esse cliente antes de salvar (o botão "Buscar" do
  // formulário) — o navegador já manda fantasia/vendedor prontos junto do
  // salvamento, então não precisa repetir a busca inteira (que é cara: lê
  // ENTREGAS, PEDIDOS, ROTAS_CLIENTES e CADASTRO_CLIENTES). Só busca nesta
  // hora se por algum motivo não vier preenchido.
  let fantasia = String(dados.fantasia || "").trim();
  let vendedor = String(dados.vendedor || "").trim();
  let rotaSugeridaFallback = "";

  if (!fantasia) {
    const info = buscarClientePorCodigo(id);
    if (!info.encontrado) return { ok: false, erro: info.erro };
    fantasia = info.fantasia;
    vendedor = info.vendedor;
    rotaSugeridaFallback = info.rotaSugerida || "";
  }

  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };

  const linha = encontrarUltimaLinhaComDados_(hist) + 1; // entra logo depois da última entrega real
  const endereco = String(dados.endereco || "").trim() || rotaSugeridaFallback || "";

  // grava tudo em só 2 idas à planilha (em vez de 9) — cada .setValue()
  // separado é uma ida e volta de rede; agrupar em .setValues() é bem
  // mais rápido pra salvar.
  hist.getRange(linha, CFG.COL_ID_HIST, 1, 7).setValues([[
    id, "", fantasia, "", "", "", vendedor
  ]]); // cols 9(I)=CLIENTE .. 15(O)=REPR., pulando as colunas do meio que não mexemos

  const previsao = dados.data ? parseDataIsoParaDate_(dados.data) : "";
  hist.getRange(linha, MAPA_CFG.COL_ENDERECO, 1, 12).setValues([[
    endereco,             // 157 FA — ENDEREÇO
    dados.motorista || "",// 158 FB — MOTORISTA
    "",                    // 159 FC
    previsao,              // 160 FD — PREVISÃO
    "",                    // 161 FE
    new Date(),            // 162 FF — ENTRADA
    "",                    // 163 FG
    "EM FILA",             // 164 FH — STATUS
    false,                 // 165 FI — OK (pedido pronto)
    "",                    // 166 FJ — CIDADE (não usada aqui)
    dados.periodo || "",   // 167 FK — PERÍODO
    dados.observacoes || ""// 168 FL — OBSERVAÇÕES
  ]]);
  SpreadsheetApp.flush();

  // se o endereço já estiver no padrão, geocodifica na hora (mesma regra
  // do onEdit) — mas se a pessoa arrastou o pino pra corrigir a posição no
  // mini-mapa, usa essa correção em vez de geocodificar de novo, e já
  // salva ela no cache pra valer pra sempre nesse endereço.
  if (endereco && enderecoNoPadrao_(endereco)) {
    if (dados.latAjustado != null && dados.lngAjustado != null) {
      salvarCorrecaoGeocodificacao_(endereco, dados.latAjustado, dados.lngAjustado);
    } else {
      geocodificarEndereco(endereco);
    }
    // guarda esse endereço na "memória" — na ROTAS_CLIENTES e no histórico
    // por código, pra ela poder escolher de novo da próxima vez
    salvarEnderecoRotaCliente_(fantasia, endereco);
    salvarEnderecoNoHistorico_(id, endereco);
  }

  return { ok: true, fantasia: fantasia, endereco: endereco };
}

/**
 * Adiciona VÁRIAS entregas de uma vez — cada item vem com o SEU PRÓPRIO
 * código, data, motorista e período (não precisa ser tudo igual, cada
 * linha do lote é independente). Agora entra SEMPRE que o código existir,
 * mesmo sem endereço padronizado ainda — o endereço vem da ROTAS_CLIENTES
 * se já tiver um lá (e já geocodifica se estiver certinho); senão fica em
 * branco na tabela, pra ela completar manualmente pelo lápis de editar
 * endereço. Só falha mesmo se o código não existir em lugar nenhum.
 */
function salvarVariasEntregas(itens) {
  const lista = Array.isArray(itens) ? itens : [];
  const sucesso = [];
  const falha = [];

  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };

  let linha = encontrarUltimaLinhaComDados_(hist) + 1;

  lista.forEach(function (item) {
    const codigo = String((item && item.codigo) || "").trim();
    if (!codigo) return;

    const info = buscarClientePorCodigo(codigo);
    if (!info.encontrado) {
      falha.push({ codigo: codigo, motivo: info.erro || "código não encontrado" });
      return;
    }

    const endereco = String(info.rotaSugerida || "").trim();
    const previsao = item.data ? parseDataIsoParaDate_(item.data) : "";

    hist.getRange(linha, CFG.COL_ID_HIST, 1, 7).setValues([[
      codigo, "", info.fantasia, "", "", "", info.vendedor
    ]]);
    hist.getRange(linha, MAPA_CFG.COL_ENDERECO, 1, 12).setValues([[
      endereco, item.motorista || "", "", previsao, "", new Date(), "", "EM FILA", false, "", item.periodo || "", ""
    ]]);

    // só geocodifica se o endereço que veio da ROTAS_CLIENTES já estiver
    // no padrão certo — senão fica sem coordenada até ela completar
    const enderecoPronto = !!(endereco && enderecoNoPadrao_(endereco));
    if (enderecoPronto) geocodificarEndereco(endereco);

    sucesso.push({ codigo: codigo, fantasia: info.fantasia, enderecoPronto: enderecoPronto });
    linha++;
  });

  SpreadsheetApp.flush();
  return { ok: true, sucesso: sucesso, falha: falha };
}

/**
 * Lista os motoristas já usados na coluna MOTORISTA (FB) da HISTORICO_ENTREGAS,
 * sem repetir — usada pra preencher o <select> de motorista no formulário e
 * na tabela editável do mapa.
 */
function listarMotoristas() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return [];

  const lastRow = encontrarUltimaLinhaComDados_(hist);
  if (lastRow < 2) return [];

  const valores = hist.getRange(2, MAPA_CFG.COL_MOTORISTA, lastRow - 1, 1).getValues().flat();
  const unicos = [...new Set(valores.map(v => String(v || "").trim()).filter(Boolean))];
  unicos.sort((a, b) => a.localeCompare(b, "pt-BR"));
  return unicos;
}

/**
 * Geocodifica um endereço só pra pré-visualização (usado pelo botão
 * "Aplicar" no formulário de nova entrega, pro mini-mapa de confirmação).
 * Não grava nada na planilha, só consulta/preenche o cache de coordenadas.
 */
function geocodificarPreview(endereco) {
  const texto = String(endereco || "").trim();
  if (!texto) return { ok: false, erro: "Digite um endereço." };
  if (!enderecoNoPadrao_(texto)) {
    return { ok: false, erro: "Endereço fora do padrão (RUA, NÚMERO - BAIRRO - CIDADE/UF)." };
  }
  const coords = geocodificarEndereco(texto);
  if (!coords) return { ok: false, erro: "Não conseguimos localizar esse endereço no mapa." };
  return { ok: true, lat: coords.lat, lng: coords.lng };
}

/**
 * Atualiza um único campo de uma entrega já existente, direto na
 * HISTORICO_ENTREGAS — usada pela tabela editável do mapa. "linha" é o
 * número de linha real na planilha (devolvido em cada entrega pela
 * obterEntregasValidas_).
 */
function atualizarCampoEntrega(linha, campo, valor) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  // getLastRow() (rápido) em vez do scan preciso — aqui só precisamos de um
  // limite de segurança, não do número exato (isso é o que deixava editar
  // uma célula lento: escanear a coluna inteira toda vez que ela mexia em
  // qualquer campo).
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida (a planilha pode ter mudado — atualize a página)." };
  }

  const colunasPorCampo = {
    motorista: MAPA_CFG.COL_MOTORISTA,
    status: MAPA_CFG.COL_STATUS,
    periodo: MAPA_CFG.COL_PERIODO,
    endereco: MAPA_CFG.COL_ENDERECO,
    observacoes: MAPA_CFG.COL_OBS,
    data: MAPA_CFG.COL_PREVISAO, // coluna FD — PREVISÃO
    representante: 15,  // col 15 — REPR.
    ok: 165             // col 165 — OK (checkbox "pedido pronto")
  };
  const coluna = colunasPorCampo[campo];
  if (!coluna) return { ok: false, erro: "Campo não editável: " + campo };

  let valorFinal = valor;
  if (campo === "data") {
    valorFinal = valor ? parseDataIsoParaDate_(valor) : "";
  }
  if (campo === "ok") {
    valorFinal = (valor === true || valor === "true");
  }

  hist.getRange(linhaNum, coluna).setValue(valorFinal);
  SpreadsheetApp.flush(); // garante que a escrita vai pra planilha de verdade antes de responder

  // se editou o endereço e ele já está no padrão, geocodifica na hora e
  // guarda na "memória" de endereços (ROTAS_CLIENTES + histórico por código)
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

  return { ok: true };
}

/**
 * Troca o código do cliente de uma entrega já existente e refaz a busca
 * (ENTREGAS > PEDIDOS > CADASTRO_CLIENTES), atualizando FANTASIA e REPR.
 * junto — usada quando a pessoa edita a coluna "Código" na tabela do mapa.
 */
function atualizarCodigoEntrega(linha, novoCodigo) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida (a planilha pode ter mudado — atualize a página)." };
  }

  const codigo = String(novoCodigo || "").trim();
  if (!codigo) return { ok: false, erro: "Código não pode ficar vazio." };

  const info = buscarClientePorCodigo(codigo);
  if (!info.encontrado) return { ok: false, erro: info.erro };

  hist.getRange(linhaNum, CFG.COL_ID_HIST).setValue(codigo); // col 9 — CLIENTE
  hist.getRange(linhaNum, 11).setValue(info.fantasia);        // col 11 — FANTASIA
  hist.getRange(linhaNum, 15).setValue(info.vendedor);        // col 15 — REPR.

  return { ok: true, codigo: codigo, fantasia: info.fantasia, representante: info.vendedor };
}

/**
 * Remove uma entrega (apaga a linha inteira) — usada pelo botão de remover
 * na tabela do mapa. Depois de remover, os números de linha de tudo que
 * vinha depois mudam, então o mapa recarrega os dados do zero em seguida
 * (feito no lado do navegador, não aqui).
 */
function removerEntrega(linha) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida (a planilha pode ter mudado — atualize a página)." };
  }
  hist.deleteRow(linhaNum);
  return { ok: true };
}

/*******************************************************************************
 * SUGESTÃO DE ROTA (OpenRouteService)
 * -----------------------------------------------------------------------------
 * Recebe os pontos (linha, lat, lng) das entregas com "OK" marcado dentro do
 * filtro que a pessoa está vendo no mapa, e devolve a ordem de visita mais
 * eficiente partindo da Perfinorte e voltando pra ela — usando o mapa de
 * ruas de verdade (não é só distância em linha reta).
 *******************************************************************************/
function sugerirRota(pontos) {
  if (!pontos || !pontos.length) {
    return { ok: false, erro: "Nenhuma entrega com \"OK\" marcado nesse filtro." };
  }
  if (!MAPA_CFG.ORS_API_KEY || !MAPA_CFG.ORS_API_KEY.trim()) {
    return { ok: false, erro: "Configure MAPA_CFG.ORS_API_KEY no code.gs." };
  }
  if (pontos.length > 50) {
    return { ok: false, erro: "Muitos pontos de uma vez (máximo 50). Filtre por data/período." };
  }

  // geocodifica a origem (Perfinorte) usando o mesmo cache/pipeline das entregas
  const origem = geocodificarEndereco(MAPA_CFG.ENDERECO_ORIGEM);
  if (!origem) {
    return { ok: false, erro: "Não foi possível localizar o endereço de origem (Perfinorte)." };
  }

  const jobs = pontos.map(function(p, i) {
    return { id: i + 1, location: [p.lng, p.lat] };
  });

  const payload = {
    jobs: jobs,
    vehicles: [{
      id: 1,
      profile: "driving-car",
      start: [origem.lng, origem.lat],
      end: [origem.lng, origem.lat]
    }],
    options: { g: true } // pede a geometria da rota (o caminho real pelas ruas)
  };

  try {
    const resp = UrlFetchApp.fetch("https://api.openrouteservice.org/optimization", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": MAPA_CFG.ORS_API_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const codigo = resp.getResponseCode();
    const texto = resp.getContentText();

    if (codigo !== 200) {
      Logger.log("❌ OpenRouteService HTTP " + codigo + ": " + texto);
      return { ok: false, erro: "OpenRouteService retornou erro (HTTP " + codigo + "). Veja o log de execução pra detalhes." };
    }

    const json = JSON.parse(texto);
    if (!json.routes || !json.routes.length) {
      return { ok: false, erro: "Não foi possível calcular uma rota com esses pontos." };
    }

    const rota = json.routes[0];
    const ordem = rota.steps
      .filter(function(s) { return s.type === "job"; })
      .map(function(s) { return pontos[s.id - 1].linha; });

    return {
      ok: true,
      ordem: ordem, // números de linha na ordem sugerida de visita
      origem: { lat: origem.lat, lng: origem.lng },
      geometria: rota.geometry || null, // polyline codificada (o caminho real pelas ruas)
      distanciaKm: Math.round(rota.distance / 100) / 10,
      duracaoMin: Math.round(rota.duration / 60)
    };
  } catch (err) {
    return { ok: false, erro: String(err) };
  }
}

/*******************************************************************************
 * CONFIRMAÇÃO DE ENTREGAS PELO MOTORISTA
 * -----------------------------------------------------------------------------
 * O motorista, pela página Motorista.html, marca uma entrega como ENTREGUE
 * ou NAO_ENTREGUE — mas isso NÃO mexe na HISTORICO_ENTREGAS. Fica guardado
 * numa aba separada (Confirmacoes_Motorista) até a Viviane revisar e
 * confirmar manualmente pelo mapa. Só aí o STATUS de verdade muda.
 *******************************************************************************/
function getOrCreateConfirmacoesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("Confirmacoes_Motorista");
  if (!sheet) {
    sheet = ss.insertSheet("Confirmacoes_Motorista");
    sheet.appendRow(["LINHA", "CLIENTE", "STATUS_MARCADO", "MOTORISTA", "DATA_HORA"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * O motorista marca uma entrega pelo celular. Se ele marcar a mesma entrega
 * de novo (mudou de ideia), substitui a marcação anterior em vez de duplicar.
 */
function marcarEntregaMotorista(linha, status, motorista) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida." };
  }
  if (status !== "ENTREGUE" && status !== "NAO_ENTREGUE") {
    return { ok: false, erro: "Status inválido." };
  }

  const cliente = hist.getRange(linhaNum, 11).getValue(); // col 11 — FANTASIA

  const sheet = getOrCreateConfirmacoesSheet_();
  const lastRow = sheet.getLastRow();
  let linhaExistente = -1;
  if (lastRow >= 2) {
    const linhas = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    linhaExistente = linhas.findIndex(function (v) { return String(v) === String(linhaNum); });
  }

  const registro = [linhaNum, cliente, status, motorista || "", new Date()];
  if (linhaExistente >= 0) {
    sheet.getRange(linhaExistente + 2, 1, 1, 5).setValues([registro]);
  } else {
    sheet.appendRow(registro);
  }

  return { ok: true };
}

/**
 * "Em andamento" marcado pelo motorista — diferente de Entregue/Não
 * entregue, essa NÃO passa pela confirmação da Viviane: escreve direto o
 * STATUS na HISTORICO_ENTREGAS na hora. É uma ação de baixo risco (só
 * avisa que ele começou a trabalhar naquela entrega), então não precisa
 * da camada extra de revisão que "Entregue" tem.
 */
function marcarEmAndamentoPeloMotorista(linha, motorista) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida (a planilha pode ter mudado — atualize a página)." };
  }

  hist.getRange(linhaNum, MAPA_CFG.COL_STATUS).setValue("EM ANDAMENTO");
  SpreadsheetApp.flush();

  return { ok: true };
}

/**
 * Lista as marcações do motorista que ainda não foram confirmadas —
 * mostradas no painel de "Confirmações" do mapa.
 */
function listarConfirmacoesPendentes() {
  const sheet = getOrCreateConfirmacoesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const dados = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return dados.map(function (r) {
    return {
      linha: r[0],
      cliente: r[1],
      statusMarcado: r[2],
      motorista: r[3],
      dataHora: r[4] instanceof Date ? r[4].toISOString() : String(r[4])
    };
  });
}

/**
 * Aplica UMA confirmação de verdade na HISTORICO_ENTREGAS: se foi marcado
 * ENTREGUE, muda o STATUS. Se foi NAO_ENTREGUE, não muda o status sozinho
 * (fica a critério da Viviane decidir o que fazer) — só registra um aviso
 * nas OBSERVAÇÕES com quem marcou e quando.
 */
function aplicarConfirmacao_(hist, registro) {
  const linhaNum = registro[0];
  const status = registro[2];
  const motorista = registro[3];
  const dataHora = registro[4];

  const ultimaLinha = hist.getLastRow();
  if (linhaNum < 2 || linhaNum > ultimaLinha) return;

  if (status === "ENTREGUE") {
    hist.getRange(linhaNum, MAPA_CFG.COL_STATUS).setValue("ENTREGUE");
    registrarConclusaoManual_(hist, linhaNum);
    SpreadsheetApp.flush(); // garante que a escrita vai pra planilha antes de seguir
  } else {
    const dataTexto = (dataHora instanceof Date)
      ? Utilities.formatDate(dataHora, Session.getScriptTimeZone(), "dd/MM HH:mm")
      : "";
    const obsAtual = hist.getRange(linhaNum, MAPA_CFG.COL_OBS).getValue();
    const novaObs = (obsAtual ? obsAtual + " | " : "")
      + "Motorista (" + motorista + ") reportou NÃO ENTREGUE em " + dataTexto;
    hist.getRange(linhaNum, MAPA_CFG.COL_OBS).setValue(novaObs);
    SpreadsheetApp.flush();
  }
}

/**
 * A Viviane confirma UMA marcação específica.
 */
function confirmarEntrega(linha) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };

  const sheet = getOrCreateConfirmacoesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, erro: "Nada pendente." };

  const dados = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const idx = dados.findIndex(function (r) { return String(r[0]) === String(linhaNum); });
  if (idx === -1) return { ok: false, erro: "Confirmação não encontrada (pode já ter sido tratada)." };

  aplicarConfirmacao_(hist, dados[idx]);
  sheet.deleteRow(idx + 2);
  return { ok: true, linha: linhaNum, statusAplicado: dados[idx][2] };
}

/**
 * A Viviane confirma TODAS as marcações pendentes de uma vez.
 */
function confirmarTodasEntregas() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };

  const sheet = getOrCreateConfirmacoesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, total: 0 };

  const dados = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  dados.forEach(function (registro) { aplicarConfirmacao_(hist, registro); });
  sheet.deleteRows(2, lastRow - 1);

  return { ok: true, total: dados.length };
}

/**
 * Lê as cores DE VERDADE que a formatação condicional da HISTORICO_ENTREGAS
 * já aplica nas linhas — em vez de chutar tons parecidos, o mapa usa
 * exatamente essas cores. Pega a cor de fundo da primeira linha encontrada
 * de cada status (e da primeira linha com "OK" marcado).
 */
/**
 * Lê as cores DE VERDADE que a formatação condicional da HISTORICO_ENTREGAS
 * já aplica nas linhas — em vez de chutar tons parecidos. Reproduz a mesma
 * lógica de prioridade que o mapa usa (Cancelar > Entregue > OK marcado >
 * Em andamento > Em fila > Atenção) e pega a cor real da primeira linha que
 * se encaixa em cada categoria, já considerando status + OK juntos (a cor
 * de "OK marcado" muda o resultado mesmo sem mudar o status).
 */
function obterCoresStatus_() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return {};

  const ultimaLinha = encontrarUltimaLinhaComDados_(hist);
  if (ultimaLinha < 2) return {};

  const statusValores = hist.getRange(2, MAPA_CFG.COL_STATUS, ultimaLinha - 1, 1).getValues();
  const statusCores = hist.getRange(2, MAPA_CFG.COL_STATUS, ultimaLinha - 1, 1).getBackgrounds();
  const okValores = hist.getRange(2, 165, ultimaLinha - 1, 1).getValues();
  const previsaoValores = hist.getRange(2, MAPA_CFG.COL_PREVISAO, ultimaLinha - 1, 1).getValues();

  const cores = {};
  for (let i = 0; i < statusValores.length; i++) {
    const status = String(statusValores[i][0] || "").trim().toUpperCase();
    const ok = okValores[i][0] === true;
    const temData = !!previsaoValores[i][0];
    const corCelula = statusCores[i][0];

    let categoria = null;
    if (status === "CANCELAR") categoria = "CANCELAR";
    else if (status === "ENTREGUE") categoria = "ENTREGUE";
    else if (ok) categoria = "OK";
    else if (status === "EM ANDAMENTO") categoria = "EM ANDAMENTO";
    else if (status === "EM FILA") categoria = "EM FILA";
    else if (!status && temData) categoria = "ATENCAO";

    if (categoria && !cores[categoria]) {
      cores[categoria] = corCelula;
    }
  }

  return cores;
}

/*******************************************************************************
 * LOCALIZAÇÃO DOS MOTORISTAS EM TEMPO REAL
 * -----------------------------------------------------------------------------
 * O Motorista.html manda a posição do GPS periodicamente enquanto ele está
 * com a sessão aberta. Guarda só a posição MAIS RECENTE de cada motorista
 * (não um histórico) — o mapa da Viviane mostra isso, e some sozinho da
 * tela se não vier atualização nova por um tempo (o motorista fechou o
 * app ou desligou o celular).
 *******************************************************************************/
function getOrCreateLocalizacaoSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("Localizacao_Motoristas");
  if (!sheet) {
    sheet = ss.insertSheet("Localizacao_Motoristas");
    sheet.appendRow(["MOTORISTA", "LAT", "LNG", "ATUALIZADO_EM"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function atualizarLocalizacaoMotorista(motorista, lat, lng) {
  const nome = String(motorista || "").trim();
  if (!nome) return { ok: false, erro: "Motorista não informado." };
  const latNum = parseFloat(lat);
  const lngNum = parseFloat(lng);
  if (isNaN(latNum) || isNaN(lngNum)) return { ok: false, erro: "Coordenadas inválidas." };

  const sheet = getOrCreateLocalizacaoSheet_();
  const lastRow = sheet.getLastRow();
  let idx = -1;
  if (lastRow >= 2) {
    const nomes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    idx = nomes.findIndex(function (n) { return String(n).trim().toUpperCase() === nome.toUpperCase(); });
  }

  const registro = [nome, latNum, lngNum, new Date()];
  if (idx >= 0) {
    sheet.getRange(idx + 2, 1, 1, 4).setValues([registro]);
  } else {
    sheet.appendRow(registro);
  }
  return { ok: true };
}

function listarLocalizacoesMotoristas() {
  const sheet = getOrCreateLocalizacaoSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const dados = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return dados.map(function (r) {
    return {
      motorista: r[0],
      lat: r[1],
      lng: r[2],
      atualizadoEm: r[3] instanceof Date ? r[3].toISOString() : String(r[3])
    };
  });
}

/**
 * Devolve as coordenadas geocodificadas da Perfinorte (MAPA_CFG.ENDERECO_ORIGEM)
 * — usado pra mostrar um ícone fixo de fábrica no mapa, sempre visível,
 * pra qualquer um se localizar em relação à empresa.
 */
function obterOrigemPerfinorte() {
  const origem = geocodificarEndereco(MAPA_CFG.ENDERECO_ORIGEM);
  if (!origem) return { ok: false, erro: "Não foi possível localizar o endereço de origem." };
  return { ok: true, lat: origem.lat, lng: origem.lng, endereco: MAPA_CFG.ENDERECO_ORIGEM };
}

/**
 * Calcula distância/tempo estimado de um ponto até a Perfinorte, usando o
 * OpenRouteService — chamada só sob demanda (quando a Viviane clica no
 * motorista no mapa), pra não gastar a cota gratuita da API o tempo todo.
 */
function calcularEtaBase(lat, lng) {
  if (!MAPA_CFG.ORS_API_KEY || !MAPA_CFG.ORS_API_KEY.trim()) {
    return { ok: false, erro: "Configure MAPA_CFG.ORS_API_KEY no code.gs." };
  }
  const latNum = parseFloat(lat);
  const lngNum = parseFloat(lng);
  if (isNaN(latNum) || isNaN(lngNum)) return { ok: false, erro: "Coordenadas inválidas." };

  const origem = geocodificarEndereco(MAPA_CFG.ENDERECO_ORIGEM);
  if (!origem) return { ok: false, erro: "Não foi possível localizar o endereço de origem." };

  const url = "https://api.openrouteservice.org/v2/directions/driving-car"
    + "?api_key=" + encodeURIComponent(MAPA_CFG.ORS_API_KEY)
    + "&start=" + lngNum + "," + latNum
    + "&end=" + origem.lng + "," + origem.lat;

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const codigo = resp.getResponseCode();
    if (codigo !== 200) {
      return { ok: false, erro: "Não foi possível calcular (HTTP " + codigo + ")." };
    }
    const json = JSON.parse(resp.getContentText());
    const resumo = json.features[0].properties.summary;
    return {
      ok: true,
      distanciaKm: Math.round(resumo.distance / 100) / 10,
      duracaoMin: Math.round(resumo.duration / 60)
    };
  } catch (err) {
    return { ok: false, erro: String(err) };
  }
}

/**
 * Calcula a rota (linha + distância/tempo) entre dois pontos quaisquer —
 * usado pelo Motorista.html pra mostrar a navegação dentro do próprio app,
 * sem precisar abrir o Google Maps (assim a aba continua em primeiro
 * plano e o rastreamento de localização não é interrompido).
 */
function calcularRotaEntrePontos(lat1, lng1, lat2, lng2) {
  if (!MAPA_CFG.ORS_API_KEY || !MAPA_CFG.ORS_API_KEY.trim()) {
    return { ok: false, erro: "Configure MAPA_CFG.ORS_API_KEY no code.gs." };
  }
  const a1 = parseFloat(lat1), o1 = parseFloat(lng1);
  const a2 = parseFloat(lat2), o2 = parseFloat(lng2);
  if ([a1, o1, a2, o2].some(function (v) { return isNaN(v); })) {
    return { ok: false, erro: "Coordenadas inválidas." };
  }

  const url = "https://api.openrouteservice.org/v2/directions/driving-car"
    + "?api_key=" + encodeURIComponent(MAPA_CFG.ORS_API_KEY)
    + "&start=" + o1 + "," + a1
    + "&end=" + o2 + "," + a2;

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const codigo = resp.getResponseCode();
    if (codigo !== 200) {
      return { ok: false, erro: "Não foi possível calcular a rota (HTTP " + codigo + ")." };
    }
    const json = JSON.parse(resp.getContentText());
    const feature = json.features[0];
    const resumo = feature.properties.summary;
    return {
      ok: true,
      // GeoJSON vem como [lng, lat] — já inverte pra [lat, lng], que é o
      // que o Leaflet espera, assim o cliente não precisa converter nada.
      pontos: feature.geometry.coordinates.map(function (c) { return [c[1], c[0]]; }),
      distanciaKm: Math.round(resumo.distance / 100) / 10,
      duracaoMin: Math.round(resumo.duration / 60)
    };
  } catch (err) {
    return { ok: false, erro: String(err) };
  }
}

/*******************************************************************************
 * MENU "FORMATAR ENDEREÇO" DENTRO DA PRÓPRIA PLANILHA
 * -----------------------------------------------------------------------------
 * Mesma ferramenta do sistema (CEP → campos → mini-mapa arrastável → copiar),
 * só que direto na barra lateral do Google Sheets — não precisa abrir o
 * link do mapa pra formatar um endereço.
 *******************************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛠 Ferramentas")
    .addItem("📍 Formatar endereço", "abrirFormatadorEnderecoSidebar_")
    .addItem("🔄 Preencher memória de endereços com o histórico", "rodarPreencherMemoriaComHistorico_")
    .addToUi();
}

function abrirFormatadorEnderecoSidebar_() {
  const html = HtmlService.createHtmlOutputFromFile("FormatarEndereco")
    .setTitle("Formatar endereço");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Menu: pergunta antes de rodar (pode demorar num histórico grande) e
 * mostra um resumo do que foi feito no final.
 */
function rodarPreencherMemoriaComHistorico_() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    "Preencher memória de endereços",
    "Isso vai olhar TODAS as entregas já cadastradas na HISTORICO_ENTREGAS e " +
    "guardar os endereços padronizados na memória — pra aparecerem como " +
    "sugestão quando buscar o cliente de novo no mapa, sem precisar salvar " +
    "de novo manualmente. Pode demorar um pouco se tiver muitas linhas. " +
    "Quer continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resposta !== ui.Button.YES) return;

  const res = preencherMemoriaComHistorico();
  if (!res.ok) {
    ui.alert("Erro: " + res.erro);
    return;
  }
  ui.alert(
    "Pronto! ✅",
    "Endereços novos guardados na memória: " + res.enderecosNovosNaMemoria + "\n" +
    "Clientes novos adicionados na ROTAS_CLIENTES: " + res.clientesNovosNaRotas + "\n" +
    "Clientes com endereço atualizado na ROTAS_CLIENTES: " + res.clientesAtualizadosNaRotas,
    ui.ButtonSet.OK
  );
}

/**
 * Varre TODA a HISTORICO_ENTREGAS uma vez e "ensina" a memória de
 * endereços com tudo que já está padronizado — assim ela não fica
 * vazia esperando cada cliente ser buscado de novo manualmente. Lê tudo
 * de uma vez só e grava em blocos (não linha por linha), pra funcionar
 * bem mesmo com milhares de entregas.
 */
function preencherMemoriaComHistorico() {
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  const rot = ss.getSheetByName("ROTAS_CLIENTES");
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };

  const lastRow = encontrarUltimaLinhaComDados_(hist);
  if (lastRow < 2) return { ok: true, enderecosNovosNaMemoria: 0, clientesNovosNaRotas: 0, clientesAtualizadosNaRotas: 0 };

  // lê código, fantasia e endereço de todo o histórico, de uma vez só
  const codigos = hist.getRange(2, CFG.COL_ID_HIST, lastRow - 1, 1).getValues();
  const fantasias = hist.getRange(2, MAPA_CFG.COL_FANTASIA, lastRow - 1, 1).getValues();
  const enderecos = hist.getRange(2, MAPA_CFG.COL_ENDERECO, lastRow - 1, 1).getValues();

  // carrega o que já está salvo na memória, pra não duplicar
  const histSheet = getOrCreateHistoricoEnderecosSheet_();
  const histLastRow = histSheet.getLastRow();
  const jaSalvos = {};
  if (histLastRow >= 2) {
    histSheet.getRange(2, 1, histLastRow - 1, 2).getValues().forEach(function (r) {
      jaSalvos[String(r[0]).trim() + "||" + String(r[1]).trim().toUpperCase()] = true;
    });
  }

  // carrega a ROTAS_CLIENTES pra saber quem já tem linha (por fantasia)
  const rotLastRow = rot ? rot.getLastRow() : 0;
  const fantasiaParaLinha = {};
  if (rot && rotLastRow >= 2) {
    rot.getRange(2, 2, rotLastRow - 1, 1).getValues().forEach(function (r, idx) {
      const f = String(r[0]).trim().toUpperCase();
      if (f) fantasiaParaLinha[f] = idx + 2;
    });
  }

  const novasLinhasMemoria = [];
  const novasLinhasRotas = [];
  const ultimaPorLinhaRotas = {}; // linha -> endereço mais recente encontrado pra ela
  const vistosNestaExecucao = {};
  const fantasiaNovaVistaNestaExecucao = {};

  for (let i = 0; i < codigos.length; i++) {
    const codigo = String(codigos[i][0] || "").trim();
    const fantasia = String(fantasias[i][0] || "").trim();
    const endereco = String(enderecos[i][0] || "").trim();

    if (!codigo || !endereco || !enderecoNoPadrao_(endereco)) continue;

    const chaveMemoria = codigo + "||" + endereco.toUpperCase();
    if (!jaSalvos[chaveMemoria] && !vistosNestaExecucao[chaveMemoria]) {
      vistosNestaExecucao[chaveMemoria] = true;
      novasLinhasMemoria.push([codigo, endereco, new Date()]);
    }

    if (fantasia) {
      const fantasiaChave = fantasia.toUpperCase();
      if (fantasiaParaLinha[fantasiaChave]) {
        ultimaPorLinhaRotas[fantasiaParaLinha[fantasiaChave]] = endereco;
      } else if (!fantasiaNovaVistaNestaExecucao[fantasiaChave]) {
        fantasiaNovaVistaNestaExecucao[fantasiaChave] = true;
        novasLinhasRotas.push([fantasia, "", endereco]); // colunas B, C(vazio), D
      }
    }
  }

  // grava em blocos — bem menos idas à planilha que uma linha de cada vez
  if (novasLinhasMemoria.length) {
    histSheet.getRange(histSheet.getLastRow() + 1, 1, novasLinhasMemoria.length, 3).setValues(novasLinhasMemoria);
  }
  if (novasLinhasRotas.length && rot) {
    rot.getRange(rot.getLastRow() + 1, 2, novasLinhasRotas.length, 3).setValues(novasLinhasRotas);
  }
  const linhasParaAtualizar = Object.keys(ultimaPorLinhaRotas);
  linhasParaAtualizar.forEach(function (linhaStr) {
    rot.getRange(parseInt(linhaStr, 10), 4).setValue(ultimaPorLinhaRotas[linhaStr]);
  });

  SpreadsheetApp.flush();

  return {
    ok: true,
    enderecosNovosNaMemoria: novasLinhasMemoria.length,
    clientesNovosNaRotas: novasLinhasRotas.length,
    clientesAtualizadosNaRotas: linhasParaAtualizar.length
  };
}









// ===== PROTEÇÃO MAPA DE ENTREGAS: AUTORIZAÇÃO EXCLUSIVA DO CLOUDFLARE WORKER =====
// 1) Cole este bloco no Apps Script.
// 2) Execute configurarSegredoWorkerUmaVez() manualmente, autorize e confira a execução.
// 3) Depois APAGUE a função configurarSegredoWorkerUmaVez() do código e salve.
// O segredo ficará armazenado em Script Properties, não no código permanente.




function validarWorkerMapa_(segredoRecebido) {
  const segredoEsperado = PropertiesService.getScriptProperties().getProperty('MAPA_WORKER_SECRET');
  if (!segredoEsperado || !segredoRecebido) return false;
  return String(segredoRecebido) === String(segredoEsperado);
}




function _toIsoDateOnly_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  if (typeof valor === 'number') {
    return _toIsoDateString(valor) || '';
  }

  const texto = String(valor).trim();
  const br = texto.match(/^(\d{2})[\/-](\d{2})[\/-](\d{4})/);

  if (br) return `${br[3]}-${br[2]}-${br[1]}`;

  return _toIsoDateString(valor) || '';
}












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

const CONFIG_EXIGIR_MOTIVO_ATRASO_ = "MAPA_EXIGIR_MOTIVO_ATRASO";

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

function ehAvisoSemCodigo_(hist, linha) {
  const endereco = String(hist.getRange(linha, MAPA_CFG.COL_ENDERECO).getValue() || "").trim().toUpperCase();
  const codigo = String(hist.getRange(linha, CFG.COL_ID_HIST).getValue() || "").trim();
  return endereco === "AVISO" && !codigo;
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
  const codigos = hist.getRange(2, CFG.COL_ID_HIST, total, 1).getValues();
  const enderecos = hist.getRange(2, MAPA_CFG.COL_ENDERECO, total, 1).getValues();
  const dados = hist.getRange(2, MAPA_CFG.COL_PREVISAO, total,
    AUDITORIA_ENTREGAS_CFG_.COL_ID - MAPA_CFG.COL_PREVISAO + 1).getValues();

  let alterou = false;
  const metadados = dados.map(function(r, i) {
    const ehAvisoSemCodigo = String(enderecos[i][0] || "").trim().toUpperCase() === "AVISO"
      && !String(codigos[i][0] || "").trim();
    const temEntrega = String(clientes[i][0] || "").trim() !== "" && !ehAvisoSemCodigo;
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
  if (ehAvisoSemCodigo_(hist, linha)) return { previsaoOriginal: "", id: "" };
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
  if (ehAvisoSemCodigo_(hist, linha)) return 0;
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

function obterConfigMotivoAtraso_(ambiente) {
  const k=normalizarAmbienteMapa_(ambiente)==="TESTE"?"teste":"producao";
  return {ok:true,ativo:obterConfigModulosFrota_().config.motivosAtraso[k]===true,ambiente:k.toUpperCase()};
}
function salvarConfigMotivoAtraso_(ativo,ambiente) {
  const env=normalizarAmbienteMapa_(ambiente), c=obterConfigModulosFrota_().config, k=env==="TESTE"?"teste":"producao";
  c.motivosAtraso[k]=ativo===true||String(ativo).trim().toLowerCase()==="true";
  const saved=salvarConfigModulosFrota_(c);
  return {ok:true,ativo:c.motivosAtraso[k],ambiente:env,config:saved.config};
}

function concluirEntrega(linha, motivo, observacao) {
  const linhaNum = parseInt(linha, 10);
  const ss = SpreadsheetApp.getActive();
  const hist = ss.getSheetByName(CFG.HIST);
  if (!hist) return { ok: false, erro: "Aba HISTORICO_ENTREGAS não encontrada." };
  if (!linhaNum || linhaNum < 2 || linhaNum > hist.getLastRow()) {
    return { ok: false, erro: "Linha inválida. Atualize a página e tente novamente." };
  }
  if (ehAvisoSemCodigo_(hist, linhaNum)) {
    return { ok: false, erro: "Registro de AVISO sem código não é uma entrega." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const verificacao = entregaEstaAtrasadaPelaPrevisaoOriginal_(hist, linhaNum);
    const motivoTexto = String(motivo || "").trim();
    const observacaoTexto = String(observacao || "").trim();
    const exigirMotivo = obterConfigMotivoAtraso_().ativo;

    if (verificacao.atrasada && exigirMotivo) {
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
      verificacao.atrasada && exigirMotivo ? motivoTexto : ""
    );
    hist.getRange(linhaNum, AUDITORIA_ENTREGAS_CFG_.COL_OBSERVACAO).setValue(
      verificacao.atrasada && exigirMotivo ? observacaoTexto : ""
    );
    SpreadsheetApp.flush();

    return {
      ok: true,
      dataHoraConclusao: dataHoraAuditoriaIso_(agora),
      atrasada: verificacao.atrasada,
      motivoAtraso: verificacao.atrasada && exigirMotivo ? motivoTexto : ""
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
      const exigirMotivo = obterConfigMotivoAtraso_().ativo;
      if (verificacao.atrasada && exigirMotivo) {
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
      return { ok: true, dataHoraConclusao: dataHoraAuditoriaIso_(agora), atrasada: verificacao.atrasada, motivoAtraso: "" };
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

    // Se retirar o status ENTREGUE, limpa os metadados da conclusão.
    // A previsão original e o histórico de reprogramações permanecem intactos.
    if (campo === "status" && String(valor || "").trim().toUpperCase() !== "ENTREGUE") {
      hist.getRange(
        linhaNum,
        AUDITORIA_ENTREGAS_CFG_.COL_CONCLUSAO,
        1,
        3
      ).clearContent();
    }

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
  if (linha < 2 || hist.getName() !== CFG.HIST || ehAvisoSemCodigo_(hist, linha)) return;

  const novoValor = e.range.getValue();
  const valorAnterior = e.oldValue || "";
  const anteriorIso = normalizarDataAuditoria_(valorAnterior);
  const novoIso = normalizarDataAuditoria_(novoValor);

  // Para a primeira edição manual antes da migração, use a data anterior
  // do evento como previsão original, não a data nova já escrita na célula.
  const celulaOriginal = hist.getRange(linha, AUDITORIA_ENTREGAS_CFG_.COL_PREVISAO_ORIGINAL);
  if (!celulaOriginal.getValue()) {
    const base = anteriorIso || novoIso;
    if (base) celulaOriginal.setValue(dataAuditoriaParaDate_(base));
  }
  garantirAuditoriaLinha_(hist, linha);

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
  if (ehAvisoSemCodigo_(hist, linha)) return;
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
    if (ehAviso && !String(codigo || "").trim()) continue;
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
      idEntrega: idEntrega,
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
